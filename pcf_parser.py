"""PCF 文件解析器 - 从 PCF 文件中提取组件数据

架构说明（重构版）：
- 使用 dict 存储组件字段（fields），而非固定位置数组
- 字段按 PCF 原始名称存储（如 "ITEM-CODE"），由 collector 的注册表负责映射到数据库列
- 新增字段无需修改解析器，只需在 collector 的 COMPONENT_FIELD_MAPS 中添加映射

对比旧架构的改进：
- 旧: row[0..7] 固定下标，字段位置冲突（如 row[7] 同时存 ANGLE 和 BOLT-LENGTH）
- 新: fields["ANGLE"] / fields["BOLT-LENGTH"] 各自独立，不会冲突
- 旧: _parse_component_line 中 30+ 个 elif 分支逐字段匹配
- 新: 通用 _extract_key_value 提取所有字段，无需维护 elif 链
"""

import re
from io import StringIO
from pcf_data_collector import PcfDataCollector


class PcfParser:
    """PCF 文件解析器，负责解析单个 PCF 文件并将数据传递给数据收集器"""

    COMPONENT_TYPES = PcfDataCollector.COMPONENT_TYPES
    COMPONENT_TYPES_SET = set(COMPONENT_TYPES)

    # 坐标点关键字（行首匹配）
    POINT_KEYWORDS = ("END-POINT", "CENTRE-POINT", "JACKET-POINT", "BRANCH1-POINT", "CO-ORDS")

    def __init__(self, data_collector: PcfDataCollector, encoding: str = "auto"):
        self.collector = data_collector
        self.encoding = encoding
        self._synthetic_counter = 0
        self._materials_map: dict[str, str] = {}

    def _open_file(self, file_path: str):
        """读取文件并返回文本流。encoding='auto' 时自动检测编码。"""
        with open(file_path, "rb") as fb:
            raw = fb.read()
        if self.encoding != "auto":
            return StringIO(raw.decode(self.encoding, errors="replace"))
        # auto: 优先按 BOM 识别 UTF-8-SIG
        if raw.startswith(b"\xef\xbb\xbf"):
            return StringIO(raw[3:].decode("utf-8", errors="replace"))
        # 无 BOM：依次尝试常见编码（整段 decode 以便捕获错误）
        for enc in ("gbk", "utf-8"):
            try:
                return StringIO(raw.decode(enc))
            except UnicodeDecodeError:
                continue
        return StringIO(raw.decode("latin-1"))

    def _scan_materials(self, f) -> dict:
        """预扫描 MATERIALS 段，返回 {ITEM-CODE: DESCRIPTION} 映射"""
        materials_map = {}
        in_materials = False
        current_code = None
        f.seek(0)
        for line in f:
            stripped = line.rstrip("\n").rstrip("\r")
            if stripped == "MATERIALS":
                in_materials = True
                continue
            if in_materials:
                if not stripped.startswith(" ") and stripped.startswith("ITEM-CODE"):
                    current_code = stripped.replace("ITEM-CODE", "").strip()
                elif current_code and stripped.lstrip().startswith("DESCRIPTION"):
                    materials_map[current_code] = stripped.lstrip().replace("DESCRIPTION", "").strip()
                    current_code = None
        f.seek(0)
        return materials_map

    def process_single_file(self, file_path: str, file_index: int):
        """处理单个 PCF 文件，每个 PIPELINE-REFERENCE 分配独立的 file_index。
        返回消耗的 file_index 数量。"""
        with self._open_file(file_path) as f:
            # 预扫描 MATERIALS 段，建立 ITEM-CODE → DESCRIPTION 映射
            self._materials_map = self._scan_materials(f)
            return self._parse_stream(f, file_index)

    def _parse_stream(self, f, file_index: int) -> int:
        """逐行解析 PCF 文本流。

        核心状态机：
        - current_type:  当前实体类型（"PIPELINE" / "ELBOW" / "PIPE" / None）
        - current_fields: 当前实体的字段字典
        - current_points: 当前实体的坐标点列表
        - report:        是否输出到数据库（MATERIAL-LIST EXCLUDE 时为 False）
        """
        current_fid = file_index
        is_first_reference = True

        current_type = None
        current_fields = {}
        current_points = []
        report = True

        for line in f:
            line = line.rstrip("\n").rstrip("\r")

            if line.startswith("PIPELINE-REFERENCE"):
                self._flush(current_type, current_fields, current_points, report, current_fid)
                current_type = "PIPELINE"
                current_fields = {"REFERENCE": line.replace("PIPELINE-REFERENCE", "").strip()}
                current_points = []
                if not is_first_reference:
                    current_fid += 1
                is_first_reference = False
                report = True

            elif self._is_component_line(line):
                self._flush(current_type, current_fields, current_points, report, current_fid)
                current_type = line.split(" ")[0].replace("-", "_")
                current_fields = {}
                current_points = []
                report = True

            elif not line.startswith(" "):
                # 非缩进的非组件行（如 ISOGEN-FILES, UNITS-* 等顶层字段）
                self._flush(current_type, current_fields, current_points, report, current_fid)
                current_type = None
                current_fields = {}
                current_points = []
                report = True

            elif current_type is not None:
                if current_type == "PIPELINE":
                    self._parse_pipeline_field(line, current_fields, current_fid)
                else:
                    result = self._parse_component_field(line, current_fields, current_points,
                                                         report, current_fid)
                    if result is not None:
                        report = result

        # 文件结束后 flush 最后一条记录
        self._flush(current_type, current_fields, current_points, report, current_fid)
        return current_fid - file_index + 1

    def _is_component_line(self, line: str) -> bool:
        """检查行是否为组件类型行（非缩进且匹配组件类型）"""
        if line.startswith(" "):
            return False
        first_word = line.split(" ")[0].replace("-", "_")
        return first_word in self.COMPONENT_TYPES_SET

    def _parse_pipeline_field(self, line, fields, file_index):
        """解析 PIPELINE 区段内的行，写入 fields 字典。

        所有字段直接存入 dict，由 collector 的 PIPELINE_FIELD_MAP 决定提取哪些。
        ATTRIBUTE 行单独传递给 collector.add_attribute。
        """
        stripped = line.lstrip()
        key, value = self._extract_key_value(stripped)
        if key is None:
            return

        if key and key.startswith("ATTRIBUTE"):
            # 匹配 ATTRIBUTE / ATTRIBUTE1 / ATTRIBUTE199 等
            reference = fields.get("REFERENCE", "")
            if reference:
                self._parse_attribute(line, f"PIPELINE:{reference}", file_index)
        else:
            fields[key] = value

    def _parse_component_field(self, line, fields, points, report, file_index):
        """解析普通组件区段内的行，写入 fields 字典或 points 列表。

        所有非特殊字段通过 _extract_key_value 自动提取，无需维护 elif 链。
        返回 None 表示不修改 report，返回 False 表示设置 report=False。
        """
        stripped = line.lstrip()

        # 坐标点
        if any(stripped.startswith(pk) for pk in self.POINT_KEYWORDS):
            points.append(line.strip())
            return None

        # MATERIAL-LIST EXCLUDE
        if stripped.startswith("MATERIAL-LIST") and line.rstrip().endswith("EXCLUDE"):
            return False

        # ATTRIBUTE 行
        if stripped.startswith("ATTRIBUTE"):
            uci = fields.get("UCI", "")
            if uci:
                self._parse_attribute(line, uci, file_index)
            return None

        # 通用字段提取：key-value 直接写入 dict
        key, value = self._extract_key_value(stripped)
        if key:
            fields[key] = value

        return None

    @staticmethod
    def _extract_key_value(stripped: str) -> tuple:
        """从已去除缩进的行中提取字段名和值。

        示例:
          'ITEM-CODE    PPPSP82671' -> ('ITEM-CODE', 'PPPSP82671')
          'FABRICATION-ITEM'         -> ('FABRICATION-ITEM', '')
          'WEIGHT    2.500'          -> ('WEIGHT', '2.500')

        Returns:
            (key, value) 元组；无法提取时返回 (None, None)
        """
        match = re.match(r'^(\S+)\s*(.*)$', stripped)
        if match:
            return match.group(1), match.group(2).strip()
        return None, None

    def _parse_attribute(self, line, uci, file_index):
        """解析 ATTRIBUTE 行"""
        stripped = line.lstrip()
        space_index = stripped.find(" ")
        if space_index > 0:
            attr_name = stripped[:space_index].strip()
            attr_value = stripped[space_index + 1:].strip()
            self.collector.add_attribute(str(file_index), uci, attr_name, attr_value)

    def _flush(self, comp_type, fields, points, report, filename):
        """将当前实体刷新到数据收集器。"""
        if comp_type is None:
            return

        if comp_type == "PIPELINE":
            if fields.get("REFERENCE"):
                self.collector.insert_pipeline_from_dict(fields, filename)
            return

        if not fields:
            return

        if not report:
            return

        # UCI 为空时自动生成唯一编号
        if not fields.get("UCI"):
            self._synthetic_counter += 1
            fields["UCI"] = f"AUTO-{filename}-{self._synthetic_counter}"

        # ITEM-DESCRIPTION 为空时从 MATERIALS 段回填
        if not fields.get("ITEM-DESCRIPTION"):
            item_code = fields.get("ITEM-CODE", "")
            if item_code and item_code in self._materials_map:
                fields["ITEM-DESCRIPTION"] = self._materials_map[item_code]

        # 特殊组件的 UCI 自动编号
        if comp_type == "MESSAGE":
            fields["UCI"] = self.collector.get_next_message_id()
        elif comp_type == "TEE_STUB":
            fields["UCI"] = self.collector.get_next_tee_stub_id()

        # 传递给 collector（dict 驱动，由注册表决定字段映射）
        self.collector.insert_component_from_dict(comp_type, fields, filename)

        # 处理坐标点数据
        uci = fields.get("UCI", "")
        for ep in points:
            input_array = re.split(r"\s+", ep)
            if 4 <= len(input_array) <= 6:
                point_type = input_array[0].replace("-", "_")
                self.collector.insert_point_from_list(point_type, input_array[1:], uci, filename)
