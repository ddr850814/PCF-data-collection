"""PCF 文件解析器 - 从 PCF 文件中提取组件数据"""

import re
from io import StringIO
from pcf_data_collector import PcfDataCollector


class PcfParser:
    """PCF 文件解析器，负责解析单个 PCF 文件并将数据传递给数据收集器"""

    COMPONENT_TYPES = PcfDataCollector.COMPONENT_TYPES
    COMPONENT_TYPES_SET = set(COMPONENT_TYPES)

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
        report = True
        title = None
        row_p = ["", "", "", "", ""]
        row = ["", "", "", "", "", "", "", ""]
        endpoint = []
        current_fid = file_index
        is_first_reference = True

        with self._open_file(file_path) as f:
            # 预扫描 MATERIALS 段，建立 ITEM-CODE → DESCRIPTION 映射
            self._materials_map = self._scan_materials(f)
            for line in f:
                line = line.rstrip("\n").rstrip("\r")
                if line.startswith("PIPELINE-REFERENCE"):
                    self._write_previously(report, current_fid, row_p, row, title, endpoint)
                    if not is_first_reference:
                        current_fid += 1
                    is_first_reference = False
                    row_p[0] = line.replace("PIPELINE-REFERENCE", "").strip()
                    title = "PIPELINE"
                elif self._is_component_line(line):
                    self._write_previously(report, current_fid, row_p, row, title, endpoint)
                    title = line.split(" ")[0].replace("-", "_")
                    report = True
                elif not line.startswith(" "):
                    self._write_previously(report, current_fid, row_p, row, title, endpoint)
                    title = None
                    report = True
                elif title is not None:
                    if title == "PIPELINE":
                        report = self._parse_pipeline_line(line, row_p, current_fid) or report
                    else:
                        result = self._parse_component_line(line, row, endpoint, report, current_fid)
                        if result is not None:
                            report = result
            # 文件结束后处理最后一条记录
            self._write_previously(report, current_fid, row_p, row, title, endpoint)

        return current_fid - file_index + 1

    def _is_component_line(self, line: str) -> bool:
        """检查行是否为组件类型行（非缩进且匹配组件类型）"""
        if line.startswith(" "):
            return False
        first_word = line.split(" ")[0].replace("-", "_")
        return first_word in self.COMPONENT_TYPES_SET

    def _parse_pipeline_line(self, line, row_p, file_index):
        """解析 PIPELINE 区段内的行"""
        stripped = line.lstrip()
        if stripped.startswith("PIPING-SPEC"):
            row_p[1] = line.replace("PIPING-SPEC", "").strip()
        elif stripped.startswith("INSULATION-SPEC"):
            row_p[2] = line.replace("INSULATION-SPEC", "").strip()
        elif stripped.startswith("PAINTING-SPEC"):
            row_p[3] = line.replace("PAINTING-SPEC", "").strip()
        elif stripped.startswith("TRACING-SPEC"):
            row_p[4] = line.replace("TRACING-SPEC", "").strip()
        elif stripped.startswith("ATTRIBUTE"):
            self._parse_attribute(line, f"PIPELINE:{row_p[0]}", file_index)

    def _parse_component_line(self, line, row, endpoint, report, file_index):
        """解析普通组件区段内的行。
        返回 None 表示不修改 report，返回 False 表示设置 report=False。
        """
        stripped = line.lstrip()
        if stripped.startswith("ANGLE"):
            row[7] = line.replace("ANGLE", "").strip()
        elif stripped.startswith("BOLT-LENGTH"):
            row[7] = line.replace("BOLT-LENGTH", "").strip()
        elif stripped.startswith("MASTER-COMPONENT-IDENTIFIER"):
            row[6] = line.replace("MASTER-COMPONENT-IDENTIFIER", "").strip()
        elif stripped.startswith("SUPPORT-TYPE"):
            row[6] = line.replace("SUPPORT-TYPE", "").strip()
        elif stripped.startswith("COMPONENT-IDENTIFIER"):
            row[5] = line.replace("COMPONENT-IDENTIFIER", "").strip()
        elif stripped.startswith("SUPPORT-DIRECTION"):
            row[5] = line.replace("SUPPORT-DIRECTION", "").strip()
        elif stripped.startswith("SKEY"):
            row[4] = line.replace("SKEY", "").strip()
        elif stripped.startswith("BOLT-QUANTITY"):
            row[4] = line.replace("BOLT-QUANTITY", "").strip()
        elif stripped.startswith("ITEM-DESCRIPTION"):
            row[3] = line.replace("ITEM-DESCRIPTION", "").strip()
        elif stripped.startswith("BOLT-ITEM-DESCRIPTION"):
            row[3] = line.replace("BOLT-ITEM-DESCRIPTION", "").strip()
        elif stripped.startswith("UCI"):
            row[2] = line.replace("UCI", "").strip()
        elif stripped.startswith("TAG"):
            row[1] = line.replace("TAG", "").strip()
        elif stripped.startswith("NAME"):
            row[1] = line.replace("NAME", "").strip()
        elif stripped.startswith("BOLT-DIA"):
            row[1] = line.replace("BOLT-DIA", "").strip()
        elif stripped.startswith("CUT-PIECE-LENGTH"):
            row[1] = line.replace("CUT-PIECE-LENGTH", "").strip()
        elif stripped.startswith("TEXT"):
            row[1] = line.replace("TEXT", "").strip()
        elif stripped.startswith("ITEM-CODE"):
            row[0] = line.replace("ITEM-CODE", "").strip()
        elif stripped.startswith("BOLT-ITEM-CODE"):
            row[0] = line.replace("BOLT-ITEM-CODE", "").strip()
        elif (stripped.startswith("END-POINT") or stripped.startswith("CENTRE-POINT") or
              stripped.startswith("JACKET-POINT") or stripped.startswith("BRANCH1-POINT") or
              stripped.startswith("CO-ORDS")):
            endpoint.append(line.strip())
        elif stripped.startswith("MATERIAL-LIST") and line.rstrip().endswith("EXCLUDE"):
            return False
        elif stripped.startswith("ATTRIBUTE"):
            current_uci = row[2]
            if current_uci:
                self._parse_attribute(line, current_uci, file_index)
        return None

    def _parse_attribute(self, line, uci, file_index):
        """解析 ATTRIBUTE 行"""
        stripped = line.lstrip()
        space_index = stripped.find(" ")
        if space_index > 0:
            attr_name = stripped[:space_index].strip()
            attr_value = stripped[space_index + 1:].strip()
            self.collector.add_attribute(str(file_index), uci, attr_name, attr_value)

    def _write_previously(self, report, filename, row_p, row, title, endpoint):
        """处理前一条记录，写入数据收集器"""
        if any(n != "" for n in row_p):
            if title == "PIPELINE":
                self.collector.insert_sqlite_table_pipeline(row_p, filename)
            # 重置 row_p
            for i in range(len(row_p)):
                row_p[i] = ""
        elif any(n != "" for n in row):
            if report:
                # UCI 为空时自动生成唯一编号，确保组件与坐标点正确关联
                if not row[2]:
                    self._synthetic_counter += 1
                    row[2] = f"AUTO-{filename}-{self._synthetic_counter}"
                # ITEM-DESCRIPTION 为空时从 MATERIALS 段回填
                if not row[3] and row[0] and row[0] in self._materials_map:
                    row[3] = self._materials_map[row[0]]
                if title == "SUPPORT":
                    self.collector.insert_sqlite_table_support(row, filename)
                elif title == "MESSAGE":
                    row[2] = self.collector.get_next_message_id()
                    self.collector.insert_sqlite_table_message(row, filename)
                elif title == "INSTRUMENT_RETURN":
                    self.collector.insert_sqlite_table_instrument_return(row, filename)
                elif title == "TRAP_OFFSET":
                    self.collector.insert_sqlite_table_trap_offset(row, filename)
                elif title == "BOLT":
                    self.collector.insert_sqlite_table_bolt(row, filename)
                elif title in ("PIPE", "PIPE_FIXED", "BEND"):
                    self.collector.insert_sqlite_table_pipe(title, row, filename)
                elif title == "TEE_STUB":
                    row[2] = self.collector.get_next_tee_stub_id()
                    self.collector.insert_sqlite_table(title, row, filename)
                else:
                    self.collector.insert_sqlite_table(title, row, filename)

                # 处理端点数据
                for ep in endpoint:
                    input_array = re.split(r"\s+", ep)
                    if 4 <= len(input_array) <= 6:
                        point_type = input_array[0].replace("-", "_")
                        self.collector.insert_sqlite_table_point(
                            point_type, input_array[1:], row[2], filename)
            # 重置 row
            for i in range(len(row)):
                row[i] = ""
        endpoint.clear()
