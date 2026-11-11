"""PCF 数据收集器 - 封装所有数据收集和数据库操作"""

from database_helper import DatabaseHelper
from typing import Optional


class PcfDataCollector:
    """PCF 数据收集器，负责数据收集、表创建、数据插入和视图创建"""

    # 组件类型常量
    COMPONENT_TYPES = [
        "INSTRUMENT", "SUPPORT", "FLANGE", "GASKET", "BOLT", "PIPE", "PIPE_FIXED", "WELD",
        "TEE_STUB", "TEE", "VALVE", "VALVE_ANGLE", "FLANGE_BLIND", "OLET", "ELBOW",
        "REDUCER_CONCENTRIC", "REDUCER_ECCENTRIC", "INSTRUMENT_ANGLE", "CAP",
        "MISC_COMPONENT", "FILTER", "VALVE_3WAY", "VALVE_4WAY", "VALVE_MULTIWAY",
        "COUPLING", "CROSS", "LAPJOINT_STUBEND", "REINFORCEMENT_PAD", "BEND",
        "_UNION", "CLAMP", "MESSAGE", "INSTRUMENT_RETURN", "TRAP_OFFSET"
    ]

    # 静态已知的坐标点表；PORT_POINT1/2/...（多端口组件）在解析时动态发现并建表
    POINT_TYPES = ["END_POINT", "CENTRE_POINT", "JACKET_POINT", "BRANCH1_POINT", "BRANCH2_POINT", "CO_ORDS"]

    # 伪组件记录（管线起止/续接/抽头/管口/测点），按前缀匹配，存入 CONNECTION_RECORD 表
    CONNECTION_RECORD_PREFIXES = ("END-CONNECTION-", "END-POSITION-", "TAP-CONNECTION", "NOZZLE", "TML")

    # ========== 字段映射注册表 ==========
    # 每种组件类型对应的 (pcf_field_name, data_type) 列表
    # data_type: "string" | "int" | "double"
    # 顺序与数据库表列顺序一致（不含 FILENAME，FILENAME 自动添加）
    # 新增字段只需在此注册表中添加一行 + 在对应 _create_*_table 方法中加列

    COMMON_FIELD_MAP = [
        ("ITEM-CODE", "string"),
        ("TAG", "string"),
        ("UCI", "string"),
        ("ITEM-DESCRIPTION", "string"),
        ("SKEY", "string"),
        ("COMPONENT-IDENTIFIER", "int"),
        ("MASTER-COMPONENT-IDENTIFIER", "int"),
        ("WEIGHT", "double"),
        ("FABRICATION-ITEM", "marker"),
        ("ERECTION-ITEM", "marker"),
        ("INSULATION", "string"),
        ("MISC-SPEC3", "string"),
        ("SPOOL-IDENTIFIER", "string"),
        ("CATEGORY", "string"),
        ("OFFSHORE-ITEM", "marker"),
    ]

    COMPONENT_FIELD_MAPS = {
        "PIPE": [
            ("ITEM-CODE", "string"),
            ("CUT-PIECE-LENGTH", "double"),
            ("CUT-PIECE-ALLOWANCE", "string"),
            ("UCI", "string"),
            ("ITEM-DESCRIPTION", "string"),
            ("SKEY", "string"),
            ("COMPONENT-IDENTIFIER", "int"),
            ("MASTER-COMPONENT-IDENTIFIER", "int"),
            ("ANGLE", "int"),
            ("SPOOL-IDENTIFIER", "string"),
            ("CATEGORY", "string"),
            ("OFFSHORE-ITEM", "marker"),
        ],
        "PIPE_FIXED": [
            ("ITEM-CODE", "string"),
            ("CUT-PIECE-LENGTH", "double"),
            ("CUT-PIECE-ALLOWANCE", "string"),
            ("UCI", "string"),
            ("ITEM-DESCRIPTION", "string"),
            ("SKEY", "string"),
            ("COMPONENT-IDENTIFIER", "int"),
            ("MASTER-COMPONENT-IDENTIFIER", "int"),
            ("ANGLE", "int"),
            ("SPOOL-IDENTIFIER", "string"),
            ("CATEGORY", "string"),
            ("OFFSHORE-ITEM", "marker"),
        ],
        "BEND": [
            ("ITEM-CODE", "string"),
            ("CUT-PIECE-LENGTH", "double"),
            ("CUT-PIECE-ALLOWANCE", "string"),
            ("UCI", "string"),
            ("ITEM-DESCRIPTION", "string"),
            ("SKEY", "string"),
            ("COMPONENT-IDENTIFIER", "int"),
            ("MASTER-COMPONENT-IDENTIFIER", "int"),
            ("ANGLE", "int"),
            ("BEND-RADIUS", "string"),
            ("SPOOL-IDENTIFIER", "string"),
            ("CATEGORY", "string"),
            ("OFFSHORE-ITEM", "marker"),
        ],
        "BOLT": [
            ("BOLT-ITEM-CODE", "string"),
            ("BOLT-DIA", "double"),
            ("UCI", "string"),
            ("BOLT-ITEM-DESCRIPTION", "string"),
            ("BOLT-QUANTITY", "int"),
            ("COMPONENT-IDENTIFIER", "int"),
            ("MASTER-COMPONENT-IDENTIFIER", "int"),
            ("BOLT-LENGTH", "int"),
        ],
        "SUPPORT": [
            ("ITEM-CODE", "string"),
            ("NAME", "string"),
            ("UCI", "string"),
            ("ITEM-DESCRIPTION", "string"),
            ("SKEY", "string"),
            ("SUPPORT-DIRECTION", "string"),
            ("SUPPORT-TYPE", "string"),
        ],
        "MESSAGE": [
            ("ITEM-CODE", "string"),
            ("TEXT", "string"),
            ("UCI", "string"),
        ],
        "INSTRUMENT_RETURN": [
            ("ITEM-CODE", "string"),
            ("TAG", "string"),
            ("UCI", "string"),
            ("ITEM-DESCRIPTION", "string"),
            ("SKEY", "string"),
            ("COMPONENT-IDENTIFIER", "int"),
            ("MASTER-COMPONENT-IDENTIFIER", "int"),
        ],
        "TRAP_OFFSET": [
            ("ITEM-CODE", "string"),
            ("TAG", "string"),
            ("UCI", "string"),
            ("ITEM-DESCRIPTION", "string"),
            ("SKEY", "string"),
            ("COMPONENT-IDENTIFIER", "int"),
            ("MASTER-COMPONENT-IDENTIFIER", "int"),
        ],
        "WELD": [
            ("ITEM-CODE", "string"),
            ("UCI", "string"),
            ("ITEM-DESCRIPTION", "string"),
            ("SKEY", "string"),
            ("COMPONENT-IDENTIFIER", "int"),
            ("MASTER-COMPONENT-IDENTIFIER", "int"),
            ("REPEAT-WELD-IDENTIFIER", "int"),
            ("WELD-REMARK-NUMBER", "marker"),
            ("WELD-ATTRIBUTE1", "string"),
            ("WELD-ATTRIBUTE2", "string"),
            ("WELD-ATTRIBUTE3", "string"),
            ("WELD-ATTRIBUTE4", "string"),
        ],
        "VALVE": COMMON_FIELD_MAP + [
            ("FLOW", "int"),
            ("SPINDLE-DIRECTION", "string"),
            ("DIRECTION", "string"),
        ],
        "REDUCER_ECCENTRIC": COMMON_FIELD_MAP + [
            ("FLAT-DIRECTION", "string"),
        ],
        "REDUCER_CONCENTRIC": COMMON_FIELD_MAP,
    }

    PIPELINE_FIELD_MAP = [
        ("REFERENCE", "string"),
        ("PIPING-SPEC", "string"),
        ("INSULATION-SPEC", "string"),
        ("PAINTING-SPEC", "string"),
        ("TRACING-SPEC", "string"),
    ]

    # SQL 数据类型映射
    _SQL_TYPE_MAP = {"string": "TEXT", "int": "INTEGER", "double": "REAL", "marker": "TEXT"}

    # ========== 注册表驱动辅助方法 ==========

    @classmethod
    def _pcf_name_to_col(cls, pcf_name):
        """PCF 字段名 → 数据库列名（去横杠大写）"""
        return pcf_name.replace("-", "").replace(" ", "").upper()

    @classmethod
    def _get_field_map_for_type(cls, comp_type):
        """获取组件类型对应的字段映射表"""
        return cls.COMPONENT_FIELD_MAPS.get(comp_type, cls.COMMON_FIELD_MAP)

    @classmethod
    def _build_create_table_sql(cls, table_name, field_map):
        """根据字段映射生成 CREATE TABLE SQL"""
        cols = ["FILENAME INTEGER"]
        for pcf_name, dtype in field_map:
            col_name = cls._pcf_name_to_col(pcf_name)
            cols.append(f"{col_name} {cls._SQL_TYPE_MAP.get(dtype, 'TEXT')}")
        return f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join(cols)})"

    @classmethod
    def _build_insert_sql(cls, table_name, field_map):
        """根据字段映射生成 INSERT SQL 和参数数量"""
        col_names = ["FILENAME"] + [cls._pcf_name_to_col(f[0]) for f in field_map]
        placeholders = ",".join(["?"] * len(col_names))
        sql = f"INSERT INTO {table_name} ({','.join(col_names)}) VALUES ({placeholders})"
        return sql, len(col_names)

    def __init__(self):
        self.db_helper: Optional[DatabaseHelper] = None
        # 数据存储
        self.pipeline_data = []
        self.support_data = []
        self.bolt_data = []
        self.message_data = []
        self.instrument_return_data = []
        self.trap_offset_data = []
        self.connection_record_data = []  # 伪组件记录（END-POSITION/END-CONNECTION/TAP-CONNECTION/NOZZLE/TML）
        self.common_components = {}  # Dict[str, List[list]]
        self.pipe_data = {}          # Dict[str, List[list]]
        self.point_data = {}         # Dict[str, List[list]]
        self.attribute_data = []
        self.file_header_data = []  # 文件头字段：[header_id, field_name, field_value]
        self.import_info_data = []  # 导入来源信息：[key, value]（SOURCE-FOLDER / SOURCE-FILES）
        # 计数器
        self.tee_stub_counter = 0
        self.message_counter = 0
        self.header_counter = 0     # 文件头 ID 计数器（每个物理文件 +1）
        # 批处理大小
        self.batch_size = 1000

    def initialize_database(self, db_path: str):
        self.db_helper = DatabaseHelper(db_path)

    @property
    def connection(self):
        return self.db_helper.connection if self.db_helper else None

    def reset(self):
        self.pipeline_data.clear()
        self.support_data.clear()
        self.bolt_data.clear()
        self.message_data.clear()
        self.instrument_return_data.clear()
        self.trap_offset_data.clear()
        self.connection_record_data.clear()
        self.common_components.clear()
        self.pipe_data.clear()
        self.point_data.clear()
        self.attribute_data.clear()
        self.file_header_data.clear()
        self.import_info_data.clear()
        self.tee_stub_counter = 0
        self.message_counter = 0
        self.header_counter = 0

    # ========== 数据添加方法 ==========

    def add_pipeline(self, data):
        self.pipeline_data.append(data)

    def add_support(self, data):
        self.support_data.append(data)

    def add_bolt(self, data):
        self.bolt_data.append(data)

    def add_message(self, data):
        self.message_data.append(data)

    def add_instrument_return(self, data):
        self.instrument_return_data.append(data)

    def add_trap_offset(self, data):
        self.trap_offset_data.append(data)

    def add_connection_record(self, data):
        self.connection_record_data.append(data)

    def add_common_component(self, component_type, data):
        if component_type not in self.common_components:
            self.common_components[component_type] = []
        self.common_components[component_type].append(data)

    def add_pipe(self, pipe_type, data):
        if pipe_type not in self.pipe_data:
            self.pipe_data[pipe_type] = []
        self.pipe_data[pipe_type].append(data)

    def add_point(self, point_type, data):
        if point_type not in self.point_data:
            self.point_data[point_type] = []
        self.point_data[point_type].append(data)

    def add_attribute(self, filename, uci, attr_name, attr_value):
        self.attribute_data.append([filename, uci, attr_name, attr_value])

    def add_file_header(self, header_id, field_name, field_value):
        """记录 PCF 文件头字段（按 header_id 分组，每个物理文件一组）"""
        self.file_header_data.append([header_id, field_name, field_value])

    def add_import_info(self, key, value):
        """记录导入来源信息（如 SOURCE-FOLDER / SOURCE-FILES），供轴测图渲染定位原始 PCF"""
        self.import_info_data.append([key, str(value)])

    def get_next_header_id(self):
        """分配一个新的文件头 ID（每个物理文件一个）"""
        hid = self.header_counter
        self.header_counter += 1
        return hid

    def get_next_tee_stub_id(self):
        id_str = f"TS{self.tee_stub_counter}"
        self.tee_stub_counter += 1
        return id_str

    def get_next_message_id(self):
        id_str = f"MS{self.message_counter}"
        self.message_counter += 1
        return id_str

    # ========== 表创建 ==========

    def create_all_tables(self):
        self._create_pipeline_table()
        self._create_support_table()
        self._create_bolt_table()
        self._create_message_table()
        self._create_instrument_return_table()
        self._create_trap_offset_table()
        self._create_connection_record_table()
        self._create_import_info_table()
        self._create_attribute_table()
        self._create_file_header_table()

        for point_type in self.POINT_TYPES:
            self._create_point_table(point_type)

        for component_type in self.COMPONENT_TYPES:
            if component_type in ("SUPPORT", "MESSAGE", "BOLT", "INSTRUMENT_RETURN", "TRAP_OFFSET"):
                continue
            elif component_type in ("PIPE", "PIPE_FIXED", "BEND"):
                self._create_pipe_table(component_type)
            else:
                self._create_common_table(component_type)

    def _create_pipeline_table(self):
        self.db_helper.drop_table_if_exists("PIPELINE")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS PIPELINE (FILENAME INTEGER,REFERENCE TEXT, "
            "PIPINGSPEC TEXT, INSULATIONSPEC TEXT,PAINTINGSPEC TEXT,TRACINGSPEC TEXT,"
            "HEADER_ID INTEGER)"
        )

    def _create_point_table(self, name):
        self.db_helper.drop_table_if_exists(name)
        self.db_helper.create_table(
            f"CREATE TABLE IF NOT EXISTS {name} (FILENAME INTEGER,UCI TEXT ,"
            f"{name}X REAL, {name}Y REAL, {name}Z REAL,NPD REAL,TYPE TEXT)"
        )

    def _create_support_table(self):
        self.db_helper.drop_table_if_exists("SUPPORT")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS SUPPORT (FILENAME INTEGER,ITEMCODE TEXT,NAME TEXT,"
            "UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,SUPPORTDIRECTION TEXT,SUPPORTTYPE TEXT)"
        )

    def _create_bolt_table(self):
        self.db_helper.drop_table_if_exists("BOLT")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS BOLT (FILENAME INTEGER,BOLTITEMCODE TEXT,BOLTDIA REAL,"
            "UCI TEXT,BOLTITEMDESCRIPTION TEXT,BOLTQUANTITY INTEGER,COMPONENTIDENTIFIER INTEGER,"
            "MASTERCOMPONENTIDENTIFIER INTEGER,BOLTLENGTH INTEGER)"
        )

    def _create_message_table(self):
        self.db_helper.drop_table_if_exists("MESSAGE")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS MESSAGE (FILENAME INTEGER,ITEMCODE TEXT,TEXT TEXT,UCI TEXT)"
        )

    def _create_instrument_return_table(self):
        self.db_helper.drop_table_if_exists("INSTRUMENT_RETURN")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS INSTRUMENT_RETURN (FILENAME INTEGER,ITEMCODE TEXT,TAG TEXT,"
            "UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER)"
        )

    def _create_trap_offset_table(self):
        self.db_helper.drop_table_if_exists("TRAP_OFFSET")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS TRAP_OFFSET (FILENAME INTEGER,ITEMCODE TEXT,TAG TEXT,"
            "UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER)"
        )

    def _create_connection_record_table(self):
        """伪组件记录表：管线起止/续接/抽头/管口/测点（RECORDTYPE 保留完整关键字）。"""
        self.db_helper.drop_table_if_exists("CONNECTION_RECORD")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS CONNECTION_RECORD (FILENAME INTEGER,RECORDTYPE TEXT,UCI TEXT,TAG TEXT,"
            "PIPELINEREFERENCE TEXT,CONNECTIONREFERENCE TEXT,"
            "REFERENCEX REAL,REFERENCEY REAL,REFERENCEZ REAL)"
        )
        self.db_helper.execute_non_query("CREATE INDEX IF NOT EXISTS idx_conn_rec_uci ON CONNECTION_RECORD(UCI)")

    def _ensure_point_table(self, name):
        """为解析期动态发现的点类型（如 PORT_POINT1）按需建表（不删除已有数据）。"""
        self.db_helper.create_table(
            f"CREATE TABLE IF NOT EXISTS {name} (FILENAME INTEGER,UCI TEXT ,"
            f"{name}X REAL, {name}Y REAL, {name}Z REAL,NPD REAL,TYPE TEXT)"
        )

    def _create_import_info_table(self):
        """导入来源信息表：记录 PCF 源文件夹与文件清单，供轴测图渲染直接定位原始文件"""
        self.db_helper.drop_table_if_exists("IMPORT_INFO")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS IMPORT_INFO (KEY TEXT PRIMARY KEY, VALUE TEXT)"
        )

    def _create_attribute_table(self):
        self.db_helper.drop_table_if_exists("COMPONENT_ATTRIBUTES")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS COMPONENT_ATTRIBUTES (FILENAME INTEGER,UCI TEXT,"
            "ATTRIBUTE_NAME TEXT,ATTRIBUTE_VALUE TEXT,PRIMARY KEY (FILENAME, UCI, ATTRIBUTE_NAME))"
        )
        self.db_helper.execute_non_query("CREATE INDEX IF NOT EXISTS idx_attr_uci ON COMPONENT_ATTRIBUTES(UCI)")
        self.db_helper.execute_non_query("CREATE INDEX IF NOT EXISTS idx_attr_name ON COMPONENT_ATTRIBUTES(ATTRIBUTE_NAME)")

    def _create_file_header_table(self):
        """PCF 文件头字段表：每个物理文件一个 HEADER_ID，按字段名/值分行存储"""
        self.db_helper.drop_table_if_exists("FILE_HEADER")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS FILE_HEADER (HEADER_ID INTEGER,FIELD_NAME TEXT,FIELD_VALUE TEXT,"
            "PRIMARY KEY (HEADER_ID, FIELD_NAME))"
        )

    def _create_pipe_table(self, name):
        """注册表驱动建表：PIPE / PIPE_FIXED / BEND 各自使用 COMPONENT_FIELD_MAPS 中的映射"""
        field_map = self.COMPONENT_FIELD_MAPS.get(name, self.COMPONENT_FIELD_MAPS["PIPE"])
        self.db_helper.drop_table_if_exists(name)
        self.db_helper.create_table(self._build_create_table_sql(name, field_map))

    def _create_common_table(self, name):
        """注册表驱动建表：通用组件表，使用对应类型的字段映射（无映射时用 COMMON_FIELD_MAP）"""
        field_map = self._get_field_map_for_type(name)
        self.db_helper.drop_table_if_exists(name)
        self.db_helper.create_table(self._build_create_table_sql(name, field_map))

    # ========== 数据插入 ==========

    def insert_all_data(self, progress_callback=None):
        def report(msg, current, total):
            if progress_callback:
                progress_callback(msg, current, total)

        report("正在插入 PIPELINE 数据...", 0, 10)
        self._bulk_insert_simple(
            "INSERT INTO PIPELINE (FILENAME,REFERENCE,PIPINGSPEC,INSULATIONSPEC,PAINTINGSPEC,TRACINGSPEC,HEADER_ID) "
            "VALUES (?,?,?,?,?,?,?)",
            self.pipeline_data, 7)

        report("正在插入 SUPPORT 数据...", 1, 10)
        self._bulk_insert_simple(
            "INSERT INTO SUPPORT (FILENAME,ITEMCODE,NAME,UCI,ITEMDESCRIPTION,SKEY,SUPPORTDIRECTION,SUPPORTTYPE) "
            "VALUES (?,?,?,?,?,?,?,?)",
            self.support_data, 8)

        report("正在插入坐标点数据...", 2, 10)
        # 动态发现的点类型（PORT_POINT1/2/...）此时才建表
        for point_type in list(self.point_data.keys()):
            if point_type not in self.POINT_TYPES:
                self._ensure_point_table(point_type)
        self._bulk_insert_dictionary(
            self.point_data,
            lambda key: f"INSERT INTO {key} (FILENAME,UCI,{key}X,{key}Y,{key}Z,NPD,TYPE) VALUES (?,?,?,?,?,?,?)",
            7)

        report("正在插入组件数据...", 3, 10)
        self._bulk_insert_registry(self.common_components)

        report("正在插入管道数据...", 4, 10)
        self._bulk_insert_registry(self.pipe_data)

        report("正在插入 BOLT 数据...", 5, 10)
        self._bulk_insert_simple(
            "INSERT INTO BOLT (FILENAME,BOLTITEMCODE,BOLTDIA,UCI,BOLTITEMDESCRIPTION,BOLTQUANTITY,"
            "COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER,BOLTLENGTH) VALUES (?,?,?,?,?,?,?,?,?)",
            self.bolt_data, 9)

        report("正在插入 MESSAGE 数据...", 6, 10)
        self._bulk_insert_simple(
            "INSERT INTO MESSAGE (FILENAME,ITEMCODE,TEXT,UCI) VALUES (?,?,?,?)",
            self.message_data, 4)

        report("正在插入 INSTRUMENT_RETURN 数据...", 7, 10)
        self._bulk_insert_simple(
            "INSERT INTO INSTRUMENT_RETURN (FILENAME,ITEMCODE,TAG,UCI,ITEMDESCRIPTION,SKEY,"
            "COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER) VALUES (?,?,?,?,?,?,?,?)",
            self.instrument_return_data, 8)

        report("正在插入 TRAP_OFFSET 数据...", 8, 10)
        self._bulk_insert_simple(
            "INSERT INTO TRAP_OFFSET (FILENAME,ITEMCODE,TAG,UCI,ITEMDESCRIPTION,SKEY,"
            "COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER) VALUES (?,?,?,?,?,?,?,?)",
            self.trap_offset_data, 8)

        report("正在插入 CONNECTION_RECORD 数据...", 9, 10)
        self._bulk_insert_simple(
            "INSERT INTO CONNECTION_RECORD (FILENAME,RECORDTYPE,UCI,TAG,PIPELINEREFERENCE,CONNECTIONREFERENCE,"
            "REFERENCEX,REFERENCEY,REFERENCEZ) VALUES (?,?,?,?,?,?,?,?,?)",
            self.connection_record_data, 9)

        report("正在插入 ATTRIBUTE 数据...", 9, 10)
        self._bulk_insert_simple(
            "INSERT OR REPLACE INTO COMPONENT_ATTRIBUTES (FILENAME,UCI,ATTRIBUTE_NAME,ATTRIBUTE_VALUE) "
            "VALUES (?,?,?,?)",
            self.attribute_data, 4)

        if self.file_header_data:
            self._bulk_insert_simple(
                "INSERT OR REPLACE INTO FILE_HEADER (HEADER_ID,FIELD_NAME,FIELD_VALUE) "
                "VALUES (?,?,?)",
                self.file_header_data, 3)

        if self.import_info_data:
            self._bulk_insert_simple(
                "INSERT OR REPLACE INTO IMPORT_INFO (KEY,VALUE) VALUES (?,?)",
                self.import_info_data, 2)

        report("数据插入完成", 10, 10)

    def _bulk_insert_simple(self, sql, data, param_count):
        if not data:
            return
        conn = self.db_helper.connection
        with self.db_helper._lock:
            cursor = conn.cursor()
            for item in data:
                params = tuple(self._string_to_db_value(item[j]) if j < len(item) else None
                               for j in range(param_count))
                cursor.execute(sql, params)
            conn.commit()
        data.clear()

    def _bulk_insert_dictionary(self, data, sql_builder, param_count):
        if not data:
            return
        conn = self.db_helper.connection
        with self.db_helper._lock:
            cursor = conn.cursor()
            for table_name, table_data in data.items():
                if not table_data:
                    continue
                sql = sql_builder(table_name)
                for item in table_data:
                    params = tuple(self._string_to_db_value(item[j]) if j < len(item) else None
                                   for j in range(param_count))
                    cursor.execute(sql, params)
                table_data.clear()
            conn.commit()
        data.clear()

    def _bulk_insert_registry(self, data_dict):
        """注册表驱动的批量插入：根据每个表的字段映射自动生成 SQL"""
        if not data_dict:
            return
        conn = self.db_helper.connection
        with self.db_helper._lock:
            cursor = conn.cursor()
            for table_name, table_data in data_dict.items():
                if not table_data:
                    continue
                field_map = self._get_field_map_for_type(table_name)
                sql, param_count = self._build_insert_sql(table_name, field_map)
                for item in table_data:
                    params = tuple(self._string_to_db_value(item[j]) if j < len(item) else None
                                   for j in range(param_count))
                    cursor.execute(sql, params)
                table_data.clear()
            conn.commit()
        data_dict.clear()

    # ========== 视图创建 ==========

    def create_total_view(self):
        parts = []
        for i, comp_type in enumerate(self.COMPONENT_TYPES):
            if i == 0:
                parts.append(
                    f"SELECT FILENAME,COALESCE(ITEMCODE,'{comp_type}') AS ITEMCODE,UCI,"
                    f"ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM {comp_type}"
                )
            elif comp_type == "BOLT":
                parts.append(
                    " UNION ALL SELECT FILENAME,COALESCE(BOLTITEMCODE,'BOLT') AS ITEMCODE,UCI,"
                    "BOLTITEMDESCRIPTION,COMPONENTIDENTIFIER FROM BOLT"
                )
            elif comp_type == "SUPPORT":
                parts.append(
                    " UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'SUPPORT') AS ITEMCODE,UCI,"
                    "ITEMDESCRIPTION,'' AS COMPONENTIDENTIFIER FROM SUPPORT"
                )
            elif comp_type == "MESSAGE":
                parts.append(
                    " UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'MESSAGE') AS ITEMCODE,UCI,"
                    "'' AS ITEMDESCRIPTION,'' AS COMPONENTIDENTIFIER FROM MESSAGE"
                )
            elif comp_type == "INSTRUMENT_RETURN":
                parts.append(
                    " UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'INSTRUMENT_RETURN') AS ITEMCODE,UCI,"
                    "ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM INSTRUMENT_RETURN"
                )
            elif comp_type == "TRAP_OFFSET":
                parts.append(
                    " UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'TRAP_OFFSET') AS ITEMCODE,UCI,"
                    "ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM TRAP_OFFSET"
                )
            else:
                parts.append(
                    f" UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'{comp_type}') AS ITEMCODE,UCI,"
                    f"ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM {comp_type}"
                )
        self.db_helper.drop_view_if_exists("TOTAL")
        self.db_helper.create_table(f"CREATE VIEW IF NOT EXISTS TOTAL AS {''.join(parts)}")

    def create_sqlite_view_all(self, types):
        self.create_total_view()

    def create_component_with_attributes_view(self):
        self.db_helper.execute_non_query("""
            CREATE VIEW IF NOT EXISTS COMPONENTS_WITH_ATTR AS
            SELECT
                c.FILENAME,
                c.UCI,
                c.ITEMCODE,
                c.ITEMDESCRIPTION,
                c.COMPONENTIDENTIFIER,
                GROUP_CONCAT(a.ATTRIBUTE_NAME || '=' || a.ATTRIBUTE_VALUE, '; ') AS ATTRIBUTES
            FROM TOTAL c
            LEFT JOIN COMPONENT_ATTRIBUTES a
                ON c.FILENAME = a.FILENAME AND c.UCI = a.UCI
            GROUP BY c.FILENAME, c.UCI
        """)
        self.db_helper.execute_non_query("""
            CREATE VIEW IF NOT EXISTS PIPELINE_WITH_ATTR AS
            SELECT
                p.FILENAME,
                p.REFERENCE,
                p.PIPINGSPEC,
                p.INSULATIONSPEC,
                p.PAINTINGSPEC,
                p.TRACINGSPEC,
                GROUP_CONCAT(a.ATTRIBUTE_NAME || '=' || a.ATTRIBUTE_VALUE, '; ') AS ATTRIBUTES
            FROM PIPELINE p
            LEFT JOIN COMPONENT_ATTRIBUTES a
                ON p.FILENAME = a.FILENAME AND a.UCI = 'PIPELINE:' || p.REFERENCE
            GROUP BY p.FILENAME, p.REFERENCE
        """)

    # ========== 值转换辅助方法 ==========

    @staticmethod
    def _value_s(value):
        return value if value != "" else None

    @staticmethod
    def _value_i(value):
        try:
            return str(int(value))
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _value_d(value):
        try:
            return str(float(value))
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _string_to_db_value(value):
        if value is None or value == "":
            return None
        return value

    # ========== 兼容方法 - 供 Form1 直接调用 ==========

    def insert_sqlite_table_pipeline(self, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_s(row[1]),
                self._value_s(row[2]), self._value_s(row[3]), self._value_s(row[4]), None]
        self.add_pipeline(data)

    def insert_sqlite_table_support(self, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_s(row[1]), self._value_s(row[2]),
                self._value_s(row[3]), self._value_s(row[4]), self._value_s(row[5]), self._value_s(row[6])]
        self.add_support(data)

    def insert_sqlite_table_message(self, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_s(row[1]), self._value_s(row[2]), self._value_s(row[3])]
        self.add_message(data)

    def insert_sqlite_table_instrument_return(self, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_s(row[1]), self._value_s(row[2]),
                self._value_s(row[3]), self._value_s(row[4]), self._value_i(row[5]), self._value_i(row[6])]
        self.add_instrument_return(data)

    def insert_sqlite_table_trap_offset(self, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_s(row[1]), self._value_s(row[2]),
                self._value_s(row[3]), self._value_s(row[4]), self._value_i(row[5]), self._value_i(row[6])]
        self.add_trap_offset(data)

    def insert_sqlite_table_bolt(self, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_d(row[1]), self._value_s(row[2]),
                self._value_s(row[3]), self._value_i(row[4]), self._value_i(row[5]),
                self._value_i(row[6]), self._value_i(row[7])]
        self.add_bolt(data)

    def insert_sqlite_table_pipe(self, pipe_type, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_d(row[1]), self._value_s(row[2]),
                self._value_s(row[3]), self._value_s(row[4]), self._value_i(row[5]),
                self._value_i(row[6]), self._value_i(row[7])]
        self.add_pipe(pipe_type, data)

    def insert_sqlite_table(self, component_type, row, filename):
        data = [str(filename), self._value_s(row[0]), self._value_s(row[1]), self._value_s(row[2]),
                self._value_s(row[3]), self._value_s(row[4]), self._value_i(row[5]), self._value_i(row[6])]
        self.add_common_component(component_type, data)

    def insert_sqlite_table_point(self, point_type, point_values, uci, filename):
        self.insert_point_from_list(point_type, point_values, uci, filename)

    # ========== dict 驱动方法（新版，供 PcfParser 调用） ==========
    # 使用 COMPONENT_FIELD_MAPS 注册表从 dict 提取字段，替代固定数组下标

    def insert_component_from_dict(self, comp_type, fields, filename):
        """从字段字典提取数据并添加到对应的数据列表。

        使用 COMPONENT_FIELD_MAPS 注册表驱动字段映射。
        新增字段只需在注册表中添加一行，无需修改此方法。
        支持 marker 类型字段（FABRICATION-ITEM 等）：键存在即为 YES。
        """
        field_map = self.COMPONENT_FIELD_MAPS.get(comp_type, self.COMMON_FIELD_MAP)
        data = [str(filename)]
        for pcf_name, data_type in field_map:
            if data_type == "marker":
                # marker 字段：键存在即为 YES，不存在为 NULL
                data.append("YES" if pcf_name in fields else None)
            else:
                raw_value = fields.get(pcf_name, "")
                # 兼容：通用组件的 TAG 为空时回退到 NAME
                if not raw_value and pcf_name == "TAG":
                    raw_value = fields.get("NAME", "")
                data.append(self._convert_value(raw_value, data_type))

        if comp_type in ("PIPE", "PIPE_FIXED", "BEND"):
            self.add_pipe(comp_type, data)
        elif comp_type == "BOLT":
            self.add_bolt(data)
        elif comp_type == "SUPPORT":
            self.add_support(data)
        elif comp_type == "MESSAGE":
            self.add_message(data)
        elif comp_type == "INSTRUMENT_RETURN":
            self.add_instrument_return(data)
        elif comp_type == "TRAP_OFFSET":
            self.add_trap_offset(data)
        else:
            self.add_common_component(comp_type, data)

    def insert_pipeline_from_dict(self, fields, filename, header_id):
        """从字段字典提取 PIPELINE 数据并添加到数据列表。header_id 指向 FILE_HEADER 表。"""
        data = [str(filename)]
        for pcf_name, data_type in self.PIPELINE_FIELD_MAP:
            raw_value = fields.get(pcf_name, "")
            data.append(self._convert_value(raw_value, data_type))
        data.append(header_id)
        self.add_pipeline(data)

    def insert_point_from_list(self, point_type, point_values, uci, filename):
        """从值列表提取坐标点数据并添加到数据列表。

        token 结构: X Y Z [bore] [flags...]
        - X Y Z 后第 1 个 token 若为数字则为 bore（公称直径）
        - 其余 token 全部为端面/连接 flags（如 "BW PW HT"），合并存入 TYPE 列
        """
        value_list = list(point_values)
        if len(value_list) < 3:
            return
        try:
            x = str(float(value_list[0]))
            y = str(float(value_list[1]))
            z = str(float(value_list[2]))
        except (ValueError, TypeError):
            return

        bore = None
        rest = value_list[3:]
        if rest:
            try:
                bore = str(float(rest[0]))
                rest = rest[1:]
            except (ValueError, TypeError):
                pass
        flags = " ".join(t for t in rest if t)

        data = [str(filename), uci, x, y, z, bore, flags if flags else None]
        self.add_point(point_type, data)

    def insert_connection_record_from_dict(self, record_type, fields, filename):
        """从字段字典提取伪组件记录（END-POSITION/END-CONNECTION/TAP-CONNECTION/NOZZLE/TML）。

        PIPELINE-REFERENCE: 对端管线名；CONNECTION-REFERENCE: 对端连接点
        （若值为 3 个数字则同时解析到 REFERENCEX/Y/Z 列）。
        """
        ref_raw = fields.get("CONNECTION-REFERENCE", "")
        ref_x = ref_y = ref_z = None
        tokens = ref_raw.split()
        if len(tokens) >= 3:
            try:
                ref_x = str(float(tokens[0]))
                ref_y = str(float(tokens[1]))
                ref_z = str(float(tokens[2]))
            except ValueError:
                ref_x = ref_y = ref_z = None

        data = [
            str(filename),
            self._value_s(record_type),
            self._value_s(fields.get("UCI", "")),
            self._value_s(fields.get("TAG", "")),
            self._value_s(fields.get("PIPELINE-REFERENCE", "")),
            self._value_s(ref_raw),
            ref_x, ref_y, ref_z,
        ]
        self.add_connection_record(data)

    @staticmethod
    def _convert_value(raw_value, data_type):
        """根据数据类型转换字段值"""
        if data_type == "string":
            return PcfDataCollector._value_s(raw_value)
        elif data_type == "int":
            return PcfDataCollector._value_i(raw_value)
        elif data_type == "double":
            return PcfDataCollector._value_d(raw_value)
        elif data_type == "marker":
            return "YES" if raw_value else None
        return raw_value if raw_value != "" else None
