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

    POINT_TYPES = ["END_POINT", "CENTRE_POINT", "JACKET_POINT", "BRANCH1_POINT", "CO_ORDS"]

    def __init__(self):
        self.db_helper: Optional[DatabaseHelper] = None
        # 数据存储
        self.pipeline_data = []
        self.support_data = []
        self.bolt_data = []
        self.message_data = []
        self.instrument_return_data = []
        self.trap_offset_data = []
        self.common_components = {}  # Dict[str, List[list]]
        self.pipe_data = {}          # Dict[str, List[list]]
        self.point_data = {}         # Dict[str, List[list]]
        self.attribute_data = []
        # 计数器
        self.tee_stub_counter = 0
        self.message_counter = 0
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
        self.common_components.clear()
        self.pipe_data.clear()
        self.point_data.clear()
        self.attribute_data.clear()
        self.tee_stub_counter = 0
        self.message_counter = 0

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
        self._create_attribute_table()

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
            "PIPINGSPEC TEXT, INSULATIONSPEC TEXT,PAINTINGSPEC TEXT,TRACINGSPEC TEXT)"
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

    def _create_attribute_table(self):
        self.db_helper.drop_table_if_exists("COMPONENT_ATTRIBUTES")
        self.db_helper.create_table(
            "CREATE TABLE IF NOT EXISTS COMPONENT_ATTRIBUTES (FILENAME INTEGER,UCI TEXT,"
            "ATTRIBUTE_NAME TEXT,ATTRIBUTE_VALUE TEXT,PRIMARY KEY (FILENAME, UCI, ATTRIBUTE_NAME))"
        )
        self.db_helper.execute_non_query("CREATE INDEX IF NOT EXISTS idx_attr_uci ON COMPONENT_ATTRIBUTES(UCI)")
        self.db_helper.execute_non_query("CREATE INDEX IF NOT EXISTS idx_attr_name ON COMPONENT_ATTRIBUTES(ATTRIBUTE_NAME)")

    def _create_pipe_table(self, name):
        self.db_helper.drop_table_if_exists(name)
        self.db_helper.create_table(
            f"CREATE TABLE IF NOT EXISTS {name} (FILENAME INTEGER,ITEMCODE TEXT,CUTPIECELENGTH REAL,"
            f"UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,"
            f"MASTERCOMPONENTIDENTIFIER INTEGER,ANGLE INTEGER)"
        )

    def _create_common_table(self, name):
        self.db_helper.drop_table_if_exists(name)
        self.db_helper.create_table(
            f"CREATE TABLE IF NOT EXISTS {name} (FILENAME INTEGER,ITEMCODE TEXT,TAG TEXT,"
            f"UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER)"
        )

    # ========== 数据插入 ==========

    def insert_all_data(self, progress_callback=None):
        def report(msg, current, total):
            if progress_callback:
                progress_callback(msg, current, total)

        report("正在插入 PIPELINE 数据...", 0, 10)
        self._bulk_insert_simple(
            "INSERT INTO PIPELINE (FILENAME,REFERENCE,PIPINGSPEC,INSULATIONSPEC,PAINTINGSPEC,TRACINGSPEC) "
            "VALUES (?,?,?,?,?,?)",
            self.pipeline_data, 6)

        report("正在插入 SUPPORT 数据...", 1, 10)
        self._bulk_insert_simple(
            "INSERT INTO SUPPORT (FILENAME,ITEMCODE,NAME,UCI,ITEMDESCRIPTION,SKEY,SUPPORTDIRECTION,SUPPORTTYPE) "
            "VALUES (?,?,?,?,?,?,?,?)",
            self.support_data, 8)

        report("正在插入坐标点数据...", 2, 10)
        self._bulk_insert_dictionary(
            self.point_data,
            lambda key: f"INSERT INTO {key} (FILENAME,UCI,{key}X,{key}Y,{key}Z,NPD,TYPE) VALUES (?,?,?,?,?,?,?)",
            7)

        report("正在插入组件数据...", 3, 10)
        self._bulk_insert_dictionary(
            self.common_components,
            lambda key: f"INSERT INTO {key} (FILENAME,ITEMCODE,TAG,UCI,ITEMDESCRIPTION,SKEY,"
                        f"COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER) VALUES (?,?,?,?,?,?,?,?)",
            8)

        report("正在插入管道数据...", 4, 10)
        self._bulk_insert_dictionary(
            self.pipe_data,
            lambda key: f"INSERT INTO {key} (FILENAME,ITEMCODE,CUTPIECELENGTH,UCI,ITEMDESCRIPTION,SKEY,"
                        f"COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER,ANGLE) VALUES (?,?,?,?,?,?,?,?,?)",
            9)

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

        report("正在插入 ATTRIBUTE 数据...", 9, 10)
        self._bulk_insert_simple(
            "INSERT OR REPLACE INTO COMPONENT_ATTRIBUTES (FILENAME,UCI,ATTRIBUTE_NAME,ATTRIBUTE_VALUE) "
            "VALUES (?,?,?,?)",
            self.attribute_data, 4)

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
                self._value_s(row[2]), self._value_s(row[3]), self._value_s(row[4])]
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
        value_list = list(point_values)
        v = [None] * 5  # v0-v4

        try:
            if len(value_list) >= 5:
                v[0] = str(float(value_list[0]))
                v[1] = str(float(value_list[1]))
                v[2] = str(float(value_list[2]))
                v[3] = self._value_d(value_list[3])
                v[4] = self._value_s(value_list[4])
            elif len(value_list) == 4:
                v[0] = str(float(value_list[0]))
                v[1] = str(float(value_list[1]))
                v[2] = str(float(value_list[2]))
                v[3] = self._value_d(value_list[3])
                v[4] = None
            elif len(value_list) == 3:
                v[0] = str(float(value_list[0]))
                v[1] = str(float(value_list[1]))
                v[2] = str(float(value_list[2]))
                v[3] = None
                v[4] = None
            else:
                return
        except (ValueError, TypeError):
            return

        data = [str(filename), uci, v[0], v[1], v[2], v[3], v[4]]
        self.add_point(point_type, data)
