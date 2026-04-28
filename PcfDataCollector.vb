Imports System.Data.SQLite
Imports System.Data

''' <summary>
''' PCF 数据收集器 - 封装所有数据收集和数据库操作
''' 替代原有的全局变量，提供更好的封装和可维护性
''' </summary>
Public Class PcfDataCollector

    ' 数据库连接
    Public Property ConnectionStringBuilder As SQLiteConnectionStringBuilder
    Public ReadOnly Property DbHelper As DatabaseHelper
        Get
            If _dbHelper Is Nothing AndAlso ConnectionStringBuilder IsNot Nothing Then
                _dbHelper = New DatabaseHelper(ConnectionStringBuilder.ConnectionString)
            End If
            Return _dbHelper
        End Get
    End Property
    Private _dbHelper As DatabaseHelper

    ' 数据存储
    Public ReadOnly Property PipelineData As List(Of String())
    Public ReadOnly Property SupportData As List(Of String())
    Public ReadOnly Property BoltData As List(Of String())
    Public ReadOnly Property MessageData As List(Of String())
    Public ReadOnly Property CommonComponents As Dictionary(Of String, List(Of String()))
    Public ReadOnly Property PipeData As Dictionary(Of String, List(Of String()))
    Public ReadOnly Property PointData As Dictionary(Of String, List(Of String()))

    ' 计数器
    Public Property TeeStubCounter As Integer = 0
    Public Property MessageCounter As Integer = 0

    ' 组件类型常量（共享字段，供 Module1 使用）
    Public Shared ReadOnly ComponentTypesField As String() = {
        "INSTRUMENT", "SUPPORT", "FLANGE", "GASKET", "BOLT", "PIPE", "PIPE_FIXED", "WELD",
        "TEE_STUB", "TEE", "VALVE", "VALVE_ANGLE", "FLANGE_BLIND", "OLET", "ELBOW",
        "REDUCER_CONCENTRIC", "REDUCER_ECCENTRIC", "INSTRUMENT_ANGLE", "CAP",
        "MISC_COMPONENT", "FILTER", "VALVE_3WAY", "VALVE_4WAY", "VALVE_MULTIWAY",
        "COUPLING", "CROSS", "LAPJOINT_STUBEND", "REINFORCEMENT_PAD", "BEND",
        "_UNION", "CLAMP", "MESSAGE"
    }

    Public Shared ReadOnly PointTypesField As String() = {
        "END_POINT", "CENTRE_POINT", "JACKET_POINT", "BRANCH1_POINT", "CO_ORDS"
    }

    ' 实例属性（使用共享字段）
    Public ReadOnly Property ComponentTypes As String()
        Get
            Return ComponentTypesField
        End Get
    End Property

    Public ReadOnly Property PointTypes As String()
        Get
            Return PointTypesField
        End Get
    End Property

    Public Sub New()
        PipelineData = New List(Of String())
        SupportData = New List(Of String())
        BoltData = New List(Of String())
        MessageData = New List(Of String())
        CommonComponents = New Dictionary(Of String, List(Of String()))
        PipeData = New Dictionary(Of String, List(Of String()))
        PointData = New Dictionary(Of String, List(Of String()))
    End Sub

    ''' <summary>
    ''' 初始化数据库连接
    ''' </summary>
    Public Sub InitializeDatabase(connectionString As String)
        ConnectionStringBuilder = New SQLiteConnectionStringBuilder() With {
            .DataSource = connectionString,
            .ForeignKeys = True
        }
        _dbHelper = New DatabaseHelper(connectionString)
    End Sub

    ''' <summary>
    ''' 重置所有数据
    ''' </summary>
    Public Sub Reset()
        PipelineData.Clear()
        SupportData.Clear()
        BoltData.Clear()
        MessageData.Clear()
        CommonComponents.Clear()
        PipeData.Clear()
        PointData.Clear()
        TeeStubCounter = 0
        MessageCounter = 0
    End Sub

    ''' <summary>
    ''' 添加 Pipeline 数据
    ''' </summary>
    Public Sub AddPipeline(data As String())
        PipelineData.Add(data)
    End Sub

    ''' <summary>
    ''' 添加 Support 数据
    ''' </summary>
    Public Sub AddSupport(data As String())
        SupportData.Add(data)
    End Sub

    ''' <summary>
    ''' 添加 Bolt 数据
    ''' </summary>
    Public Sub AddBolt(data As String())
        BoltData.Add(data)
    End Sub

    ''' <summary>
    ''' 添加 Message 数据
    ''' </summary>
    Public Sub AddMessage(data As String())
        MessageData.Add(data)
    End Sub

    ''' <summary>
    ''' 添加通用组件数据
    ''' </summary>
    Public Sub AddCommonComponent(componentType As String, data As String())
        If Not CommonComponents.ContainsKey(componentType) Then
            CommonComponents.Add(componentType, New List(Of String()))
        End If
        CommonComponents(componentType).Add(data)
    End Sub

    ''' <summary>
    ''' 添加管道数据
    ''' </summary>
    Public Sub AddPipe(pipeType As String, data As String())
        If Not PipeData.ContainsKey(pipeType) Then
            PipeData.Add(pipeType, New List(Of String()))
        End If
        PipeData(pipeType).Add(data)
    End Sub

    ''' <summary>
    ''' 添加坐标点数据
    ''' </summary>
    Public Sub AddPoint(pointType As String, data As String())
        If Not PointData.ContainsKey(pointType) Then
            PointData.Add(pointType, New List(Of String()))
        End If
        PointData(pointType).Add(data)
    End Sub

    ''' <summary>
    ''' 获取下一个 TeeStub 编号
    ''' </summary>
    Public Function GetNextTeeStubId() As String
        Dim id = $"TS{TeeStubCounter}"
        TeeStubCounter += 1
        Return id
    End Function

    ''' <summary>
    ''' 获取下一个 Message 编号
    ''' </summary>
    Public Function GetNextMessageId() As String
        Dim id = $"MS{MessageCounter}"
        MessageCounter += 1
        Return id
    End Function

    ''' <summary>
    ''' 创建所有数据表
    ''' </summary>
    Public Sub CreateAllTables()
        CreatePipelineTable()
        CreateSupportTable()
        CreateBoltTable()
        CreateMessageTable()

        For Each pointType In PointTypesField
            CreatePointTable(pointType)
        Next

        For Each componentType In ComponentTypesField
            Select Case componentType
                Case "SUPPORT"
                    ' 已创建
                Case "MESSAGE"
                    ' 已创建
                Case "BOLT"
                    ' 已创建
                Case "PIPE", "PIPE_FIXED", "BEND"
                    CreatePipeTable(componentType)
                Case Else
                    CreateCommonTable(componentType)
            End Select
        Next
    End Sub

    Private Sub CreatePipelineTable()
        DbHelper.DropTableIfExists("PIPELINE")
        DbHelper.CreateTable("CREATE TABLE IF NOT EXISTS PIPELINE (FILENAME INTEGER,REFERENCE TEXT, PIPINGSPEC TEXT, INSULATIONSPEC TEXT,PAINTINGSPEC TEXT,TRACINGSPEC TEXT)")
    End Sub

    Private Sub CreatePointTable(name As String)
        DbHelper.DropTableIfExists(name)
        DbHelper.CreateTable($"CREATE TABLE IF NOT EXISTS {name} (FILENAME INTEGER,UCI TEXT ,{name}X REAL, {name}Y REAL, {name}Z REAL,NPD REAL,TYPE TEXT)")
    End Sub

    Private Sub CreateSupportTable()
        DbHelper.DropTableIfExists("SUPPORT")
        DbHelper.CreateTable("CREATE TABLE IF NOT EXISTS SUPPORT (FILENAME INTEGER,ITEMCODE TEXT,NAME TEXT,UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,SUPPORTDIRECTION TEXT,SUPPORTTYPE TEXT)")
    End Sub

    Private Sub CreateBoltTable()
        DbHelper.DropTableIfExists("BOLT")
        DbHelper.CreateTable("CREATE TABLE IF NOT EXISTS BOLT (FILENAME INTEGER,BOLTITEMCODE TEXT,BOLTDIA REAL,UCI TEXT,BOLTITEMDESCRIPTION TEXT,BOLTQUANTITY INTEGER,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER,BOLTLENGTH INTEGER)")
    End Sub

    Private Sub CreateMessageTable()
        DbHelper.DropTableIfExists("MESSAGE")
        DbHelper.CreateTable("CREATE TABLE IF NOT EXISTS MESSAGE (FILENAME INTEGER,ITEMCODE TEXT,TEXT TEXT,UCI TEXT)")
    End Sub

    Private Sub CreatePipeTable(name As String)
        DbHelper.DropTableIfExists(name)
        DbHelper.CreateTable($"CREATE TABLE IF NOT EXISTS {name} (FILENAME INTEGER,ITEMCODE TEXT,CUTPIECELENGTH REAL,UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER,ANGLE INTEGER)")
    End Sub

    Private Sub CreateCommonTable(name As String)
        DbHelper.DropTableIfExists(name)
        DbHelper.CreateTable($"CREATE TABLE IF NOT EXISTS {name} (FILENAME INTEGER,ITEMCODE TEXT,TAG TEXT,UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER)")
    End Sub

    ' 默认批处理大小（每批插入的记录数）
    Public Property BatchSize As Integer = 1000

    ''' <summary>
    ''' 将所有数据插入数据库（分批处理，避免内存不足）
    ''' </summary>
    Public Sub InsertAllData(Optional progressCallback As Action(Of String, Integer, Integer) = Nothing)
        ' 插入 PIPELINE 数据
        progressCallback?.Invoke("正在插入 PIPELINE 数据...", 0, 7)
        BulkInsertSimple(
            "INSERT INTO PIPELINE (FILENAME,REFERENCE,PIPINGSPEC,INSULATIONSPEC,PAINTINGSPEC,TRACINGSPEC) VALUES (@p0,@p1,@p2,@p3,@p4,@p5)",
            PipelineData, 6, progressCallback)

        ' 插入 SUPPORT 数据
        progressCallback?.Invoke("正在插入 SUPPORT 数据...", 1, 7)
        BulkInsertSimple(
            "INSERT INTO SUPPORT (FILENAME,ITEMCODE,NAME,UCI,ITEMDESCRIPTION,SKEY,SUPPORTDIRECTION,SUPPORTTYPE) VALUES (@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7)",
            SupportData, 8, progressCallback)

        ' 插入坐标点数据
        progressCallback?.Invoke("正在插入坐标点数据...", 2, 7)
        BulkInsertDictionary(PointData,
            Function(key) $"INSERT INTO {key} (FILENAME,UCI,{key}X,{key}Y,{key}Z,NPD,TYPE) VALUES (@p0,@p1,@p2,@p3,@p4,@p5,@p6)", 7, progressCallback)

        ' 插入通用组件数据
        progressCallback?.Invoke("正在插入组件数据...", 3, 7)
        BulkInsertDictionary(CommonComponents,
            Function(key) $"INSERT INTO {key} (FILENAME,ITEMCODE,TAG,UCI,ITEMDESCRIPTION,SKEY,COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER) VALUES (@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7)", 8, progressCallback)

        ' 插入管道数据
        progressCallback?.Invoke("正在插入管道数据...", 4, 7)
        BulkInsertDictionary(PipeData,
            Function(key) $"INSERT INTO {key} (FILENAME,ITEMCODE,CUTPIECELENGTH,UCI,ITEMDESCRIPTION,SKEY,COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER,ANGLE) VALUES (@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8)", 9, progressCallback)

        ' 插入 BOLT 数据
        progressCallback?.Invoke("正在插入 BOLT 数据...", 5, 7)
        BulkInsertSimple(
            "INSERT INTO BOLT (FILENAME,BOLTITEMCODE,BOLTDIA,UCI,BOLTITEMDESCRIPTION,BOLTQUANTITY,COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER,BOLTLENGTH) VALUES (@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8)",
            BoltData, 9, progressCallback)

        ' 插入 MESSAGE 数据
        progressCallback?.Invoke("正在插入 MESSAGE 数据...", 6, 7)
        BulkInsertSimple(
            "INSERT INTO MESSAGE (FILENAME,ITEMCODE,TEXT,UCI) VALUES (@p0,@p1,@p2,@p3)",
            MessageData, 4, progressCallback)

        progressCallback?.Invoke("数据插入完成", 7, 7)
    End Sub

    ''' <summary>
    ''' 通用批量插入方法 - 简单列表（分批处理，避免内存不足）
    ''' </summary>
    Private Sub BulkInsertSimple(sql As String, data As List(Of String()), paramCount As Integer, Optional progressCallback As Action(Of String, Integer, Integer) = Nothing)
        If data Is Nothing OrElse data.Count = 0 Then Return

        Dim totalCount = data.Count
        Dim batchCount = CInt(Math.Ceiling(totalCount / BatchSize))

        Using transaction = DbHelper.BeginTransaction()
            Using cmd As New SQLiteCommand(sql, DbHelper.Connection, transaction)
                ' 预创建参数
                Dim parameters(paramCount - 1) As SQLiteParameter
                For i = 0 To paramCount - 1
                    parameters(i) = New SQLiteParameter($"@p{i}", DbType.String)
                    cmd.Parameters.Add(parameters(i))
                Next

                ' 分批处理
                For batchIndex = 0 To batchCount - 1
                    Dim startIndex = batchIndex * BatchSize
                    Dim endIndex = Math.Min(startIndex + BatchSize, totalCount)

                    ' 处理当前批次
                    For i = startIndex To endIndex - 1
                        Dim item = data(i)
                        For j = 0 To paramCount - 1
                            parameters(j).Value = StringToDbValue(If(j < item.Length, item(j), Nothing))
                        Next
                        cmd.ExecuteNonQuery()
                    Next

                    ' 每处理10批提交一次，释放内存
                    If batchIndex Mod 10 = 0 AndAlso batchIndex > 0 Then
                        transaction.Commit()
                        ' 重新开启新事务
                        Using newTransaction = DbHelper.BeginTransaction()
                            cmd.Transaction = newTransaction
                        End Using
                    End If

                    ' 报告进度
                    If batchIndex Mod 5 = 0 Then
                        progressCallback?.Invoke($"批量插入进度: {endIndex}/{totalCount}", batchIndex, batchCount)
                    End If
                Next

                transaction.Commit()
            End Using
        End Using

        ' 清空已插入的数据，释放内存
        data.Clear()
        GC.Collect()
    End Sub

    ''' <summary>
    ''' 通用批量插入方法 - 字典（分批处理，避免内存不足）
    ''' </summary>
    Private Sub BulkInsertDictionary(data As Dictionary(Of String, List(Of String())), sqlBuilder As Func(Of String, String), paramCount As Integer, Optional progressCallback As Action(Of String, Integer, Integer) = Nothing)
        If data Is Nothing OrElse data.Count = 0 Then Return

        Dim totalTables = data.Count
        Dim currentTable = 0

        Using transaction = DbHelper.BeginTransaction()
            For Each kvp In data
                If kvp.Value Is Nothing OrElse kvp.Value.Count = 0 Then Continue For

                Dim tableName = kvp.Key
                Dim tableData = kvp.Value
                Dim totalCount = tableData.Count
                Dim batchCount = CInt(Math.Ceiling(totalCount / BatchSize))

                Dim sql = sqlBuilder(tableName)
                Using cmd As New SQLiteCommand(sql, DbHelper.Connection, transaction)
                    ' 预创建参数
                    Dim parameters(paramCount - 1) As SQLiteParameter
                    For i = 0 To paramCount - 1
                        parameters(i) = New SQLiteParameter($"@p{i}", DbType.String)
                        cmd.Parameters.Add(parameters(i))
                    Next

                    ' 分批处理当前表的数据
                    For batchIndex = 0 To batchCount - 1
                        Dim startIndex = batchIndex * BatchSize
                        Dim endIndex = Math.Min(startIndex + BatchSize, totalCount)

                        For i = startIndex To endIndex - 1
                            Dim item = tableData(i)
                            For j = 0 To paramCount - 1
                                parameters(j).Value = StringToDbValue(If(j < item.Length, item(j), Nothing))
                            Next
                            cmd.ExecuteNonQuery()
                        Next

                        ' 报告进度
                        If batchIndex Mod 5 = 0 Then
                            progressCallback?.Invoke($"插入 {tableName}: {endIndex}/{totalCount}", currentTable, totalTables)
                        End If
                    Next
                End Using

                ' 清空已插入的数据，释放内存
                tableData.Clear()
                currentTable += 1
            Next
            transaction.Commit()
        End Using

        ' 清空字典，释放内存
        data.Clear()
        GC.Collect()
    End Sub

    ''' <summary>
    ''' 创建 TOTAL 视图
    ''' </summary>
    Public Sub CreateTotalView()
        Dim part = ""
        For i = 0 To ComponentTypesField.Count - 1
            Dim type = ComponentTypesField(i)
            If i = 0 Then
                part &= $"SELECT FILENAME,COALESCE(ITEMCODE,'{type}') AS ITEMCODE,UCI,ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM {type}"
            ElseIf type = "BOLT" Then
                part &= $" UNION ALL SELECT FILENAME,COALESCE(BOLTITEMCODE,'BOLT') AS ITEMCODE,UCI,BOLTITEMDESCRIPTION,COMPONENTIDENTIFIER FROM BOLT"
            ElseIf type = "SUPPORT" Then
                part &= $" UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'SUPPORT') AS ITEMCODE,UCI,ITEMDESCRIPTION,'' AS COMPONENTIDENTIFIER FROM SUPPORT"
            ElseIf type = "MESSAGE" Then
                part &= $" UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'MESSAGE') AS ITEMCODE,UCI,'' AS ITEMDESCRIPTION,'' AS COMPONENTIDENTIFIER FROM MESSAGE"
            Else
                part &= $" UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'{type}') AS ITEMCODE,UCI,ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM {type}"
            End If
        Next

        DbHelper.DropViewIfExists("TOTAL")
        DbHelper.CreateTable($"CREATE VIEW IF NOT EXISTS TOTAL AS {part}")
    End Sub

#Region "类型转换辅助方法"

    ''' <summary>
    ''' 字符串值转换：空字符串转为 Nothing
    ''' </summary>
    Private Function ValueS(value As String) As String
        If value <> "" Then
            Return value
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' 整数值转换：尝试解析为整数，失败返回 Nothing
    ''' </summary>
    Private Function ValueI(value As String) As String
        Dim i0 As Integer
        If Integer.TryParse(value, i0) Then
            Return i0.ToString()
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' 双精度值转换：尝试解析为双精度，失败返回 Nothing
    ''' </summary>
    Private Function ValueD(value As String) As String
        Dim d0 As Double
        If Double.TryParse(value, d0) Then
            Return d0.ToString()
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' 将字符串值转换为数据库参数值
    ''' </summary>
    Private Function StringToDbValue(value As String) As Object
        Return If(String.IsNullOrEmpty(value), DirectCast(DBNull.Value, Object), value)
    End Function

#End Region

#Region "兼容层 - 供 Form1 直接调用的方法"

    ''' <summary>
    ''' 插入 PIPELINE 数据（兼容方法）
    ''' </summary>
    Public Sub InsertSQLiteTablePIPELINE(row As List(Of String), filename As Integer)
        Dim data As String() = {filename.ToString(), ValueS(row(0)), ValueS(row(1)), ValueS(row(2)), ValueS(row(3)), ValueS(row(4))}
        AddPipeline(data)
    End Sub

    ''' <summary>
    ''' 插入 SUPPORT 数据（兼容方法）
    ''' </summary>
    Public Sub InsertSQLiteTableSupport(row As List(Of String), filename As Integer)
        Dim data As String() = {filename.ToString(), ValueS(row(0)), ValueS(row(1)), ValueS(row(2)), ValueS(row(3)), ValueS(row(4)), ValueS(row(5)), ValueS(row(6))}
        AddSupport(data)
    End Sub

    ''' <summary>
    ''' 插入 MESSAGE 数据（兼容方法）
    ''' </summary>
    Public Sub InsertSQLiteTableMessage(row As List(Of String), filename As Integer)
        Dim data As String() = {filename.ToString(), ValueS(row(0)), ValueS(row(1)), ValueS(row(2)), ValueS(row(3))}
        AddMessage(data)
    End Sub

    ''' <summary>
    ''' 插入 BOLT 数据（兼容方法）
    ''' </summary>
    Public Sub InsertSQLiteTableBolt(row As List(Of String), filename As Integer)
        ' BOLT: FILENAME, BOLTITEMCODE, BOLTDIA, UCI, BOLTITEMDESCRIPTION, BOLTQUANTITY, COMPONENTIDENTIFIER, MASTERCOMPONENTIDENTIFIER, BOLTLENGTH
        Dim data As String() = {filename.ToString(), ValueS(row(0)), ValueD(row(1)), ValueS(row(2)), ValueS(row(3)), ValueI(row(4)), ValueI(row(5)), ValueI(row(6)), ValueI(row(7))}
        AddBolt(data)
    End Sub

    ''' <summary>
    ''' 插入管道数据（兼容方法）
    ''' </summary>
    Public Sub InsertSQLiteTablePipe(pipeType As String, row As List(Of String), filename As Integer)
        ' PIPE: FILENAME, ITEMCODE, CUTPIECELENGTH, UCI, ITEMDESCRIPTION, SKEY, COMPONENTIDENTIFIER, MASTERCOMPONENTIDENTIFIER, ANGLE
        Dim data As String() = {filename.ToString(), ValueS(row(0)), ValueD(row(1)), ValueS(row(2)), ValueS(row(3)), ValueS(row(4)), ValueI(row(5)), ValueI(row(6)), ValueI(row(7))}
        AddPipe(pipeType, data)
    End Sub

    ''' <summary>
    ''' 插入通用组件数据（兼容方法）
    ''' </summary>
    Public Sub InsertSQLiteTable(componentType As String, row As List(Of String), filename As Integer)
        ' 通用组件: FILENAME, ITEMCODE, TAG, UCI, ITEMDESCRIPTION, SKEY, COMPONENTIDENTIFIER, MASTERCOMPONENTIDENTIFIER
        Dim data As String() = {filename.ToString(), ValueS(row(0)), ValueS(row(1)), ValueS(row(2)), ValueS(row(3)), ValueS(row(4)), ValueI(row(5)), ValueI(row(6))}
        AddCommonComponent(componentType, data)
    End Sub

    ''' <summary>
    ''' 插入坐标点数据（兼容方法）
    ''' </summary>
    Public Sub InsertSQLiteTablePOINT(pointType As String, pointValues As IEnumerable(Of String), uci As String, filename As Integer)
        Dim valueList = pointValues.ToList()
        Dim a0 As Double, a1 As Double, a2 As Double
        Dim v0 As String = Nothing, v1 As String = Nothing, v2 As String = Nothing, v3 As String = Nothing, v4 As String = Nothing

        Select Case valueList.Count
            Case 5
                If Double.TryParse(valueList(0), a0) AndAlso Double.TryParse(valueList(1), a1) AndAlso Double.TryParse(valueList(2), a2) Then
                    v0 = a0.ToString()
                    v1 = a1.ToString()
                    v2 = a2.ToString()
                    v3 = ValueD(valueList(3))
                    v4 = ValueS(valueList(4))
                Else
                    Exit Sub
                End If
            Case 4
                If Double.TryParse(valueList(0), a0) AndAlso Double.TryParse(valueList(1), a1) AndAlso Double.TryParse(valueList(2), a2) Then
                    v0 = a0.ToString()
                    v1 = a1.ToString()
                    v2 = a2.ToString()
                    v3 = ValueD(valueList(3))
                    v4 = Nothing
                Else
                    Exit Sub
                End If
            Case 3
                If Double.TryParse(valueList(0), a0) AndAlso Double.TryParse(valueList(1), a1) AndAlso Double.TryParse(valueList(2), a2) Then
                    v0 = a0.ToString()
                    v1 = a1.ToString()
                    v2 = a2.ToString()
                    v3 = Nothing
                    v4 = Nothing
                Else
                    Exit Sub
                End If
        End Select

        Dim data As String() = {filename.ToString(), uci, v0, v1, v2, v3, v4}
        AddPoint(pointType, data)
    End Sub

    ''' <summary>
    ''' 创建 TOTAL 视图（兼容方法）
    ''' </summary>
    Public Sub CreateSQLiteViewAll(types As String())
        CreateTotalView()
    End Sub

#End Region

End Class
