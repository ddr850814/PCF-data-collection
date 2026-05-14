Imports System.ComponentModel
Imports System.Data.SQLite
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Public Class Form1
    Private WithEvents BgWorker As New BackgroundWorker()
    Dim count As Integer
    Dim Path1 As String
    Dim Path2 As String

    ' 当前数据收集器实例
    Private _dataCollector As PcfDataCollector

    ' 错误日志记录
    Private _errorLog As New List(Of String)
    Private _successCount As Integer = 0
    Private _errorCount As Integer = 0

    ''' <summary>
    ''' 动态属性描述符（用于PropertyGrid动态显示属性）
    ''' </summary>
    Public Class DynamicPropertyDescriptor
        Inherits PropertyDescriptor

        Private _value As Object
        Private _category As String

        ' 共享方法：构建包含值的描述
        Private Shared Function BuildDescription(description As String, value As Object) As String
            Dim valueStr As String = If(value IsNot Nothing, value.ToString(), "")
            If TypeOf value Is CoordinatePoint Then
                Dim pt = DirectCast(value, CoordinatePoint)
                valueStr = $"({pt.X}, {pt.Y}, {pt.Z}) NPD:{pt.NPD} Type:{pt.Type}"
            End If
            Return valueStr
        End Function

        Public Sub New(name As String, value As Object, category As String, description As String)
            MyBase.New(name, New Attribute() {
                New DisplayNameAttribute(name),
                New DescriptionAttribute(BuildDescription(description, value)),
                New CategoryAttribute(category)
            })
            _value = value
            _category = category
        End Sub

        Public Overrides ReadOnly Property ComponentType As Type
            Get
                Return GetType(DynamicPropertyBag)
            End Get
        End Property

        Public Overrides ReadOnly Property PropertyType As Type
            Get
                Return If(_value IsNot Nothing, _value.GetType(), GetType(String))
            End Get
        End Property

        Public Overrides ReadOnly Property IsReadOnly As Boolean
            Get
                Return False
            End Get
        End Property

        Public Overrides Function GetValue(component As Object) As Object
            Return _value
        End Function

        Public Overrides Sub SetValue(component As Object, value As Object)
            _value = value
        End Sub

        Public Overrides Function CanResetValue(component As Object) As Boolean
            Return False
        End Function

        Public Overrides Sub ResetValue(component As Object)
        End Sub

        Public Overrides Function ShouldSerializeValue(component As Object) As Boolean
            Return False
        End Function
    End Class

    ''' <summary>
    ''' 动态属性包（支持动态添加属性到PropertyGrid）
    ''' </summary>
    Public Class DynamicPropertyBag
        Implements ICustomTypeDescriptor

        Private _properties As New Dictionary(Of String, Object)()

        Public Sub AddProperty(name As String, value As Object, category As String, description As String)
            _properties(name) = value
        End Sub

        ' 无参数版本（接口要求）- 显式接口实现
        Private Function GetPropertiesNoArgs() As PropertyDescriptorCollection Implements ICustomTypeDescriptor.GetProperties
            Return GetProperties(Nothing)
        End Function

        ' 带参数版本（接口实现）
        Public Function GetProperties(attributes() As Attribute) As PropertyDescriptorCollection Implements ICustomTypeDescriptor.GetProperties
            Dim props As New List(Of PropertyDescriptor)()

            ' 添加基本属性
            If _properties.ContainsKey("FILENAME") Then
                props.Add(New DynamicPropertyDescriptor("文件名", _properties("FILENAME"), "基本信息", "PCF文件名"))
            End If
            If _properties.ContainsKey("REFERENCE") Then
                props.Add(New DynamicPropertyDescriptor("管线编号", _properties("REFERENCE"), "基本信息", "管线参考编号"))
            End If
            If _properties.ContainsKey("PIPINGSPEC") Then
                props.Add(New DynamicPropertyDescriptor("管道等级", _properties("PIPINGSPEC"), "基本信息", "管道规格等级"))
            End If
            If _properties.ContainsKey("INSULATIONSPEC") Then
                props.Add(New DynamicPropertyDescriptor("保温等级", _properties("INSULATIONSPEC"), "基本信息", "保温规格等级"))
            End If
            If _properties.ContainsKey("PAINTINGSPEC") Then
                props.Add(New DynamicPropertyDescriptor("油漆等级", _properties("PAINTINGSPEC"), "基本信息", "油漆规格等级"))
            End If
            If _properties.ContainsKey("TRACINGSPEC") Then
                props.Add(New DynamicPropertyDescriptor("伴热等级", _properties("TRACINGSPEC"), "基本信息", "伴热规格等级"))
            End If

            ' 添加动态属性（COMPONENT_ATTRIBUTES）
            For Each kvp As KeyValuePair(Of String, Object) In _properties
                If Not {"FILENAME", "REFERENCE", "PIPINGSPEC", "INSULATIONSPEC", "PAINTINGSPEC", "TRACINGSPEC"}.Contains(kvp.Key) Then
                    props.Add(New DynamicPropertyDescriptor(kvp.Key, kvp.Value, "属性", kvp.Key))
                End If
            Next

            Return New PropertyDescriptorCollection(props.ToArray())
        End Function

        Public Function GetAttributes() As AttributeCollection Implements ICustomTypeDescriptor.GetAttributes
            Return TypeDescriptor.GetAttributes(Me, True)
        End Function

        Public Function GetClassName() As String Implements ICustomTypeDescriptor.GetClassName
            Return TypeDescriptor.GetClassName(Me, True)
        End Function

        Public Function GetComponentName() As String Implements ICustomTypeDescriptor.GetComponentName
            Return TypeDescriptor.GetComponentName(Me, True)
        End Function

        Public Function GetConverter() As TypeConverter Implements ICustomTypeDescriptor.GetConverter
            Return TypeDescriptor.GetConverter(Me, True)
        End Function

        Public Function GetDefaultEvent() As EventDescriptor Implements ICustomTypeDescriptor.GetDefaultEvent
            Return TypeDescriptor.GetDefaultEvent(Me, True)
        End Function

        Public Function GetDefaultProperty() As PropertyDescriptor Implements ICustomTypeDescriptor.GetDefaultProperty
            Return TypeDescriptor.GetDefaultProperty(Me, True)
        End Function

        Public Function GetEditor(editorBaseType As Type) As Object Implements ICustomTypeDescriptor.GetEditor
            Return TypeDescriptor.GetEditor(Me, editorBaseType, True)
        End Function

        Public Function GetEvents() As EventDescriptorCollection Implements ICustomTypeDescriptor.GetEvents
            Return TypeDescriptor.GetEvents(Me, True)
        End Function

        Public Function GetEvents(attributes() As Attribute) As EventDescriptorCollection Implements ICustomTypeDescriptor.GetEvents
            Return TypeDescriptor.GetEvents(Me, True)
        End Function

        Public Function GetPropertyOwner(pd As PropertyDescriptor) As Object Implements ICustomTypeDescriptor.GetPropertyOwner
            Return Me
        End Function
    End Class

    ''' <summary>
    ''' Pipeline 属性类（用于PropertyGrid显示）
    ''' </summary>
    Public Class PipelineProperties
        <DisplayName("文件名")>
        <Description("PCF文件名")>
        Public Property FileName As String

        <DisplayName("管线编号")>
        <Description("管线参考编号")>
        Public Property Reference As String

        <DisplayName("管道等级")>
        <Description("管道规格等级")>
        Public Property PipingSpec As String

        <DisplayName("保温等级")>
        <Description("保温规格等级")>
        Public Property InsulationSpec As String

        <DisplayName("油漆等级")>
        <Description("油漆规格等级")>
        Public Property PaintingSpec As String

        <DisplayName("伴热等级")>
        <Description("伴热规格等级")>
        Public Property TracingSpec As String

        ' 动态属性存储
        <Browsable(False)>
        Public Property DynamicProperties As New Dictionary(Of String, String)
    End Class

    ''' <summary>
    ''' 坐标点详情类（用于PropertyGrid展开显示XYZ、NPD和TYPE）
    ''' </summary>
    <TypeConverter(GetType(ExpandableObjectConverter))>
    Public Class CoordinatePoint
        <DisplayName("X坐标")>
        <Description("X轴坐标值")>
        Public Property X As String

        <DisplayName("Y坐标")>
        <Description("Y轴坐标值")>
        Public Property Y As String

        <DisplayName("Z坐标")>
        <Description("Z轴坐标值")>
        Public Property Z As String

        <DisplayName("公称直径")>
        <Description("NPD - 公称管道直径")>
        Public Property NPD As String

        <DisplayName("端面形式")>
        <Description("TYPE - 端面连接形式")>
        Public Property Type As String

        Public Overrides Function ToString() As String
            Return ""  ' 父属性在属性列表中不显示值，子属性已显示详细信息
        End Function
    End Class

    ''' <summary>
    ''' 获取或创建数据收集器
    ''' </summary>
    Private ReadOnly Property DataCollector As PcfDataCollector
        Get
            If _dataCollector Is Nothing Then
                _dataCollector = GetOrCreateDataCollector()
            End If
            Return _dataCollector
        End Get
    End Property

    ''' <summary>
    ''' 窗体加载事件 - 初始化控件状态
    ''' </summary>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 未打开数据库时，下拉列表不可用
        ComboBox1.Enabled = False
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("COUNT")
        ComboBox1.SelectedIndex = 0

        ' 初始化 PropertyGrid
        PropertyGrid1.SelectedObject = Nothing
        PropertyGrid1.HelpVisible = True
        PropertyGrid1.ToolbarVisible = False
    End Sub

    Private Sub WritePreviously(report As Boolean, filename As Integer, ByRef rowP As List(Of String), ByRef row As List(Of String), Title As String, ByRef ENDPOINT As List(Of String))
        If rowP.Any(Function(n) n <> "") Then
            If Title = "PIPELINE" Then
                DataCollector.InsertSQLiteTablePIPELINE(rowP, filename)
            End If
            rowP = New List(Of String) From {"", "", "", "", ""}
        ElseIf row.Any(Function(n) n <> "") Then
            If report = True Then
                Select Case Title
                    Case "SUPPORT"
                        DataCollector.InsertSQLiteTableSupport(row, filename)
                    Case "MESSAGE"
                        row(2) = DataCollector.GetNextMessageId()
                        DataCollector.InsertSQLiteTableMessage(row, filename)
                    Case "INSTRUMENT_RETURN"
                        DataCollector.InsertSQLiteTableInstrumentReturn(row, filename)
                    Case "TRAP_OFFSET"
                        DataCollector.InsertSQLiteTableTrapOffset(row, filename)
                    Case "BOLT"
                        DataCollector.InsertSQLiteTableBolt(row, filename)
                    Case "PIPE", "PIPE_FIXED", "BEND"
                        DataCollector.InsertSQLiteTablePipe(Title, row, filename)
                    Case "TEE_STUB"
                        row(2) = DataCollector.GetNextTeeStubId()
                        DataCollector.InsertSQLiteTable(Title, row, filename)
                    Case Else
                        DataCollector.InsertSQLiteTable(Title, row, filename)
                End Select
                For Each InputArray In From InputS In ENDPOINT Select Regex.Split(InputS, "\s+")
                    If InputArray.Count() >= 4 And InputArray.Count() <= 6 Then
                        DataCollector.InsertSQLiteTablePOINT(InputArray(0).Replace("-", "_"), InputArray.Skip(1), row(2), filename)
                    End If
                Next
            End If
            row = New List(Of String) From {"", "", "", "", "", "", "", ""}
        End If
        ENDPOINT.Clear()
    End Sub

    ''' <summary>
    ''' 处理单个PCF文件
    ''' </summary>
    Private Sub ProcessSingleFile(filePath As String, fileIndex As Integer)
        Dim report = True
        Dim Title As String = Nothing
        Dim rowP As New List(Of String) From {"", "", "", "", ""}
        Dim row As New List(Of String) From {"", "", "", "", "", "", "", ""}
        Dim ENDPOINT As New List(Of String)

        Using reader As New StreamReader(filePath, Encoding.GetEncoding("GBK"))
            Dim line As String = reader.ReadLine()
            While line IsNot Nothing
                If line.StartsWith("PIPELINE-REFERENCE") Then
                    rowP(0) = line.Replace("PIPELINE-REFERENCE", "").Trim
                    Title = "PIPELINE"
                ElseIf ComponentTypes.Contains(line.Split(" "c)(0).Replace("-", "_")) Then
                    WritePreviously(report, fileIndex, rowP, row, Title, ENDPOINT)
                    Title = line.Split(" "c)(0).Replace("-", "_")
                    report = True
                ElseIf Not line.StartsWith(" ") Then
                    WritePreviously(report, fileIndex, rowP, row, Title, ENDPOINT)
                    Title = Nothing
                    report = True
                ElseIf Title IsNot Nothing Then
                    If Title = "PIPELINE" Then
                        If line.TrimStart.StartsWith("PIPING-SPEC") Then
                            rowP(1) = line.Replace("PIPING-SPEC", "").Trim
                        ElseIf line.TrimStart.StartsWith("INSULATION-SPEC") Then
                            rowP(2) = line.Replace("INSULATION-SPEC", "").Trim
                        ElseIf line.TrimStart.StartsWith("PAINTING-SPEC") Then
                            rowP(3) = line.Replace("PAINTING-SPEC", "").Trim
                        ElseIf line.TrimStart.StartsWith("TRACING-SPEC") Then
                            rowP(4) = line.Replace("TRACING-SPEC", "").Trim
                        ElseIf line.TrimStart.StartsWith("ATTRIBUTE") Then
                            Dim trimmedLine = line.TrimStart()
                            Dim spaceIndex = trimmedLine.IndexOf(" "c)
                            If spaceIndex > 0 Then
                                Dim attrName = trimmedLine.Substring(0, spaceIndex).Trim()
                                Dim attrValue = trimmedLine.Substring(spaceIndex + 1).Trim()
                                Dim pipelineRef = rowP(0)
                                If Not String.IsNullOrEmpty(pipelineRef) Then
                                    DataCollector.AddAttribute(fileIndex.ToString(), "PIPELINE:" & pipelineRef, attrName, attrValue)
                                End If
                            End If
                        End If
                    Else
                        If line.TrimStart.StartsWith("ANGLE") Then
                            row(7) = line.Replace("ANGLE", "").Trim
                        ElseIf line.TrimStart.StartsWith("BOLT-LENGTH") Then
                            row(7) = line.Replace("BOLT-LENGTH", "").Trim
                        ElseIf line.TrimStart.StartsWith("MASTER-COMPONENT-IDENTIFIER") Then
                            row(6) = line.Replace("MASTER-COMPONENT-IDENTIFIER", "").Trim
                        ElseIf line.TrimStart.StartsWith("SUPPORT-TYPE") Then
                            row(6) = line.Replace("SUPPORT-TYPE", "").Trim
                        ElseIf line.TrimStart.StartsWith("COMPONENT-IDENTIFIER") Then
                            row(5) = line.Replace("COMPONENT-IDENTIFIER", "").Trim
                        ElseIf line.TrimStart.StartsWith("SUPPORT-DIRECTION") Then
                            row(5) = line.Replace("SUPPORT-DIRECTION", "").Trim
                        ElseIf line.TrimStart.StartsWith("SKEY") Then
                            row(4) = line.Replace("SKEY", "").Trim
                        ElseIf line.TrimStart.StartsWith("BOLT-QUANTITY") Then
                            row(4) = line.Replace("BOLT-QUANTITY", "").Trim
                        ElseIf line.TrimStart.StartsWith("ITEM-DESCRIPTION") Then
                            row(3) = line.Replace("ITEM-DESCRIPTION", "").Trim
                        ElseIf line.TrimStart.StartsWith("BOLT-ITEM-DESCRIPTION") Then
                            row(3) = line.Replace("BOLT-ITEM-DESCRIPTION", "").Trim
                        ElseIf line.TrimStart.StartsWith("UCI") Then
                            row(2) = line.Replace("UCI", "").Trim
                        ElseIf line.TrimStart.StartsWith("TAG") Then
                            row(1) = line.Replace("TAG", "").Trim
                        ElseIf line.TrimStart.StartsWith("NAME") Then
                            row(1) = line.Replace("NAME", "").Trim
                        ElseIf line.TrimStart.StartsWith("BOLT-DIA") Then
                            row(1) = line.Replace("BOLT-DIA", "").Trim
                        ElseIf line.TrimStart.StartsWith("CUT-PIECE-LENGTH") Then
                            row(1) = line.Replace("CUT-PIECE-LENGTH", "").Trim
                        ElseIf line.TrimStart.StartsWith("TEXT") Then
                            row(1) = line.Replace("TEXT", "").Trim
                        ElseIf line.TrimStart.StartsWith("ITEM-CODE") Then
                            row(0) = line.Replace("ITEM-CODE", "").Trim
                        ElseIf line.TrimStart.StartsWith("BOLT-ITEM-CODE") Then
                            row(0) = line.Replace("BOLT-ITEM-CODE", "").Trim
                        ElseIf line.TrimStart.StartsWith("END-POINT") Or line.TrimStart.StartsWith("CENTRE-POINT") Or line.TrimStart.StartsWith("JACKET-POINT") Or line.TrimStart.StartsWith("BRANCH1-POINT") Or line.TrimStart.StartsWith("CO-ORDS") Then
                            ENDPOINT.Add(line.Trim)
                        ElseIf line.TrimStart.StartsWith("MATERIAL-LIST") And line.TrimEnd.EndsWith("EXCLUDE") Then
                            report = False
                        ElseIf line.TrimStart.StartsWith("ATTRIBUTE") Then
                            Dim trimmedLine = line.TrimStart()
                            Dim spaceIndex = trimmedLine.IndexOf(" "c)
                            If spaceIndex > 0 Then
                                Dim attrName = trimmedLine.Substring(0, spaceIndex).Trim()
                                Dim attrValue = trimmedLine.Substring(spaceIndex + 1).Trim()
                                Dim currentUci = row(2)
                                If Not String.IsNullOrEmpty(currentUci) Then
                                    DataCollector.AddAttribute(fileIndex.ToString(), currentUci, attrName, attrValue)
                                End If
                            End If
                        End If
                    End If
                End If
                line = reader.ReadLine()
            End While
            WritePreviously(report, fileIndex, rowP, row, Title, ENDPOINT)
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim FolderDialog As New FolderBrowserDialog
            Dim result = FolderDialog.ShowDialog
            Path1 = FolderDialog.SelectedPath
            If result = DialogResult.OK Then
                If Not Directory.GetFiles(Path1, "*.pcf", SearchOption.AllDirectories).Any Then
                    MsgBox("木有发现pcf")
                Else
                    BgWorker.WorkerSupportsCancellation = True
                    BgWorker.RunWorkerAsync()
                End If
                Path2 = Path1 & "\pcf.db"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub BgWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles BgWorker.DoWork
        Dim divisibleByThree = Directory.GetFiles(Path1, "*.pcf", SearchOption.AllDirectories)
        count = 0

        Invoke(Sub()
                   Enabled = False
                   Button1.Text = count & "/" & divisibleByThree.Count
               End Sub)

        ' 初始化数据库连接
        ConnectionStringBuilder = New SQLiteConnectionStringBuilder() With {.DataSource = Path1 & "\pcf.db", .ForeignKeys = True}

        ' 初始化数据库助手
        InitializeDatabaseHelper()

        ' 创建新的数据收集器实例
        _dataCollector = CreateDataCollector()

        ' 使用数据收集器创建所有表（统一方法）
        DataCollector.CreateAllTables()

        ' 初始化错误计数器
        _successCount = 0
        _errorCount = 0
        _errorLog.Clear()

        For i = 0 To divisibleByThree.Count() - 1
            Try
                ProcessSingleFile(divisibleByThree(i), i)
                _successCount += 1
            Catch ex As Exception
                _errorCount += 1
                Dim errorMsg = $"文件 {i + 1}: {Path.GetFileName(divisibleByThree(i))} - {ex.Message}"
                _errorLog.Add(errorMsg)
                ' 继续处理下一个文件
            End Try

            ' 更新进度显示
            Dim currentIndex = i + 1
            Dim totalCount = divisibleByThree.Count
            Invoke(Sub()
                       Button1.Text = $"{currentIndex}/{totalCount} (成功:{_successCount}, 失败:{_errorCount})"
                   End Sub)
        Next
        Invoke(Sub()
                   Button1.Text = "插入数据表"
               End Sub)

        ' 使用数据收集器插入所有数据
        DataCollector.InsertAllData()

        Invoke(Sub()
                   Button1.Text = "插入视图"
               End Sub)
        DataCollector.CreateSQLiteViewAll(ComponentTypes)

        ' 使用 DbHelper 创建视图
        DbHelper.ExecuteNonQuery($"DROP VIEW IF EXISTS COUNT;CREATE VIEW IF NOT EXISTS COUNT AS SELECT PIPELINE.FILENAME, REFERENCE, TOTAL.ITEMCODE, ITEMDESCRIPTION, COUNT( * ) As COUNT 
        FROM
        (SELECT * FROM TOTAL GROUP BY UCI,FILENAME) AS TOTAL 
        Join PIPELINE ON PIPELINE.FILENAME = TOTAL.FILENAME
        Group BY REFERENCE, TOTAL.ITEMCODE, ITEMDESCRIPTION 
        ORDER BY REFERENCE")

        For Each Type In ComponentTypes
            Dim sql As String
            Select Case Type
                Case "PIPE"
                    sql = $"DROP VIEW IF EXISTS {"X" & Type};CREATE VIEW IF NOT EXISTS {"X" & Type} AS SELECT PIPELINE.FILENAME,UCI, REFERENCE, ITEMCODE, ITEMDESCRIPTION, CUTPIECELENGTH
            FROM 
            	( SELECT * FROM {Type} GROUP BY UCI,FILENAME) AS {Type} 
                JOIN PIPELINE ON PIPELINE.FILENAME = {Type}.FILENAME
            ORDER BY REFERENCE"
                Case "BEND"
                    sql = "DROP VIEW IF EXISTS XBEND;CREATE VIEW IF NOT EXISTS XBEND AS SELECT PIPELINE.FILENAME,UCI, REFERENCE,ITEMCODE, ITEMDESCRIPTION,CUTPIECELENGTH,ANGLE
            FROM 
            	( SELECT * FROM BEND GROUP BY UCI,FILENAME) AS BEND 
                JOIN PIPELINE ON PIPELINE.FILENAME = BEND.FILENAME
            ORDER BY REFERENCE"
                Case "INSTRUMENT", "INSTRUMENT_ANGLE"
                    sql = $"DROP VIEW IF EXISTS {"X" & Type};CREATE VIEW IF NOT EXISTS {"X" & Type} AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, TAG
            FROM 
            	( SELECT * FROM {Type} GROUP BY UCI,FILENAME) AS {Type} 
                JOIN PIPELINE ON PIPELINE.FILENAME = {Type}.FILENAME
            ORDER BY REFERENCE"
                Case "INSTRUMENT_RETURN"
                    sql = "DROP VIEW IF EXISTS XINSTRUMENT_RETURN;CREATE VIEW IF NOT EXISTS XINSTRUMENT_RETURN AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, TAG, ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM INSTRUMENT_RETURN GROUP BY UCI,FILENAME) AS INSTRUMENT_RETURN 
                JOIN PIPELINE ON PIPELINE.FILENAME = INSTRUMENT_RETURN.FILENAME
            ORDER BY REFERENCE"
                Case "TRAP_OFFSET"
                    sql = "DROP VIEW IF EXISTS XTRAP_OFFSET;CREATE VIEW IF NOT EXISTS XTRAP_OFFSET AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, TAG, ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM TRAP_OFFSET GROUP BY UCI,FILENAME) AS TRAP_OFFSET 
                JOIN PIPELINE ON PIPELINE.FILENAME = TRAP_OFFSET.FILENAME
            ORDER BY REFERENCE"
                Case "SUPPORT"
                    sql = "DROP VIEW IF EXISTS XSUPPORT;CREATE VIEW IF NOT EXISTS XSUPPORT AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,NAME, ITEMDESCRIPTION,SUPPORTDIRECTION,SUPPORTTYPE
            FROM 
            	( SELECT * FROM SUPPORT GROUP BY UCI,FILENAME) AS SUPPORT 
                JOIN PIPELINE ON PIPELINE.FILENAME = SUPPORT.FILENAME
            ORDER BY REFERENCE"
                Case "GASKET"
                    sql = "DROP VIEW IF EXISTS XGASKET;CREATE VIEW IF NOT EXISTS XGASKET AS SELECT PIPELINE.FILENAME,GASKET.UCI,REFERENCE,GASKET.ITEMCODE, GASKET.ITEMDESCRIPTION,TOTAL.ITEMCODE,TOTAL.ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM GASKET GROUP BY UCI,FILENAME) AS GASKET 
                JOIN PIPELINE ON PIPELINE.FILENAME = GASKET.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = GASKET.FILENAME AND TOTAL.COMPONENTIDENTIFIER=GASKET.MASTERCOMPONENTIDENTIFIER
            ORDER BY REFERENCE"
                Case "BOLT"
                    sql = "DROP VIEW IF EXISTS XBOLT;CREATE VIEW IF NOT EXISTS XBOLT AS SELECT PIPELINE.FILENAME,BOLT.UCI,REFERENCE,BOLTITEMCODE,BOLTITEMDESCRIPTION,BOLTDIA, BOLTQUANTITY, BOLTLENGTH, ITEMCODE,ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM BOLT GROUP BY UCI,FILENAME) AS BOLT 
                JOIN PIPELINE ON PIPELINE.FILENAME = BOLT.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = BOLT.FILENAME AND TOTAL.COMPONENTIDENTIFIER=BOLT.MASTERCOMPONENTIDENTIFIER
            ORDER BY REFERENCE"
                Case "WELD"
                    sql = "DROP VIEW IF EXISTS XWELD;CREATE VIEW IF NOT EXISTS XWELD AS SELECT
	PIPELINE.FILENAME,
	WELD.UCI,
	REFERENCE,
	Q.NPD,
	TOTAL.ITEMCODE,
	TOTAL.ITEMDESCRIPTION 
FROM
	( SELECT * FROM WELD GROUP BY UCI, FILENAME ) AS WELD
	JOIN PIPELINE ON PIPELINE.FILENAME = WELD.FILENAME
	JOIN TOTAL ON TOTAL.FILENAME = WELD.FILENAME 
	AND TOTAL.COMPONENTIDENTIFIER = WELD.MASTERCOMPONENTIDENTIFIER
	JOIN ( SELECT * FROM END_POINT GROUP BY FILENAME, UCI ) AS Q ON Q.FILENAME = WELD.FILENAME 
	AND Q.UCI = WELD.UCI 
ORDER BY
	REFERENCE"
                Case "TEE_STUB"
                    sql = $"DROP VIEW IF EXISTS XTEE_STUB;CREATE VIEW IF NOT EXISTS XTEE_STUB AS SELECT PIPELINE.FILENAME,UCI,REFERENCE
            FROM 
            	( SELECT * FROM TEE_STUB GROUP BY UCI,FILENAME) AS TEE_STUB 
                JOIN PIPELINE ON PIPELINE.FILENAME = TEE_STUB.FILENAME
            ORDER BY REFERENCE"
                Case "_UNION"
                    sql = "DROP VIEW IF EXISTS X_UNION;CREATE VIEW IF NOT EXISTS X_UNION AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM _UNION GROUP BY UCI,FILENAME) AS _UNION 
                JOIN PIPELINE ON PIPELINE.FILENAME = _UNION.FILENAME
            ORDER BY REFERENCE"
                Case "MESSAGE"
                    sql = $"DROP VIEW IF EXISTS XMESSAGE;CREATE VIEW IF NOT EXISTS XMESSAGE AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,TEXT
            FROM 
            	( SELECT * FROM MESSAGE GROUP BY UCI,FILENAME) AS MESSAGE 
                JOIN PIPELINE ON PIPELINE.FILENAME = MESSAGE.FILENAME
            ORDER BY REFERENCE"
                Case Else
                    sql = $"DROP VIEW IF EXISTS {"X" & Type};CREATE VIEW IF NOT EXISTS {"X" & Type} AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM {Type} GROUP BY UCI,FILENAME) AS {Type} 
                JOIN PIPELINE ON PIPELINE.FILENAME = {Type}.FILENAME
            ORDER BY REFERENCE"
            End Select
            DbHelper.ExecuteNonQuery(sql)
        Next

        ' 创建带属性的组件视图
        DataCollector.CreateComponentWithAttributesView()

        Invoke(Sub()
                   Enabled = True
                   SelectCount()
               End Sub)
    End Sub
    Private Sub BgWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BgWorker.RunWorkerCompleted
        Invoke(Sub()
                   Button1.Text = "写入数据库"

                   ' 显示处理结果和错误日志
                   If _errorCount > 0 Then
                       Dim errorSummary = $"处理完成。成功: {_successCount}, 失败: {_errorCount}" & vbCrLf & vbCrLf &
                                         "错误详情 (前10条):" & vbCrLf &
                                         String.Join(vbCrLf, _errorLog.Take(10))

                       If _errorLog.Count > 10 Then
                           errorSummary &= vbCrLf & $"... 还有 {_errorLog.Count - 10} 个错误"
                       End If

                       MessageBox.Show(errorSummary, "PCF处理结果", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                   ElseIf _successCount > 0 Then
                       MessageBox.Show($"所有 {_successCount} 个文件处理成功！", "PCF处理完成", MessageBoxButtons.OK, MessageBoxIcon.Information)
                   End If
               End Sub)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            Dim oOpenFileDialog As New OpenFileDialog With {
            .Filter = "SQLite|*.db"
        }
            Dim result = oOpenFileDialog.ShowDialog
            Path2 = oOpenFileDialog.FileName
            If result = DialogResult.OK Then
                ConnectionStringBuilder = New SQLiteConnectionStringBuilder() With {.DataSource = Path2, .ForeignKeys = True}
                InitializeDatabaseHelper()
                SelectCount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            RemoveHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
            If ConnectionStringBuilder IsNot Nothing AndAlso ComboBox1.SelectedItem IsNot Nothing Then
                DataGridView1.DataSource = Nothing
                PropertyGrid1.SelectedObject = Nothing
                Dim selectedType = ComboBox1.SelectedItem.ToString()
                Dim sql As String

                ' 根据选择的类型构建查询
                If selectedType = "COUNT" Then
                    ' 统计信息
                    sql = "SELECT * FROM COUNT;"
                    DataGridView1.DataSource = DbHelper.FillDataTable(sql)
                    DataGridView1.Columns.Item(0).Visible = False
                    
                    ' 主动触发一次属性加载（选中第一行）
                    If DataGridView1.Rows.Count > 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(1)
                    End If
                Else
                    ' 组件类型，将显示名称转换回数据库名称
                    Dim dbType = selectedType.Replace("-", "_")

                    ' 特殊处理_UNION类型（在数据库中存储为_UNION，显示为UNION）
                    If dbType = "UNION" Then
                        dbType = "_UNION"
                    End If

                    Dim viewName = "X" & dbType
                    sql = $"SELECT * FROM {viewName};"
                    DataGridView1.DataSource = DbHelper.FillDataTable(sql)

                    ' 隐藏FILENAME和UCI列（如果存在）
                    If DataGridView1.Columns.Count > 0 Then
                        DataGridView1.Columns.Item(0).Visible = False
                        If DataGridView1.Columns.Count > 1 Then
                            DataGridView1.Columns.Item(1).Visible = False
                        End If
                    End If
                End If

                AddHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub SelectCount()
        ' 动态填充ComboBox，只显示有数据的组件类型
        PopulateComboBoxWithData()

        ' 启用下拉列表
        ComboBox1.Enabled = True

        If ComboBox1.Items.Count > 0 Then
            If ComboBox1.SelectedIndex = 0 Then
                RemoveHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
                DataGridView1.DataSource = Nothing
                PropertyGrid1.SelectedObject = Nothing
                Const sql = "SELECT * FROM COUNT;"
                DataGridView1.DataSource = DbHelper.FillDataTable(sql)
                DataGridView1.Columns.Item(0).Visible = False
                AddHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
                
                ' 主动触发一次属性加载（选中第一行）
                If DataGridView1.Rows.Count > 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Rows(0).Cells(1)
                End If
            Else
                ComboBox1.SelectedIndex = 0
            End If
        End If
    End Sub

    ''' <summary>
    ''' 根据数据库中实际数据动态填充ComboBox，只显示有数据的组件类型
    ''' </summary>
    Private Sub PopulateComboBoxWithData()
        ' 保存当前选中项
        Dim currentSelection = If(ComboBox1.SelectedItem?.ToString(), "COUNT")

        ' 清空ComboBox
        ComboBox1.Items.Clear()

        ' 首先添加COUNT选项（统计信息）
        ComboBox1.Items.Add("COUNT")

        ' 检查每个组件类型是否有数据
        For Each componentType In ComponentTypes
            Try
                ' 确定视图名称（_UNION特殊处理）
                Dim viewName As String
                If componentType = "_UNION" Then
                    viewName = "X_UNION"
                Else
                    viewName = "X" & componentType
                End If

                ' 检查对应的X视图是否有数据
                Dim countSql = $"SELECT COUNT(*) FROM {viewName}"
                Dim countObj = DbHelper.ExecuteScalar(Of Object)(countSql)
                Dim countResult As Long = 0
                If countObj IsNot Nothing AndAlso Not IsDBNull(countObj) Then
                    Long.TryParse(countObj.ToString(), countResult)
                End If

                If countResult > 0 Then
                    ' 将下划线替换为连字符显示（与原始格式一致）
                    Dim displayName = componentType.Replace("_", "-")
                    ComboBox1.Items.Add(displayName)
                End If
            Catch ex As Exception
                ' 如果视图不存在或查询失败，记录错误但不中断
                ' 在调试模式下显示错误信息
                #If DEBUG Then
                System.Diagnostics.Debug.WriteLine($"检查视图 {componentType} 失败: {ex.Message}")
                #End If
                Continue For
            End Try
        Next

        ' 恢复之前的选择（如果还在列表中）
        If ComboBox1.Items.Contains(currentSelection) Then
            ComboBox1.SelectedItem = currentSelection
        ElseIf ComboBox1.Items.Count > 0 Then
            ComboBox1.SelectedIndex = 0
        End If
    End Sub

    Private Sub DataGridView1_CurrentCellChanged(sender As Object, e As EventArgs)
        Try
            If DataGridView1.CurrentRow IsNot Nothing AndAlso ComboBox1.SelectedItem IsNot Nothing Then
                Dim selectedType = ComboBox1.SelectedItem.ToString()

                If selectedType = "COUNT" And DataGridView1.ColumnCount >= 1 Then
                    ' COUNT视图显示pipeline信息
                    Dim FileName = DataGridView1.CurrentRow.Cells.Item(0).Value
                    If Not IsDBNull(FileName) Then
                        LoadPipelineSpecs(FileName)
                    End If
                ElseIf DataGridView1.ColumnCount >= 2 Then
                    ' 其他组件类型显示坐标点信息
                    Dim FileName = DataGridView1.CurrentRow.Cells.Item(0).Value
                    Dim UCI = DataGridView1.CurrentRow.Cells.Item(1).Value
                    If Not IsDBNull(FileName) And Not IsDBNull(UCI) Then
                        LoadPointData(FileName, UCI)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub LoadPipelineSpecs(fileName As Object)
        ' 创建动态属性包
        Dim propBag As New DynamicPropertyBag()

        ' 查询 PIPELINE 基本信息
        Dim sql As String = "SELECT PIPINGSPEC, INSULATIONSPEC, PAINTINGSPEC, TRACINGSPEC, REFERENCE " &
                           "FROM PIPELINE WHERE FILENAME=@FileName"
        Dim param As New SQLiteParameter("@FileName", fileName)
        Dim dt As DataTable = DbHelper.FillDataTable(sql, param)

        If dt.Rows.Count > 0 Then
            Dim row As DataRow = dt.Rows(0)

            ' 添加基本属性
            propBag.AddProperty("FILENAME", fileName.ToString(), "基本信息", "PCF文件名")
            propBag.AddProperty("REFERENCE", If(IsDBNull(row("REFERENCE")), "", row("REFERENCE").ToString()), "基本信息", "管线参考编号")
            propBag.AddProperty("PIPINGSPEC", If(IsDBNull(row("PIPINGSPEC")), "", row("PIPINGSPEC").ToString()), "基本信息", "管道规格等级")
            propBag.AddProperty("INSULATIONSPEC", If(IsDBNull(row("INSULATIONSPEC")), "", row("INSULATIONSPEC").ToString()), "基本信息", "保温规格等级")
            propBag.AddProperty("PAINTINGSPEC", If(IsDBNull(row("PAINTINGSPEC")), "", row("PAINTINGSPEC").ToString()), "基本信息", "油漆规格等级")
            propBag.AddProperty("TRACINGSPEC", If(IsDBNull(row("TRACINGSPEC")), "", row("TRACINGSPEC").ToString()), "基本信息", "伴热规格等级")

            ' 查询并添加所有 ATTRIBUTE（动态属性）
            Dim attrSql As String = "SELECT ATTRIBUTE_NAME, ATTRIBUTE_VALUE FROM COMPONENT_ATTRIBUTES " &
                                   "WHERE FILENAME=@FileName AND UCI LIKE 'PIPELINE:%' ORDER BY ATTRIBUTE_NAME"
            Dim attrTable As DataTable = DbHelper.FillDataTable(attrSql, param)

            For Each attrRow As DataRow In attrTable.Rows
                Dim attrName As String = If(IsDBNull(attrRow("ATTRIBUTE_NAME")), "", attrRow("ATTRIBUTE_NAME").ToString())
                Dim attrValue As String = If(IsDBNull(attrRow("ATTRIBUTE_VALUE")), "", attrRow("ATTRIBUTE_VALUE").ToString())
                If Not String.IsNullOrEmpty(attrName) Then
                    ' 使用 ATTRIBUTE_NAME 作为属性名，ATTRIBUTE_VALUE 作为属性值
                    propBag.AddProperty(attrName, attrValue, "属性", attrName)
                End If
            Next
        End If

        ' 使用 PropertyGrid 显示动态属性包
        PropertyGrid1.SelectedObject = propBag
    End Sub

    Private Sub LoadPointData(fileName As Object, uci As Object)
        ' 创建动态属性包显示坐标点
        Dim propBag As New DynamicPropertyBag()

        Dim params As SQLiteParameter() = {
            New SQLiteParameter("@UCI", uci),
            New SQLiteParameter("@FileName", fileName)
        }

        ' 查询 END_POINT（可能有多个）
        Dim endPointSql As String = "SELECT END_POINTX, END_POINTY, END_POINTZ, NPD, TYPE FROM END_POINT WHERE UCI=@UCI AND FILENAME=@FileName"
        Dim endPointDt As DataTable = DbHelper.FillDataTable(endPointSql, params)
        For i As Integer = 0 To endPointDt.Rows.Count - 1
            Dim row As DataRow = endPointDt.Rows(i)
            Dim pointName As String = If(i = 0, "端点", $"端点{i + 1}")
            Dim pointObj As New CoordinatePoint() With {
                .X = If(IsDBNull(row("END_POINTX")), "", row("END_POINTX").ToString()),
                .Y = If(IsDBNull(row("END_POINTY")), "", row("END_POINTY").ToString()),
                .Z = If(IsDBNull(row("END_POINTZ")), "", row("END_POINTZ").ToString()),
                .NPD = If(IsDBNull(row("NPD")), "", row("NPD").ToString()),
                .Type = If(IsDBNull(row("TYPE")), "", row("TYPE").ToString())
            }
            propBag.AddProperty(pointName, pointObj, "坐标点", "ENDPOINT坐标")
        Next

        ' 查询 CENTRE_POINT（可能有多个）
        Dim centreSql As String = "SELECT CENTRE_POINTX, CENTRE_POINTY, CENTRE_POINTZ, NPD, TYPE FROM CENTRE_POINT WHERE UCI=@UCI AND FILENAME=@FileName"
        Dim centreDt As DataTable = DbHelper.FillDataTable(centreSql, params)
        For i As Integer = 0 To centreDt.Rows.Count - 1
            Dim row As DataRow = centreDt.Rows(i)
            Dim pointName As String = If(i = 0, "中心点", $"中心点{i + 1}")
            Dim pointObj As New CoordinatePoint() With {
                .X = If(IsDBNull(row("CENTRE_POINTX")), "", row("CENTRE_POINTX").ToString()),
                .Y = If(IsDBNull(row("CENTRE_POINTY")), "", row("CENTRE_POINTY").ToString()),
                .Z = If(IsDBNull(row("CENTRE_POINTZ")), "", row("CENTRE_POINTZ").ToString()),
                .NPD = If(IsDBNull(row("NPD")), "", row("NPD").ToString()),
                .Type = If(IsDBNull(row("TYPE")), "", row("TYPE").ToString())
            }
            propBag.AddProperty(pointName, pointObj, "坐标点", "CENTREPOINT坐标")
        Next

        ' 查询 BRANCH1_POINT（可能有多个）
        Dim branchSql As String = "SELECT BRANCH1_POINTX, BRANCH1_POINTY, BRANCH1_POINTZ, NPD, TYPE FROM BRANCH1_POINT WHERE UCI=@UCI AND FILENAME=@FileName"
        Dim branchDt As DataTable = DbHelper.FillDataTable(branchSql, params)
        For i As Integer = 0 To branchDt.Rows.Count - 1
            Dim row As DataRow = branchDt.Rows(i)
            Dim pointName As String = If(i = 0, "分支点", $"分支点{i + 1}")
            Dim pointObj As New CoordinatePoint() With {
                .X = If(IsDBNull(row("BRANCH1_POINTX")), "", row("BRANCH1_POINTX").ToString()),
                .Y = If(IsDBNull(row("BRANCH1_POINTY")), "", row("BRANCH1_POINTY").ToString()),
                .Z = If(IsDBNull(row("BRANCH1_POINTZ")), "", row("BRANCH1_POINTZ").ToString()),
                .NPD = If(IsDBNull(row("NPD")), "", row("NPD").ToString()),
                .Type = If(IsDBNull(row("TYPE")), "", row("TYPE").ToString())
            }
            propBag.AddProperty(pointName, pointObj, "坐标点", "BRANCH1POINT坐标")
        Next

        ' 查询 JACKET_POINT（可能有多个）
        Dim jacketSql As String = "SELECT JACKET_POINTX, JACKET_POINTY, JACKET_POINTZ, NPD, TYPE FROM JACKET_POINT WHERE UCI=@UCI AND FILENAME=@FileName"
        Dim jacketDt As DataTable = DbHelper.FillDataTable(jacketSql, params)
        For i As Integer = 0 To jacketDt.Rows.Count - 1
            Dim row As DataRow = jacketDt.Rows(i)
            Dim pointName As String = If(i = 0, "夹套点", $"夹套点{i + 1}")
            Dim pointObj As New CoordinatePoint() With {
                .X = If(IsDBNull(row("JACKET_POINTX")), "", row("JACKET_POINTX").ToString()),
                .Y = If(IsDBNull(row("JACKET_POINTY")), "", row("JACKET_POINTY").ToString()),
                .Z = If(IsDBNull(row("JACKET_POINTZ")), "", row("JACKET_POINTZ").ToString()),
                .NPD = If(IsDBNull(row("NPD")), "", row("NPD").ToString()),
                .Type = If(IsDBNull(row("TYPE")), "", row("TYPE").ToString())
            }
            propBag.AddProperty(pointName, pointObj, "坐标点", "JACKETPOINT坐标")
        Next

        ' 查询 CO_ORDS（可能有多个）
        Dim coordsSql As String = "SELECT CO_ORDSX, CO_ORDSY, CO_ORDSZ, NPD, TYPE FROM CO_ORDS WHERE UCI=@UCI AND FILENAME=@FileName"
        Dim coordsDt As DataTable = DbHelper.FillDataTable(coordsSql, params)
        For i As Integer = 0 To coordsDt.Rows.Count - 1
            Dim row As DataRow = coordsDt.Rows(i)
            Dim pointName As String = If(i = 0, "坐标点", $"坐标点{i + 1}")
            Dim pointObj As New CoordinatePoint() With {
                .X = If(IsDBNull(row("CO_ORDSX")), "", row("CO_ORDSX").ToString()),
                .Y = If(IsDBNull(row("CO_ORDSY")), "", row("CO_ORDSY").ToString()),
                .Z = If(IsDBNull(row("CO_ORDSZ")), "", row("CO_ORDSZ").ToString()),
                .NPD = If(IsDBNull(row("NPD")), "", row("NPD").ToString()),
                .Type = If(IsDBNull(row("TYPE")), "", row("TYPE").ToString())
            }
            propBag.AddProperty(pointName, pointObj, "坐标点", "CO-ORDS坐标")
        Next

        ' 使用 PropertyGrid 显示动态属性包
        PropertyGrid1.SelectedObject = propBag
    End Sub

    Private Sub 导出ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 导出ToolStripMenuItem.Click
        Try
            '数据接下来输出到EXCEL表格
            Dim wk As IWorkbook '新建IWorkbook对象
            Dim localFilePath = "D:\\TEST.xlsx"
            ' 调用一个系统自带的保存文件对话框 写一个EXCEL
            Dim FilePathDialog As New FolderBrowserDialog() '新建winform自带保存文件对话框对象
            'saveFileDialog.Filter = "Excel Office2007及以上(*.xlsx)|*.xlsx" '过滤只能存储的对象
            Dim result As DialogResult = FilePathDialog.ShowDialog() '显示对话框
            If result = DialogResult.OK Then
                localFilePath = FilePathDialog.SelectedPath.ToString()
                '创建工作簿
                Dim destinationFolder As String = localFilePath
                wk = New XSSFWorkbook()
                Dim sql As String
                Dim oDataTable As DataTable
                Dim tb As ISheet
                Dim row As IRow
                
                ' 第一步：创建PIPELINE表格，包含所有pipeline信息和属性（每个属性作为单独列）
                ' 先获取所有可能的属性名
                Dim attrNamesSql As String = "SELECT DISTINCT ATTRIBUTE_NAME FROM COMPONENT_ATTRIBUTES WHERE UCI LIKE 'PIPELINE:%' ORDER BY ATTRIBUTE_NAME"
                Dim attrNamesTable As DataTable = DbHelper.FillDataTable(attrNamesSql)
                Dim attributeNames As New List(Of String)
                For Each dr As DataRow In attrNamesTable.Rows
                    attributeNames.Add(dr("ATTRIBUTE_NAME").ToString())
                Next
                
                ' 获取pipeline数据，包含属性信息
                Dim pipelineSql As String = "SELECT p.FILENAME, p.REFERENCE, p.PIPINGSPEC, p.INSULATIONSPEC, p.PAINTINGSPEC, p.TRACINGSPEC FROM PIPELINE p"
                Dim pipelineDataTable As DataTable = DbHelper.FillDataTable(pipelineSql)
                
                If pipelineDataTable.Columns.Count > 0 Then
                    tb = wk.CreateSheet("PIPELINE")
                    
                    ' 创建列标题
                    row = tb.CreateRow(0)
                    ' 基本字段（移除FILENAME列）
                    row.CreateCell(0).SetCellValue("REFERENCE")
                    row.CreateCell(1).SetCellValue("PIPINGSPEC")
                    row.CreateCell(2).SetCellValue("INSULATIONSPEC")
                    row.CreateCell(3).SetCellValue("PAINTINGSPEC")
                    row.CreateCell(4).SetCellValue("TRACINGSPEC")
                    ' 动态属性列
                    Dim attrStartCol As Integer = 5
                    For i = 0 To attributeNames.Count - 1
                        row.CreateCell(attrStartCol + i).SetCellValue(attributeNames(i))
                    Next
                    
                    ' 填充数据行
                    For i = 0 To pipelineDataTable.Rows.Count - 1
                        row = tb.CreateRow(i + 1)
                        Dim pipelineRow As DataRow = pipelineDataTable.Rows(i)
                        
                        ' 基本字段（移除FILENAME列）
                        row.CreateCell(0).SetCellValue(pipelineRow("REFERENCE").ToString())
                        row.CreateCell(1).SetCellValue(If(pipelineRow("PIPINGSPEC") Is DBNull.Value, "", pipelineRow("PIPINGSPEC").ToString()))
                        row.CreateCell(2).SetCellValue(If(pipelineRow("INSULATIONSPEC") Is DBNull.Value, "", pipelineRow("INSULATIONSPEC").ToString()))
                        row.CreateCell(3).SetCellValue(If(pipelineRow("PAINTINGSPEC") Is DBNull.Value, "", pipelineRow("PAINTINGSPEC").ToString()))
                        row.CreateCell(4).SetCellValue(If(pipelineRow("TRACINGSPEC") Is DBNull.Value, "", pipelineRow("TRACINGSPEC").ToString()))
                        
                        ' 获取该pipeline的属性并填充到对应列
                        Dim pipelineUci As String = "PIPELINE:" & pipelineRow("REFERENCE").ToString()
                        Dim attrSql As String = "SELECT ATTRIBUTE_VALUE FROM COMPONENT_ATTRIBUTES WHERE UCI = @UCI AND ATTRIBUTE_NAME = @ATTR_NAME"
                        
                        For j = 0 To attributeNames.Count - 1
                            Dim attrName As String = attributeNames(j)
                            Dim attrParams As SQLiteParameter() = {
                                New SQLiteParameter("@UCI", pipelineUci),
                                New SQLiteParameter("@ATTR_NAME", attrName)
                            }
                            Dim attrValueTable As DataTable = DbHelper.FillDataTable(attrSql, attrParams)
                            
                            If attrValueTable.Rows.Count > 0 Then
                                row.CreateCell(attrStartCol + j).SetCellValue(attrValueTable.Rows(0)("ATTRIBUTE_VALUE").ToString())
                            Else
                                row.CreateCell(attrStartCol + j).SetCellValue("") ' 空值
                            End If
                        Next
                    Next
                End If
                
                ' 第二步：创建各个组件表格（只创建有数据的工作簿）
                For Each Type In ComponentTypes
                    sql = $"SELECT * FROM {"X" & Type} ;"
                    oDataTable = DbHelper.FillDataTable(sql)
                    ' 检查是否有数据行（至少有一行数据）
                    If oDataTable.Rows.Count > 0 And oDataTable.Columns.Count > 0 And oDataTable.Columns.Contains("FILENAME") Then
                        tb = wk.CreateSheet(Type)
                        oDataTable.Columns.Remove("FILENAME")
                        For i = 0 To oDataTable.Rows.Count
                            row = tb.CreateRow(i)
                            For j = 0 To oDataTable.Columns.Count - 1
                                If i = 0 Then
                                    row.CreateCell(j).SetCellValue(oDataTable.Columns.Item(j).Caption)
                                Else
                                    row.CreateCell(j).SetCellValue(oDataTable.Rows(i - 1).Item(j).ToString())
                                End If
                            Next
                        Next
                    End If
                Next
                Dim validFileName As String = Regex.Replace(Path.GetFileNameWithoutExtension(Path2), "[\x00-\x1F\x7F<>?*:\""/\\|]", "", RegexOptions.None)
                ' 确保文件名不以点开始，不以空格或连字符结束
                validFileName = validFileName.TrimEnd(" .-".ToCharArray()).TrimStart(".".ToCharArray())
                '创建文件
                Using fs As New FileStream(Path.Combine(destinationFolder, validFileName & ".xlsx"), FileMode.Create, FileAccess.Write) '打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
                    wk.Write(fs) '文件IO 创建EXCEL
                    MessageBox.Show("成功输出文本到Excel ^_^！")
                    fs.Close()
                End Using
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ContextMenuStrip1_Opened(sender As Object, e As EventArgs) Handles ContextMenuStrip1.Opened
        If ConnectionStringBuilder Is Nothing Or Path2 Is Nothing Then
            导出ToolStripMenuItem.Enabled = False
        Else
            导出ToolStripMenuItem.Enabled = True
        End If
    End Sub
End Class
