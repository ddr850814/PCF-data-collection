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
    Dim TEESTUB As Integer
    Dim MESSAGE As Integer

    Private Sub WritePreviously(report As Boolean, filename As Integer, ByRef rowP As List(Of String), ByRef row As List(Of String), Title As String, ByRef ENDPOINT As List(Of String))
        If rowP.Any(Function(n) n <> "") Then
            If Title = "PIPELINE" Then
                InsertSQLiteTablePIPELINE(rowP, filename)
            End If
            rowP = New List(Of String) From {"", "", "", "", ""}
        ElseIf row.Any(Function(n) n <> "") Then
            If report = True Then
                Select Case Title
                    Case "SUPPORT"
                        InsertSQLiteTableSupport(row, filename)
                    Case "MESSAGE"
                        row(2) = "MS" & MESSAGE
                        MESSAGE += 1
                        InsertSQLiteTableMessage(row, filename)
                    Case "BOLT"
                        InsertSQLiteTableBolt(row, filename)
                    Case "PIPE", "PIPE_FIXED", "BEND"
                        InsertSQLiteTablePipe(Title, row, filename)
                    Case "TEE_STUB"
                        row(2) = "TS" & TEESTUB
                        TEESTUB += 1
                        InsertSQLiteTable(Title, row, filename)
                    Case Else
                        InsertSQLiteTable(Title, row, filename)
                End Select
                For Each InputArray In From InputS In ENDPOINT Select Regex.Split(InputS, "\s+")
                    If InputArray.Count() >= 4 And InputArray.Count() <= 6 Then
                        InsertSQLiteTablePOINT(InputArray(0).Replace("-", "_"), InputArray.Skip(1), row(2), filename)
                    End If
                Next
            End If
            row = New List(Of String) From {"", "", "", "", "", "", "", ""}
        End If
        ENDPOINT.Clear()
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
        TEESTUB = 0
        MESSAGE = 0
        Invoke(Sub()
                   Enabled = False
                   Button1.Text = count & "/" & divisibleByThree.Count
               End Sub)
        ConnectionStringBuilder = New SQLiteConnectionStringBuilder() With {.DataSource = Path1 & "\pcf.db", .ForeignKeys = True}
        CreateSQLiteTablePIPELINE()
        SPipeline = New List(Of String())()
        Dim Types = New String() _
                {"INSTRUMENT", "SUPPORT", "FLANGE", "GASKET", "BOLT", "PIPE", "PIPE_FIXED", "WELD", "TEE_STUB", "TEE",
                 "VALVE", "VALVE_ANGLE", "FLANGE_BLIND", "OLET", "ELBOW", "REDUCER_CONCENTRIC", "REDUCER_ECCENTRIC",
                 "INSTRUMENT_ANGLE", "CAP", "MISC_COMPONENT", "FILTER", "VALVE_3WAY", "VALVE_4WAY", "VALVE_MULTIWAY", "COUPLING", "CROSS",
                 "LAPJOINT_STUBEND", "REINFORCEMENT_PAD", "BEND", "_UNION", "CLAMP", "MESSAGE"}
        Dim Points = New String() {"END_POINT", "CENTRE_POINT", "JACKET_POINT", "BRANCH1_POINT", "CO_ORDS"}
        Spoint = New Dictionary(Of String, List(Of String()))()
        For Each Point In Points
            CreateSQLiteTablePOINT(Point)
            Spoint.Add(Point, New List(Of String()))
        Next
        Ssupport = New List(Of String())()
        SBolt = New List(Of String())()
        SMessage = New List(Of String())()
        SPipe = New Dictionary(Of String, List(Of String()))()
        Scommon = New Dictionary(Of String, List(Of String()))()
        For Each Type In Types
            Select Case Type
                Case "SUPPORT"
                    CreateSQLiteTableSupport()
                Case "MESSAGE"
                    CreateSQLiteTableMessage()
                Case "BOLT"
                    CreateSQLiteTableBolt()
                Case "PIPE", "PIPE_FIXED", "BEND"
                    CreateSQLiteTablePipe(Type)
                    SPipe.Add(Type, New List(Of String()))
                Case Else
                    CreateSQLiteTable(Type)
                    Scommon.Add(Type, New List(Of String()))
            End Select
        Next
        For i = 0 To divisibleByThree.Count() - 1
            Dim report = True
            Dim Title As String = Nothing
            Dim rowP As New List(Of String) From {"", "", "", "", ""}
            Dim row As New List(Of String) From {"", "", "", "", "", "", "", ""}
            Dim ENDPOINT As New List(Of String)
            Using reader As New StreamReader(divisibleByThree(i), Encoding.GetEncoding("GBK"))
                Dim line As String = reader.ReadLine()
                While line IsNot Nothing
                    If line.StartsWith("PIPELINE-REFERENCE") Then
                        rowP(0) = line.Replace("PIPELINE-REFERENCE", "").Trim
                        Title = "PIPELINE"
                    ElseIf Types.Contains(line.Split(" "c)(0).Replace("-", "_")) Then
                        WritePreviously(report, i, rowP, row, Title, ENDPOINT)
                        Title = line.Split(" "c)(0).Replace("-", "_")
                        report = True
                    ElseIf Not line.StartsWith(" ") Then
                        WritePreviously(report, i, rowP, row, Title, ENDPOINT)
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
                            End If
                        End If
                    End If
                    line = reader.ReadLine()
                End While
                WritePreviously(report, i, rowP, row, Title, ENDPOINT)
            End Using
            count += 1
            Invoke(Sub()
                       Button1.Text = count & "/" & divisibleByThree.Count
                   End Sub)
        Next
        Invoke(Sub()
                   Button1.Text = "插入数据表"
               End Sub)
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Using transaction As SQLiteTransaction = con.BeginTransaction()
                Dim command As New SQLiteCommand(con)
                command.CommandText = "INSERT INTO PIPELINE (FILENAME,REFERENCE,PIPINGSPEC,INSULATIONSPEC,PAINTINGSPEC,TRACINGSPEC) VALUES (@Column1, @Column2,@Column3,@Column4,@Column5,@Column6)"
                For Each item In SPipeline
                    For i = 0 To 5
                        command.Parameters.AddWithValue($"@Column{i + 1}", item(i))
                    Next
                    command.ExecuteNonQuery()
                    command.Parameters.Clear()  ' Clear parameters for the next iteration to avoid reuse issues
                Next
                transaction.Commit()
            End Using
        End Using
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Using transaction As SQLiteTransaction = con.BeginTransaction()
                Dim command As New SQLiteCommand(con)
                command.CommandText = "INSERT INTO SUPPORT (FILENAME,ITEMCODE,NAME,UCI,ITEMDESCRIPTION,SKEY,SUPPORTDIRECTION,SUPPORTTYPE) VALUES (@Column1, @Column2,@Column3,@Column4,@Column5,@Column6,@Column7,@Column8)"
                For Each item In Ssupport
                    For i = 0 To 7
                        command.Parameters.AddWithValue($"@Column{i + 1}", item(i))
                    Next
                    command.ExecuteNonQuery()
                    command.Parameters.Clear()  ' Clear parameters for the next iteration to avoid reuse issues
                Next
                transaction.Commit()
            End Using
        End Using
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Using transaction As SQLiteTransaction = con.BeginTransaction()
                Dim command As New SQLiteCommand(con)
                For Each item In Spoint
                    command.CommandText = $"INSERT INTO {item.Key} (FILENAME,UCI,{item.Key}X,{item.Key}Y,{item.Key}Z,NPD,TYPE) VALUES (@Column1, @Column2,@Column3,@Column4,@Column5,@Column6,@Column7)"
                    For Each oPoint In item.Value
                        For i = 0 To 6
                            command.Parameters.AddWithValue($"@Column{i + 1}", oPoint(i))
                        Next
                        command.ExecuteNonQuery()
                        command.Parameters.Clear()  ' Clear parameters for the next iteration to avoid reuse issues
                    Next
                Next
                transaction.Commit()
            End Using
        End Using
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Using transaction As SQLiteTransaction = con.BeginTransaction()
                Dim command As New SQLiteCommand(con)
                For Each item In Scommon
                    command.CommandText = $"INSERT INTO {item.Key} (FILENAME,ITEMCODE,TAG,UCI,ITEMDESCRIPTION,SKEY,COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER) VALUES (@Column1, @Column2,@Column3,@Column4,@Column5,@Column6,@Column7,@Column8)"
                    For Each oComponent In item.Value
                        For i = 0 To 7
                            command.Parameters.AddWithValue($"@Column{i + 1}", oComponent(i))
                        Next
                        command.ExecuteNonQuery()
                        command.Parameters.Clear()  ' Clear parameters for the next iteration to avoid reuse issues
                    Next
                Next
                transaction.Commit()
            End Using
        End Using
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Using transaction As SQLiteTransaction = con.BeginTransaction()
                Dim command As New SQLiteCommand(con)
                For Each item In SPipe
                    command.CommandText = $"INSERT INTO {item.Key} (FILENAME,ITEMCODE,CUTPIECELENGTH,UCI,ITEMDESCRIPTION,SKEY,COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER,ANGLE) VALUES (@Column1, @Column2,@Column3,@Column4,@Column5,@Column6,@Column7,@Column8,@Column9)"
                    For Each oPipe In item.Value
                        For i = 0 To 8
                            command.Parameters.AddWithValue($"@Column{i + 1}", oPipe(i))
                        Next
                        command.ExecuteNonQuery()
                        command.Parameters.Clear()  ' Clear parameters for the next iteration to avoid reuse issues
                    Next
                Next
                transaction.Commit()
            End Using
        End Using
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Using transaction As SQLiteTransaction = con.BeginTransaction()
                Dim command As New SQLiteCommand(con)
                command.CommandText = "INSERT INTO BOLT (FILENAME,BOLTITEMCODE,BOLTDIA,UCI,BOLTITEMDESCRIPTION,BOLTQUANTITY,COMPONENTIDENTIFIER,MASTERCOMPONENTIDENTIFIER,BOLTLENGTH) VALUES (@Column1, @Column2,@Column3,@Column4,@Column5,@Column6,@Column7,@Column8,@Column9)"
                For Each item In SBolt
                    For i = 0 To 8
                        command.Parameters.AddWithValue($"@Column{i + 1}", item(i))
                    Next
                    command.ExecuteNonQuery()
                    command.Parameters.Clear()  ' Clear parameters for the next iteration to avoid reuse issues
                Next
                transaction.Commit()
            End Using
        End Using
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Using transaction As SQLiteTransaction = con.BeginTransaction()
                Dim command As New SQLiteCommand(con)
                command.CommandText = "INSERT INTO MESSAGE (FILENAME,ITEMCODE,TEXT,UCI) VALUES (@Column1, @Column2,@Column3,@Column4)"
                For Each item In SMessage
                    For i = 0 To 3
                        command.Parameters.AddWithValue($"@Column{i + 1}", item(i))
                    Next
                    command.ExecuteNonQuery()
                    command.Parameters.Clear()  ' Clear parameters for the next iteration to avoid reuse issues
                Next
                transaction.Commit()
            End Using
        End Using
        Invoke(Sub()
                   Button1.Text = "插入视图"
               End Sub)
        CreateSQLiteViewAll(Types)
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP VIEW IF EXISTS COUNT;CREATE VIEW IF NOT EXISTS COUNT AS SELECT PIPELINE.FILENAME, REFERENCE, TOTAL.ITEMCODE, ITEMDESCRIPTION, COUNT( * ) As count 
        FROM
        (SELECT * FROM TOTAL GROUP BY UCI,FILENAME) AS TOTAL 
        Join PIPELINE ON PIPELINE.FILENAME = TOTAL.FILENAME
        Group BY REFERENCE, TOTAL.ITEMCODE, ITEMDESCRIPTION 
        ORDER BY REFERENCE;"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
            For Each Type In Types
                Select Case Type
                    Case "PIPE"
                        sql = $"DROP VIEW IF EXISTS {"X" & Type};CREATE VIEW IF NOT EXISTS {"X" & Type} AS SELECT PIPELINE.FILENAME,UCI, REFERENCE, ITEMCODE, ITEMDESCRIPTION, CUTPIECELENGTH
            FROM 
            	( SELECT * FROM {Type} GROUP BY UCI,FILENAME) AS {Type} 
                JOIN PIPELINE ON PIPELINE.FILENAME = {Type}.FILENAME
            ORDER BY REFERENCE;"
                    Case "BEND"
                        sql = $"DROP VIEW IF EXISTS XBEND;CREATE VIEW IF NOT EXISTS XBEND AS SELECT PIPELINE.FILENAME,UCI, REFERENCE,ITEMCODE, ITEMDESCRIPTION,CUTPIECELENGTH,ANGLE
            FROM 
            	( SELECT * FROM BEND GROUP BY UCI,FILENAME) AS BEND 
                JOIN PIPELINE ON PIPELINE.FILENAME = BEND.FILENAME
            ORDER BY REFERENCE;"
                    Case "INSTRUMENT", "INSTRUMENT_ANGLE"
                        sql = $"DROP VIEW IF EXISTS {"X" & Type};CREATE VIEW IF NOT EXISTS {"X" & Type} AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, TAG
            FROM 
            	( SELECT * FROM {Type} GROUP BY UCI,FILENAME) AS {Type} 
                JOIN PIPELINE ON PIPELINE.FILENAME = {Type}.FILENAME
            ORDER BY REFERENCE;"
                    Case "SUPPORT"
                        sql = "DROP VIEW IF EXISTS XSUPPORT;CREATE VIEW IF NOT EXISTS XSUPPORT AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,NAME, ITEMDESCRIPTION,SUPPORTDIRECTION,SUPPORTTYPE
            FROM 
            	( SELECT * FROM SUPPORT GROUP BY UCI,FILENAME) AS SUPPORT 
                JOIN PIPELINE ON PIPELINE.FILENAME = SUPPORT.FILENAME
            ORDER BY REFERENCE;"
                    Case "GASKET"
                        sql = "DROP VIEW IF EXISTS XGASKET;CREATE VIEW IF NOT EXISTS XGASKET AS SELECT PIPELINE.FILENAME,GASKET.UCI,REFERENCE,GASKET.ITEMCODE, GASKET.ITEMDESCRIPTION,TOTAL.ITEMCODE,TOTAL.ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM GASKET GROUP BY UCI,FILENAME) AS GASKET 
                JOIN PIPELINE ON PIPELINE.FILENAME = GASKET.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = GASKET.FILENAME AND TOTAL.COMPONENTIDENTIFIER=GASKET.MASTERCOMPONENTIDENTIFIER
            ORDER BY REFERENCE;"
                    Case "BOLT"
                        sql = "DROP VIEW IF EXISTS XBOLT;CREATE VIEW IF NOT EXISTS XBOLT AS SELECT PIPELINE.FILENAME,BOLT.UCI,REFERENCE,BOLTITEMCODE,BOLTITEMDESCRIPTION,BOLTDIA, BOLTQUANTITY, BOLTLENGTH, ITEMCODE,ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM BOLT GROUP BY UCI,FILENAME) AS BOLT 
                JOIN PIPELINE ON PIPELINE.FILENAME = BOLT.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = BOLT.FILENAME AND TOTAL.COMPONENTIDENTIFIER=BOLT.MASTERCOMPONENTIDENTIFIER
            ORDER BY REFERENCE;"
                    Case "WELD"
                        sql = "DROP VIEW IF EXISTS XWELD;CREATE VIEW IF NOT EXISTS XWELD AS SELECT PIPELINE.FILENAME,WELD.UCI,REFERENCE,TOTAL.ITEMCODE,TOTAL.ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM WELD GROUP BY UCI,FILENAME) AS WELD 
                JOIN PIPELINE ON PIPELINE.FILENAME = WELD.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = WELD.FILENAME AND TOTAL.COMPONENTIDENTIFIER=WELD.MASTERCOMPONENTIDENTIFIER
            ORDER BY REFERENCE;"
                    Case "TEE_STUB"
                        sql = $"DROP VIEW IF EXISTS XTEE_STUB;CREATE VIEW IF NOT EXISTS XTEE_STUB AS SELECT PIPELINE.FILENAME,UCI,REFERENCE
            FROM 
            	( SELECT * FROM TEE_STUB GROUP BY UCI,FILENAME) AS TEE_STUB 
                JOIN PIPELINE ON PIPELINE.FILENAME = TEE_STUB.FILENAME
            ORDER BY REFERENCE;"
                    Case "_UNION"
                        sql = "DROP VIEW IF EXISTS X_UNION;CREATE VIEW IF NOT EXISTS X_UNION AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM _UNION GROUP BY UCI,FILENAME) AS _UNION 
                JOIN PIPELINE ON PIPELINE.FILENAME = _UNION.FILENAME
            ORDER BY REFERENCE;"
                    Case "MESSAGE"
                        sql = $"DROP VIEW IF EXISTS XMESSAGE;CREATE VIEW IF NOT EXISTS XMESSAGE AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,TEXT
            FROM 
            	( SELECT * FROM MESSAGE GROUP BY UCI,FILENAME) AS MESSAGE 
                JOIN PIPELINE ON PIPELINE.FILENAME = MESSAGE.FILENAME
            ORDER BY REFERENCE;"
                    Case Else
                        sql = $"DROP VIEW IF EXISTS {"X" & Type};CREATE VIEW IF NOT EXISTS {"X" & Type} AS SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, ITEMDESCRIPTION
            FROM 
            	( SELECT * FROM {Type} GROUP BY UCI,FILENAME) AS {Type} 
                JOIN PIPELINE ON PIPELINE.FILENAME = {Type}.FILENAME
            ORDER BY REFERENCE;"
                End Select
                Using command As New SQLiteCommand(sql, con)
                    command.ExecuteNonQuery()
                End Using
            Next
        End Using
        Invoke(Sub()
                   Enabled = True
                   SelectCount()
               End Sub)
    End Sub
    Private Sub BgWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BgWorker.RunWorkerCompleted
        Invoke(Sub()
                   Button1.Text = "写入数据库"
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
                SelectCount()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            RemoveHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
            If ConnectionStringBuilder IsNot Nothing Then
                DataGridView1.DataSource = Nothing
                DataGridView2.DataSource = Nothing
                Dim Type = ComboBox1.SelectedItem.ToString.Replace("-", "_")
                Dim sql As String
                Select Case ComboBox1.SelectedIndex
                    Case 0
                        sql = "SELECT * FROM COUNT;"
                        DataGridView1.DataSource = FillDataTable(sql)
                        DataGridView1.Columns.Item(0).Visible = False
                    Case 30
                        sql = $"SELECT * FROM {"X_" & Type} ;"
                        DataGridView1.DataSource = FillDataTable(sql)
                        DataGridView1.Columns.Item(0).Visible = False
                        DataGridView1.Columns.Item(1).Visible = False
                    Case Else
                        sql = $"SELECT * FROM {"X" & Type} ;"
                        DataGridView1.DataSource = FillDataTable(sql)
                        DataGridView1.Columns.Item(0).Visible = False
                        DataGridView1.Columns.Item(1).Visible = False
                End Select
                AddHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub SelectCount()
        If ComboBox1.SelectedIndex = 0 Then
            RemoveHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
            DataGridView1.DataSource = Nothing
            DataGridView2.DataSource = Nothing
            Const sql = "SELECT * FROM COUNT;"
            DataGridView1.DataSource = FillDataTable(sql)
            DataGridView1.Columns.Item(0).Visible = False
            AddHandler DataGridView1.CurrentCellChanged, AddressOf DataGridView1_CurrentCellChanged
        Else
            ComboBox1.SelectedIndex = 0
        End If
    End Sub

    Private Sub DataGridView1_CurrentCellChanged(sender As Object, e As EventArgs)
        Try
            Dim sql As String
            If DataGridView1.CurrentRow IsNot Nothing Then
                If ComboBox1.SelectedIndex = 0 And DataGridView1.ColumnCount >= 1 Then
                    Dim FileName = DataGridView1.CurrentRow.Cells.Item(0).Value
                    If Not IsDBNull(FileName) Then
                        sql = $"SELECT  PIPINGSPEC,INSULATIONSPEC,PAINTINGSPEC,TRACINGSPEC
            FROM 
            	PIPELINE
                WHERE FILENAME={FileName}"
                        DataGridView2.DataSource = FillDataTable(sql)
                    End If
                ElseIf DataGridView1.ColumnCount >= 2 Then
                    Dim FileName = DataGridView1.CurrentRow.Cells.Item(0).Value
                    Dim UCI = DataGridView1.CurrentRow.Cells.Item(1).Value
                    If Not IsDBNull(FileName) And Not IsDBNull(UCI) Then
                        sql = $"SELECT  'ENDPOINT' AS POINTTYPE,END_POINTX AS POINTX,END_POINTY AS POINTY,END_POINTZ AS POINTZ,NPD,TYPE
            FROM 
            	END_POINT
                WHERE UCI='{UCI}' AND FILENAME={FileName}
              UNION ALL
             SELECT  'CENTREPOINT' AS POINTTYPE,CENTRE_POINTX AS POINTX,CENTRE_POINTY AS POINTY,CENTRE_POINTZ AS POINTZ,NPD,TYPE
            FROM 
            	CENTRE_POINT
                WHERE UCI='{UCI}'  AND FILENAME={FileName}              
              UNION ALL
             SELECT  'BRANCH1EPOINT' AS POINTTYPE,BRANCH1_POINTX AS POINTX,BRANCH1_POINTY AS POINTY,BRANCH1_POINTZ AS POINTZ,NPD,TYPE
            FROM 
            	BRANCH1_POINT
                WHERE UCI='{UCI}'  AND FILENAME={FileName}
              UNION ALL
             SELECT  'JACKETPOINT' AS POINTTYPE,JACKET_POINTX AS POINTX,JACKET_POINTY AS POINTY,JACKET_POINTZ AS POINTZ,NPD,TYPE
            FROM 
            	JACKET_POINT
                WHERE UCI='{UCI}' AND FILENAME={FileName}
              UNION ALL
             SELECT 'CO-ORDS' AS POINTTYPE,CO_ORDSX AS POINTX,CO_ORDSY AS POINTY,CO_ORDSZ AS POINTZ,NPD,TYPE
            FROM
                CO_ORDS
                WHERE UCI ='{UCI}' AND FILENAME={FileName};"
                        DataGridView2.DataSource = FillDataTable(sql)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
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
                Dim Types = New String() _
                        {"INSTRUMENT", "SUPPORT", "FLANGE", "GASKET", "BOLT", "PIPE", "PIPE_FIXED", "WELD", "TEE_STUB", "TEE",
                         "VALVE", "VALVE_ANGLE", "FLANGE_BLIND", "OLET", "ELBOW", "REDUCER_CONCENTRIC", "REDUCER_ECCENTRIC",
                         "INSTRUMENT_ANGLE", "CAP", "MISC_COMPONENT", "FILTER", "VALVE_3WAY", "VALVE_4WAY", "VALVE_MULTIWAY", "COUPLING", "CROSS",
                         "LAPJOINT_STUBEND", "REINFORCEMENT_PAD", "BEND", "_UNION", "CLAMP", "MESSAGE"}
                For Each Type In Types
                    sql = $"SELECT * FROM {"X" & Type} ;"
                    oDataTable = FillDataTable(sql)
                    If oDataTable.Columns.Count > 0 And oDataTable.Columns.Contains("FILENAME") Then
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
