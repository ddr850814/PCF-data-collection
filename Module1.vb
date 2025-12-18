Imports System.Data.SQLite

Module Module1
    Public ConnectionStringBuilder As SQLiteConnectionStringBuilder
    Public Ssupport As List(Of String()), SBolt As List(Of String()), SPipeline As List(Of String()), SMessage As List(Of String())
    Public Scommon As Dictionary(Of String, List(Of String())), SPipe As Dictionary(Of String, List(Of String())), Spoint As Dictionary(Of String, List(Of String()))
    Dim Value0 As String, Value1 As String, Value2 As String, Value3 As String, Value4 As String, Value5 As String, Value6 As String, Value7 As String
    'Public Function ExecuteNonQuery(sql As String, ParamArray parameters As SQLiteParameter()) As Integer
    '    Dim r As Integer
    '    Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
    '        con.Open()
    '        Using cmd As SQLiteCommand = con.CreateCommand()
    '            cmd.CommandText = sql
    '            cmd.Parameters.AddRange(parameters)
    '            r = cmd.ExecuteNonQuery()
    '        End Using
    '    End Using
    '    Return r
    'End Function
    'Public Function ExecuteScalar(sql As String, ParamArray parameters As SQLiteParameter()) As Object
    '    Dim r As Object
    '    Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
    '        con.Open()
    '        Using cmd As SQLiteCommand = con.CreateCommand()
    '            cmd.CommandText = sql
    '            cmd.Parameters.AddRange(parameters)
    '            r = cmd.ExecuteScalar()
    '        End Using
    '    End Using
    '    Return r
    'End Function
    'Public Function GetDataTable(sql As String) As DataTable
    '    Dim dt As New DataTable()
    '    Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
    '        con.Open()
    '        Using cmd As SQLiteCommand = con.CreateCommand()
    '            cmd.CommandText = sql
    '            Using dr As SQLiteDataReader = cmd.ExecuteReader()
    '                Dim schemaTable As DataTable = dr.GetSchemaTable()
    '                For Each row As DataRow In schemaTable.Rows
    '                    dt.Columns.Add(row.Item(0))
    '                Next
    '                While dr.Read
    '                    Dim r As DataRow = dt.NewRow()
    '                    For Each col As DataColumn In r.Table.Columns
    '                        r.Item(col.ColumnName) = dr.Item(col.ColumnName)
    '                    Next
    '                    dt.Rows.Add(r)
    '                End While
    '            End Using
    '        End Using
    '    End Using
    '    Return dt
    'End Function
    Function FillDataTable(sql As String) As DataTable
        Dim dt As New DataTable()
        Using adp As New SQLiteDataAdapter(sql, ConnectionStringBuilder.ConnectionString)
            adp.Fill(dt)
        End Using
        Return dt
    End Function
    Function FillDataTable(sql As String, starRecord As Integer, maxRecords As Integer) As DataTable
        Dim dt As New DataTable()
        Using adp As New SQLiteDataAdapter(sql, ConnectionStringBuilder.ConnectionString)
            adp.Fill(starRecord, maxRecords, dt)
        End Using
        Return dt
    End Function

    Public Sub CreateSQLiteTablePIPELINE()
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP TABLE IF EXISTS PIPELINE;CREATE TABLE IF NOT EXISTS PIPELINE (FILENAME INTEGER,REFERENCE TEXT, PIPINGSPEC TEXT, INSULATIONSPEC TEXT,PAINTINGSPEC TEXT,TRACINGSPEC TEXT);"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub CreateSQLiteTablePOINT(Name As String)
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP TABLE IF EXISTS {Name};CREATE TABLE IF NOT EXISTS {Name} (FILENAME INTEGER,UCI TEXT ,{Name}X REAL, {Name}Y REAL, {Name}Z REAL,NPD REAL,TYPE TEXT);"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub CreateSQLiteTableSupport()
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP TABLE IF EXISTS SUPPORT;CREATE TABLE IF NOT EXISTS SUPPORT (FILENAME INTEGER,ITEMCODE TEXT,NAME TEXT,UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,SUPPORTDIRECTION TEXT,SUPPORTTYPE TEXT);"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Sub CreateSQLiteTableBolt()
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP TABLE IF EXISTS BOLT;CREATE TABLE IF NOT EXISTS BOLT (FILENAME INTEGER,BOLTITEMCODE TEXT,BOLTDIA REAL,UCI TEXT,BOLTITEMDESCRIPTION TEXT,BOLTQUANTITY INTEGER,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER,BOLTLENGTH INTEGER);"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Sub CreateSQLiteTableMessage()
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP TABLE IF EXISTS MESSAGE;CREATE TABLE IF NOT EXISTS MESSAGE (FILENAME INTEGER,ITEMCODE TEXT,TEXT TEXT,UCI TEXT);"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Sub CreateSQLiteTablePipe(Name As String)
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP TABLE IF EXISTS {Name};CREATE TABLE IF NOT EXISTS {Name} (FILENAME INTEGER,ITEMCODE TEXT,CUTPIECELENGTH REAL,UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER,ANGLE INTEGER);"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub CreateSQLiteTable(Name As String)
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP TABLE IF EXISTS {Name};CREATE TABLE IF NOT EXISTS {Name} (FILENAME INTEGER,ITEMCODE TEXT,TAG TEXT,UCI TEXT,ITEMDESCRIPTION TEXT,SKEY TEXT,COMPONENTIDENTIFIER INTEGER,MASTERCOMPONENTIDENTIFIER INTEGER);"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub CreateSQLiteViewAll(Types As String())
        Dim part = ""
        For i = 0 To Types.Count - 1
            If i = 0 Then
                part &= $"SELECT FILENAME,COALESCE(ITEMCODE,'{Types(i)}') AS ITEMCODE,UCI,ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM {Types(i)}"
            ElseIf Types(i) = "BOLT" Then
                part &= $" UNION ALL SELECT FILENAME,COALESCE(BOLTITEMCODE,'BOLT') AS ITEMCODE,UCI,BOLTITEMDESCRIPTION,COMPONENTIDENTIFIER FROM BOLT"
            ElseIf Types(i) = "SUPPORT" Then
                part &= $" UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'SUPPORT') AS ITEMCODE,UCI,ITEMDESCRIPTION,'' AS COMPONENTIDENTIFIER FROM SUPPORT"
            ElseIf Types(i) = "MESSAGE" Then
                part &= $" UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'MESSAGE') AS ITEMCODE,UCI,'' AS ITEMDESCRIPTION,'' AS COMPONENTIDENTIFIER FROM MESSAGE"
            Else
                part &= $" UNION ALL SELECT FILENAME,COALESCE(ITEMCODE,'{Types(i)}') AS ITEMCODE,UCI,ITEMDESCRIPTION,COMPONENTIDENTIFIER FROM {Types(i)}"
            End If
        Next
        Using con As New SQLiteConnection(ConnectionStringBuilder.ConnectionString)
            con.Open()
            Dim sql As String = $"DROP VIEW IF EXISTS TOTAL;CREATE VIEW IF NOT EXISTS TOTAL AS
              {part};"
            Using command As New SQLiteCommand(sql, con)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub InsertSQLiteTable(Name As String, value As List(Of String), filename As Integer)
        value0 = ValueS(value(0))
        value1 = ValueS(value(1))
        value2 = ValueS(value(2))
        value3 = ValueS(value(3))
        value4 = ValueS(value(4))
        value5 = ValueI(value(5))
        Value6 = ValueI(value(6))
        Scommon.Item(Name).Add(New String() {filename, Value0, Value1, Value2, Value3, Value4, Value5, Value6})
    End Sub

    Public Sub InsertSQLiteTablePipe(Name As String, value As List(Of String), filename As Integer)
        value0 = ValueS(value(0))
        value1 = ValueD(value(1))
        value2 = ValueS(value(2))
        value3 = ValueS(value(3))
        value4 = ValueS(value(4))
        value5 = ValueI(value(5))
        value6 = ValueI(value(6))
        Value7 = ValueI(value(7))
        SPipe.Item(Name).Add(New String() {filename, Value0, Value1, Value2, Value3, Value4, Value5, Value6, Value7})
    End Sub

    Public Sub InsertSQLiteTablePIPELINE(value As List(Of String), filename As Integer)
        value0 = ValueS(value(0))
        value1 = ValueS(value(1))
        value2 = ValueS(value(2))
        value3 = ValueS(value(3))
        Value4 = ValueS(value(4))
        SPipeline.Add(New String() {filename, Value0, Value1, Value2, Value3, Value4})
    End Sub

    Public Sub InsertSQLiteTablePOINT(Name As String, value As IEnumerable(Of String), UCI As String, filename As Integer)
        Dim a0 As Double, a1 As Double, a2 As Double
        Select Case value.Count
            Case 5
                If Double.TryParse(value(0), a0) And Double.TryParse(value(1), a1) And Double.TryParse(value(2), a2) Then
                    Value0 = a0.ToString
                    Value1 = a1.ToString
                    Value2 = a2.ToString
                    Value3 = ValueD(value(3))
                    Value4 = ValueS(value(4))
                Else
                    Exit Sub
                End If
            Case 4
                If Double.TryParse(value(0), a0) And Double.TryParse(value(1), a1) And Double.TryParse(value(2), a2) Then
                    Value0 = a0.ToString
                    Value1 = a1.ToString
                    Value2 = a2.ToString
                    Value3 = ValueD(value(3))
                    Value4 = Nothing
                Else
                    Exit Sub
                End If
            Case 3
                If Double.TryParse(value(0), a0) And Double.TryParse(value(1), a1) And Double.TryParse(value(2), a2) Then
                    Value0 = a0.ToString
                    Value1 = a1.ToString
                    Value2 = a2.ToString
                    Value3 = Nothing
                    Value4 = Nothing
                Else
                    Exit Sub
                End If
        End Select
        Spoint.Item(Name).Add(New String() {filename, UCI, Value0, Value1, Value2, Value3, Value4})
    End Sub

    Public Sub InsertSQLiteTableSupport(value As List(Of String), filename As Integer)
        value0 = ValueS(value(0))
        value1 = ValueS(value(1))
        value2 = ValueS(value(2))
        value3 = ValueS(value(3))
        value4 = ValueS(value(4))
        value5 = ValueS(value(5))
        Value6 = ValueS(value(6))
        Ssupport.Add(New String() {filename, Value0, Value1, Value2, Value3, Value4, Value5, Value6})
    End Sub
    Public Sub InsertSQLiteTableBolt(value As List(Of String), filename As Integer)
        value0 = ValueS(value(0))
        value1 = ValueD(value(1))
        value2 = ValueS(value(2))
        value3 = ValueS(value(3))
        value4 = ValueI(value(4))
        value5 = ValueI(value(5))
        value6 = ValueI(value(6))
        Value7 = ValueI(value(7))
        SBolt.Add(New String() {filename, Value0, Value1, Value2, Value3, Value4, Value5, Value6, Value7})
    End Sub
    Public Sub InsertSQLiteTableMessage(value As List(Of String), filename As Integer)
        Value0 = ValueS(value(0))
        Value1 = ValueS(value(1))
        Value2 = ValueS(value(2))
        SMessage.Add(New String() {filename, Value0, Value1, Value2})
    End Sub
    Private Function ValueS(value As String) As String
        If value <> "" Then
            Return value
        Else
            Return Nothing
        End If
    End Function

    Private Function ValueI(value As String) As String
        Dim i0 As Integer
        If Integer.TryParse(value, i0) Then
            Return i0.ToString
        Else
            Return Nothing
        End If
    End Function

    Private Function ValueD(value As String) As String
        Dim i0 As Double
        If Double.TryParse(value, i0) Then
            Return i0.ToString
        Else
            Return Nothing
        End If
    End Function
End Module
