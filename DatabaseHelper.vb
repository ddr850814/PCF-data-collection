Imports System.Data.SQLite
Imports System.Data

Public Class DatabaseHelper
    Implements IDisposable

    Private _connection As SQLiteConnection
    Private _connectionString As String
    Private _isDisposed As Boolean = False

    Public Sub New(connectionString As String)
        _connectionString = connectionString
    End Sub

    Public ReadOnly Property Connection As SQLiteConnection
        Get
            If _connection Is Nothing Then
                _connection = New SQLiteConnection(_connectionString)
            End If
            If _connection.State <> ConnectionState.Open Then
                _connection.Open()
            End If
            Return _connection
        End Get
    End Property

    Public Sub EnsureConnection()
        Dim unused = Connection
    End Sub

    ''' <summary>
    ''' 执行非查询 SQL，使用显式类型的参数
    ''' </summary>
    Public Function ExecuteNonQuery(sql As String, ParamArray parameters As SQLiteParameter()) As Integer
        Using cmd As New SQLiteCommand(sql, Connection)
            If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                cmd.Parameters.AddRange(parameters)
            End If
            Return cmd.ExecuteNonQuery()
        End Using
    End Function

    ''' <summary>
    ''' 执行非查询 SQL，使用简单参数值（内部使用显式 String 类型）
    ''' </summary>
    Public Function ExecuteNonQueryWithValues(sql As String, ParamArray values As Object()) As Integer
        Using cmd As New SQLiteCommand(sql, Connection)
            If values IsNot Nothing AndAlso values.Length > 0 Then
                For i = 0 To values.Length - 1
                    ' 使用显式 String 类型，避免 AddWithValue 的类型推断
                    Dim param As New SQLiteParameter($"@p{i}", DbType.String)
                    param.Value = If(values(i) Is Nothing, DirectCast(DBNull.Value, Object), values(i).ToString())
                    cmd.Parameters.Add(param)
                Next
            End If
            Return cmd.ExecuteNonQuery()
        End Using
    End Function

    ''' <summary>
    ''' 执行标量查询，使用显式类型的参数
    ''' </summary>
    Public Function ExecuteScalar(Of T)(sql As String, ParamArray parameters As SQLiteParameter()) As T
        Using cmd As New SQLiteCommand(sql, Connection)
            If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                cmd.Parameters.AddRange(parameters)
            End If
            Dim result = cmd.ExecuteScalar()
            If result Is DBNull.Value OrElse result Is Nothing Then
                Return Nothing
            End If
            Return DirectCast(result, T)
        End Using
    End Function

    ''' <summary>
    ''' 填充 DataTable，使用显式类型的参数
    ''' </summary>
    Public Function FillDataTable(sql As String, ParamArray parameters As SQLiteParameter()) As DataTable
        Dim dt As New DataTable()
        Using cmd As New SQLiteCommand(sql, Connection)
            If parameters IsNot Nothing AndAlso parameters.Length > 0 Then
                cmd.Parameters.AddRange(parameters)
            End If
            Using reader = cmd.ExecuteReader()
                dt.Load(reader)
            End Using
        End Using
        Return dt
    End Function

    ''' <summary>
    ''' 填充 DataTable，使用简单参数值（内部使用显式 String 类型）
    ''' </summary>
    Public Function FillDataTableWithValues(sql As String, ParamArray values As Object()) As DataTable
        Dim dt As New DataTable()
        Using cmd As New SQLiteCommand(sql, Connection)
            If values IsNot Nothing AndAlso values.Length > 0 Then
                For i = 0 To values.Length - 1
                    Dim param As New SQLiteParameter($"@p{i}", DbType.String)
                    param.Value = If(values(i) Is Nothing, DirectCast(DBNull.Value, Object), values(i).ToString())
                    cmd.Parameters.Add(param)
                Next
            End If
            Using reader = cmd.ExecuteReader()
                dt.Load(reader)
            End Using
        End Using
        Return dt
    End Function

    Public Function BeginTransaction() As SQLiteTransaction
        Return Connection.BeginTransaction()
    End Function

    Public Sub CreateTable(sql As String)
        ExecuteNonQuery(sql)
    End Sub

    Public Sub DropTableIfExists(tableName As String)
        ExecuteNonQuery($"DROP TABLE IF EXISTS {tableName}")
    End Sub

    Public Sub DropViewIfExists(viewName As String)
        ExecuteNonQuery($"DROP VIEW IF EXISTS {viewName}")
    End Sub

    ''' <summary>
    ''' 创建参数（显式指定类型）
    ''' </summary>
    Public Shared Function CreateParameter(name As String, dbType As DbType, value As Object) As SQLiteParameter
        Dim param As New SQLiteParameter(name, dbType)
        param.Value = If(value Is Nothing, DirectCast(DBNull.Value, Object), value)
        Return param
    End Function

    ''' <summary>
    ''' 创建 String 参数
    ''' </summary>
    Public Shared Function CreateStringParam(name As String, value As String) As SQLiteParameter
        Return CreateParameter(name, DbType.String, value)
    End Function

    ''' <summary>
    ''' 创建 Int32 参数
    ''' </summary>
    Public Shared Function CreateIntParam(name As String, value As Integer?) As SQLiteParameter
        Return CreateParameter(name, DbType.Int32, If(value.HasValue, DirectCast(value.Value, Object), Nothing))
    End Function

    ''' <summary>
    ''' 创建 Decimal 参数
    ''' </summary>
    Public Shared Function CreateDecimalParam(name As String, value As Decimal?) As SQLiteParameter
        Return CreateParameter(name, DbType.Decimal, If(value.HasValue, DirectCast(value.Value, Object), Nothing))
    End Function

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not _isDisposed Then
            If disposing Then
                If _connection IsNot Nothing Then
                    If _connection.State = ConnectionState.Open Then
                        _connection.Close()
                    End If
                    _connection.Dispose()
                    _connection = Nothing
                End If
            End If
            _isDisposed = True
        End If
    End Sub
End Class
