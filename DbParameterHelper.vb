Imports System.Data.SQLite
Imports System.Data

Public Module DbParameterHelper

    ''' <summary>
    ''' 创建 Int32 参数
    ''' </summary>
    Public Function CreateIntParam(name As String, value As Integer?) As SQLiteParameter
        Dim param As New SQLiteParameter(name, DbType.Int32)
        param.Value = If(value.HasValue, DirectCast(value.Value, Object), DBNull.Value)
        Return param
    End Function

    ''' <summary>
    ''' 创建 Int64 参数
    ''' </summary>
    Public Function CreateInt64Param(name As String, value As Long?) As SQLiteParameter
        Dim param As New SQLiteParameter(name, DbType.Int64)
        param.Value = If(value.HasValue, DirectCast(value.Value, Object), DBNull.Value)
        Return param
    End Function

    ''' <summary>
    ''' 创建 String 参数（自动处理空值）
    ''' </summary>
    Public Function CreateStringParam(name As String, value As String, Optional maxLength As Integer = -1) As SQLiteParameter
        Dim param As SQLiteParameter
        If maxLength > 0 Then
            param = New SQLiteParameter(name, DbType.String, maxLength)
        Else
            param = New SQLiteParameter(name, DbType.String)
        End If
        param.Value = If(String.IsNullOrEmpty(value), DirectCast(DBNull.Value, Object), value)
        Return param
    End Function

    ''' <summary>
    ''' 创建 Decimal 参数
    ''' </summary>
    Public Function CreateDecimalParam(name As String, value As Decimal?) As SQLiteParameter
        Dim param As New SQLiteParameter(name, DbType.Decimal)
        param.Value = If(value.HasValue, DirectCast(value.Value, Object), DBNull.Value)
        Return param
    End Function

    ''' <summary>
    ''' 创建 Double 参数
    ''' </summary>
    Public Function CreateDoubleParam(name As String, value As Double?) As SQLiteParameter
        Dim param As New SQLiteParameter(name, DbType.Double)
        param.Value = If(value.HasValue, DirectCast(value.Value, Object), DBNull.Value)
        Return param
    End Function

    ''' <summary>
    ''' 创建 DateTime 参数（SQLite 中存储为 TEXT）
    ''' </summary>
    Public Function CreateDateTimeParam(name As String, value As DateTime?) As SQLiteParameter
        Dim param As New SQLiteParameter(name, DbType.String)
        If value.HasValue Then
            param.Value = value.Value.ToString("yyyy-MM-dd HH:mm:ss")
        Else
            param.Value = DBNull.Value
        End If
        Return param
    End Function

    ''' <summary>
    ''' 创建 Boolean 参数（SQLite 中存储为 INTEGER 0/1）
    ''' </summary>
    Public Function CreateBoolParam(name As String, value As Boolean?) As SQLiteParameter
        Dim param As New SQLiteParameter(name, DbType.Int32)
        If value.HasValue Then
            param.Value = If(value.Value, 1, 0)
        Else
            param.Value = DBNull.Value
        End If
        Return param
    End Function

    ''' <summary>
    ''' 安全地解析字符串为 Int32
    ''' </summary>
    Public Function SafeParseInt(value As String) As Integer?
        Dim result As Integer
        If Integer.TryParse(value, result) Then
            Return result
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 安全地解析字符串为 Int64
    ''' </summary>
    Public Function SafeParseInt64(value As String) As Long?
        Dim result As Long
        If Long.TryParse(value, result) Then
            Return result
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 安全地解析字符串为 Decimal
    ''' </summary>
    Public Function SafeParseDecimal(value As String) As Decimal?
        Dim result As Decimal
        If Decimal.TryParse(value, result) Then
            Return result
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 安全地解析字符串为 Double
    ''' </summary>
    Public Function SafeParseDouble(value As String) As Double?
        Dim result As Double
        If Double.TryParse(value, result) Then
            Return result
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 安全地解析字符串为 DateTime
    ''' </summary>
    Public Function SafeParseDateTime(value As String) As DateTime?
        Dim result As DateTime
        If DateTime.TryParse(value, result) Then
            Return result
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 将字符串值转换为数据库参数值（自动处理空字符串为 DBNull）
    ''' </summary>
    Public Function StringToDbValue(value As String) As Object
        Return If(String.IsNullOrEmpty(value), DirectCast(DBNull.Value, Object), value)
    End Function

End Module
