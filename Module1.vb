Imports System.Data.SQLite
Imports System.Data

''' <summary>
''' 模块级别的共享资源
''' 注意：尽量减少全局状态，优先使用 PcfDataCollector 类
''' </summary>
Module Module1

    ' 数据库连接（保持为模块级别，因为需要在多个地方使用）
    Public ConnectionStringBuilder As SQLiteConnectionStringBuilder
    Public DbHelper As DatabaseHelper

    ' 当前活动的 PCF 数据收集器实例
    Public CurrentDataCollector As PcfDataCollector

    ''' <summary>
    ''' 初始化数据库连接
    ''' </summary>
    Public Sub InitializeDatabaseHelper()
        If ConnectionStringBuilder IsNot Nothing Then
            DbHelper = New DatabaseHelper(ConnectionStringBuilder.ConnectionString)
        End If
    End Sub

    ''' <summary>
    ''' 创建新的数据收集器实例
    ''' </summary>
    Public Function CreateDataCollector() As PcfDataCollector
        CurrentDataCollector = New PcfDataCollector()
        If ConnectionStringBuilder IsNot Nothing Then
            CurrentDataCollector.ConnectionStringBuilder = ConnectionStringBuilder
        End If
        Return CurrentDataCollector
    End Function

    ''' <summary>
    ''' 获取或创建数据收集器
    ''' </summary>
    Public Function GetOrCreateDataCollector() As PcfDataCollector
        If CurrentDataCollector Is Nothing Then
            Return CreateDataCollector()
        End If
        Return CurrentDataCollector
    End Function

    ''' <summary>
    ''' 重置当前数据收集器
    ''' </summary>
    Public Sub ResetDataCollector()
        If CurrentDataCollector IsNot Nothing Then
            CurrentDataCollector.Reset()
        End If
    End Sub

    ' 组件类型常量
    Public ReadOnly ComponentTypes As String() = PcfDataCollector.ComponentTypesField
    Public ReadOnly PointTypes As String() = PcfDataCollector.PointTypesField

    ''' <summary>
    ''' 填充 DataTable（无参数版本）
    ''' </summary>
    Public Function FillDataTable(sql As String) As DataTable
        Return DbHelper.FillDataTable(sql)
    End Function

End Module
