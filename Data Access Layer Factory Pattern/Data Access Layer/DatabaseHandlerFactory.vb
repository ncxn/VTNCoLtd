Imports System.Configuration

Public Class DatabaseHandlerFactory
    Private connectionStringSettings As ConnectionStringSettings

    Public Sub New(ByVal connectionStringName As String)
        connectionStringSettings = ConfigurationManager.ConnectionStrings(connectionStringName)
    End Sub

    Public Function CreateDatabase() As IDatabaseHandler
        Dim database As IDatabaseHandler = Nothing

        Select Case connectionStringSettings.ProviderName.ToLower()
            Case "system.data.sqlclient"
                database = New SqlDataAccess(connectionStringSettings.ConnectionString)
            Case "Oracle.ManagedDataAccess.Client"
                database = New OracleDataAccess(connectionStringSettings.ConnectionString)
            Case "system.data.oleDb"
                database = New OledbDataAccess(connectionStringSettings.ConnectionString)
            Case "MySql.Data.MySqlClient"
                database = New MySQLDataAccess(connectionStringSettings.ConnectionString)
        End Select

        Return database
    End Function

    Public Function GetProviderName() As String
        Return connectionStringSettings.ProviderName
    End Function
End Class

