Imports System
Imports System.Data
Imports System.Data.Common


Public Class DatabaseHelper
        Public Property ProviderManager As ProviderManager
        Public Property ConnectionString As String

        Public Sub New()
            ConnectionString = ConfigurationSettings.ConnectionString
            ProviderManager = New ProviderManager()
        End Sub

        Public Sub New(ByVal connectionName As String)
            ConnectionString = ConfigurationSettings.GetConnectionString(connectionName)
            ProviderManager = New ProviderManager(ConfigurationSettings.GetProviderName(connectionName))
        End Sub

        Public Sub New(ByVal connectionString As String, ByVal providerName As String)
            connectionString = connectionString
            ProviderManager = New ProviderManager(providerName)
        End Sub

        Public Function GetConnection() As IDbConnection
            Try
                Dim connection = ProviderManager.Factory.CreateConnection()
                connection.ConnectionString = ConnectionString
                connection.Open()
                Return connection
            Catch __unusedException1__ As Exception
                Throw New Exception("Error occured while creating connection. Please check connection string and provider name.")
            End Try
        End Function

        Public Sub CloseConnection(ByVal connection As IDbConnection)
            connection.Close()
        End Sub

        Public Function GetCommand(ByVal commandText As String, ByVal connection As IDbConnection, ByVal commandType As CommandType) As IDbCommand
            Try
                Dim command As IDbCommand = ProviderManager.Factory.CreateCommand()
                command.CommandText = commandText
                command.Connection = connection
                command.CommandType = commandType
                Return command
            Catch __unusedException1__ As Exception
                Throw New Exception("Invalid parameter 'commandText'.")
            End Try
        End Function

        Public Function GetDataAdapter(ByVal command As IDbCommand) As DbDataAdapter
            Dim adapter As DbDataAdapter = ProviderManager.Factory.CreateDataAdapter()
            adapter.SelectCommand = CType(command, DbCommand)
            adapter.InsertCommand = CType(command, DbCommand)
            adapter.UpdateCommand = CType(command, DbCommand)
            adapter.DeleteCommand = CType(command, DbCommand)
            Return adapter
        End Function
    Public Function GetParameter(ByVal name As String, ByVal value As Object) As IDbDataParameter
        Try
            Dim dbParam As DbParameter = ProviderManager.Factory.CreateParameter()
            dbParam.ParameterName = name
            dbParam.Value = value
            dbParam.Direction = ParameterDirection.Input
            'dbParam.DbType = DbType
            Return dbParam
        Catch __unusedException1__ As Exception
            Throw New Exception("Invalid parameter or type.")
        End Try
    End Function
    Public Function GetParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType) As DbParameter
        Try
            Dim dbParam As DbParameter = ProviderManager.Factory.CreateParameter()
            dbParam.ParameterName = name
            dbParam.Value = value
            dbParam.Direction = ParameterDirection.Input
            dbParam.DbType = dbType
            Return dbParam
        Catch __unusedException1__ As Exception
            Throw New Exception("Invalid parameter or type.")
        End Try
    End Function

    Public Function GetParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType, ByVal parameterDirection As ParameterDirection) As DbParameter
            Try
                Dim dbParam As DbParameter = ProviderManager.Factory.CreateParameter()
                dbParam.ParameterName = name
                dbParam.Value = value
                dbParam.Direction = parameterDirection
                dbParam.DbType = dbType
                Return dbParam
            Catch __unusedException1__ As Exception
                Throw New Exception("Invalid parameter or type.")
            End Try
        End Function

        Public Function GetParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType, ByVal size As Integer, ByVal parameterDirection As ParameterDirection) As DbParameter
            Try
                Dim dbParam As DbParameter = ProviderManager.Factory.CreateParameter()
                dbParam.ParameterName = name
                dbParam.Value = value
                dbParam.Size = size
                dbParam.Direction = parameterDirection
                dbParam.DbType = dbType
                Return dbParam
            Catch __unusedException1__ As Exception
                Throw New Exception("Invalid parameter or type.")
            End Try
        End Function
    End Class
