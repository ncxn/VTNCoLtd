Imports System.Data.SqlClient
Public Class SqlDataAccess
    Implements IDatabaseHandler

    Private Property _ConnectionString As String

    Public Sub New(ByVal connectionString As String)
        _ConnectionString = connectionString
    End Sub

    Public Function CreateConnection() As IDbConnection Implements IDatabaseHandler.CreateConnection
        Return New SqlConnection(_ConnectionString)
    End Function

    Public Sub CloseConnection(ByVal connection As IDbConnection) Implements IDatabaseHandler.CloseConnection
        Dim sqlConnection = CType(connection, SqlConnection)
        sqlConnection.Close()
        sqlConnection.Dispose()
    End Sub

    Public Function CreateCommand(ByVal commandText As String, ByVal commandType As CommandType, ByVal connection As IDbConnection) As IDbCommand Implements IDatabaseHandler.CreateCommand
        Return New SqlCommand With {
            .CommandText = commandText,
            .Connection = CType(connection, SqlConnection),
            .CommandType = commandType
        }
    End Function

    Public Function CreateAdapter(ByVal command As IDbCommand) As IDataAdapter Implements IDatabaseHandler.CreateAdapter
        Return New SqlDataAdapter(CType(command, SqlCommand))
    End Function

    Public Function CreateParameter(ByVal command As IDbCommand) As IDbDataParameter Implements IDatabaseHandler.CreateParameter
        Dim SQLcommand As SqlCommand = CType(command, SqlCommand)
        Return SQLcommand.CreateParameter()
    End Function

End Class
