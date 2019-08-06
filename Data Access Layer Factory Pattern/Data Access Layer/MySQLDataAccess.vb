Imports System.Data
Imports MySql.Data.MySqlClient
Public Class MySQLDataAccess
        Implements IDatabaseHandler

    Private Property _ConnectionString As String

    Public Sub New(ByVal connectionString As String)
        _ConnectionString = connectionString
    End Sub

        Public Function CreateConnection() As IDbConnection Implements IDatabaseHandler.CreateConnection
        Return New MySqlConnection(_ConnectionString)
    End Function

        Public Sub CloseConnection(ByVal connection As IDbConnection) Implements IDatabaseHandler.CloseConnection
            Dim MySqlConnection = CType(connection, MySqlConnection)
            MySqlConnection.Close()
            MySqlConnection.Dispose()
        End Sub

        Public Function CreateCommand(ByVal commandText As String, ByVal commandType As CommandType, ByVal connection As IDbConnection) As IDbCommand Implements IDatabaseHandler.CreateCommand
            Return New MySqlCommand With {
                .CommandText = commandText,
                .Connection = CType(connection, MySqlConnection),
                .CommandType = commandType
            }
        End Function

        Public Function CreateAdapter(ByVal command As IDbCommand) As IDataAdapter Implements IDatabaseHandler.CreateAdapter
            Return New MySqlDataAdapter(CType(command, MySqlCommand))
        End Function

        Public Function CreateParameter(ByVal command As IDbCommand) As IDbDataParameter Implements IDatabaseHandler.CreateParameter
            Dim SQLcommand As MySqlCommand = CType(command, MySqlCommand)
            Return SQLcommand.CreateParameter()
        End Function
    End Class

