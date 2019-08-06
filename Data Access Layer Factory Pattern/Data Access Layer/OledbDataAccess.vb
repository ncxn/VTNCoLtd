Imports System.Data
Imports System.Data.OleDb

Public Class OledbDataAccess
    Implements IDatabaseHandler

    Private Property _ConnectionString As String

    Public Sub New(ByVal connectionString As String)
        _ConnectionString = connectionString
    End Sub

    Public Function CreateConnection() As IDbConnection Implements IDatabaseHandler.CreateConnection
        Return New OleDbConnection(_ConnectionString)
    End Function

    Public Sub CloseConnection(ByVal connection As IDbConnection) Implements IDatabaseHandler.CloseConnection
        Dim oleDbConnection = CType(connection, OleDbConnection)
        oleDbConnection.Close()
        oleDbConnection.Dispose()
    End Sub

    Public Function CreateCommand(ByVal commandText As String, ByVal commandType As CommandType, ByVal connection As IDbConnection) As IDbCommand Implements IDatabaseHandler.CreateCommand
        Return New OleDbCommand With {
            .CommandText = commandText,
            .Connection = CType(connection, OleDbConnection),
            .CommandType = commandType
        }
    End Function

    Public Function CreateAdapter(ByVal command As IDbCommand) As IDataAdapter Implements IDatabaseHandler.CreateAdapter
        Return New OleDbDataAdapter(CType(command, OleDbCommand))
    End Function

    Public Function CreateParameter(ByVal command As IDbCommand) As IDbDataParameter Implements IDatabaseHandler.CreateParameter
        Dim SQLcommand As OleDbCommand = CType(command, OleDbCommand)
        Return SQLcommand.CreateParameter()
    End Function
End Class

