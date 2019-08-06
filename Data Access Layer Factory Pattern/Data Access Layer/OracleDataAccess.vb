Imports System.Data
Imports Oracle.ManagedDataAccess.Client
Public Class OracleDataAccess
    Implements IDatabaseHandler

    Private Property _ConnectionString As String

    Public Sub New(ByVal _ConnectionString As String)
        _ConnectionString = _ConnectionString
    End Sub

    Public Function CreateConnection() As IDbConnection Implements IDatabaseHandler.CreateConnection
        Return New OracleConnection(_ConnectionString)
    End Function

    Public Sub CloseConnection(ByVal connection As IDbConnection) Implements IDatabaseHandler.CloseConnection
        Dim oracleConnection = CType(connection, OracleConnection)
        oracleConnection.Close()
        oracleConnection.Dispose()
    End Sub

    Public Function CreateCommand(ByVal commandText As String, ByVal commandType As CommandType, ByVal connection As IDbConnection) As IDbCommand Implements IDatabaseHandler.CreateCommand
        Return New OracleCommand With {
                .CommandText = commandText,
                .Connection = CType(connection, OracleConnection),
                .CommandType = commandType
            }
    End Function

    Public Function CreateAdapter(ByVal command As IDbCommand) As IDataAdapter Implements IDatabaseHandler.CreateAdapter
        Return New OracleDataAdapter(CType(command, OracleCommand))
    End Function

    Public Function CreateParameter(ByVal command As IDbCommand) As IDbDataParameter Implements IDatabaseHandler.CreateParameter
        Dim SQLcommand As OracleCommand = CType(command, OracleCommand)
        Return SQLcommand.CreateParameter()
    End Function
End Class


