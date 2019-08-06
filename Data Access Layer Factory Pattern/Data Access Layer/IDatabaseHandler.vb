Public Interface IDatabaseHandler
    Function CreateConnection() As IDbConnection
    Sub CloseConnection(ByVal connection As IDbConnection)
    Function CreateCommand(ByVal commandText As String, ByVal commandType As CommandType, ByVal connection As IDbConnection) As IDbCommand
    Function CreateAdapter(ByVal command As IDbCommand) As IDataAdapter
    Function CreateParameter(ByVal command As IDbCommand) As IDbDataParameter
End Interface
