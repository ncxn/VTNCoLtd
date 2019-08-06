Imports System
Imports System.Data
Imports System.Runtime.InteropServices

Public Class DBManager
    Private dbFactory As DatabaseHandlerFactory
    Private database As IDatabaseHandler
    Private providerName As String

    Public Sub New(ByVal connectionStringName As String)
        dbFactory = New DatabaseHandlerFactory(connectionStringName)
        database = dbFactory.CreateDatabase()
        providerName = dbFactory.GetProviderName()
    End Sub

    Public Function GetDatabasecOnnection() As IDbConnection
        Return database.CreateConnection()
    End Function

    Public Sub CloseConnection(ByVal connection As IDbConnection)
        database.CloseConnection(connection)
    End Sub

    Public Function CreateParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType) As IDbDataParameter
        Return DataParameterManager.CreateParameter(providerName, name, value, dbType, ParameterDirection.Input)
    End Function

    Public Function CreateParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType) As IDbDataParameter
        Return DataParameterManager.CreateParameter(providerName, name, size, value, dbType, ParameterDirection.Input)
    End Function

    Public Function CreateParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return DataParameterManager.CreateParameter(providerName, name, size, value, dbType, direction)
    End Function

    Public Function GetDataTable(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As IDbDataParameter() = Nothing) As DataTable
        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Dim dataset = New DataSet()
                Dim dataAdaper = database.CreateAdapter(command)
                dataAdaper.Fill(dataset)
                Return dataset.Tables(0)
            End Using
        End Using
    End Function

    Public Function GetDataSet(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As IDbDataParameter() = Nothing) As DataSet
        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Dim dataset = New DataSet()
                Dim dataAdaper = database.CreateAdapter(command)
                dataAdaper.Fill(dataset)
                Return dataset
            End Using
        End Using
    End Function

    Public Function GetDataReader(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As IDbDataParameter(), <Out> ByRef connection As IDbConnection) As IDataReader
        Dim reader As IDataReader = Nothing
        connection = database.CreateConnection()
        connection.Open()
        Dim command = database.CreateCommand(commandText, commandType, connection)

        If parameters IsNot Nothing Then

            For Each parameter In parameters
                command.Parameters.Add(parameter)
            Next
        End If

        reader = command.ExecuteReader()
        Return reader
    End Function

    Public Sub Delete(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As IDbDataParameter() = Nothing)
        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub Insert(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As IDbDataParameter())
        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function Insert(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As IDbDataParameter(), <Out> ByRef lastId As Integer) As Integer
        lastId = 0

        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Dim newId As Object = command.ExecuteScalar()
                lastId = Convert.ToInt32(newId)
            End Using
        End Using

        Return lastId
    End Function

    Public Function Insert(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As IDbDataParameter(), <Out> ByRef lastId As Long) As Long
        lastId = 0

        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Dim newId As Object = command.ExecuteScalar()
                lastId = Convert.ToInt64(newId)
            End Using
        End Using

        Return lastId
    End Function

    Public Sub InsertWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As IDbDataParameter())
        Dim transactionScope As IDbTransaction = Nothing

        Using connection = database.CreateConnection()
            connection.Open()
            transactionScope = connection.BeginTransaction()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Try
                    command.ExecuteNonQuery()
                    transactionScope.Commit()
                Catch __unusedException1__ As Exception
                    transactionScope.Rollback()
                Finally
                    connection.Close()
                End Try
            End Using
        End Using
    End Sub

    Public Sub InsertWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal isolationLevel As IsolationLevel, ByVal parameters As IDbDataParameter())
        Dim transactionScope As IDbTransaction = Nothing

        Using connection = database.CreateConnection()
            connection.Open()
            transactionScope = connection.BeginTransaction(isolationLevel)

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Try
                    command.ExecuteNonQuery()
                    transactionScope.Commit()
                Catch __unusedException1__ As Exception
                    transactionScope.Rollback()
                Finally
                    connection.Close()
                End Try
            End Using
        End Using
    End Sub

    Public Sub Update(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As IDbDataParameter())
        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub UpdateWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As IDbDataParameter())
        Dim transactionScope As IDbTransaction = Nothing

        Using connection = database.CreateConnection()
            connection.Open()
            transactionScope = connection.BeginTransaction()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Try
                    command.ExecuteNonQuery()
                    transactionScope.Commit()
                Catch __unusedException1__ As Exception
                    transactionScope.Rollback()
                Finally
                    connection.Close()
                End Try
            End Using
        End Using
    End Sub

    Public Sub UpdateWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal isolationLevel As IsolationLevel, ByVal parameters As IDbDataParameter())
        Dim transactionScope As IDbTransaction = Nothing

        Using connection = database.CreateConnection()
            connection.Open()
            transactionScope = connection.BeginTransaction(isolationLevel)

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Try
                    command.ExecuteNonQuery()
                    transactionScope.Commit()
                Catch __unusedException1__ As Exception
                    transactionScope.Rollback()
                Finally
                    connection.Close()
                End Try
            End Using
        End Using
    End Sub

    Public Function GetScalarValue(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As IDbDataParameter() = Nothing) As Object
        Using connection = database.CreateConnection()
            connection.Open()

            Using command = database.CreateCommand(commandText, commandType, connection)

                If parameters IsNot Nothing Then

                    For Each parameter In parameters
                        command.Parameters.Add(parameter)
                    Next
                End If

                Return command.ExecuteScalar()
            End Using
        End Using
    End Function
End Class

