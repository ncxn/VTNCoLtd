Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices

Namespace DataAccessHandler
    Public Class MSSqlHelper
        Private Property ConnectionString As String

        Public Sub New(ByVal connectionString As String)
            connectionString = connectionString
        End Sub

        Public Sub CloseConnection(ByVal connection As SqlConnection)
            connection.Close()
        End Sub

        Public Function CreateParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType) As SqlParameter
            Return CreateParameter(name, 0, value, dbType, ParameterDirection.Input)
        End Function

        Public Function CreateParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType) As SqlParameter
            Return CreateParameter(name, size, value, dbType, ParameterDirection.Input)
        End Function

        Public Function CreateParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As SqlParameter
            Return New SqlParameter With {
                .DbType = dbType,
                .ParameterName = name,
                .Size = size,
                .Direction = direction,
                .Value = value
            }
        End Function

        Public Function GetDataTable(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As SqlParameter() = Nothing) As DataTable
            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

                    If parameters IsNot Nothing Then

                        For Each parameter In parameters
                            command.Parameters.Add(parameter)
                        Next
                    End If

                    Dim dataset = New DataSet()
                    Dim dataAdaper = New SqlDataAdapter(command)
                    dataAdaper.Fill(dataset)
                    Return dataset.Tables(0)
                End Using
            End Using
        End Function

        Public Function GetDataSet(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As SqlParameter() = Nothing) As DataSet
            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

                    If parameters IsNot Nothing Then

                        For Each parameter In parameters
                            command.Parameters.Add(parameter)
                        Next
                    End If

                    Dim dataset = New DataSet()
                    Dim dataAdaper = New SqlDataAdapter(command)
                    dataAdaper.Fill(dataset)
                    Return dataset
                End Using
            End Using
        End Function

        Public Function GetDataReader(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As SqlParameter(), <Out> ByRef connection As SqlConnection) As IDataReader
            Dim reader As IDataReader = Nothing
            connection = New SqlConnection(ConnectionString)
            connection.Open()
            Dim command = New SqlCommand(commandText, connection)
            command.CommandType = commandType

            If parameters IsNot Nothing Then

                For Each parameter In parameters
                    command.Parameters.Add(parameter)
                Next
            End If

            reader = command.ExecuteReader()
            Return reader
        End Function

        Public Sub Delete(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As SqlParameter() = Nothing)
            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

                    If parameters IsNot Nothing Then

                        For Each parameter In parameters
                            command.Parameters.Add(parameter)
                        Next
                    End If

                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub Insert(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As SqlParameter())
            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

                    If parameters IsNot Nothing Then

                        For Each parameter In parameters
                            command.Parameters.Add(parameter)
                        Next
                    End If

                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Function Insert(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As SqlParameter(), <Out> ByRef lastId As Integer) As Integer
            lastId = 0

            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

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

        Public Function Insert(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As SqlParameter(), <Out> ByRef lastId As Long) As Long
            lastId = 0

            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

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

        Public Sub InsertWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As SqlParameter())
            Dim transactionScope As SqlTransaction = Nothing

            Using connection = New SqlConnection(ConnectionString)
                connection.Open()
                transactionScope = connection.BeginTransaction()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

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

        Public Sub InsertWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal isolationLevel As IsolationLevel, ByVal parameters As SqlParameter())
            Dim transactionScope As SqlTransaction = Nothing

            Using connection = New SqlConnection(ConnectionString)
                connection.Open()
                transactionScope = connection.BeginTransaction(isolationLevel)

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

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

        Public Sub Update(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As SqlParameter())
            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

                    If parameters IsNot Nothing Then

                        For Each parameter In parameters
                            command.Parameters.Add(parameter)
                        Next
                    End If

                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Public Sub UpdateWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal parameters As SqlParameter())
            Dim transactionScope As SqlTransaction = Nothing

            Using connection = New SqlConnection(ConnectionString)
                connection.Open()
                transactionScope = connection.BeginTransaction()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

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

        Public Sub UpdateWithTransaction(ByVal commandText As String, ByVal commandType As CommandType, ByVal isolationLevel As IsolationLevel, ByVal parameters As SqlParameter())
            Dim transactionScope As SqlTransaction = Nothing

            Using connection = New SqlConnection(ConnectionString)
                connection.Open()
                transactionScope = connection.BeginTransaction(isolationLevel)

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

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

        Public Function GetScalarValue(ByVal commandText As String, ByVal commandType As CommandType, ByVal Optional parameters As SqlParameter() = Nothing) As Object
            Using connection = New SqlConnection(ConnectionString)
                connection.Open()

                Using command = New SqlCommand(commandText, connection)
                    command.CommandType = commandType

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
End Namespace
