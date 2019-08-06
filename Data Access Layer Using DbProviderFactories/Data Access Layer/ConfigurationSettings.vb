Imports System
Imports System.Configuration


Friend Module ConfigurationSettings
        Public ReadOnly Property DefaultConnection As String
            Get
                Return ConfigurationManager.ConnectionStrings("DefaultConnection").ToString()
            End Get
        End Property

        Public ReadOnly Property ProviderName As String
            Get
                Return ConfigurationManager.ConnectionStrings(DefaultConnection).ProviderName
            End Get
        End Property

        Public ReadOnly Property ConnectionString As String
            Get

                Try
                    Return ConfigurationManager.ConnectionStrings(DefaultConnection).ConnectionString
                Catch __unusedException1__ As Exception
                    Throw New Exception(String.Format("Connection string '{0}' not found.", DefaultConnection))
                End Try
            End Get
        End Property

        Function GetConnectionString(ByVal connectionName As String) As String
            Try
                Return ConfigurationManager.ConnectionStrings(connectionName).ConnectionString
            Catch __unusedException1__ As Exception
                Throw New Exception(String.Format("Connection string '{0}' not found.", connectionName))
            End Try
        End Function

        Function GetProviderName(ByVal connectionName As String) As String
            Try
                Return ConfigurationManager.ConnectionStrings(connectionName).ProviderName
            Catch __unusedException1__ As Exception
                Throw New Exception(String.Format("Connection string '{0}' not found.", connectionName))
            End Try
        End Function
    End Module

