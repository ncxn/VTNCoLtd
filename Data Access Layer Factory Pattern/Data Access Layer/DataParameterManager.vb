Imports System.Data
Imports Oracle.ManagedDataAccess.Client
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Data.Odbc
Imports MySql.Data.MySqlClient

Public Class DataParameterManager
    Public Shared Function CreateParameter(ByVal providerName As String, ByVal name As String, ByVal value As Object, ByVal dbType As DbType, ByVal Optional direction As ParameterDirection = ParameterDirection.Input) As IDbDataParameter
        Dim parameter As IDbDataParameter = Nothing

        Select Case providerName.ToLower()
            Case "system.data.sqlclient"
                Return CreateSqlParameter(name, value, dbType, direction)
            Case "system.data.oracleclient"
                Return CreateOracleParameter(name, value, dbType, direction)
            Case "system.data.oleDb"
            Case "mysql.data.mysqlclient"
                Return CreateMySqlParameter(name, value, dbType, direction)
        End Select

        Return parameter
    End Function

    Public Shared Function CreateParameter(ByVal providerName As String, ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType, ByVal Optional direction As ParameterDirection = ParameterDirection.Input) As IDbDataParameter
        Dim parameter As IDbDataParameter = Nothing

        Select Case providerName.ToLower()
            Case "system.data.sqlclient"
                Return CreateSqlParameter(name, size, value, dbType, direction)
            Case "system.data.oracleclient"
                Return CreateOracleParameter(name, size, value, dbType, direction)
            Case "system.data.oleDb"
            Case "MySql.Data.MySqlClient"
                Return CreateMySqlParameter(name, size, value, dbType, direction)
        End Select

        Return parameter
    End Function

    Private Shared Function CreateSqlParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New SqlParameter With {
            .DbType = dbType,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function

    Private Shared Function CreateSqlParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New SqlParameter With {
            .DbType = dbType,
            .Size = size,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function

    Private Shared Function CreateOracleParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New OracleParameter With {
            .DbType = dbType,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function

    Private Shared Function CreateOracleParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New OracleParameter With {
            .DbType = dbType,
            .Size = size,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function
    Private Shared Function CreateMySqlParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New MySqlParameter With {
            .DbType = dbType,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function

    Private Shared Function CreateMySqlParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New MySqlParameter With {
            .DbType = dbType,
            .Size = size,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function
    Private Shared Function CreateOleDbParameter(ByVal name As String, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New OleDbParameter With {
            .DbType = dbType,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function

    Private Shared Function CreateOleDbParameter(ByVal name As String, ByVal size As Integer, ByVal value As Object, ByVal dbType As DbType, ByVal direction As ParameterDirection) As IDbDataParameter
        Return New OleDbParameter With {
            .DbType = dbType,
            .Size = size,
            .ParameterName = name,
            .Direction = direction,
            .Value = value
        }
    End Function
End Class

