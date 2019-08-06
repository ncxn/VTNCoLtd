Public Class Program
    Public Shared Sub Main(ByVal args As String())
        Console.WriteLine("================ Using Database Factory =================" & vbLf & vbLf & vbLf)
        UsingDatabaseFactory()
        Console.ReadKey()
    End Sub

    Private Shared Sub UsingDatabaseFactory()
        Dim dbManager = New DBManager("DBConnection")

        Dim user = New User With {
                .FirstName = "First",
                .LastName = "Last",
                .Dob = DateTime.Now.AddDays(-3000),
                .IsActive = True
            }
        Dim parameters = New List(Of IDbDataParameter)()
        parameters.Add(dbManager.CreateParameter("@FirstName", 50, user.FirstName, DbType.String))
        parameters.Add(dbManager.CreateParameter("@LastName", user.LastName, DbType.String))
        parameters.Add(dbManager.CreateParameter("@Dob", user.Dob, DbType.DateTime))
        parameters.Add(dbManager.CreateParameter("@IsActive", 50, user.IsActive, DbType.Boolean))
        Dim lastId As Integer = 0
        dbManager.Insert("DAH_User_Insert", CommandType.StoredProcedure, parameters.ToArray(), lastId)
        Console.WriteLine(vbLf & "INSERTED ID: " & lastId)
        Dim dataTable = dbManager.GetDataTable("DAH_User_GetAll", CommandType.StoredProcedure)
        Console.WriteLine(vbLf & "TOTAL ROWS IN TABLE: " & dataTable.Rows.Count)
        Dim connection As IDbConnection = Nothing
        Dim dataReader = dbManager.GetDataReader("DAH_User_GetAll", CommandType.StoredProcedure, Nothing, connection)

        Try
            user = New User()

            While dataReader.Read()
                user.FirstName = dataReader("FirstName").ToString()
                user.LastName = dataReader("LastName").ToString()
            End While

            Console.WriteLine(String.Format(vbLf & "DATA READER VALUES FirstName: {0} LastName: {1}", user.FirstName, user.LastName))
        Catch __unusedException1__ As Exception
        Finally
            dataReader.Close()
            dbManager.CloseConnection(connection)
        End Try

        Dim scalar As Object = dbManager.GetScalarValue("DAH_User_Scalar", CommandType.StoredProcedure)
        Console.WriteLine(vbLf & "SCALAR VALUE: " & scalar.ToString())
        user = New User With {
                .Id = lastId,
                .FirstName = "First1",
                .LastName = "Last1",
                .Dob = DateTime.Now.AddDays(-5000)
            }
        parameters = New List(Of IDbDataParameter)()
        parameters.Add(dbManager.CreateParameter("@Id", user.Id, DbType.Int32))
        parameters.Add(dbManager.CreateParameter("@FirstName", 50, user.FirstName, DbType.String))
        parameters.Add(dbManager.CreateParameter("@LastName", user.LastName, DbType.String))
        parameters.Add(dbManager.CreateParameter("@Dob", user.Dob, DbType.DateTime))
        dbManager.Update("DAH_User_Update", CommandType.StoredProcedure, parameters.ToArray())
        dataTable = dbManager.GetDataTable("DAH_User_GetAll", CommandType.StoredProcedure)
        Console.WriteLine(String.Format(vbLf & "UPADTED VALUES FirstName: {0} LastName: {1}", dataTable.Rows(0)("FirstName").ToString(), dataTable.Rows(0)("LastName").ToString()))
        parameters = New List(Of IDbDataParameter)()
        parameters.Add(dbManager.CreateParameter("@Id", user.Id, DbType.Int32))
        dbManager.Delete("DAH_User_Delete", CommandType.StoredProcedure, parameters.ToArray())
        Console.WriteLine(vbLf & "DELETED RECORD FOR ID: " & user.Id)
        dataTable = dbManager.GetDataTable("DAH_User_GetAll", CommandType.StoredProcedure)
        Console.WriteLine(vbLf & "TOTAL ROWS IN TABLE: " & dataTable.Rows.Count)
    End Sub
End Class