Imports System.Data.SqlClient

Module Module1
    Public cnx As SqlConnection

    Public Sub connex()
        Try
            cnx = New SqlConnection("Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\hassan.douiri\source\repos\WindowsApp2\WindowsApp2\Database1.mdf;Integrated Security=True")

            cnx.Open()
        Catch ex As Exception
            MsgBox(ex.Message)
            End
        End Try
    End Sub

End Module
