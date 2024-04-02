Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Form2.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim result As DialogResult = MessageBox.Show("هل تريدون فعلا الخروج؟ ", "التاكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            End
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        connex()
    End Sub
    Private Sub ExporterDonnees()
        Dim exportationsFilePath As String = "Exportations_" & DateTime.Now.Year.ToString() & ".xlsx"
        Dim importationsFilePath As String = "Importations_" & DateTime.Now.Year.ToString() & ".xlsx"

        Try
            ' Exportation des données de la table "exportations"
            ExporterExportations(exportationsFilePath)

            ' Exportation des données de la table "importations"
            ExporterImportations(importationsFilePath)

            MessageBox.Show("Les données ont été exportées avec succès.", "Exportation réussie", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Une erreur s'est produite lors de l'exportation des données : " & ex.Message, "Erreur d'exportation", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Try
            Dim cmdSuppressionImportations As New SqlCommand("DELETE FROM importations", cnx)
            cmdSuppressionImportations.ExecuteNonQuery()

            Dim cmdSuppressionExportations As New SqlCommand("DELETE FROM exportations", cnx)
            cmdSuppressionExportations.ExecuteNonQuery()

            MessageBox.Show("Exportation terminée. La base de données a été vidée.", "Exportation", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Une erreur s'est produite lors de la suppression des enregistrements de la base de données.", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExporterExportations(filePath As String)
        Dim excelApp As Excel.Application = New Excel.Application()
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim excelWorksheet As Excel.Worksheet = excelWorkbook.ActiveSheet

        ' Connexion à la base de données
        Dim connectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\hassan.douiri\Desktop\application\Database1.mdf;Integrated Security=True"
        Dim query As String = "SELECT * FROM exportations"
        Using connection As SqlConnection = New SqlConnection(connectionString)
            connection.Open()

            ' Exécution de la requête
            Dim command As SqlCommand = New SqlCommand(query, connection)
            Dim dataReader As SqlDataReader = command.ExecuteReader()

            ' Remplissage des données dans la feuille Excel
            Dim row As Integer = 1
            While dataReader.Read()
                For col As Integer = 0 To dataReader.FieldCount - 1
                    excelWorksheet.Cells(row, col + 1) = dataReader(col).ToString()
                Next
                row += 1
            End While

            dataReader.Close()
        End Using

        ' Enregistrement du fichier Excel
        excelWorkbook.SaveAs(filePath)

        ' Fermeture des objets Excel
        excelWorkbook.Close()
        excelApp.Quit()
        ReleaseObject(excelWorksheet)
        ReleaseObject(excelWorkbook)
        ReleaseObject(excelApp)
    End Sub

    Private Sub ExporterImportations(filePath As String)
        Dim excelApp As Excel.Application = New Excel.Application()
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim excelWorksheet As Excel.Worksheet = excelWorkbook.ActiveSheet

        ' Connexion à la base de données
        Dim connectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Users\hassan.douiri\Desktop\application\Database1.mdf;Integrated Security=True"

        Dim query As String = "SELECT * FROM importations"
        Using connection As SqlConnection = New SqlConnection(connectionString)
            connection.Open()

            ' Exécution de la requête
            Dim command As SqlCommand = New SqlCommand(query, connection)
            Dim dataReader As SqlDataReader = command.ExecuteReader()

            ' Remplissage des données dans la feuille Excel
            Dim row As Integer = 1
            While dataReader.Read()
                For col As Integer = 0 To dataReader.FieldCount - 1
                    excelWorksheet.Cells(row, col + 1) = dataReader(col).ToString()
                Next
                row += 1
            End While

            dataReader.Close()
        End Using

        ' Enregistrement du fichier Excel
        excelWorkbook.SaveAs(filePath)

        ' Fermeture des objets Excel
        excelWorkbook.Close()
        excelApp.Quit()
        ReleaseObject(excelWorksheet)
        ReleaseObject(excelWorkbook)
        ReleaseObject(excelApp)
    End Sub

    Private Sub ReleaseObject(obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Dim result As DialogResult = MessageBox.Show("هل تريدون تصدير البيانات؟ ", "التاكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ExporterDonnees()
        End If

    End Sub
End Class
