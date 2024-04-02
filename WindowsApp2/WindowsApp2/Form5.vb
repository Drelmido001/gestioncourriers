Imports System.Data.SqlClient

Public Class Form5
    Private Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            Dim cmd1 As New SqlCommand("select num_enreg as[رقم_التسجيل],date_enreg as[تاريخ_التسجيل],source as[المرجع],outil as [الطريقة],type as[نوع الوجهة],destination as[الوجهة],doc as[عدد المرفقات],sujet as[الموضوع],wd as[الوضعية],des as[وحدة_الاصدار] from exportations ", cnx)
            Dim dtr As SqlDataReader
            dtr = cmd1.ExecuteReader

            Dim dt As New DataTable
            '   dt.Clear()
            dt.Load(dtr)

            DataGridView1.DataSource = dt
            'dt.Clear()
            dtr.Close()
        Else
            Try
                Dim cmd As New SqlCommand("select num_enreg as[رقم_التسجيل],date_enreg as[تاريخ_التسجيل],source as[المرجع],outil as [الطريقة],type as[نوع الوجهة],destination as[الوجهة],doc as[عدد المرفقات],sujet as[الموضوع],wd as[الوضعية] ,des as[وحدة_الاصدار]from exportations where num_enreg=" & TextBox2.Text, cnx)
                Dim dtr As SqlDataReader
                dtr = cmd.ExecuteReader
                If dtr.HasRows Then
                    Dim dt As New DataTable
                    dt.Load(dtr)
                    DataGridView1.DataSource = dt
                    dtr.Close()
                Else
                    MsgBox("لايوجد")
                    dtr.Close()
                End If
                dtr.Close()
            Catch ex As Exception
                MsgBox("!يوجد مشكل")

            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form3.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim selectedYear As String = DateTimePicker3.Value.Year.ToString() ' Obtenez l'année sélectionnée
            Dim selectedMonth As String = DateTimePicker3.Value.Month.ToString() ' Obtenez le mois sélectionné

            Dim cmd As New SqlCommand("SELECT num_enreg AS [رقم_التسجيل], date_enreg AS [تاريخ_التسجيل], source AS [المرجع], outil AS [الطريقة], type AS [نوع الوجهة], destination AS [الوجهة], doc AS [عدد المرفقات], sujet AS [الموضوع], wd AS [الوضعية], des AS [وحدة_الاصدار] FROM exportations WHERE YEAR(date_enreg) = @selectedYear AND MONTH(date_enreg) = @selectedMonth", cnx)
            cmd.Parameters.AddWithValue("@selectedYear", selectedYear)
            cmd.Parameters.AddWithValue("@selectedMonth", selectedMonth)

            Dim dtr As SqlDataReader
            dtr = cmd.ExecuteReader

            If dtr.HasRows Then
                Dim dt As New DataTable
                dt.Load(dtr)
                DataGridView1.DataSource = dt
            Else
                MsgBox("لا يوجد بيانات")
            End If

            dtr.Close()
        Catch ex As Exception
            MsgBox("حدثت مشكلة!")
        End Try
    End Sub

    Private Sub Form5_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim cmd As New SqlCommand("select num_enreg as[رقم_التسجيل],date_enreg as[تاريخ_التسجيل],source as[المرجع],outil as [الطريقة],type as[نوع الوجهة],destination as[الوجهة],doc as[عدد المرفقات],sujet as[الموضوع],wd as[الوضعية],des as[وحدة_الاصدار] from exportations ", cnx)
        Dim dtr As SqlDataReader
        dtr = cmd.ExecuteReader()

        Dim dt As New DataTable
        '   dt.Clear()
        dt.Load(dtr)
        DataGridView1.DataSource = dt
        dtr.Close()
        TextBox2.Text = ""
    End Sub
End Class
