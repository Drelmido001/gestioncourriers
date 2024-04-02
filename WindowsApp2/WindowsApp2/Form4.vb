Imports System.Data.SqlClient

Public Class Form4


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = "" Then
            Dim cmd1 As New SqlCommand("select num as[رقم_التسجيل],date as[تاريخ_التسجيل],type as [نوعها],source as[المصدر],num_enreg as [رقم_الصدور],date_e as [تاريخ_الصدور],sujet as [الموضوع],num_doc as[عدد_المرفقات],pour as[لأجل],delais as[اخر_اجل],outil as[ورد_عن_طريق],destination as[الوجهة_الادارية] from importations ", cnx)
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
                Dim cmd As New SqlCommand("select num as[رقم_التسجيل],date as[تاريخ_التسجيل],type as[نوعها],source as[المصدر],num_enreg as[رقم_الصدور],date_e as[تاريخ_الصدور],sujet as[الموضوع],num_doc as[عدد_المرفقات],pour as[لأجل],delais as[اخر_اجل],outil as[ورد_عن_طريق],destination as[الوجهة_الادارية] from importations where num=" & TextBox2.Text, cnx)
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



    Private Sub Button3_Click(sender As Object, e As EventArgs)


    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        Form2.Show()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim selectedYear As String = DateTimePicker3.Value.Year.ToString() ' Obtenez l'année sélectionnée
            Dim selectedMonth As String = DateTimePicker3.Value.Month.ToString() ' Obtenez le mois sélectionné
            Dim cmd As New SqlCommand("SELECT num AS [رقم_التسجيل], date AS [تاريخ_التسجيل], type AS [نوعها], source AS [المصدر], num_enreg AS [رقم_الصدور], date_e AS [تاريخ_الصدور], sujet AS [الموضوع], num_doc AS [عدد_المرفقات], pour AS [لأجل], delais AS [اخر_اجل], outil AS [ورد_عن_طريق], destination AS [الوجهة_الادارية] FROM importations WHERE YEAR(date) = " & selectedYear & " AND MONTH(date) = " & selectedMonth, cnx)
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
    End Sub

    Private Sub Form4_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim cmd As New SqlCommand("select num as[رقم_التسجيل],date as[تاريخ_التسجيل],type as [نوعها],source as[المصدر],num_enreg as [رقم_الصدور],date_e as [تاريخ_الصدور],sujet as [الموضوع],num_doc as[عدد_المرفقات],pour as[لأجل],delais as[اخر_اجل],outil as[ورد_عن_طريق],destination as[الوجهة_الادارية] from importations ", cnx)
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