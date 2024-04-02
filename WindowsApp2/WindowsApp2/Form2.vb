Imports System.Data.SqlClient
Imports System.Diagnostics.Eventing.Reader

Public Class Form2


    Private Sub vide()
        TextBox2.Text = ""
        TextBox5.Text = ""
        TextBox3.Text = ""
        ComboBox1.SelectedItem = ""
        ComboBox4.SelectedItem = ""
        ComboBox2.SelectedItem = ""
        ComboBox3.SelectedItem = ""
        RichTextBox1.Text = " "
        ComboBox5.SelectedItem = ""
        DateTimePicker3.Value = DateTime.Now
        DateTimePicker2.Value = DateTime.Now
        DateTimePicker1.Value = DateTime.Now
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub



    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub


    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Hide()
        Form4.Show()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex = 1 Then
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("")
            ComboBox4.Items.Add("المفتشية العامة للشؤون الإدارية")
            ComboBox4.Items.Add("المفتشية العامة للشؤون التربوية")
            ComboBox4.Items.Add("مديرية التقويم وتنظيم الحياة المدرسية والتكوينات")
            ComboBox4.Items.Add("المركز الوطني للتقويم والامتحانات")
            ComboBox4.Items.Add("المركز الوطني للتجديد والتجريب")
            ComboBox4.Items.Add("مديرية الاستراتيجية والاحصاء والتخطيط")
            ComboBox4.Items.Add("مديرية الشؤون القانونية والمنازعات")
            ComboBox4.Items.Add("مديرية الموارد البشرية وتكوين الأطر")
            ComboBox4.Items.Add("مديرية المناهج")
            ComboBox4.Items.Add("مديرية الارتقاء بالرياضة المدرسية")
            ComboBox4.Items.Add("مديرية إدارة منظومة الإعلام")
            ComboBox4.Items.Add("مديرية الشؤون العامة والميزانية والممتلكات")
            ComboBox4.Items.Add("مديرية التعاون والارتقاء بالتعليم المدرسي الخصوصي")
            ComboBox4.Items.Add("الكتابة العامة – قسم الاتصال")
        ElseIf ComboBox3.SelectedIndex = 2 Then
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("")
            ComboBox4.Items.Add("مديرية طنجة أصيلة")
            ComboBox4.Items.Add("مديرية الفحص أنجرة")
            ComboBox4.Items.Add("مديرية العرائش")
            ComboBox4.Items.Add("مديرية تطوان")
            ComboBox4.Items.Add("مديرية المضيق الفنيدق")
            ComboBox4.Items.Add("مديرية شفشاون")
            ComboBox4.Items.Add("مديرية وزان")
            ComboBox4.Items.Add("مديرية الحسيمة")

        ElseIf ComboBox3.SelectedIndex = 3 Then
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("")
            ComboBox4.Items.Add("أكاديمية جهة الداخلة واد الذهب")
            ComboBox4.Items.Add("أكاديمية العيون الساقية الحمراء")
            ComboBox4.Items.Add("أكاديمية جهة كلميم واد نون")
            ComboBox4.Items.Add("أكاديمية جهة سوس ماسة")
            ComboBox4.Items.Add("أكاديمية جهة مراكش آسفي")
            ComboBox4.Items.Add("أكاديمية جهة الشرق")
            ComboBox4.Items.Add("أكاديمية جهة الدار البيضاء سطات")
            ComboBox4.Items.Add("أكاديمية جهة الرباط سلا القنيطرة")
            ComboBox4.Items.Add("أكاديمية جهة بني ملال خنيفرة")
            ComboBox4.Items.Add("أكاديمية جهة فاس مكناس")
            ComboBox4.Items.Add("أكاديمية درعة تافيلالت")
        ElseIf ComboBox3.SelectedIndex = 4 Then
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("خارجية")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox2.Text = "" Then
            MessageBox.Show("ادخل رقم التسجيل", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim cmd As New SqlCommand("SELECT num FROM importations WHERE num = " & TextBox2.Text, cnx)
            Dim dtr As SqlDataReader = cmd.ExecuteReader()

            If dtr.Read() Then
                dtr.Close()

                Dim result As DialogResult = MessageBox.Show("هل تريدون ان تحذفوا هذه الواردة؟", "التاكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                If result = DialogResult.Yes Then
                    Try
                        Dim cmd2 As New SqlCommand("DELETE FROM importations WHERE num = " & TextBox2.Text, cnx)
                        cmd2.ExecuteNonQuery()

                        MsgBox("تم الحذف بنجاح")
                    Catch ex As Exception
                        MsgBox("!يوجد مشكل")
                    End Try
                Else
                    MessageBox.Show("لم يتم حذف الواردة")
                End If
            Else
                dtr.Close()
                MessageBox.Show("لا يوجد واردة بهذا الرقم")
            End If
        End If
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
            Try
            If DateTimePicker2.Value < DateTimePicker1.Value Or DateTimePicker2.Value = DateTimePicker1.Value Then
                Throw New Exception("خطا في مراعاة التسلسل الزمني بين تاريخي الصدور و التسجيل")
            ElseIf DateTimePicker3.Value < DateTimePicker2.Value Then
                Throw New Exception("خطا في مراعاة التسلسل الزمني بين تاريخي التسجيل و اخر اجل")
            ElseIf String.IsNullOrWhiteSpace(TextBox2.Text) Or String.IsNullOrWhiteSpace(TextBox3.Text) Or String.IsNullOrWhiteSpace(TextBox5.Text) Or ComboBox1.SelectedItem Is Nothing Or ComboBox2.SelectedItem Is Nothing Or ComboBox3.SelectedItem Is Nothing Or ComboBox4.SelectedItem Is Nothing Or ComboBox5.SelectedItem Is Nothing Or String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                Throw New Exception("يجب ان تملأ جميع الحقول")
            Else

                Dim cmd As New SqlCommand("SELECT num FROM importations WHERE num = " & TextBox2.Text, cnx)
                Dim dtr As SqlDataReader = cmd.ExecuteReader()
                If dtr.Read() Then
                    dtr.Close()
                    Throw New Exception("الرقم الذي ادخلتموه يوجد")
                End If
                dtr.Close()
                ' Formater les valeurs de date avant de les insérer dans la requête
                Dim date1 As String = DateTimePicker1.Value.ToString("yyyy-MM-dd")
                Dim date2 As String = DateTimePicker2.Value.ToString("yyyy-MM-dd")
                Dim date3 As String = DateTimePicker3.Value.ToString("yyyy-MM-dd")


                Dim cmd1 As New SqlCommand("INSERT INTO importations VALUES (" & TextBox2.Text & ", '" & date1 & "', N'" & ComboBox3.SelectedItem & "', N'" & ComboBox4.SelectedItem & "', " & TextBox3.Text & ", '" & date2 & "', N'" & RichTextBox1.Text & "', " & TextBox5.Text & ", N'" & ComboBox2.SelectedItem & "', '" & date3 & "', N'" & ComboBox5.SelectedItem & "', N'" & ComboBox1.SelectedItem & "')", cnx)
                cmd1.ExecuteNonQuery()
                vide()
            End If
        Catch ex As Exception
                MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

            End Try
        End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As DialogResult = MessageBox.Show("هل تريدون ان تحينوا هذه الواردة؟ ", "التاكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Try
                Dim updateQuery As String = "UPDATE importations SET "

                If DateTimePicker1.Value <> DateTimePicker1.MinDate Then
                    updateQuery += "[date] = @Date, "
                End If

                If ComboBox3.SelectedIndex <> -1 Then
                    updateQuery += "[type] = @Type, "
                End If

                If ComboBox4.SelectedIndex <> -1 Then
                    updateQuery += "source = @Source, "
                End If

                If Not String.IsNullOrWhiteSpace(TextBox3.Text) Then
                    updateQuery += "num_enreg = @NumEnreg, "
                End If

                If DateTimePicker2.Value <> DateTimePicker2.MinDate Then
                    updateQuery += "date_e = @DateE, "
                End If

                If Not String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                    updateQuery += "sujet = @Sujet, "
                End If

                If Not String.IsNullOrWhiteSpace(TextBox5.Text) Then
                    updateQuery += "num_doc = @NumDoc, "
                End If

                If ComboBox2.SelectedIndex <> -1 Then
                    updateQuery += "pour = @Pour, "
                End If

                If DateTimePicker3.Value <> DateTimePicker3.MinDate Then
                    updateQuery += "delais = @Delais, "
                End If

                If ComboBox5.SelectedIndex <> -1 Then
                    updateQuery += "outil = @Outil, "
                End If

                If ComboBox1.SelectedIndex <> -1 Then
                    updateQuery += "destination = @Destination, "
                End If

                If Not String.IsNullOrWhiteSpace(updateQuery) Then
                    ' Remove the trailing comma and space
                    updateQuery = updateQuery.Trim().TrimEnd(","c, " "c)

                    updateQuery += " WHERE num = @Num"

                    Dim cmd As New SqlCommand(updateQuery, cnx)

                    If DateTimePicker1.Value <> DateTimePicker1.MinDate Then
                        cmd.Parameters.AddWithValue("@Date", DateTimePicker1.Value)
                    End If

                    If ComboBox3.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@Type", ComboBox3.SelectedItem.ToString())
                    End If

                    If ComboBox4.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@Source", ComboBox4.SelectedItem.ToString())
                    End If

                    If Not String.IsNullOrWhiteSpace(TextBox3.Text) Then
                        cmd.Parameters.AddWithValue("@NumEnreg", Integer.Parse(TextBox3.Text))
                    End If

                    If DateTimePicker2.Value <> DateTimePicker2.MinDate Then
                        cmd.Parameters.AddWithValue("@DateE", DateTimePicker2.Value)
                    End If

                    If Not String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                        cmd.Parameters.AddWithValue("@Sujet", RichTextBox1.Text)
                    End If

                    If Not String.IsNullOrWhiteSpace(TextBox5.Text) Then
                        cmd.Parameters.AddWithValue("@NumDoc", Integer.Parse(TextBox5.Text))
                    End If

                    If ComboBox2.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@Pour", ComboBox2.SelectedItem.ToString())
                    End If

                    If DateTimePicker3.Value <> DateTimePicker3.MinDate Then
                        cmd.Parameters.AddWithValue("@Delais", DateTimePicker3.Value)
                    End If

                    If ComboBox5.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@Outil", ComboBox5.SelectedItem.ToString())
                    End If

                    If ComboBox1.SelectedIndex <> -1 Then
                        cmd.Parameters.AddWithValue("@Destination", ComboBox1.SelectedItem.ToString())
                    End If

                    cmd.Parameters.AddWithValue("@Num", Integer.Parse(TextBox2.Text))

                    cmd.ExecuteNonQuery()

                    MessageBox.Show("تمت العملية بنجاح", "عملية ناجحة", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("لم يتم إجراء أي تعديلات", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

            Catch ex As SqlException
                MessageBox.Show("خطأ في قاعدة البيانات: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Catch ex As FormatException
                MessageBox.Show("تنسيق غير صحيح: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

            Catch ex As Exception
                MessageBox.Show("خطأ غير متوقع: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
            Dim cmd As New SqlCommand("SELECT * FROM importations WHERE num > '" & TextBox2.Text & "' ORDER BY num ASC", cnx)
            Dim dtr As SqlDataReader = cmd.ExecuteReader()
            If dtr.Read() Then
                TextBox2.Text = dtr(0)
                DateTimePicker2.Value = dtr(1)
                ComboBox3.SelectedItem = dtr(2)
                ComboBox4.SelectedItem = dtr(3)
                TextBox3.Text = dtr(4)
                DateTimePicker2.Value = dtr(5)
                RichTextBox1.Text = dtr(6)
                TextBox5.Text = dtr(7)
                ComboBox2.SelectedItem = dtr(8)
                DateTimePicker3.Value = dtr(9)
                ComboBox5.SelectedItem = dtr(10)
                ComboBox1.SelectedItem = dtr(11)
            Else
            vide()

        End If
            dtr.Close()
        End Sub

        Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click

            Dim cmd As New SqlCommand("SELECT TOP 1 * FROM importations ORDER BY num DESC", cnx)
            Dim dtr As SqlDataReader = cmd.ExecuteReader()
            If dtr.Read() Then
                TextBox2.Text = dtr(0)
                DateTimePicker2.Value = dtr(1)
                ComboBox3.SelectedItem = dtr(2)
                ComboBox4.SelectedItem = dtr(3)
                TextBox3.Text = dtr(4)
                DateTimePicker2.Value = dtr(5)
                RichTextBox1.Text = dtr(6)
                TextBox5.Text = dtr(7)
                ComboBox2.SelectedItem = dtr(8)
                DateTimePicker3.Value = dtr(9)
                ComboBox5.SelectedItem = dtr(10)
                ComboBox1.SelectedItem = dtr(11)
            End If
            dtr.Close()
        End Sub

        Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
            Dim cmd As New SqlCommand("SELECT TOP 1 * FROM importations ORDER BY num ASC", cnx)
            Dim dtr As SqlDataReader = cmd.ExecuteReader()
            If dtr.Read() Then
                TextBox2.Text = dtr(0)
                DateTimePicker2.Value = dtr(1)
                ComboBox3.SelectedItem = dtr(2)
                ComboBox4.SelectedItem = dtr(3)
                TextBox3.Text = dtr(4)
                DateTimePicker2.Value = dtr(5)
                RichTextBox1.Text = dtr(6)
                TextBox5.Text = dtr(7)
                ComboBox2.SelectedItem = dtr(8)
                DateTimePicker3.Value = dtr(9)
                ComboBox5.SelectedItem = dtr(10)
                ComboBox1.SelectedItem = dtr(11)
            End If
            dtr.Close()
        End Sub

        Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
            Dim cmd As New SqlCommand("SELECT * FROM importations WHERE num < '" & TextBox2.Text & "' ORDER BY num DESC", cnx)
            Dim dtr As SqlDataReader = cmd.ExecuteReader()
            If dtr.Read() Then
                TextBox2.Text = dtr(0)
                DateTimePicker2.Value = dtr(1)
                ComboBox3.SelectedItem = dtr(2)
                ComboBox4.SelectedItem = dtr(3)
                TextBox3.Text = dtr(4)
                DateTimePicker2.Value = dtr(5)
                RichTextBox1.Text = dtr(6)
                TextBox5.Text = dtr(7)
                ComboBox2.SelectedItem = dtr(8)
                DateTimePicker3.Value = dtr(9)
                ComboBox5.SelectedItem = dtr(10)
                ComboBox1.SelectedItem = dtr(11)
            Else
                vide()
            End If
            dtr.Close()
        End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        vide()
    End Sub
End Class

