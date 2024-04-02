Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form3

    Private Sub vide()
        TextBox2.Text = ""
        DateTimePicker3.Value = DateTime.Now
        TextBox1.Text = ""
        ComboBox5.SelectedItem = ""

        ComboBox1.SelectedItem = ""

        ComboBox3.SelectedItem = ""
        ComboBox4.SelectedItem = ""
        TextBox5.Text = ""
        RichTextBox1.Text = ""
        ComboBox5.SelectedItem = ""
    End Sub


    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If

    End Sub



    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If

    End Sub



    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click

        Dim source As String = TextBox1.Text

        Try
            If String.IsNullOrWhiteSpace(TextBox2.Text) Or String.IsNullOrWhiteSpace(TextBox1.Text) Or String.IsNullOrWhiteSpace(TextBox5.Text) Or ComboBox5.SelectedItem Is Nothing Or ComboBox3.SelectedItem Is Nothing Or ComboBox4.SelectedItem Is Nothing Or String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                Throw New Exception("يجب ان تملأ جميع الحقول")
            Else
                Dim cmd As New SqlCommand("SELECT num_enreg FROM exportations WHERE num_enreg = " & TextBox2.Text, cnx)
                Dim dtr As SqlDataReader = cmd.ExecuteReader()
                If dtr.Read() Then
                    dtr.Close()
                    Throw New Exception("الرقم الذي ادخلتموه يوجد")
                End If
                dtr.Close()

                Dim datee As String = DateTimePicker3.Value.ToString("yyyy-MM-dd")

                Dim cmd1 As New SqlCommand("INSERT INTO exportations (num_enreg, date_enreg, source, outil, type, destination, doc, sujet, des) VALUES (" & TextBox2.Text & ", '" & datee & "', " & TextBox1.Text & ", N'" & ComboBox5.SelectedItem & "', N'" & ComboBox3.SelectedItem & "', N'" & ComboBox4.SelectedItem & "', " & TextBox5.Text & ", N'" & RichTextBox1.Text & "', N'" & ComboBox1.SelectedItem & "')", cnx)
                cmd1.ExecuteNonQuery()

                MessageBox.Show("تمت العملية بنجاح", "عملية ناجحة", MessageBoxButtons.OK, MessageBoxIcon.Information)
                vide()
                Dim cmdCheck As New SqlCommand("SELECT COUNT(*) FROM importations WHERE num = '" & source & "'", cnx)
                Dim count As Integer = Convert.ToInt32(cmdCheck.ExecuteScalar())
                If count > 0 Then
                    Dim dateEnreg As DateTime = DateTimePicker3.Value.ToString("yyyy-MM-dd")
                    MessageBox.Show("document répondu le " & dateEnreg.ToString(), "Succès", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' Mettre à jour le champ "wd" avec la chaîne "done"
                    Dim cmdUpdate As New SqlCommand("UPDATE exportations SET wd = N'منجز' WHERE source = @source", cnx)
                    cmdUpdate.Parameters.AddWithValue("@source", source)
                    cmdUpdate.ExecuteNonQuery()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try



    End Sub

    Private Sub ComboBox3_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sourcee As String = TextBox1.Text
        Dim result As DialogResult = MessageBox.Show("هل تريدون ان تحينوا هذه الصادرة؟ ", "التاكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            If result = DialogResult.Yes Then
                Try
                    Dim updateQuery As String = "UPDATE exportations SET "

                    If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
                        updateQuery += "source = @Source, "
                    End If

                    If ComboBox5.SelectedIndex <> -1 Then
                        updateQuery += "outil = @Outil, "
                    End If

                    If ComboBox1.SelectedIndex <> -1 Then
                        updateQuery += "type = @Type, "
                    End If

                    If Not String.IsNullOrWhiteSpace(TextBox5.Text) Then
                        updateQuery += "doc = @Doc, "
                    End If

                    If Not String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                        updateQuery += "sujet = @Sujet, "
                    End If

                    If ComboBox1.SelectedIndex <> -1 Then
                        updateQuery += "des = @Destination, "
                    End If

                    If ComboBox1.SelectedIndex <> -1 Then
                        updateQuery += "destination = @Destination, "
                    End If

                    If Not String.IsNullOrWhiteSpace(updateQuery) Then
                        ' Remove the trailing comma and space
                        updateQuery = updateQuery.Trim().TrimEnd(","c, " "c)

                        updateQuery += " WHERE num_enreg = @NumEnreg"

                        Dim cmd As New SqlCommand(updateQuery, cnx)

                        If Not String.IsNullOrWhiteSpace(TextBox1.Text) Then
                            cmd.Parameters.AddWithValue("@Source", TextBox1.Text)
                        End If

                        If ComboBox5.SelectedIndex <> -1 Then
                            cmd.Parameters.AddWithValue("@Outil", ComboBox5.SelectedItem.ToString())
                        End If

                        If ComboBox1.SelectedIndex <> -1 Then
                            cmd.Parameters.AddWithValue("@Type", ComboBox1.SelectedItem.ToString())
                        End If

                        If Not String.IsNullOrWhiteSpace(TextBox5.Text) Then
                            cmd.Parameters.AddWithValue("@Doc", TextBox5.Text)
                        End If

                        If Not String.IsNullOrWhiteSpace(RichTextBox1.Text) Then
                            cmd.Parameters.AddWithValue("@Sujet", RichTextBox1.Text)
                        End If

                        If ComboBox1.SelectedIndex <> -1 Then
                            cmd.Parameters.AddWithValue("@Destination", ComboBox1.SelectedItem.ToString())
                        End If

                        cmd.Parameters.AddWithValue("@NumEnreg", TextBox2.Text)

                        cmd.ExecuteNonQuery()

                        MessageBox.Show("تمت العملية بنجاح", "عملية ناجحة", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("لم يتم إجراء أي تعديلات", "تنبيه", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Catch ex As SqlException
                    MessageBox.Show("خطأ في قاعدة البيانات: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Catch ex As Exception
                    MessageBox.Show("خطأ غير متوقع: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If


        End If




    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Hide()
        Form5.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
        Form1.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim cmd As New SqlCommand("SELECT * FROM exportations WHERE num_enreg < '" & TextBox2.Text & "' ORDER BY num_enreg DESC", cnx)
        Dim dtr As SqlDataReader = cmd.ExecuteReader()
        If dtr.Read() Then
            TextBox2.Text = dtr(0)
            DateTimePicker3.Value = If(Not Convert.IsDBNull(dtr(1)), CDate(dtr(1)), DateTimePicker3.MinDate)
            TextBox1.Text = If(Not Convert.IsDBNull(dtr(2)), CStr(dtr(2)), String.Empty)
            ComboBox5.SelectedItem = If(Not Convert.IsDBNull(dtr(3)), CStr(dtr(3)), Nothing)
            ComboBox3.SelectedItem = If(Not Convert.IsDBNull(dtr(4)), CStr(dtr(4)), Nothing)
            ComboBox4.SelectedItem = If(Not Convert.IsDBNull(dtr(5)), CStr(dtr(5)), Nothing)
            TextBox5.Text = If(Not Convert.IsDBNull(dtr(6)), CStr(dtr(6)), String.Empty)
            RichTextBox1.Text = If(Not Convert.IsDBNull(dtr(7)), CStr(dtr(7)), String.Empty)
            ComboBox1.SelectedItem = If(Not Convert.IsDBNull(dtr(8)), CStr(dtr(8)), Nothing)
        Else
            vide()
        End If
        dtr.Close()
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim cmd As New SqlCommand("SELECT TOP 1 * FROM exportations ORDER BY num_enreg DESC", cnx)
        Dim dtr As SqlDataReader = cmd.ExecuteReader()
        If dtr.Read() Then
            TextBox2.Text = dtr(0)
            DateTimePicker3.Value = If(Not Convert.IsDBNull(dtr(1)), CDate(dtr(1)), DateTimePicker3.MinDate)
            TextBox1.Text = If(Not Convert.IsDBNull(dtr(2)), CStr(dtr(2)), String.Empty)
            ComboBox5.SelectedItem = If(Not Convert.IsDBNull(dtr(3)), CStr(dtr(3)), Nothing)
            ComboBox3.SelectedItem = If(Not Convert.IsDBNull(dtr(4)), CStr(dtr(4)), Nothing)
            ComboBox4.SelectedItem = If(Not Convert.IsDBNull(dtr(5)), CStr(dtr(5)), Nothing)
            TextBox5.Text = If(Not Convert.IsDBNull(dtr(6)), CStr(dtr(6)), String.Empty)
            RichTextBox1.Text = If(Not Convert.IsDBNull(dtr(7)), CStr(dtr(7)), String.Empty)
            ComboBox1.SelectedItem = If(Not Convert.IsDBNull(dtr(8)), CStr(dtr(8)), Nothing)
        End If
        dtr.Close()
    End Sub



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim cmd As New SqlCommand("SELECT TOP 1 * FROM exportations ORDER BY num_enreg ASC", cnx)
        Dim dtr As SqlDataReader = cmd.ExecuteReader()
        If dtr.Read() Then
            TextBox2.Text = dtr(0)
            DateTimePicker3.Value = If(Not Convert.IsDBNull(dtr(1)), CDate(dtr(1)), DateTimePicker3.MinDate)
            TextBox1.Text = If(Not Convert.IsDBNull(dtr(2)), CStr(dtr(2)), String.Empty)
            ComboBox5.SelectedItem = If(Not Convert.IsDBNull(dtr(3)), CStr(dtr(3)), Nothing)
            ComboBox3.SelectedItem = If(Not Convert.IsDBNull(dtr(4)), CStr(dtr(4)), Nothing)
            ComboBox4.SelectedItem = If(Not Convert.IsDBNull(dtr(5)), CStr(dtr(5)), Nothing)
            TextBox5.Text = If(Not Convert.IsDBNull(dtr(6)), CStr(dtr(6)), String.Empty)
            RichTextBox1.Text = If(Not Convert.IsDBNull(dtr(7)), CStr(dtr(7)), String.Empty)
            ComboBox1.SelectedItem = If(Not Convert.IsDBNull(dtr(8)), CStr(dtr(8)), Nothing)
        End If
        dtr.Close()
    End Sub






    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim cmd As New SqlCommand("SELECT * FROM exportations WHERE num_enreg >'" & TextBox2.Text & "' ORDER BY num_enreg ASC", cnx)
        Dim dtr As SqlDataReader = cmd.ExecuteReader()
        If dtr.Read() Then
            TextBox2.Text = dtr(0)
            DateTimePicker3.Value = If(Not Convert.IsDBNull(dtr(1)), CDate(dtr(1)), DateTimePicker3.MinDate)
            TextBox1.Text = If(Not Convert.IsDBNull(dtr(2)), CStr(dtr(2)), String.Empty)
            ComboBox5.SelectedItem = If(Not Convert.IsDBNull(dtr(3)), CStr(dtr(3)), Nothing)
            ComboBox3.SelectedItem = If(Not Convert.IsDBNull(dtr(4)), CStr(dtr(4)), Nothing)
            ComboBox4.SelectedItem = If(Not Convert.IsDBNull(dtr(5)), CStr(dtr(5)), Nothing)
            TextBox5.Text = If(Not Convert.IsDBNull(dtr(6)), CStr(dtr(6)), String.Empty)
            RichTextBox1.Text = If(Not Convert.IsDBNull(dtr(7)), CStr(dtr(7)), String.Empty)
            ComboBox1.SelectedItem = If(Not Convert.IsDBNull(dtr(8)), CStr(dtr(8)), Nothing)
        Else
            vide()
        End If
        dtr.Close()
    End Sub
    Private Sub Form3_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        vide()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 1 Then
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
        ElseIf ComboBox1.SelectedIndex = 2 Then
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

        ElseIf ComboBox1.SelectedIndex = 3 Then
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
        ElseIf ComboBox1.SelectedIndex = 4 Then
            ComboBox4.Items.Add("")
            ComboBox4.Items.Clear()
            ComboBox4.Items.Add("خارجية")
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim cmd1 As New SqlCommand("UPDATE exportations SET date_enreg = NULL, source = NULL, outil = NULL, type = NULL, destination = NULL, doc = NULL, sujet = NULL, des = NULL WHERE num_enreg = " & TextBox2.Text, cnx)
            cmd1.ExecuteNonQuery()

            MessageBox.Show("تمت العملية بنجاح", "عملية ناجحة", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As SqlException
            MessageBox.Show("خطأ في قاعدة البيانات: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As FormatException
            MessageBox.Show("تنسيق غير صحيح: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Catch ex As Exception
            MessageBox.Show("خطأ غير متوقع: " & ex.Message, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class