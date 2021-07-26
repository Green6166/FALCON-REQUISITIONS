Imports System.ComponentModel
Imports System.Net.Mail
Imports Microsoft.Office.Interop.Word
Public Class Form1
    Dim dater = TimeOfDay
    Dim from
    Dim jerry
    Dim objdoc
    Function emailbot(email As String) As Action
        Dim mailer = New Net.Mail.SmtpClient

        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New System.Net.Mail.MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential(TextBox4from.Text, TextBox4authorisation.Text)
            Smtp_Server.Port = 587
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "smtp.live.com"
            Try
                e_mail.Attachments.Add(New Net.Mail.Attachment("C:\attachment.pdf"))

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            e_mail = New System.Net.Mail.MailMessage()
            e_mail.From = New MailAddress(TextBox4from.Text)
            e_mail.To.Add(email)
            e_mail.Subject = "requestions prototype"
            e_mail.IsBodyHtml = False
            e_mail.Body = (jerry)
            Smtp_Server.Send(e_mail)
            MsgBox("Mail Sent")


        Catch ex As Exception


        End Try

    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox2.Items.Add("btbox")
        ComboBox2.Items.Add("dhl")
        ComboBox2.Items.Add("collection by engineer")
        ComboBox2.Items.Add("alt post service (add name of alt service here)")

        Try
            Dim SR As System.IO.StreamReader = IO.File.OpenText("C:\falcon_requestioins_email.txt")
            While SR.Peek <> -1
                TextBox2.Text = (SR.ReadToEnd)
            End While
            SR.Close()
        Catch

        End Try
        Try
            Dim SR2 As System.IO.StreamReader = IO.File.OpenText("C:\falcon_options_combobox1.txt")
            While SR2.Peek <> -1
                TextBox3.Text = (SR2.ReadToEnd)
            End While
            SR2.Close()
        Catch

        End Try
        Try
            Dim SR3 As System.IO.StreamReader = IO.File.OpenText("C:\falcon_extra_detes.txt")
            While SR3.Peek <> -1
                TextBox1.Text = (SR3.ReadToEnd)
            End While
            SR3.Close()
        Catch
        End Try

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            My.Computer.FileSystem.DeleteFile("C:\falcon_requestioins email.txt")
            My.Computer.FileSystem.DeleteFile("C:\falcon_extra_detes.txt")
            My.Computer.FileSystem.DeleteFile("C:\falcon_options_combobox1.txt")
        Catch
        End Try
        Dim SW2 As IO.StreamWriter = IO.File.CreateText("C:\falcon_options_combobox1.txt")
        Try
            SW2.Write(TextBox3.Text)
            SW2.Close()


            Dim SW As IO.StreamWriter = IO.File.CreateText("C:\falcon_requestioins_email.txt")

            SW.Write(TextBox2.Text)

            SW.Close()

            Dim SW3 As IO.StreamWriter = IO.File.CreateText("C:\falcon_extra_detes.txt")

            SW3.Write(TextBox1.Text)

            SW3.Close()
        Catch ex As Exception
            MsgBox("error" & ex.Message)
        End Try
    End Sub

    Function kalcus() As Action
        Dim objWord As Microsoft.Office.Interop.Word.Application
        Dim objDoc As Microsoft.Office.Interop.Word.Document
        Dim objTable As Microsoft.Office.Interop.Word.Table
        objWord = CreateObject("Word.Application")
        objWord.Visible = True
        objDoc = objWord.Documents.Add
        Dim r As Integer, c As Integer
        Try





            objTable = objDoc.Tables.Add(objDoc.Bookmarks.Item("\endofdoc").Range, 15, 5)
            objTable.Range.ParagraphFormat.SpaceAfter = 6
            objTable.Cell(1, 1).Range.Text = "TIME"
            objTable.Cell(1, 2).Range.Text = "TC.NO:"
            objTable.Cell(1, 3).Range.Text = "item"
            objTable.Cell(1, 4).Range.Text = "transport method"
            objTable.Cell(1, 5).Range.Text = "extra notes/address"
            Try
                For r = 2 To 15
                    objTable.Cell(r, 1).Range.Text = TimeOfDay.ToLocalTime.ToLongTimeString
                    objTable.Cell(r, 2).Range.Text = ListBox1.Items(3).ToString
                    objTable.Cell(r, 3).Range.Text = (TextBox3.Lines.GetValue(r - 2))
                    objTable.Cell(r, 4).Range.Text = (ComboBox2.Text)
                    objTable.Cell(r, 5).Range.Text = (TextBox1.Lines.GetValue(r - 2))
                Next
            Catch

            End Try
            objTable.Rows.Item(1).Range.Font.Bold = True
            objTable.Rows.Item(1).Range.Font.Italic = True
            objTable.Borders.Shadow = True



        Catch ex As Exception
            MsgBox(ex.Message & ex.Source & "
" & ex.HelpLink)
        Finally
            Try
                ' Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF used in test

                objdoc.ExportAsFixedFormat("C:\attachment.pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF)
                objdoc.Close(False)
            Catch ex As Exception
                MsgBox("finaliser code is bad!      " & ex.Message)
                Try
                    objdoc.close()
                Catch
                    MsgBox("your really in the shnitzel if you get this error contact jgree6661@gmail.com your computer should be ok unless something tampers with my code")
                    MsgBox("id like to reasure you that the program has simply failed you your rig is in NO danger but im more concered how you got this far.")
                End Try
            End Try
            End Try
    End Function
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        ListBox1.Items.Remove(ListBox1.Items)

        Dim emailslist = TextBox2.Text
        ListBox1.Items.Add(TimeOfDay.ToString)
        ListBox1.Items.Add(TextBox3.Text)
        ListBox1.Items.Add(TextBox1.Text)
        ListBox1.Items.Add("TC NO." & NumericUpDown1.Value)
        ListBox1.Items.Add("" + ComboBox2.Text)
        Dim i = 0

        For Each item In ListBox1.Items
            jerry = (jerry & "

" & item)
        Next
        While TextBox2.Lines.Count > i
            ListBox1.Items.Add(TextBox2.Lines.ElementAt(i))
            i = i + 1
        End While
        If My.Computer.FileSystem.FileExists(("C:\attachment.pdf")) = True Then
            My.Computer.FileSystem.DeleteFile(("C:\attachment.pdf"))
        End If
        kalcus()
        While My.Computer.FileSystem.FileExists(("C:\attachment.pdf")) = False
            '// make the buttone wait for kalkus to complete the file And then call the email function bot
        End While
        For Each item In ListBox1.Items
                Try


                    emailbot(item)
                Catch
                End Try

        Next



    End Sub


End Class
