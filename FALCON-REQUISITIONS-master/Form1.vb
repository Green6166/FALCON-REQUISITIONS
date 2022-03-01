Imports System.ComponentModel
Imports System.Net.Mail
Imports Microsoft.Office.Interop.Word

Public Class Form1
    Dim dater = TimeOfDay
    Dim from
    Dim jerry
    Dim objdoc
    Dim docstr = Environment.GetEnvironmentVariable("TEMP") 'gets directory locant of the environmental variable %TEMP%

    Function emailbot(email As String) As Action ''assumes the str from listbox 1 is a email in a string variable
        Dim mailer = New Net.Mail.SmtpClient 'starts a smtp client
        If Tim.Enabled = True Then 'if timer is active when emailbot is active it will pause the timer
            Tim.Stop() 'stops timer


        End If
        If email.Contains("@") Then 'filter the jargon as it scans alot of crap aswell
            Try
                Tim.Interval = 4000 'sets time to 4secs
                Dim Smtp_Server As New SmtpClient 'defines new client
                Dim e_mail As New System.Net.Mail.MailMessage() 'starts a new message
                Smtp_Server.UseDefaultCredentials = False 'no default credentials as there are custom details to add to the client
                Smtp_Server.Credentials = New Net.NetworkCredential(TextBox4from.Text, TextBox4authorisation.Text) 'takes the password and username of a microsoft email to use as a smtp bot boi
                Smtp_Server.Port = 587   'normal port 
                Smtp_Server.EnableSsl = True 'protects from tampering
                Smtp_Server.DeliveryMethod = SmtpDeliveryMethod.Network 'specifies delivery method
                Smtp_Server.Host = "smtp.gmail.com" 'this is te domain url for smtpbots. gmail ones do not work correctly they are blocked by google use @live
                Try

                Catch ex As Exception
                    MsgBox(ex.Message & "24") 'on error t says the approximate line number where the code went wrong good for traceback
                End Try


                e_mail = New System.Net.Mail.MailMessage() 'started new message is back in the picture
                e_mail.From = New MailAddress(TextBox4from.Text) 'defines where the mail is from
                e_mail.To.Add(email) 'adds the target of the email from the values given to emailbot
                e_mail.Subject = "requestions" 'subject of the email is important
                e_mail.IsBodyHtml = False 'says that th body of the email is html or not better off false for security reasons
                e_mail.Body = (jerry) 'dumps the concatingnated string of listbox1 o the body to use s traceback or incase the code failes

                Dim fileTXT As String = docstr & "\attachment.pdf" 'the full directory of the file is compiled here
                Dim data As Net.Mail.Attachment = New Net.Mail.Attachment(fileTXT) 'attempts to add document made by kalcus and stowed in c:/ as a attachment
                data.Name = "Requisition.pdf" 'renames file??
                e_mail.Attachments.Add(data) 'attaches file
                Smtp_Server.Send(e_mail) 'sends eamil
                'gives user a notification that email is sent  - used to got pretty annoying nower days its much less intruding on work flow
                'timer for waiting for other email functions
                Tim.Start() 'starts the timer
                NotifyIcon1.Text = "Email Has Been Sent!" 'sets the text to the confirmation message
                'shows confirmation message for 2 seconds or 2000 milliseconds


            Catch ex As Exception
                NotifyIcon1.Text = "There is a problem..." 'sets the text to the error message
                'shows error message for 2 seconds or 2000 milliseconds
                MsgBox("54" & ex.ToString) 'msgbox straight after with more dat
                Tim.Start()
                'this is nothing because there is so many non emails in the code it is inefficent to put a message there everytime it comes across a non email and fails ard 
            End Try
            NotifyIcon1.ShowBalloonTip(2000)
        ElseIf TextBox2.Text.Contains("@") Then 'more filtering
            Tim.Start() 'started the timer


        Else
            NotifyIcon1.Text = "theres no valid emails to target!" 'sets message inregard to no valid emails
            NotifyIcon1.ShowBalloonTip(3000) 'shows message above for 3 seconds
            Tim.Start() 'starts the timer
        End If
    End Function
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox2.Items.Add("btbox") 'adds items to te combobox during load cause im too dumb todo it myself
        ComboBox2.Items.Add("dhl")  'adds items to te combobox during load cause im too dumb todo it myself
        ComboBox2.Items.Add("collection by engineer") 'adds items to te combobox during load cause im too dumb todo it myself
        ComboBox2.Items.Add("alt post service (add name of alt service here)") 'adds items to te combobox during load cause im too dumb todo it myself

        Try
            Dim SR As System.IO.StreamReader = IO.File.OpenText(docstr & "\falcon_requestioins_email.txt") 'opens the email list of people from the last closure of the program
            While SR.Peek <> -1
                TextBox2.Text = (SR.ReadToEnd) ' scans eachline  then trnasferrs eachline to each line on textbox
            End While
            SR.Close()
        Catch ex As Exception
            MsgBox(ex.Message & " 57")
        End Try
        Try
            Dim SR2 As System.IO.StreamReader = IO.File.OpenText(docstr & "\falcon_options_combobox1.txt") 'list of option from the last closure of the program
            While SR2.Peek <> -1
                TextBox3.Text = (SR2.ReadToEnd) ' scans eachline  then transferrs eachline to each line on textbox
            End While
            SR2.Close()
        Catch ex As Exception
            MsgBox(ex.Message & " 66")

        End Try
        Try
            Dim SR3 As System.IO.StreamReader = IO.File.OpenText(docstr & "\falcon_extra_detes.txt") 'list of extra details from the last closure of the program
            While SR3.Peek <> -1
                TextBox1.Text = (SR3.ReadToEnd) ' scans eachline  then trnasferrs eachline to each line on textbox
            End While
            SR3.Close()
        Catch ex As Exception
            MsgBox(ex.Message & " 76")
        End Try

    End Sub


    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            My.Computer.FileSystem.DeleteFile(docstr & "\falcon_requestioins email.txt") 'wipes the following config files so it can apped them
            My.Computer.FileSystem.DeleteFile(docstr & "\falcon_extra_detes.txt") 'wipes the following config files so it can apped them
            My.Computer.FileSystem.DeleteFile(docstr & "\falcon_options_combobox1.txt") 'wipes the following config files so it can apped them
        Catch ex As Exception

        End Try
        Try
            Dim SW2 As IO.StreamWriter = IO.File.CreateText(docstr & "\falcon_options_combobox1.txt") 'once wiped the code takes each line from the following text box and adds it to the representitive configfile as a .txt  

            SW2.Write(TextBox3.Text)
            SW2.Close()


            Dim SW As IO.StreamWriter = IO.File.CreateText(docstr & "\falcon_requestioins_email.txt")  'once wiped the code takes each line from the following text box and adds it to the representitive configfile as a .txt 

            SW.Write(TextBox2.Text)

            SW.Close()

            Dim SW3 As IO.StreamWriter = IO.File.CreateText(docstr & "\falcon_extra_detes.txt")  'once wiped the code takes each line from the following text box and adds it to the representitive configfile as a .txt 

            SW3.Write(TextBox1.Text)

            SW3.Close()
        Catch ex As Exception
            MsgBox("error" & ex.Message) 'catch and try statment to reduce crash risk
        End Try
    End Sub

    Function kalcus() As Action  'kalcus is a misselling of calcus derived from calculate   
        Dim objWord As Microsoft.Office.Interop.Word.Application 'defines MSword as a application variable in itself
        Dim objDoc As Microsoft.Office.Interop.Word.Document ' defines the documents as a variable for defining other element positions
        Dim objTable As Microsoft.Office.Interop.Word.Table 'defines the table element for the document 
        objWord = CreateObject("Word.Application") 'makes a word document
        objWord.Visible = False 'hides it from the user good 4 debuggin thats for damn sure!
        objDoc = objWord.Documents.Add 'no clue what this does
        Dim r As Integer ' R is for ROW! row row row ur boat gently down the stream! 
        Try



            objTable = objDoc.Tables.Add(objDoc.Bookmarks.Item("\endofdoc").Range, 15, 5) 'defines the ranges of the table / axis length
            objTable.Range.ParagraphFormat.SpaceAfter = 6 'pixalated spacing 
            objTable.Cell(1, 1).Range.Text = "TIME"  'sets the first row of data as the names of the boxes
            objTable.Cell(1, 2).Range.Text = "TC.NO:" 'sets the first row of data as the names of the boxes
            objTable.Cell(1, 3).Range.Text = "item" 'sets the first row of data as the names of the boxes
            objTable.Cell(1, 4).Range.Text = "transport method" 'sets the first row of data as the names of the boxes
            objTable.Cell(1, 5).Range.Text = "extra notes/address" 'sets the first row of data as the names of the boxes
            Try
                For r = 2 To 15 ' 2 is there instead of 1 because tables in word are not zero based index and we want the next bottem from the top bar/row
                    objTable.Cell(r, 1).Range.Text = TimeOfDay.ToLocalTime.ToLongTimeString 'time of day but as a string 
                    objTable.Cell(r, 2).Range.Text = ListBox1.Items(3).ToString
                    objTable.Cell(r, 3).Range.Text = (TextBox3.Lines.GetValue(r - 2)) 'gets val from a line in textbox the -2 is to convert it to the correct index base
                    objTable.Cell(r, 4).Range.Text = (ComboBox2.Text) 'uses val from selected text in combobox2
                    objTable.Cell(r, 5).Range.Text = (TextBox1.Lines.GetValue(r - 2)) 'gets val from a line in textbox the -2 is to convert it to the correct index base
                Next
            Catch ex As Exception
                'MsgBox(ex.Message & ex.Source & " 142") 'pulls dump of mistakes ALL THE DAM TIME OMG

            End Try
            objTable.Rows.Item(1).Range.Font.Bold = True 'sets stuf to bold
            objTable.Rows.Item(1).Range.Font.Italic = True 'sets stuff to italic
            objTable.Borders.Shadow = True 'adds border shadow for edginess

        Catch ex As Exception
            'MsgBox(ex.Message & ex.Source & ex.StackTrace & "148" & ex.HelpLink)

        Finally
            Try

                If My.Computer.FileSystem.FileExists((docstr & "\attachment.pdf")) = True Then 'if the old files exists then it will be deleted
                    My.Computer.FileSystem.DeleteFile((docstr & "\attachment.pdf")) 'if the old files exists then it will be deleted
                End If
                objDoc.ExportAsFixedFormat(docstr & "\attachment.pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF) 'exports in pdf format called attachment.pdf in c/:drive
                objDoc.Close(False, False, False) 'CLOSE THE DOCUMENT DOWN FIRST THEN THE APPLICATION
                objWord.Quit(False, False, False) 'closes document with no consideration to saving the file etc basically saying shut up and shut down
            Catch ex As Exception
                MsgBox("finaliser code is bad!157      " & ex.Message) 'if unsuccessful it wont crash but give me understanding
                Try
                    objDoc.ExportAsFixedFormat(docstr & "\attachment.pdf", Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF) 'exports in pdf format called attachment.pdf in c/:drive
                    objDoc.Close(False, False, False)
                    objWord.Quit(False, False, False) 'attempts to close document and application regardless of outcome of exportation
                    Me.Close()
                Catch except As Exception
                    MsgBox("cant close properly i advise a restart to flush ram of hidden word applications " & except.Message) 'worst case senario if it cant even close
                    Me.Close()
                End Try
            End Try
        End Try
    End Function
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        Dim emailslist = TextBox2.Text
        ListBox1.Items.Add(TimeOfDay.ToString) 'time of day to string not complex
        ListBox1.Items.Add(TextBox3.Text) 'adds text from textbox 3 as a item  NOTE: not line by line as in the whole text of the box disregarding lines are extrapolated
        ListBox1.Items.Add(TextBox1.Text) 'same as textbox 3
        ListBox1.Items.Add("TC NO." & NumericUpDown1.Value) 'adds the unique tc number 
        ListBox1.Items.Add("" + ComboBox2.Text) ' takes the selected option for combobox and adds it to the listbox
        Dim i = 0

        For Each item In ListBox1.Items
            jerry = (jerry & "

" & item) 'concatenates data from listbox1
        Next
        While TextBox2.Lines.Count > i
            ListBox1.Items.Add(TextBox2.Lines.ElementAt(i))
            i = i + 1
        End While
        'If My.Computer.FileSystem.FileExists(("C:\attachment.pdf")) = True Then 'if the old files exists then it will be deleted
        'My.Computer.FileSystem.DeleteFile(("C:\attachment.pdf")) 'if the old files exists then it will be deleted
        'End If
        kalcus() 'Kalcus document assembler function is called in
        While My.Computer.FileSystem.FileExists((docstr & "\attachment.pdf")) = False 'waits for kalcus to finish by making the file exist
            '// make the buttone wait for kalkus to complete the file And then call the email function bot
        End While

        For Each item In ListBox1.Items 'spamms the emailbot function to all items in listbox1  
            Try


                emailbot(item) 'invokes emailbot function but in this context.  repeatedly in the "for each" statment






            Catch EX As Exception
                MsgBox("211" & EX.Message)
            End Try

        Next



    End Sub
    Private Sub tim_tick(sender As Object, e As EventArgs) Handles Tim.Tick
        Try
            My.Computer.FileSystem.DeleteFile(docstr & "\falcon_requestioins email.txt") 'wipes the following config files so it can append them
            My.Computer.FileSystem.DeleteFile(docstr & "\falcon_extra_detes.txt") 'wipes the following config files so it can append them
            My.Computer.FileSystem.DeleteFile(docstr & "\falcon_options_combobox1.txt") 'wipes the following config files so it can append them
        Catch ex As Exception

        End Try
        Try
            Dim SW2 As IO.StreamWriter = IO.File.CreateText(docstr & "\falcon_options_combobox1.txt") 'once wiped the code takes each line from the following text box and adds it to the representitive configfile as a .txt  

            SW2.Write(TextBox3.Text)
            SW2.Close()


            Dim SW As IO.StreamWriter = IO.File.CreateText(docstr & "\falcon_requestioins_email.txt")  'once wiped the code takes each line from the following text box and adds it to the representitive configfile as a .txt 

            SW.Write(TextBox2.Text)

            SW.Close()

            Dim SW3 As IO.StreamWriter = IO.File.CreateText(docstr & "\falcon_extra_detes.txt")  'once wiped the code takes each line from the following text box and adds it to the representitive configfile as a .txt 

            SW3.Write(TextBox1.Text)

            SW3.Close()
        Catch ex As Exception
            MsgBox("error" & ex.Message) 'catch and try statment to reduce crash risk
        End Try
        System.Windows.Forms.Application.Restart()
    End Sub


End Class
