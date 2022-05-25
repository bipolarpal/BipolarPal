
Imports System.ComponentModel
Imports System.Net
Imports System.Net.Mail

Public Class ResetPass
    Public MailTo As String
    Public Subject As String = "BipolarPal Password Reset"
    Public Body As String


    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        If (TextBox2.Text = "") Then
            Label2.Text = "Email is empty!"
            Label2.ForeColor = Color.Red
            Label2.Visible = True
            Exit Sub

        End If


        If (TextBox1.Text = "" And TextBox1.Visible = True) Then
            Label2.Text = "Reset Code is empty!"
            Label2.ForeColor = Color.Red
            Label2.Visible = True
            Exit Sub

        End If

        If Not RegexUtilities.IsValidEmail(TextBox2.Text) Then
            Label2.Text = "Wrong Email format!"
            Label2.ForeColor = Color.Red
            Label2.Visible = True
            Exit Sub

        End If
        Dim conn As New OleDbConnection
        Dim cmd As New OleDbCommand
        'Dim connString = "Driver={Microsoft Access Driver (*.mdb)};Data Source=BipolarPal.mdb;User ID=Admin;Password=bRa2hAchAveb7iswUthO;"
        Dim connString As String = My.Settings.conString

        '  Dim connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BipolarPal.accdb;Persist Security Info=False;Jet OLEDB:Database Password=bRa2hAchAveb7iswUthO"
        MailTo = TextBox2.Text

        Dim GetCurrentDate As Date = Date.Now.ToString

        If TextBox1.Visible Then
            Dim strCheckResetCode As String = "SELECT resetcode FROM accounts WHERE email=@email AND resetcode=@resetCode AND codeexpiray > @CrrentDate"

            conn.ConnectionString = connString
            cmd.Connection = conn
            conn.Open()
            cmd.CommandText = strCheckResetCode

            cmd.Parameters.AddWithValue("@email", MailTo)
            cmd.Parameters.AddWithValue("@resetCode", TextBox1.Text)
            cmd.Parameters.AddWithValue("@CrrentDate", GetCurrentDate)

            Dim rowCount As Integer = CInt(cmd.ExecuteScalar())
            If rowCount > 0 Then
                Label2.Text = "Reset Code has been verified!"
                Label2.ForeColor = Color.Green
                Label2.Visible = True
                Timer1.Start()

            Else
                Label2.Text = "Wrong or expired Reset Code!"
                Label2.ForeColor = Color.Red
                Label2.Visible = True
                Label10.Visible = True

                TextBox1.Visible = False



            End If
        Else

            Dim strQryCheckLogin As String = "SELECT username, email FROM accounts WHERE email=@email "


            conn.ConnectionString = connString
            cmd.Connection = conn
            conn.Open()
            cmd.CommandText = strQryCheckLogin


            cmd.Parameters.AddWithValue("@email", MailTo)

            'query.ExecuteNonQuery()
            Dim dr = cmd.ExecuteReader()
            Dim EmailExsist As Boolean
            Dim uName As String = ""
            ' Dim NoEmail As Boolean
            If dr.HasRows Then
                While dr.Read()

                    'Dim IsPasswordCorrect As Boolean = VerifyHash(sha512Hash, TextBox3.Text, dr("passwordhash").ToString)
                    If dr("email").ToString = MailTo Then
                        uName = dr("username")
                        EmailExsist = True




                    End If

                End While
                dr.Close()
            Else
                Label2.Text = "Email doesn't exsist!"
                Label2.ForeColor = Color.Red
                Label2.Visible = True
            End If

            If EmailExsist Then
                ' Dim rCode As Integer


                Randomize()

                Dim rCode As Integer = Int((100000 * Rnd()) + 1)

                Dim cExpiray As Date = Date.Now.AddMinutes(30).ToString





                cmd.Parameters.Clear()
                Dim strSaveResetCode As String = "UPDATE accounts SET resetcode=@rCode, codeexpiray=@cExpiray WHERE email=@email"
                cmd.CommandText = strSaveResetCode

                cmd.Parameters.AddWithValue("@rCode", rCode)
                cmd.Parameters.AddWithValue("@cExpiray", cExpiray)
                cmd.Parameters.AddWithValue("@email", MailTo)







                Dim CommandStatus As Integer = cmd.ExecuteNonQuery()


                If CommandStatus > 0 Then


                    PictureBox1.Visible = True
                    PictureBox1.Refresh()


                    Body = "<table align=""center"" cellspacing=""0"" style=""border-collapse:collapse; width:60%""> <tbody> <tr> <td> <hr /> <p><img alt="""" src=""https://bipolarpal.com/maillogo.png"" style=""float:left; height:70px; width:357px"" /></p> <p>&nbsp;</p> <p>&nbsp;</p> <p>&nbsp;</p> <hr /> <p><span style=""background-color:white""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""color:#222222""><span style=""font-size:0.875rem""><span style=""background-color:#ffffff""><span style=""font-family:&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-size:16px""><span style=""font-family:&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""color:#141823""><span style=""font-size:15px"">Hi " & uName & ", </span></span></span></span></span></span></span></span></span></span></span></span></p> <p><span style=""background-color:white""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""color:#222222""><span style=""font-size:0.875rem""><span style=""background-color:#ffffff""><span style=""font-family:&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-size:16px""><span style=""font-size:15px"">We received a request to reset your BipolarPal application password.</span></span></span></span></span></span></span></span></span></span></p> <p><span style=""background-color:white""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""color:#222222""><span style=""font-size:0.875rem""><span style=""background-color:#ffffff""><span style=""font-family:&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-size:16px""><span style=""font-size:15px"">Enter the following password reset&nbsp;code in the BipolarPal Application:</span></span></span></span></span></span></span></span></span></span></p> <table align=""center"" border=""1"" cellpadding=""0"" cellspacing=""0"" style=""border-collapse:collapse; margin-bottom:20px; margin-top:20px; width:100px""> <tbody> <tr> <td style=""text-align:center""><span style=""color:#27ae60""><strong>" & rCode & "</strong></span></td> </tr> </tbody> </table><br/><p>This code will expire in 30 minutes.</p> <p><span style=""background-color:white""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""color:#222222""><span style=""font-size:0.875rem""><span style=""background-color:#ffffff""><span style=""font-family:&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-size:16px""><span style=""font-size:15px""><span style=""color:#333333""><strong>Didn't request this change?</strong></span></span></span></span></span></span></span></span></span></span></span></p> <span style=""background-color:white""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""color:#222222""><span style=""font-size:0.875rem""><span style=""background-color:#ffffff""><span style=""font-family:&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-size:16px""><span style=""font-size:15px"">If you didn't request a new password,&nbsp;<a href=""https://bipolarpal.com/abuse"" target=""_blank"">let us know</a>.&nbsp;</span></span></span></span></span></span></span></span></span></span></td> </tr> <tr> <td> <p>&nbsp;</p> <p><span style=""background-color:white""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""color:#222222""><span style=""font-size:0.875rem""><span style=""background-color:#ffffff""><span style=""font-family:&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""font-family:Roboto,RobotoDraft,Helvetica,Arial,sans-serif""><span style=""font-size:12px""><span style=""font-family:Roboto-Regular,Roboto,-apple-system,BlinkMacSystemFont,&quot;Helvetica Neue&quot;,Helvetica,&quot;Lucida Grande&quot;,tahoma,verdana,arial,sans-serif""><span style=""color:#8a8d91"">This message was sent to&nbsp;<a href=""mailto:" & MailTo & """ target=""_blank"">" & MailTo & "</a>&nbsp;at your request.<br /> <a href=""https://bipolarpal.com/"" target=""_blank"">www.BipolarPal.com</a></span></span></span></span></span></span></span></span></span></span></p> </td> </tr> </tbody> </table>"


                    ' SurroundingSub(email, "BipolarPal Password Reset", sBody)
                    'PictureBox1.Visible = False

                    BackgroundWorker1.RunWorkerAsync()



                    Label2.Text = "Please check your email."
                    Label2.ForeColor = Color.Green
                    Label2.Visible = True
                    Label10.Visible = False
                    TextBox1.Visible = True
                    'Timer1.Start()


                Else
                    Label9.Text = "Sorry! Something went wrong"
                    Label9.Visible = True

                End If



                conn.Close()
            End If
        End If

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Me.Close()

    End Sub
    Dim counter As Integer = 1
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        counter += 1

        If counter = 4 Then
            ChangePassword.Email = TextBox2.Text
            ChangePassword.ResetCode = TextBox1.Text

            ChangePassword.Show()
            Me.Close()

        End If
    End Sub




    Private Sub BackgroundWorker1_DoWork_1(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        Dim mail As New MailMessage()
        Dim SmtpServer As New SmtpClient()
        mail.From = New MailAddress("kazemguru@gmail.com")
        mail.[To].Add(MailTo)
        mail.Subject = Subject
        mail.IsBodyHtml = True

        mail.Body = Body
        SmtpServer.UseDefaultCredentials = False
        Dim NetworkCred As New NetworkCredential("kazemguru@gmail.com", "bjkqbamgaczdhqrm")
        SmtpServer.Credentials = NetworkCred
        SmtpServer.EnableSsl = True
        SmtpServer.Port = 25
        SmtpServer.Host = "smtp.gmail.com"
        SmtpServer.Send(mail)


        If BackgroundWorker1.CancellationPending Then
            e.Cancel = True
        End If
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        PictureBox1.Visible = False

    End Sub
End Class