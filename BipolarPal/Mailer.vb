Imports System.Net
Imports System.Net.Mail

Module Mailer


    Public Sub SurroundingSub(ByVal mailTo As String, ByVal Subject As String, ByVal Body As String)
        Dim mail As New MailMessage()
        Dim SmtpServer As New SmtpClient()
        mail.From = New MailAddress("kazemguru@gmail.com")
        mail.[To].Add(mailTo)
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
    End Sub



End Module
