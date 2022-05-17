Imports System.Data.OleDb
Imports System.Net.Mail

Public Class Login


    'Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
    '    Me.Close()

    'End Sub

    'Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    'End Sub

    'Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
    '    ResetPass.Show()
    '    Me.Close()

    'End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click


        If (TextBox2.Text = "") Then
            Label10.Text = "Email is empty!"
            Label10.ForeColor = Color.Red
            Label10.Visible = True
            Exit Sub

        End If


        If (TextBox3.Text = "") Then
            Label10.Text = "Password is empty!"
            Label10.ForeColor = Color.Red
            Label10.Visible = True
            Exit Sub

        End If

        If Not RegexUtilities.IsValidEmail(TextBox2.Text) Then
            Label10.Text = "Wrong Email format!"
            Label10.ForeColor = Color.Red
            Label10.Visible = True
            Exit Sub

        End If
        Dim conn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BipolarPal.accdb;Persist Security Info=False;Jet OLEDB:Database Password=bRa2hAchAveb7iswUthO"

        Dim email As String = TextBox2.Text
        ' Dim salt As String = "temp hash"

        ' Dim pass As String = SHA512Hasher.Base64Hash(TextBox3.Text, salt)
        ' Dim pass As String = Web.Helpers.Crypto.HashPassword(TextBox3.Text)
        ' Dim verifiedPass As String = Web.Helpers.Crypto.VerifyHashedPassword(pass, TextBox3.Text)
        Dim strQryCheckLogin As String = "SELECT email,passwordhash FROM accounts WHERE email=@email "


        conn.ConnectionString = connString
        cmd.Connection = conn
        conn.Open()
        cmd.CommandText = strQryCheckLogin


        cmd.Parameters.AddWithValue("@email", email)

        'query.ExecuteNonQuery()
        Dim dr = cmd.ExecuteReader()
        While dr.Read()
            'Dim IsPasswordCorrect As Boolean = Web.Helpers.Crypto.VerifyHashedPassword(dr("passwordhash").ToString, TextBox3.Text)
            Dim sha512Hash As Security.Cryptography.SHA512 = Security.Cryptography.SHA512.Create()

            Dim IsPasswordCorrect As Boolean = VerifyHash(sha512Hash, TextBox3.Text, dr("passwordhash").ToString)
            If dr("email").ToString = email And IsPasswordCorrect Then
                Main.Show()
                Me.Close()
            Else
                Label10.Text = "Wrong Email or Password!"
                Label10.ForeColor = Color.Red
                Label10.Visible = True

            End If

        End While

        If dr.HasRows = False Then
            Label10.Text = "Please signup first!"
            Label10.ForeColor = Color.Red
            Label10.Visible = True

            Timer1.Start()


        End If


    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Me.Close()

    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        ResetPass.Show()
        Me.Close()

    End Sub
    Dim counter As Integer = 1

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        counter += 1

        If counter = 4 Then
            Timer1.Enabled = False
            SignUp.Show()

            Me.Close()

        End If
    End Sub
End Class