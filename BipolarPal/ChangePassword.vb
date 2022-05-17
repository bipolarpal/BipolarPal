Imports System.Data.OleDb

Public Class ChangePassword
    Private _Email As String = ""
    Friend Property Email As String
        Get
            Return _Email
        End Get
        Set(ByVal value As String)
            _Email = value
        End Set
    End Property
    Private _ResetCode As String = ""
    Friend Property ResetCode As String
        Get
            Return _ResetCode
        End Get
        Set(ByVal value As String)
            _ResetCode = value
        End Set
    End Property

    Public Sub DoChangePWD(Email As String, Resetcode As String)

    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Me.Close()

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        ProgressBarCustomColor1.Visible = True

        CalculateMeter(TextBox2.Text)
    End Sub

    Private ReadOnly StrengthWords() As String = {"Very Very Weak", "Very Weak", "Weak", "Better", "Medium", "Strong", "Strongest"}
    Dim PassAccepted As Boolean

    Private Sub CalculateMeter(pass As String)
        Dim score As Integer

        Dim password As String = pass
        If (password.Length > 6) Then score += 1 'Length more than 6
        If System.Text.RegularExpressions.Regex.IsMatch(password, "[a-z]") And System.Text.RegularExpressions.Regex.IsMatch(password, "[A-Z]") Then
            score += 1 'upper and lower case
        End If
        If System.Text.RegularExpressions.Regex.IsMatch(password, "\d+") Then
            score += 1 'number
        End If


        If System.Text.RegularExpressions.Regex.IsMatch(password, "[!,@,#,$,%,^,&,*,?,_,~,-,/""]") Then
            score += 1 'special character
        End If

        If (password.Length >= 10) Then score += 1 'length more than 9 
        If (password.Length > 15) Then score += 1 'length more than 15 
        ProgressBarCustomColor1.Value = score / 6 * 100 'finding percentage to increase 
        'Label10.Width = 50 * score 'label width is not auto so seeting it to show color amount 
        'Label10.Text = StrengthWords(score) 'Getting strength word from string array declarred above 
        'Label10.TextAlign = ContentAlignment.MiddleCenter 'alignning to center can be done one time in design 
        'Label10.BackColor = GetColor(score) 'Getting color and setting 
        ProgressBarCustomColor1.ForeColor = GetColor(score) 'does not work unless you disable Visual Styles from application properties 


        If score = 3 Then
            PassAccepted = True

        End If



    End Sub
    Shared Function GetColor(ByVal score As Integer) As Color
        Select Case score
            Case 0
                Return Color.Red
            Case 1
                Return Color.Red
            Case 2
                Return Color.Purple
            Case 3
                Return Color.LightGreen
            Case 4
                Return Color.MediumSeaGreen
            Case 5
                Return Color.Green
            Case 6
                Return Color.DarkGreen

        End Select
    End Function

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        If (TextBox2.Text = "") Then
            Label2.Text = "Password is empty!"
            Label2.ForeColor = Color.Red
            Label2.Visible = True
            Exit Sub

        End If


        If (TextBox1.Text = "") Then
            Label2.Text = "Password Confirmation is empty!"
            Label2.ForeColor = Color.Red
            Label2.Visible = True
            Exit Sub

        End If


        Dim hPwd As String = PasswordHash(TextBox2.Text)
        Dim GetCurrentDate As Date = Date.Now.ToString


        Dim conn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BipolarPal.accdb;Persist Security Info=False;Jet OLEDB:Database Password=bRa2hAchAveb7iswUthO"

        Dim strUpdatePWD As String = "UPDATE accounts SET passwordhash = @hPwd WHERE email = @mail AND codeexpiray > @CurrentTime AND resetcode = @resetcode"


        conn.ConnectionString = connString
        cmd.Connection = conn
        conn.Open()
        cmd.CommandText = strUpdatePWD


        cmd.Parameters.AddWithValue("@hPwd", hPwd)
        cmd.Parameters.AddWithValue("@email", Email)
        cmd.Parameters.AddWithValue("@CurrentTime", GetCurrentDate)
        cmd.Parameters.AddWithValue("@resetcode", ResetCode)


        Dim CommandStatus As Integer = cmd.ExecuteNonQuery()


        If CommandStatus > 0 Then
            Label2.Text = "Password has been updated!"


            Label2.ForeColor = Color.Green

            Label2.Visible = True
            PictureBox1.Visible = True
            Timer1.Start()

        Else
            Label2.Text = "Sorry! Something went wrong"
            Label2.ForeColor = Color.Red

            Label2.Visible = True

        End If



        conn.Close()



    End Sub
    Dim counter As Integer = 1


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        counter = counter + 1

        If counter = 4 Then
            Timer1.Enabled = False

            Login.Show()
            Me.Close()

        End If
    End Sub
End Class