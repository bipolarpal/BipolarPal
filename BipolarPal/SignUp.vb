
Imports System.Data.OleDb





Public Class SignUp
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Me.Close()

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Dim bytes() As Byte = System.Text.Encoding.Unicode.GetBytes("1234")



    End Sub



    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

        If (TextBox1.Text = "") Then
            Label9.Text = "Nick Name is empty!"
            Label9.ForeColor = Color.Red

            Label9.Visible = True
            Exit Sub

        End If
        If (TextBox2.Text = "") Then
            Label9.Text = "Email is empty!"
            Label9.ForeColor = Color.Red
            Label9.Visible = True
            Exit Sub

        End If


        If (TextBox3.Text = "") Then
            Label9.Text = "Password is empty!"
            Label9.ForeColor = Color.Red
            Label9.Visible = True
            Exit Sub

        End If

        If (TextBox4.Text = "") Then
            Label9.Text = "Password verify is empty!"
            Label9.ForeColor = Color.Red
            Label9.Visible = True
            Exit Sub

        End If

        If (Not RegexUtilities.IsValidEmail(TextBox2.Text)) Then
            Label9.Text = "Wrong Email format!"
            Label9.ForeColor = Color.Red
            Label9.Visible = True
            Exit Sub

        End If


        If (Not PassAccepted) Then
            Label9.Text = "Not accepted Password! Weak."
            Label9.ForeColor = Color.Red
            Label9.Visible = True
            Exit Sub

        End If



        Dim email As String = TextBox2.Text








        If (TextBox3.Text = TextBox4.Text) Then
                Dim conn As New OleDbConnection
                Dim cmd As New OleDbCommand

            Dim query As String = "SELECT count(email) FROM accounts WHERE email=@email"


            '  Dim connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=BipolarPal.accdb;Persist Security Info=False;Jet OLEDB:Database Password=bRa2hAchAveb7iswUthO"


            'Dim connString = "Driver={Microsoft Access Driver (*.mdb)};Data Source=BipolarPal.mdb;User ID=Admin;Password=bRa2hAchAveb7iswUthO;"
            Dim connString As String = My.Settings.conString

            conn.ConnectionString = connString
                cmd.Connection = conn
                conn.Open()
                cmd.CommandText = query


                cmd.Parameters.AddWithValue("@email", email)

            'query.ExecuteNonQuery()
            Dim rowCount As Integer = CInt(cmd.ExecuteScalar())
            If rowCount > 0 Then
                Label9.Visible = False

                Label10.Visible = True
                Label11.Visible = True
                Label12.Visible = True

                'Timer1.Start()

                conn.Close()

                Exit Sub

            ElseIf rowCount = 0 Then


                Dim strQuerySave As String = "INSERT Into Accounts (username, email, passwordhash) VALUES (@username, @email, @passwordhash)"

                If (conn.State = ConnectionState.Closed) Then
                    conn.Open()
                End If


                ' Dim HashedPass As String = SHA512Hasher.Base64Hash(Me.TextBox3.Text, salt)
                ' Dim HashedPass As String = Web.Helpers.Crypto.HashPassword(TextBox3.Text)
                Dim HashedPass As String = PasswordHash(TextBox3.Text)
                cmd.Parameters.Clear()


                cmd.CommandText = strQuerySave
                cmd.Parameters.AddWithValue("@username", Me.TextBox1.Text)
                cmd.Parameters.AddWithValue("@email", Me.TextBox2.Text)
                cmd.Parameters.AddWithValue("@passwordhash", HashedPass)





                Dim CommandStatus As Integer = cmd.ExecuteNonQuery()


                If CommandStatus > 0 Then
                    Label9.Text = "Congratulation! Account created."
                    Label9.ForeColor = Color.Green

                    Label9.Visible = True
                    PictureBox1.Visible = True
                Else
                    Label9.Text = "Sorry! Something went wrong"
                    Label9.Visible = True

                End If



                conn.Close()

                End If




                '=========================








                Timer1.Start()
            Else
                Label9.ForeColor = Color.Red

            Label9.Text = "Passwords do NOT match"
            Label9.Visible = True

            End If
        'Catch ex As Exception





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





    Private StrengthWords() As String = {"Very Very Weak", "Very Weak", "Weak", "Better", "Medium", "Strong", "Strongest"}
    Private Sub TextBox3_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyUp
        ProgressBarCustomColor1.Visible = True

        CalculateMeter(TextBox3.Text)
    End Sub
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

    Private Function GetColor(ByVal score As Integer) As Color
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

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        Login.Show()
        Me.Close()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub


End Class
