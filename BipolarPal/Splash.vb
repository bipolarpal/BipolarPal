Public Class Splash


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try


            ProgressBar1.Increment(1)
            If ProgressBar1.Value >= 100 Then
                If My.Settings.isFirstRun Then
                    My.Settings.isFirstRun = False
                    My.Settings.Save()
                    Timer1.Enabled = False
                    SignUp.Show()
                    Me.Close()
                Else
                    Login.Show()
                    Me.Close()
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Splash_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = "©" & Year(Now) & ", Version " & Application.ProductVersion


        Timer1.Start()

    End Sub
End Class