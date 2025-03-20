Imports System.Data.SqlClient

Public Class Form2
    Dim connectionString As String = "Data Source=MICHAELPC\SQLEXPRESS;Initial Catalog=nmvloginform;Integrated Security=True;Encrypt=False;"

    Private Sub BtnLogin_Click(sender As Object, e As EventArgs) Handles BtnLogin.Click
        Try
            Using con As New SqlConnection(connectionString)
                Dim query As String = "SELECT * FROM login WHERE username = @username AND password = @password"
                Using cmd As New SqlCommand(query, con)
                    cmd.Parameters.AddWithValue("@username", TxtUser.Text)
                    cmd.Parameters.AddWithValue("@password", TxtPass.Text)

                    con.Open()
                    Dim reader As SqlDataReader = cmd.ExecuteReader()

                    If reader.HasRows Then
                        MessageBox.Show("Login successful!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Me.Hide()
                        Dim mainForm As New Form1()
                        mainForm.ShowDialog()
                        Me.Close()
                    Else
                        MessageBox.Show("Invalid username or password.", "Login Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error connecting to database: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub ChkShowPass_CheckedChanged(sender As Object, e As EventArgs) Handles ChkShowPass.CheckedChanged
        If ChkShowPass.Checked Then
            TxtPass.PasswordChar = ControlChars.NullChar
        Else
            TxtPass.PasswordChar = "*"
        End If

    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click
        Application.Exit()
    End Sub

End Class
