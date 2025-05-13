Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1

    Private isOpenEye As Boolean = True
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadEmailToComboBox()
        Guna2TextBox1.Focus()
        Panel4.Visible = False
        Panel4.Visible = False
        Me.KeyPreview = True
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim username As String = Guna2TextBox1.Text.Trim()
        Dim password As String = Guna2TextBox2.Text.Trim()
        Dim fullName As String = String.Empty
        Dim isInactive As Boolean = False

        If IsUserLockedOut(username) Then
            Return
        End If

        If AuthenticateUser(username, password, fullName, isInactive) Then
            ' Reset failed attempts after successful login
            ResetFailedAttempts(username)

            Dim role As String = GetUserRole(username)
            userFullName = fullName
            userRole = role

            MessageBox.Show("Login successful.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Guna2TextBox1.Clear()
            Guna2TextBox2.Clear()

            Dim form2 As New Form2(userRole, fullName)
            Me.Hide()
            form2.Show()
            Clear()
        ElseIf isInactive Then
            MessageBox.Show("Your account is inactive. Please contact the administrator.", "Account Inactive", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Clear()
        Else
            ' Invalid login
            IncrementFailedAttempts(username)
            MessageBox.Show("Invalid username or password.", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Clear()
        End If
    End Sub


    Public Function AuthenticateUser(username As String, password As String, ByRef fullName As String, ByRef isInactive As Boolean) As Boolean
        ' SQL query to validate user and retrieve full name and status
        Dim query As String = "
    SELECT FullName, Status FROM UserAccountTb 
    WHERE 
        UserName COLLATE Latin1_General_CS_AS = @Username AND 
        Password COLLATE Latin1_General_CS_AS = @Password"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Username", username)
                cmd.Parameters.AddWithValue("@Password", password)

                con.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()

                If reader.HasRows Then
                    ' Read the user details
                    While reader.Read()
                        fullName = reader("FullName").ToString()

                        ' Check if the user status is INACTIVE
                        If reader("Status").ToString().ToUpper() = "INACTIVE" Then
                            isInactive = True
                            Return False ' Inactive users cannot log in
                        End If
                    End While
                    Return True ' User found and is active
                Else
                    Return False ' Invalid username or password
                End If
            End Using
        End Using
    End Function

    Public Sub IncrementFailedAttempts(username As String)
        Dim query As String = "
        UPDATE UserAccountTb
        SET 
            FailedAttempts = FailedAttempts + 1,
            LockoutTime = CASE 
                WHEN FailedAttempts + 1 >= 5 THEN DATEADD(MINUTE, 1, GETDATE())
                ELSE NULL
            END
        WHERE UserName = @Username"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Username", username)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Sub ResetFailedAttempts(username As String)
        Dim query As String = "UPDATE UserAccountTb SET FailedAttempts = 0, LockoutTime = NULL WHERE UserName = @Username"
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Username", username)
                con.Open()
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Function IsUserLockedOut(username As String) As Boolean
        Dim query As String = "SELECT FailedAttempts, LockoutTime FROM UserAccountTb WHERE UserName = @Username"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Username", username)
                con.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim lockoutObj = reader("LockoutTime")
                        If Not IsDBNull(lockoutObj) Then
                            Dim lockoutTime As DateTime = Convert.ToDateTime(lockoutObj)
                            If DateTime.Now < lockoutTime Then
                                ' First show generic lockout message
                                MessageBox.Show("Too many failed attempts.", "Login Locked", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                                ' Then show specific remaining time
                                Dim remainingTime As TimeSpan = lockoutTime - DateTime.Now
                                Dim lockoutMessage As String = $"Please wait {remainingTime.Minutes:D2}:{remainingTime.Seconds:D2} before trying again."
                                MessageBox.Show(lockoutMessage, "Login Locked", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                Return True
                            End If
                        End If
                    End If
                End Using
            End Using
        End Using

        Return False
    End Function




    Public Function GetUserRole(username As String) As String
        Dim query As String = "SELECT Role FROM UserAccountTb WHERE UserName = @UserName"
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@UserName", username)
                con.Open()
                Dim role As String = Convert.ToString(cmd.ExecuteScalar())
                Return role
            End Using
        End Using
    End Function


    ' Clear button
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Clear()
    End Sub

    ' Password recovery button (using email)
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        ' Username from TextBox4
        Dim email As String = Guna2ComboBox1.Text.Trim()      ' Email from Guna2ComboBox1
        Dim appPassword As String = Guna2TextBox3.Text.Trim() ' App password from Guna2TextBox3

        If SendExistingPassword(email, appPassword) Then
            ClearEmail()

        Else
            MessageBox.Show("Please provide valid details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ClearEmail()

        End If
    End Sub

    ' Function to retrieve the password from the database and send it via email
    Private Function SendExistingPassword(email As String, appPassword As String) As Boolean
        Dim query As String = "SELECT Password FROM UserAccountTb WHERE Email = @Email"
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)

                cmd.Parameters.AddWithValue("@Email", email)
                con.Open()
                Dim password As Object = cmd.ExecuteScalar()

                If password IsNot Nothing Then
                    ' Send email if password exists
                    SendEmail(email, "Your Password", "Your password is: " & password.ToString(), appPassword)
                    Return True
                Else
                    Return False
                End If
            End Using
        End Using
    End Function


    Private Sub SendEmail(toEmail As String, subject As String, body As String, appPassword As String)

        Dim smtpClient As New SmtpClient("smtp.gmail.com", 587) With {
            .Credentials = New Net.NetworkCredential(toEmail, appPassword),
            .EnableSsl = True
        }


        Dim mailMessage As New MailMessage() With {
            .From = New MailAddress(toEmail),
            .Subject = subject,
            .Body = body,
            .IsBodyHtml = False
        }

        mailMessage.To.Add(toEmail)

        Try
            smtpClient.Send(mailMessage)
            MessageBox.Show("Password sent successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Please provide valid details", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' Toggle password visibility
    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        Guna2TextBox2.PasswordChar = If(Guna2TextBox2.PasswordChar = "*"c, ControlChars.NullChar, "*"c)

    End Sub

    ' Show the password recovery panel
    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Panel4.Visible = True
        Guna2Panel3.Visible = False
        PictureBox1.Visible = False
        Label6.Visible = False
    End Sub

    ' Hide the password recovery panel
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Label6.Visible = True
        Panel4.Visible = False
        Guna2Panel3.Visible = True
        PictureBox1.Visible = True
    End Sub

    ' Exit the application
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to exit?", "Exit Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Application.ExitThread()
        End If
    End Sub



    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        Guna2TextBox3.PasswordChar = If(Guna2TextBox3.PasswordChar = "*"c, ControlChars.NullChar, "*"c)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Guna2TextBox3.Clear()
    End Sub




    Private Sub Guna2TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles Guna2TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button6.PerformClick()
        End If
    End Sub


    Public Sub Clear()
        Guna2TextBox1.Clear()
        Guna2TextBox2.Clear()
        Guna2TextBox1.Focus()
    End Sub
    Public Sub ClearEmail()

        Guna2TextBox3.Clear()
    End Sub

    Private Sub Guna2TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles Guna2TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1.PerformClick()
        End If
    End Sub
    Private Sub LoadEmailToComboBox()
        Guna2ComboBox1.Items.Clear()

        Dim query As String = "SELECT DISTINCT Email FROM UserAccountTb WHERE Email IS NOT NULL AND Email <> ''"
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Try
                    con.Open()
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Guna2ComboBox1.Items.Add(reader("Email").ToString())
                        End While
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error loading emails: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' Check if the Esc key is pressed
        If e.KeyCode = Keys.Escape Then
            Button4.PerformClick()

        End If
    End Sub
End Class
