Imports System.Data.SqlClient

Public Class Form9
    Private Sub Form9_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Guna2CustomGradientPanel1.Visible = False
        AccountBind()


    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        ' Hide the panel
        Guna2CustomGradientPanel1.Visible = False

        ' Clear the fields inside the panel
        Clear()

        ' Clear selection in DataGridView
        DataGridView1.ClearSelection()

    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Guna2CustomGradientPanel1.Visible = True
        Guna2Button2.Enabled = False
    End Sub
    Private Function IsFieldValueExists(columnName As String, value As String, Optional caseSensitive As Boolean = False) As Boolean
        Dim query As String
        If caseSensitive AndAlso columnName = "UserName" Then
            query = $"SELECT COUNT(*) FROM UserAccountTb WHERE {columnName} COLLATE Latin1_General_CS_AS = @Value"
        Else
            query = $"SELECT COUNT(*) FROM UserAccountTb WHERE {columnName} = @Value"
        End If

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Value", value)
                con.Open()
                Return Convert.ToInt32(cmd.ExecuteScalar()) > 0
            End Using
        End Using
    End Function


    'Insert function
    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        ' Get values from input fields
        Dim UserName As String = Guna2TextBox2.Text.Trim()
        Dim Password As String = Guna2TextBox3.Text.Trim()
        Dim Role As String = If(Guna2ComboBox1.SelectedItem Is Nothing, String.Empty, Guna2ComboBox1.SelectedItem.ToString().Trim())
        Dim FullName As String = Guna2TextBox4.Text.Trim()
        Dim ContactNo As String = Guna2TextBox6.Text.Trim()
        Dim Address As String = Guna2TextBox7.Text.Trim()
        Dim Email As String = Guna2TextBox5.Text.Trim()
        Dim Status As String = If(Guna2ComboBox2.SelectedItem Is Nothing, String.Empty, Guna2ComboBox2.SelectedItem.ToString().Trim())

        ' Validate if all required fields are filled
        If String.IsNullOrEmpty(UserName) OrElse String.IsNullOrEmpty(Password) OrElse String.IsNullOrEmpty(Role) OrElse
       String.IsNullOrEmpty(FullName) OrElse String.IsNullOrEmpty(ContactNo) OrElse String.IsNullOrEmpty(Address) OrElse
       String.IsNullOrEmpty(Email) OrElse String.IsNullOrEmpty(Status) Then
            MessageBox.Show("Please fill in all required fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If IsFieldValueExists("UserName", UserName, caseSensitive:=True) Then
            MessageBox.Show("Username already exists. Please choose a different username.", "Duplicate Username", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Clear()
            Exit Sub
        End If

        If IsFieldValueExists("Email", Email) Then
            MessageBox.Show("Email already exists. Please use a different email.", "Duplicate Email", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Clear()
            Exit Sub
        End If


        ' SQL query to insert data
        Dim query As String = "INSERT INTO UserAccountTb (UserName, Password, Role, FullName, ContactNo, Email, Address, Status) " &
                          "VALUES (@UserName, @Password, @Role, @FullName, @ContactNo, @Email, @Address, @Status)"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                ' Add parameters to prevent SQL injection
                cmd.Parameters.AddWithValue("@UserName", UserName)
                cmd.Parameters.AddWithValue("@Password", Password)
                cmd.Parameters.AddWithValue("@Role", Role)
                cmd.Parameters.AddWithValue("@FullName", FullName)
                cmd.Parameters.AddWithValue("@ContactNo", ContactNo)
                cmd.Parameters.AddWithValue("@Email", Email)
                cmd.Parameters.AddWithValue("@Address", Address)
                cmd.Parameters.AddWithValue("@Status", Status)

                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    MessageBox.Show("User account has been created successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    AccountBind()
                    Clear()
                Catch ex As SqlException
                    MessageBox.Show("SQL error occurred: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    con.Close()
                End Try
            End Using
        End Using
    End Sub


    'Display data
    Private Sub AccountBind(Optional searchName As String = "")
        Dim query As String = "SELECT Id, UserName, Password, Role, FullName, ContactNo, Email, Address, Status FROM UserAccountTb ORDER BY Id DESC"

        If Not String.IsNullOrEmpty(searchName) Then
            query &= " WHERE UserName LIKE @Name"
        End If

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Dim column1 As New AutoCompleteStringCollection

                ' Correct parameter name for searching
                If Not String.IsNullOrEmpty(searchName) Then
                    cmd.Parameters.AddWithValue("@Name", "%" & searchName & "%")
                End If

                Using da As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        da.Fill(dt)

                        ' Set AutoComplete for Guna2TextBox1
                        For Each row As DataRow In dt.Rows
                            column1.Add(row("UserName").ToString())
                        Next
                        Guna2TextBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                        Guna2TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
                        Guna2TextBox1.AutoCompleteCustomSource = column1

                        ' Bind the DataGridView
                        DataGridView1.DataSource = dt

                        ' Mask Password and Role columns with asterisks (*)
                        For Each row As DataGridViewRow In DataGridView1.Rows
                            If row.Cells("Password").Value IsNot Nothing Then
                                row.Cells("Password").Value = New String("*"c, row.Cells("Password").Value.ToString().Length)
                            End If

                            If row.Cells("Role").Value IsNot Nothing Then
                                row.Cells("Role").Value = New String("*"c, row.Cells("Role").Value.ToString().Length)
                            End If
                        Next

                    End Using
                End Using
            End Using
        End Using
    End Sub


    Private Function GetRoleFromDatabase(userName As String) As String
        Dim role As String = String.Empty
        Dim query As String = "SELECT Role FROM UserAccountTb WHERE UserName = @UserName"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@UserName", userName)

                con.Open()
                role = cmd.ExecuteScalar().ToString()
                con.Close()
            End Using
        End Using

        Return role
    End Function

    Private Function GetPasswordFromDatabase(userName As String) As String
        Dim password As String = String.Empty
        Dim query As String = "SELECT Password FROM UserAccountTb WHERE UserName = @UserName"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@UserName", userName)

                con.Open()
                password = cmd.ExecuteScalar().ToString()
                con.Close()
            End Using
        End Using

        Return password
    End Function

    Private Sub Guna2TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles Guna2TextBox1.TextChanged
        Module1.FilterAccount(Guna2TextBox1.Text, DataGridView1, ConnectionString)
    End Sub



    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        AccountBind()

    End Sub


    'Fetch Data
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Ensure a valid row index is selected
        If e.RowIndex >= 0 Then
            ' Fetch the data from the selected row
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            Guna2CustomGradientPanel1.Visible = True
            Guna2Button2.Enabled = False
            ' Fill the corresponding fields in Guna2CustomGradientPanel1
            Guna2TextBox2.Text = row.Cells("UserName").Value.ToString() ' UserName
            Guna2ComboBox1.Text = row.Cells("Role").Value.ToString() ' Role
            Guna2ComboBox2.Text = row.Cells("Status").Value.ToString() ' Role
            Guna2TextBox4.Text = row.Cells("FullName").Value.ToString() ' FullName
            Guna2TextBox6.Text = row.Cells("ContactNo").Value.ToString() ' ContactNo
            Guna2TextBox7.Text = row.Cells("Address").Value.ToString() ' Address
            Guna2TextBox5.Text = row.Cells("Email").Value.ToString() ' Email
        End If
    End Sub

    'Update function
    Private Sub Button1_Click(sender As Object, e As EventArgs)
        ' Validate if a row is selected
        If DataGridView1.SelectedRows.Count = 0 Then
            MessageBox.Show("Please select a user to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Get selected row and ID
        Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
        Dim idToUpdate As Integer = Convert.ToInt32(selectedRow.Cells("Id").Value)

        ' Extract values from controls
        Dim UserName As String = Guna2TextBox2.Text.Trim()
        Dim Password As String = Guna2TextBox3.Text.Trim()
        Dim Role As String = If(Guna2ComboBox1.SelectedItem Is Nothing, String.Empty, Guna2ComboBox1.SelectedItem.ToString())
        Dim Status As String = If(Guna2ComboBox2.SelectedItem Is Nothing, String.Empty, Guna2ComboBox2.SelectedItem.ToString())
        Dim FullName As String = Guna2TextBox4.Text.Trim()
        Dim ContactNo As String = Guna2TextBox6.Text.Trim()
        Dim Address As String = Guna2TextBox7.Text.Trim()
        Dim Email As String = Guna2TextBox5.Text.Trim()

        ' Validate required fields
        If String.IsNullOrEmpty(UserName) OrElse String.IsNullOrEmpty(Password) OrElse
       String.IsNullOrEmpty(Role) OrElse String.IsNullOrEmpty(Status) OrElse
       String.IsNullOrEmpty(FullName) OrElse String.IsNullOrEmpty(ContactNo) Then

            MessageBox.Show("Please fill in all required fields.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Check for duplicate username (excluding current user)
        If IsFieldValueExists("UserName", UserName, True) AndAlso Not UserName.Equals(selectedRow.Cells("UserName").Value.ToString(), StringComparison.Ordinal) Then
            MessageBox.Show("Username already exists. Please choose a different one.", "Duplicate Username", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Check for duplicate email (excluding current user)
        If IsFieldValueExists("Email", Email) AndAlso Not Email.Equals(selectedRow.Cells("Email").Value.ToString(), StringComparison.OrdinalIgnoreCase) Then
            MessageBox.Show("Email already exists. Please use a different one.", "Duplicate Email", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Confirm update
        If MessageBox.Show("Are you sure you want to update this user?", "Confirm Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
            Return
        End If

        ' SQL update with Status
        Dim query As String = "UPDATE UserAccountTb SET UserName = @UserName, Password = @Password, Role = @Role, Status = @Status, FullName = @FullName, " &
                          "ContactNo = @ContactNo, Email = @Email, Address = @Address WHERE Id = @Id"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@UserName", UserName)
                cmd.Parameters.AddWithValue("@Password", Password)
                cmd.Parameters.AddWithValue("@Role", Role)
                cmd.Parameters.AddWithValue("@Status", Status)
                cmd.Parameters.AddWithValue("@FullName", FullName)
                cmd.Parameters.AddWithValue("@ContactNo", ContactNo)
                cmd.Parameters.AddWithValue("@Email", Email)
                cmd.Parameters.AddWithValue("@Address", Address)
                cmd.Parameters.AddWithValue("@Id", idToUpdate)

                Try
                    con.Open()
                    Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        MessageBox.Show("The user details have been updated successfully.", "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        AccountBind()
                    Else
                        MessageBox.Show("No record found with the given ID.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                Catch ex As SqlException
                    MessageBox.Show("A SQL error occurred: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    con.Close()
                End Try
            End Using
        End Using

        Clear()
    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Clear()
    End Sub


    Private Sub Guna2TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles Guna2TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            Guna2Button1.PerformClick()
        End If
    End Sub
    Public Sub Clear()
        Guna2TextBox2.Clear()
        Guna2TextBox5.Clear()
        Guna2TextBox5.Clear()
        Guna2TextBox3.Clear()
        Guna2TextBox4.Clear()
        Guna2TextBox6.Clear()
        Guna2TextBox7.Clear()
        Guna2ComboBox1.SelectedIndex = -1
        Guna2ComboBox2.SelectedIndex = -1
        Guna2TextBox5.Focus()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Guna2CircleButton2_Click(sender As Object, e As EventArgs) Handles Guna2CircleButton2.Click

        Guna2CustomGradientPanel1.Visible = False
        Clear()
        DataGridView1.ClearSelection()
    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
        Clear()
    End Sub
End Class