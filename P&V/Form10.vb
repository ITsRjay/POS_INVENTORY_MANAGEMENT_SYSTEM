Imports System.Data.SqlClient

Public Class Form10

	Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
		' Get values from textboxes
		Dim Name As String = Guna2TextBox2.Text.Trim()
		Dim Brand As String = Guna2TextBox3.Text.Trim()
		Dim Contact As String = Guna2TextBox4.Text.Trim()
		Dim SocMed As String = Guna2TextBox5.Text.Trim()
		Dim EmailAddress As String = Guna2TextBox6.Text.Trim()
		Dim Address As String = Guna2TextBox7.Text.Trim()

		' Check if any field is empty
		If String.IsNullOrEmpty(Name) OrElse String.IsNullOrEmpty(Brand) OrElse String.IsNullOrEmpty(Contact) OrElse
	   String.IsNullOrEmpty(SocMed) OrElse String.IsNullOrEmpty(EmailAddress) OrElse String.IsNullOrEmpty(Address) Then
			MessageBox.Show("Please fill in all required fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			Exit Sub
		End If

		' SQL Query to insert data
		Dim query As String = "INSERT INTO SupplierTb (Name, Brand, Contact, SocMed, EmailAddress, Address) " &
						  "VALUES (@Name, @Brand, @Contact, @SocMed, @EmailAddress, @Address)"

		Using con As New SqlConnection(ConnectionString)
			Using cmd As New SqlCommand(query, con)
				' Add parameters to prevent SQL injection
				cmd.Parameters.AddWithValue("@Name", Name)
				cmd.Parameters.AddWithValue("@Brand", Brand)
				cmd.Parameters.AddWithValue("@Contact", Contact)
				cmd.Parameters.AddWithValue("@SocMed", SocMed)
				cmd.Parameters.AddWithValue("@EmailAddress", EmailAddress)
				cmd.Parameters.AddWithValue("@Address", Address)

				Try
					con.Open()
					cmd.ExecuteNonQuery()
					MessageBox.Show("Supplier data has been inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

					' Refresh the supplier data after insertion
					SupplierBind()
					Clear()
					Guna2TextBox2.Focus()
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


	Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Guna2CustomGradientPanel1.Visible = False
		SupplierBind()

	End Sub
	Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
		Guna2CustomGradientPanel1.Visible = Not Guna2CustomGradientPanel1.Visible
		Guna2Button2.Enabled = False
	End Sub


	Private Sub SupplierBind(Optional searchName As String = "")
		Dim query As String = "SELECT Id, Name, Brand, Contact, SocMed, EmailAddress, Address FROM SupplierTb"

		If Not String.IsNullOrEmpty(searchName) Then
			query &= " WHERE Name LIKE @Name"
		End If

		query &= " ORDER BY Id DESC" ' Ensure ORDER BY comes last

		Using con As New SqlConnection(ConnectionString)
			Using cmd As New SqlCommand(query, con)
				Dim column1 As New AutoCompleteStringCollection
				If Not String.IsNullOrEmpty(searchName) Then
					cmd.Parameters.AddWithValue("@Name", "%" & searchName & "%")
				End If

				Using da As New SqlDataAdapter(cmd)
					Using dt As New DataTable()
						da.Fill(dt)

						' Set AutoComplete for Guna2TextBox1
						For Each row As DataRow In dt.Rows
							column1.Add(row("Name").ToString())
						Next
						Guna2TextBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend
						Guna2TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
						Guna2TextBox1.AutoCompleteCustomSource = column1

						' Bind the data to DataGridView
						DataGridView1.DataSource = dt

						' Hide the Id column AFTER binding
						If DataGridView1.Columns.Contains("Id") Then
							DataGridView1.Columns("Id").Visible = False
						End If
					End Using
				End Using
			End Using
		End Using
	End Sub


	Private Sub Button2_Click(sender As Object, e As EventArgs)
		SupplierBind()
	End Sub
	Private Sub Guna2TextBox1_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox1.TextChanged

		Module1.FilterSupplier(Guna2TextBox1.Text, DataGridView1, ConnectionString)
	End Sub

	Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
		SupplierBind()
	End Sub

	Private Sub Label9_Click(sender As Object, e As EventArgs)

	End Sub



	' Populates the textboxes when a row in DataGridView1 is clicked
	Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
		If e.RowIndex >= 0 Then
			' Get the selected row
			Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
			Guna2CustomGradientPanel1.Visible = True
			Guna2Button2.Enabled = True
			' Populate the textboxes with the data from the selected row
			Guna2TextBox2.Text = selectedRow.Cells("Name").Value.ToString()
			Guna2TextBox3.Text = selectedRow.Cells("Brand").Value.ToString()
			Guna2TextBox4.Text = selectedRow.Cells("Contact").Value.ToString()
			Guna2TextBox5.Text = selectedRow.Cells("SocMed").Value.ToString()
			Guna2TextBox6.Text = selectedRow.Cells("EmailAddress").Value.ToString()
			Guna2TextBox7.Text = selectedRow.Cells("Address").Value.ToString()
		End If
	End Sub

	Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
		Clear()

	End Sub
	Public Sub Clear()
		Guna2TextBox2.Clear()
		Guna2TextBox3.Clear()
		Guna2TextBox4.Clear()
		Guna2TextBox5.Clear()
		Guna2TextBox6.Clear()
		Guna2TextBox7.Clear()
	End Sub
	Private Sub Guna2TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles Guna2TextBox7.KeyDown
		If e.KeyCode = Keys.Enter Then
			Guna2Button1.PerformClick()
		End If
	End Sub

	Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
		If DataGridView1.SelectedRows.Count = 0 Then
			MessageBox.Show("Please select a row to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			Exit Sub
		End If

		' Get selected row and ID
		Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
		Dim idToUpdate As Integer = Convert.ToInt32(selectedRow.Cells("Id").Value)

		' Get and trim values
		Dim Name As String = Guna2TextBox2.Text.Trim()
		Dim Brand As String = Guna2TextBox3.Text.Trim()
		Dim Contact As String = Guna2TextBox4.Text.Trim()
		Dim SocMed As String = Guna2TextBox5.Text.Trim()
		Dim EmailAddress As String = Guna2TextBox6.Text.Trim()
		Dim Address As String = Guna2TextBox7.Text.Trim()

		' Basic validation
		If String.IsNullOrWhiteSpace(Name) OrElse String.IsNullOrWhiteSpace(Brand) Then
			MessageBox.Show("Name and Brand fields are required.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
			Exit Sub
		End If

		Using con As New SqlConnection(ConnectionString)
			con.Open()

			' Duplicate check (excluding the current record by Id)
			Dim checkQuery As String = "SELECT COUNT(*) FROM SupplierTb WHERE Name = @Name AND Brand = @Brand AND Id <> @Id"
			Using checkCmd As New SqlCommand(checkQuery, con)
				checkCmd.Parameters.AddWithValue("@Name", Name)
				checkCmd.Parameters.AddWithValue("@Brand", Brand)
				checkCmd.Parameters.AddWithValue("@Id", idToUpdate)

				Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
				If count > 0 Then
					MessageBox.Show("A supplier with the same Name and Brand already exists.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
					Exit Sub
				End If
			End Using

			' Proceed with update
			Dim updateQuery As String = "UPDATE SupplierTb SET Name = @Name, Brand = @Brand, Contact = @Contact, " &
									"SocMed = @SocMed, EmailAddress = @EmailAddress, Address = @Address WHERE Id = @Id"
			Using updateCmd As New SqlCommand(updateQuery, con)
				updateCmd.Parameters.AddWithValue("@Name", Name)
				updateCmd.Parameters.AddWithValue("@Brand", Brand)
				updateCmd.Parameters.AddWithValue("@Contact", Contact)
				updateCmd.Parameters.AddWithValue("@SocMed", SocMed)
				updateCmd.Parameters.AddWithValue("@EmailAddress", EmailAddress)
				updateCmd.Parameters.AddWithValue("@Address", Address)
				updateCmd.Parameters.AddWithValue("@Id", idToUpdate)

				Try
					Dim rowsAffected As Integer = updateCmd.ExecuteNonQuery()
					If rowsAffected > 0 Then
						MessageBox.Show("Supplier information updated successfully.", "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
					Else
						MessageBox.Show("Update failed. Record not found.", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning)
					End If

					SupplierBind() ' Refresh DataGridView
				Catch ex As SqlException
					MessageBox.Show("A database error occurred: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
				Catch ex As Exception
					MessageBox.Show("An unexpected error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
				End Try
			End Using
		End Using
	End Sub


	Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
		Guna2CustomGradientPanel1.Visible = False
		DataGridView1.ClearSelection()
	End Sub
End Class