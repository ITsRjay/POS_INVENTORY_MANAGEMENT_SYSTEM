Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form12

    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Guna2CustomGradientPanel1.Visible = False
        VarietyBind()
        Autocomplete()
        LoadCategories()

    End Sub
    Private Sub Guna2Button1_Click(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        ' Assign "-" automatically if textbox is empty
        Dim Category As String = If(String.IsNullOrWhiteSpace(Guna2TextBox2.Text), "-", Guna2TextBox2.Text)
        Dim Brand As String = If(String.IsNullOrWhiteSpace(Guna2TextBox4.Text), "-", Guna2TextBox4.Text)
        Dim Model As String = If(String.IsNullOrWhiteSpace(Guna2TextBox5.Text), "-", Guna2TextBox5.Text)
        Dim ProductName As String = If(String.IsNullOrWhiteSpace(Guna2TextBox3.Text), "-", Guna2TextBox3.Text)

        Using con As New SqlConnection(ConnectionString)
            con.Open()

            ' Only check duplicates if the value is not "-"

            If Category <> "-" Then
                Dim checkBrand As New SqlCommand("SELECT COUNT(*) FROM VarietyTb WHERE Category = @Category", con)
                checkBrand.Parameters.AddWithValue("@Category", Category)
                If Convert.ToInt32(checkBrand.ExecuteScalar()) > 0 Then
                    MessageBox.Show("Category already exists!", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End If

            If Brand <> "-" Then
                Dim checkBrand As New SqlCommand("SELECT COUNT(*) FROM VarietyTb WHERE Brand = @Brand", con)
                checkBrand.Parameters.AddWithValue("@Brand", Brand)
                If Convert.ToInt32(checkBrand.ExecuteScalar()) > 0 Then
                    MessageBox.Show("Brand already exists!", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End If

            If Model <> "-" Then
                Dim checkModel As New SqlCommand("SELECT COUNT(*) FROM VarietyTb WHERE Model = @Model", con)
                checkModel.Parameters.AddWithValue("@Model", Model)
                If Convert.ToInt32(checkModel.ExecuteScalar()) > 0 Then
                    MessageBox.Show("Model already exists!", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End If

            If ProductName <> "-" Then
                Dim checkProductName As New SqlCommand("SELECT COUNT(*) FROM VarietyTb WHERE ProductName = @ProductName", con)
                checkProductName.Parameters.AddWithValue("@ProductName", ProductName)
                If Convert.ToInt32(checkProductName.ExecuteScalar()) > 0 Then
                    MessageBox.Show("Product Name already exists!", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End If

            ' If all checks pass, proceed to insert
            Dim insertQuery As String = "INSERT INTO VarietyTb (Category, Brand, Model, ProductName) VALUES (@Category, @Brand, @Model, @ProductName)"
            Using insertCmd As New SqlCommand(insertQuery, con)
                insertCmd.Parameters.AddWithValue("@Category", Category)
                insertCmd.Parameters.AddWithValue("@Brand", Brand)
                insertCmd.Parameters.AddWithValue("@Model", Model)
                insertCmd.Parameters.AddWithValue("@ProductName", ProductName)

                Try
                    insertCmd.ExecuteNonQuery()
                    MessageBox.Show("Data has been inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    VarietyBind()
                Catch ex As SqlException
                    MessageBox.Show("SQL error occurred: " & ex.Message)
                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message)
                End Try
            End Using
        End Using

        Clear()
    End Sub



    Private Sub VarietyBind()

        Dim query As String = "SELECT Id, Category, ProductName, Brand, Model FROM VarietyTb ORDER BY Id DESC"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Using da As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt

                    End Using
                End Using
            End Using
        End Using

    End Sub

    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        If DataGridView1.CurrentRow IsNot Nothing Then
            Dim id As Integer = Convert.ToInt32(DataGridView1.CurrentRow.Cells("Id").Value)
            Dim category As String = Guna2TextBox2.Text.Trim()
            Dim productName As String = Guna2TextBox3.Text.Trim()
            Dim brand As String = Guna2TextBox4.Text.Trim()
            Dim model As String = Guna2TextBox5.Text.Trim()

            Dim query As String = "UPDATE VarietyTb SET Category = @Category, ProductName = @ProductName, Brand = @Brand, Model = @Model WHERE Id = @Id"

            Using con As New SqlConnection(ConnectionString)
                Using cmd As New SqlCommand(query, con)
                    cmd.Parameters.AddWithValue("@Category", category)
                    cmd.Parameters.AddWithValue("@ProductName", productName)
                    cmd.Parameters.AddWithValue("@Brand", brand)
                    cmd.Parameters.AddWithValue("@Model", model)
                    cmd.Parameters.AddWithValue("@Id", id)

                    con.Open()
                    cmd.ExecuteNonQuery()
                    con.Close()
                End Using
            End Using

            ' Refresh DataGridView and hide panel
            VarietyBind()
            Guna2CustomGradientPanel1.Visible = False

            MessageBox.Show("Record updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Please select a record to update.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Guna2CustomGradientPanel1.Visible = Not Guna2CustomGradientPanel1.Visible
        Guna2Button2.Enabled = False

    End Sub


    Sub Autocomplete()
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()


                Dim cmd = New SqlCommand("SELECT Id, Category, ProductName, Brand, Model  FROM VarietyTb", con)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Dim column1 As New AutoCompleteStringCollection
                For Each row As DataRow In dt.Rows
                    Dim combinedValue As String = row("Category").ToString() & " " & row("ProductName").ToString() & "" & row("Brand").ToString() & " " & row("Model").ToString()
                    column1.Add(combinedValue)
                Next
                Guna2TextBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                Guna2TextBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
                Guna2TextBox1.AutoCompleteCustomSource = column1
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub Guna2TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles Guna2TextBox1.TextChanged
        Module1.FilterDataGridView3(Guna2TextBox1.Text, DataGridView1, ConnectionString)
    End Sub



    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        VarietyBind()
        Autocomplete()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Guna2CustomGradientPanel1.Visible = False
        DataGridView1.ClearSelection()
        Clear()
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim selectedCategory As String = ComboBox3.SelectedItem.ToString()
        FilterByCategory(selectedCategory)
    End Sub
    Public Sub FilterByCategory(selectedCategory As String)
        Dim query As String = "SELECT Id, Category, ProductName, Brand, Model FROM VarietyTb WHERE Category = @Category"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Category", selectedCategory)
                Using da As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
                    End Using
                End Using
            End Using
        End Using
    End Sub
    Public Sub LoadCategories()
        ComboBox3.Items.Clear()

        Using con As New SqlConnection(ConnectionString)
            con.Open()

            Dim query As String = "SELECT DISTINCT Category FROM VarietyTb"
            Using cmd As New SqlCommand(query, con)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)

                    For Each row As DataRow In dt.Rows
                        ComboBox3.Items.Add(row("Category").ToString())
                    Next

                End Using
            End Using

            con.Close()
        End Using
    End Sub

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
        Clear()

    End Sub
    Public Sub Clear()
        Guna2TextBox2.Clear()
        Guna2TextBox4.Clear()
        Guna2TextBox5.Clear()
        Guna2TextBox3.Clear()
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Ensure a valid row is selected
        If e.RowIndex >= 0 Then

            Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            Guna2TextBox3.ForeColor = Color.Black
            Guna2Button2.Enabled = True

            Guna2TextBox2.Text = selectedRow.Cells("Category").Value.ToString()
            Guna2TextBox3.Text = selectedRow.Cells("ProductName").Value.ToString()
            Guna2TextBox4.Text = selectedRow.Cells("Brand").Value.ToString()
            Guna2TextBox5.Text = selectedRow.Cells("Model").Value.ToString()

            ' Show Guna2CustomGradientPanel1
            Guna2CustomGradientPanel1.Visible = True
        End If
    End Sub

    Private Sub Guna2TextBox5_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox5.TextChanged

    End Sub
    Private Sub Guna2TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles Guna2TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            Guna2Button1.PerformClick()
        End If
    End Sub



    Private Sub Guna2CircleButton2_Click(sender As Object, e As EventArgs) Handles Guna2CircleButton2.Click
        Guna2CustomGradientPanel1.Visible = False
        DataGridView1.ClearSelection()
        Clear()
    End Sub
End Class