Imports System.Data.SqlClient
Imports MessagingToolkit.Barcode

Public Class Form4

    Private barcodeValue As String
    Private barcodeQty As Integer




    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Guna2CustomGradientPanel1.Visible = False
        AddHandler DataGridView1.DataBindingComplete, AddressOf DataGridView1_DataBindingComplete

        BindData()
        PopulateComboBoxes()
        Combo1()
        LoadCategories()
        Autocomplete()


    End Sub



    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click

        BindData()
        PopulateComboBoxes()
        Autocomplete()
        Combo1()
        Guna2TextBox1.Clear()
    End Sub

    Private Sub BindData()
        Dim query As String = "SELECT Barcode, Category, ProductName, Brand, Model, CapitalPrice, Price, Quantity, Expense, VatStatus, ProductStatus, DateIn, Supplier FROM InventoryTb ORDER BY DateIn DESC"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Using da As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt

                        If DataGridView1.Columns.Contains("Expense") Then
                            DataGridView1.Columns("Expense").Visible = False
                        End If

                    End Using
                End Using
            End Using
        End Using
    End Sub

    Private Sub ApplyRowColorFormatting()
        ' Apply conditional formatting to each row
        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.IsNewRow Then Continue For ' Skip new row placeholder

            Dim quantity As Integer = 0
            Dim status As String = ""

            ' Parse Quantity
            If Not IsDBNull(row.Cells("Quantity").Value) Then
                Integer.TryParse(row.Cells("Quantity").Value.ToString(), quantity)
            End If

            ' Parse ProductStatus
            If Not IsDBNull(row.Cells("ProductStatus").Value) Then
                status = row.Cells("ProductStatus").Value.ToString().ToUpper()
            End If

            ' Apply color based on conditions
            If quantity = 0 OrElse status = "INACTIVE" OrElse status = "DAMAGED" Then
                row.DefaultCellStyle.BackColor = Color.IndianRed ' Redish color
                row.DefaultCellStyle.ForeColor = Color.White
            ElseIf quantity <= 5 Then
                row.DefaultCellStyle.BackColor = Color.Tan ' Orange color for low quantity
                row.DefaultCellStyle.ForeColor = Color.White
            Else
                row.DefaultCellStyle.BackColor = Color.White ' Default white background
                row.DefaultCellStyle.ForeColor = Color.Black
            End If

            If status = "LOSS PRODUCT" Then
                row.DefaultCellStyle.BackColor = Color.Gray ' Gray for loss products
                row.DefaultCellStyle.ForeColor = Color.White
            End If
        Next
    End Sub

    Private Sub DataGridView1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
        ApplyRowColorFormatting()
    End Sub

    Private Sub Guna2Button1_Click_1(sender As Object, e As EventArgs) Handles Guna2Button1.Click
        ' Check for empty fields
        If String.IsNullOrWhiteSpace(Guna2TextBox2.Text) OrElse
        Guna2ComboBox1.SelectedIndex = -1 OrElse
        Guna2ComboBox3.SelectedIndex = -1 OrElse
        Guna2ComboBox5.SelectedIndex = -1 OrElse
        String.IsNullOrWhiteSpace(Guna2TextBox3.Text) OrElse
        String.IsNullOrWhiteSpace(Guna2TextBox4.Text) OrElse
        String.IsNullOrWhiteSpace(Guna2TextBox5.Text) OrElse
        String.IsNullOrWhiteSpace(Guna2TextBox6.Text) Then

            MessageBox.Show("Please fill in all fields before submitting.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Get values from controls
        Dim Barcode As String = Guna2TextBox2.Text.Trim()
        Dim ProductStatus As String = Guna2ComboBox1.SelectedItem.ToString()
        Dim Category As String = Guna2ComboBox3.SelectedItem.ToString()
        Dim Brand As String = If(Guna2ComboBox4.SelectedItem IsNot Nothing, Guna2ComboBox4.SelectedItem.ToString(), String.Empty)
        Dim Supplier As String = Guna2ComboBox5.SelectedItem.ToString()
        Dim Model As String = If(Guna2ComboBox6.SelectedItem IsNot Nothing, Guna2ComboBox6.SelectedItem.ToString(), String.Empty)
        Dim ProductName As String = If(Guna2ComboBox7.SelectedItem IsNot Nothing, Guna2ComboBox7.SelectedItem.ToString(), String.Empty)
        Dim Price As Decimal = Convert.ToDecimal(Guna2TextBox3.Text)
        Dim Quantity As String = Guna2TextBox4.Text
        Dim Expense As String = Guna2TextBox5.Text
        Dim CapitalPrice As Decimal = Convert.ToDecimal(Guna2TextBox6.Text)
        Dim VatStatus As String

        If Guna2CustomRadioButton1.Checked Then
            VatStatus = "VAT"
        ElseIf Guna2CustomRadioButton2.Checked Then
            VatStatus = "NonVAT"
        Else
            MessageBox.Show("Please select a VAT status.", "Missing VAT Status", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Using con As New SqlConnection(ConnectionString)
            con.Open()

            ' Check if barcode already exists
            Dim checkQuery As String = "SELECT COUNT(*) FROM InventoryTb WHERE Barcode = @Barcode"
            Using checkCmd As New SqlCommand(checkQuery, con)
                checkCmd.Parameters.AddWithValue("@Barcode", Barcode)
                Dim count As Integer = Convert.ToInt32(checkCmd.ExecuteScalar())
                If count > 0 Then
                    MessageBox.Show("This barcode already exists in the inventory.", "Duplicate Barcode", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
            End Using

            ' Begin transaction
            Using transaction = con.BeginTransaction()
                Try
                    Dim queryDecreasing As String = "INSERT INTO InventoryTb (Barcode, Category, Brand, Model, ProductName, Price, Quantity, Expense, CapitalPrice, DateIn, Supplier, VatStatus, ProductStatus) " &
                                                "VALUES (@Barcode, @Category, @Brand, @Model, @ProductName, @Price, @Quantity, @Expense, @CapitalPrice, GETDATE(), @Supplier, @VatStatus, @ProductStatus)"

                    Dim queryNonDecreasing As String = "INSERT INTO InventoryRecord (Barcode, Category, Brand, Model, ProductName, Price, Quantity, Expense, CapitalPrice, DateIn, Supplier, VatStatus) " &
                                                   "VALUES (@Barcode, @Category, @Brand, @Model, @ProductName, @Price, @Quantity, @Expense, @CapitalPrice, GETDATE(), @Supplier, @VatStatus)"

                    ' Insert into InventoryTb
                    Using cmd As New SqlCommand(queryDecreasing, con, transaction)
                        cmd.Parameters.AddWithValue("@Barcode", Barcode)
                        cmd.Parameters.AddWithValue("@Category", Category)
                        cmd.Parameters.AddWithValue("@Brand", Brand)
                        cmd.Parameters.AddWithValue("@Model", Model)
                        cmd.Parameters.AddWithValue("@ProductName", ProductName)
                        cmd.Parameters.AddWithValue("@Price", Price)
                        cmd.Parameters.AddWithValue("@Quantity", Quantity)
                        cmd.Parameters.AddWithValue("@Expense", Expense)
                        cmd.Parameters.AddWithValue("@CapitalPrice", CapitalPrice)
                        cmd.Parameters.AddWithValue("@Supplier", Supplier)
                        cmd.Parameters.AddWithValue("@VatStatus", VatStatus)
                        cmd.Parameters.AddWithValue("@ProductStatus", ProductStatus)
                        cmd.ExecuteNonQuery()
                    End Using

                    ' Insert into InventoryRecord
                    Using cmd As New SqlCommand(queryNonDecreasing, con, transaction)
                        cmd.Parameters.AddWithValue("@Barcode", Barcode)
                        cmd.Parameters.AddWithValue("@Category", Category)
                        cmd.Parameters.AddWithValue("@Brand", Brand)
                        cmd.Parameters.AddWithValue("@Model", Model)
                        cmd.Parameters.AddWithValue("@ProductName", ProductName)
                        cmd.Parameters.AddWithValue("@Price", Price)
                        cmd.Parameters.AddWithValue("@Quantity", Quantity)
                        cmd.Parameters.AddWithValue("@Expense", Expense)
                        cmd.Parameters.AddWithValue("@CapitalPrice", CapitalPrice)
                        cmd.Parameters.AddWithValue("@Supplier", Supplier)
                        cmd.Parameters.AddWithValue("@VatStatus", VatStatus)
                        cmd.ExecuteNonQuery()
                    End Using

                    transaction.Commit()
                    MessageBox.Show("The product has been inserted successfully.", "Insertion Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    BindData() ' Refresh data after insertion
                    LoadCategories()
                    Clear()
                Catch ex As SqlException
                    transaction.Rollback()
                    MessageBox.Show("SQL error occurred: " & ex.Message)
                Catch ex As Exception
                    transaction.Rollback()
                    MessageBox.Show("An error occurred: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub


    Public Sub Clear()
        ' Clear textboxes after insertion
        Guna2TextBox2.Clear()
        Guna2TextBox3.Clear()
        Guna2TextBox4.Clear()
        Guna2TextBox5.Clear()
        Guna2TextBox6.Clear()
        Guna2ComboBox1.SelectedIndex = -1
        Guna2ComboBox3.SelectedIndex = -1
        Guna2ComboBox4.SelectedIndex = -1
        Guna2ComboBox5.SelectedIndex = -1
        Guna2ComboBox6.SelectedIndex = -1
        Guna2ComboBox7.SelectedIndex = -1
        Guna2CustomRadioButton1.Checked = False
        Guna2CustomRadioButton2.Checked = False
        Guna2TextBox6.Focus()

    End Sub



    Private Function GenerateBarcode() As String

        Dim rand As New Random()
        Dim timestamp As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim randomPart As String = rand.Next(1000, 9999).ToString()

        Return timestamp & randomPart
    End Function

    Private Sub Guna2Button3_Click(sender As Object, e As EventArgs) Handles Guna2Button3.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)
            Dim idToDelete As String = selectedRow.Cells("Barcode").Value.ToString()

            ' Query to check if the item has already been transacted
            Dim queryCheckTransaction As String = "SELECT COUNT(*) FROM PaymentTb WHERE Barcode = @Barcode"

            Using con As New SqlConnection(ConnectionString)
                Using cmdCheck As New SqlCommand(queryCheckTransaction, con)
                    cmdCheck.Parameters.AddWithValue("@Barcode", idToDelete)

                    Try
                        con.Open()
                        Dim transactionCount As Integer = Convert.ToInt32(cmdCheck.ExecuteScalar())

                        If transactionCount > 0 Then
                            ' If the item has been transacted, show error message and exit
                            MessageBox.Show("This item cannot be deleted because it has already been transacted.", "Deletion Restricted", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            Exit Sub
                        End If
                    Catch ex As SqlException
                        MessageBox.Show("An error occurred while checking transactions: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                End Using
            End Using

            ' Show confirmation dialog before deleting
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to delete this item?", "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result <> DialogResult.Yes Then Exit Sub

            ' Queries to delete from both tables
            Dim queryDeleteInventoryTb As String = "DELETE FROM InventoryTb WHERE Barcode = @Barcode"
            Dim queryDeleteInventoryRecord As String = "DELETE FROM InventoryRecord WHERE Barcode = @Barcode"

            Using con As New SqlConnection(ConnectionString)
                Using cmd1 As New SqlCommand(queryDeleteInventoryTb, con)
                    Using cmd2 As New SqlCommand(queryDeleteInventoryRecord, con)
                        ' Add parameter to both commands
                        cmd1.Parameters.AddWithValue("@Barcode", idToDelete)
                        cmd2.Parameters.AddWithValue("@Barcode", idToDelete)

                        Try
                            con.Open()
                            ' Begin a transaction
                            Using transaction = con.BeginTransaction()
                                cmd1.Transaction = transaction
                                cmd2.Transaction = transaction

                                ' Execute both commands
                                cmd1.ExecuteNonQuery()
                                cmd2.ExecuteNonQuery()

                                ' Commit the transaction
                                transaction.Commit()
                            End Using

                            ' Update UI and notify the user
                            PopulateComboBoxes()
                            BindData()
                            MessageBox.Show("The selected record has been deleted successfully from both tables.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Catch ex As SqlException
                            MessageBox.Show("An error occurred while deleting the record: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    End Using
                End Using
            End Using
        Else
            MessageBox.Show("Please select a row to delete.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub


    Public isGuna2CustomGradientPanel1Visible As Boolean = False
    Public isPanel2Visible As Boolean = False

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        isGuna2CustomGradientPanel1Visible = Not isGuna2CustomGradientPanel1Visible

        If isGuna2CustomGradientPanel1Visible Then
            Guna2CustomGradientPanel1.Visible = True
            Guna2Button3.Enabled = False ' Disable Delete
            Guna2Button2.Enabled = False ' Disable Delete
            DataGridView1.ClearSelection() ' Unselect any selected row

            Clear()

            ' Enable fields for new input
            Guna2ComboBox3.Enabled = True
            Guna2ComboBox4.Enabled = True
            Guna2ComboBox5.Enabled = True
            Guna2ComboBox6.Enabled = True
            Guna2ComboBox7.Enabled = True
            Guna2TextBox3.Enabled = True
            Guna2TextBox4.Enabled = True
            Guna2TextBox5.Enabled = True
            Guna2TextBox2.Enabled = True
            Guna2CircleButton2.Enabled = True
            Guna2CustomRadioButton1.Enabled = True
            Guna2CustomRadioButton2.Enabled = True

        Else
            Guna2CustomGradientPanel1.Visible = False
        End If
    End Sub




    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Guna2CustomGradientPanel1.Visible = False
    End Sub

    Private Sub MultiplyValues()
        ' Check if the values in Guna2TextBox3 and Guna2TextBox4 are numeric
        Dim value1 As Decimal
        Dim value2 As Decimal

        ' Try to parse the values as Decimal
        If Decimal.TryParse(Guna2TextBox6.Text, value1) AndAlso Decimal.TryParse(Guna2TextBox4.Text, value2) Then
            ' Perform the multiplication and show the result in Guna2TextBox5
            Guna2TextBox5.Text = (value1 * value2).ToString()
        Else
            Guna2TextBox5.Text = ""
        End If
    End Sub

    Public Sub Combo1()
        Using con As New SqlConnection(ConnectionString)
            con.Open()

            Dim query As String = "SELECT DISTINCT Name FROM SupplierTb"
            Using cmd As New SqlCommand(query, con)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)

                    ' Save the current selection
                    Dim selectedItem As String = If(Guna2ComboBox5.SelectedItem IsNot Nothing, Guna2ComboBox5.SelectedItem.ToString(), String.Empty)

                    ' Loop through the DataTable and add only new items to ComboBox
                    For Each row As DataRow In dt.Rows
                        Dim supplierName As String = row("Name").ToString()
                        If Not Guna2ComboBox5.Items.Contains(supplierName) Then
                            Guna2ComboBox5.Items.Add(supplierName)
                        End If
                    Next

                    ' Restore the previously selected item (if exists)
                    If Not String.IsNullOrEmpty(selectedItem) Then
                        Guna2ComboBox5.SelectedItem = selectedItem
                    End If
                End Using
            End Using

            con.Close()
        End Using
    End Sub


    Private Sub PopulateComboBoxes()
        Dim query As String = "SELECT DISTINCT Category, Brand, Model, ProductName FROM VarietyTb"

        ' Declare HashSet to ensure unique values
        Dim categorySet As New HashSet(Of String)()
        Dim brandSet As New HashSet(Of String)()
        Dim modelSet As New HashSet(Of String)()
        Dim flavorSet As New HashSet(Of String)()

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Try
                    con.Open()
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        ' Clear ComboBoxes before adding new items
                        Guna2ComboBox3.Items.Clear()
                        Guna2ComboBox4.Items.Clear()
                        Guna2ComboBox6.Items.Clear()
                        Guna2ComboBox7.Items.Clear()

                        While reader.Read()
                            ' Add Category to HashSet if it's not null
                            If Not IsDBNull(reader("Category")) Then
                                Dim category As String = reader("Category").ToString()
                                If categorySet.Add(category) Then
                                    Guna2ComboBox3.Items.Add(category) ' Only add if it's unique
                                End If
                            End If

                            ' Add Brand to HashSet if it's not null
                            If Not IsDBNull(reader("Brand")) Then
                                Dim brand As String = reader("Brand").ToString()
                                If brandSet.Add(brand) Then
                                    Guna2ComboBox4.Items.Add(brand) ' Only add if it's unique
                                End If
                            End If

                            ' Add Model to HashSet if it's not null
                            If Not IsDBNull(reader("Model")) Then
                                Dim model As String = reader("Model").ToString()
                                If modelSet.Add(model) Then
                                    Guna2ComboBox6.Items.Add(model) ' Only add if it's unique
                                End If
                            End If

                            ' Add ProductName to HashSet if it's not null
                            If Not IsDBNull(reader("ProductName")) Then
                                Dim flavor As String = reader("ProductName").ToString()
                                If flavorSet.Add(flavor) Then
                                    Guna2ComboBox7.Items.Add(flavor) ' Only add if it's unique
                                End If
                            End If
                        End While
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message)
                Finally
                    con.Close()
                End Try
            End Using
        End Using
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)

        Guna2CustomGradientPanel1.Visible = True
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs)
        BindData()
    End Sub


    Private Sub Guna2Button5_Click(sender As Object, e As EventArgs) Handles Guna2Button5.Click
        Clear()
    End Sub


    'Fetch data to its certain fields
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            ' Enable all fields for editing
            Guna2ComboBox3.Enabled = False  ' Category
            Guna2ComboBox4.Enabled = False ' Brand
            Guna2ComboBox5.Enabled = False  ' Supplier
            Guna2ComboBox6.Enabled = False ' Model
            Guna2ComboBox7.Enabled = False ' ProductName
            Guna2TextBox3.Enabled = True   ' Price
            Guna2TextBox4.Enabled = True   ' Quantity
            Guna2TextBox5.Enabled = True  ' Expense
            Guna2TextBox2.Enabled = False ' Barcode
            Guna2CircleButton2.Enabled = True
            Guna2CustomGradientPanel1.Visible = True
            Guna2Button3.Enabled = True 'Enable Delete button
            Guna2Button2.Enabled = True ' Enable Delete button
            Guna2CustomRadioButton1.Enabled = False 'Vat
            Guna2CustomRadioButton2.Enabled = False  'Non-Vat


            ' Populate fields from the selected row
            Guna2TextBox2.Text = row.Cells("Barcode").Value.ToString()
            Guna2ComboBox1.SelectedItem = row.Cells("ProductStatus").Value.ToString()
            Guna2ComboBox3.SelectedItem = row.Cells("Category").Value.ToString()
            Guna2ComboBox4.SelectedItem = row.Cells("Brand").Value.ToString()
            Guna2ComboBox5.SelectedItem = row.Cells("Supplier").Value.ToString()
            Guna2ComboBox6.SelectedItem = row.Cells("Model").Value.ToString()
            Guna2ComboBox7.SelectedItem = row.Cells("ProductName").Value.ToString()
            Guna2TextBox3.Text = row.Cells("Price").Value.ToString()
            Guna2TextBox4.Text = row.Cells("Quantity").Value.ToString()
            Guna2TextBox5.Text = row.Cells("Expense").Value.ToString()
            Guna2TextBox6.Text = row.Cells("CapitalPrice").Value.ToString()

            Dim VatStatus As String = row.Cells("VatStatus").Value.ToString()
            If VatStatus = "VAT" Then
                Guna2CustomRadioButton1.Checked = True
            ElseIf VatStatus = "Non-VAT" Then
                Guna2CustomRadioButton2.Checked = True
            Else
                Guna2CustomRadioButton1.Checked = False
                Guna2CustomRadioButton2.Checked = False
            End If
        End If
    End Sub



    Private Sub Button13_Click(sender As Object, e As EventArgs)
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
    End Sub

    Private Sub Guna2ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Guna2ComboBox5.SelectedIndexChanged
        Combo1()
    End Sub

    Private Sub Guna2TextBox1_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox1.TextChanged
        Module1.FilterInventory(Guna2TextBox1.Text, DataGridView1, ConnectionString)
        ApplyRowColorFormatting()
    End Sub


    'For search bar
    Sub Autocomplete()
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()


                Dim cmd = New SqlCommand("SELECT Barcode FROM InventoryTb", con)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Dim column1 As New AutoCompleteStringCollection
                For Each row As DataRow In dt.Rows
                    Dim combinedValue As String = row("Barcode").ToString()
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

    Private Sub Guna2ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Guna2ComboBox2.SelectedIndexChanged
        Dim selectedCategory As String = Guna2ComboBox2.SelectedItem.ToString()
        FilterByCategory(selectedCategory)
    End Sub

    Public Sub FilterByCategory(selectedCategory As String)
        Dim query As String = "SELECT Barcode, Category, ProductName, Brand, Model, CapitalPrice, Price, Quantity, Expense, VatStatus, ProductStatus, DateIn, Supplier FROM InventoryTb WHERE Category = @Category ORDER BY DateIn DESC"

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
        Guna2ComboBox2.Items.Clear()

        Using con As New SqlConnection(ConnectionString)
            con.Open()

            Dim query As String = "SELECT DISTINCT Category FROM InventoryTb"
            Using cmd As New SqlCommand(query, con)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)

                    For Each row As DataRow In dt.Rows
                        Guna2ComboBox2.Items.Add(row("Category").ToString())
                    Next
                    BindData()
                End Using
            End Using

            con.Close()
        End Using
    End Sub


    '' Button1 Click Event - To save DataGridView data to a selected CSV file
    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    SaveDataGridViewToCSV()
    'End Sub

    'Private Sub SaveDataGridViewToCSV()
    '    ' Initialize and configure the SaveFileDialog
    '    Dim sfd As New SaveFileDialog With {
    '    .Filter = "CSV files (*.csv)|*.csv", ' Only allow CSV files
    '    .Title = "Save a CSV File",
    '    .FileName = "PrintedDataGridView.csv" ' Default file name
    '}

    '    ' Show the dialog and continue only if the user clicks Save
    '    If sfd.ShowDialog() = DialogResult.OK Then
    '        Try
    '            ' Create a StreamWriter to write to the selected file path
    '            Using writer As New System.IO.StreamWriter(sfd.FileName)
    '                ' Write column headers
    '                For col As Integer = 0 To DataGridView1.Columns.Count - 1
    '                    writer.Write(DataGridView1.Columns(col).HeaderText)
    '                    If col < DataGridView1.Columns.Count - 1 Then writer.Write(",")
    '                Next
    '                writer.WriteLine()

    '                ' Write rows
    '                For row As Integer = 0 To DataGridView1.Rows.Count - 1
    '                    For col As Integer = 0 To DataGridView1.Columns.Count - 1
    '                        writer.Write(DataGridView1.Rows(row).Cells(col).Value?.ToString())
    '                        If col < DataGridView1.Columns.Count - 1 Then writer.Write(",")
    '                    Next
    '                    writer.WriteLine()
    '                Next
    '            End Using

    '            ' Notify success
    '            MessageBox.Show("Data saved to CSV file successfully!", "Save Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

    '        Catch ex As Exception
    '            MessageBox.Show("An error occurred while saving the data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        End Try
    '    End If
    'End Sub


    Private Sub Button7_Click_1(sender As Object, e As EventArgs)
        Guna2CustomGradientPanel1.Visible = Not Guna2CustomGradientPanel1.Visible
        DataGridView1.ClearSelection()
        Clear()
    End Sub
    Private Sub Guna2Button2_Click(sender As Object, e As EventArgs) Handles Guna2Button2.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.SelectedRows(0)

            Dim Barcode As String = Guna2TextBox2.Text
            Dim NewProductStatus As String = If(Guna2ComboBox1.SelectedItem IsNot Nothing, Guna2ComboBox1.SelectedItem.ToString(), String.Empty)
            Dim Category As String = If(Guna2ComboBox3.SelectedItem IsNot Nothing, Guna2ComboBox3.SelectedItem.ToString(), String.Empty)
            Dim Brand As String = If(Guna2ComboBox4.SelectedItem IsNot Nothing, Guna2ComboBox4.SelectedItem.ToString(), String.Empty)
            Dim Supplier As String = If(Guna2ComboBox5.SelectedItem IsNot Nothing, Guna2ComboBox5.SelectedItem.ToString(), String.Empty)
            Dim Model As String = If(Guna2ComboBox6.SelectedItem IsNot Nothing, Guna2ComboBox6.SelectedItem.ToString(), String.Empty)
            Dim ProductName As String = If(Guna2ComboBox7.SelectedItem IsNot Nothing, Guna2ComboBox7.SelectedItem.ToString(), String.Empty)
            Dim Price As Decimal = Convert.ToDecimal(Guna2TextBox3.Text)
            Dim NewQuantity As Integer = Convert.ToInt32(Guna2TextBox4.Text)
            Dim CapitalPrice As Decimal = Convert.ToDecimal(Guna2TextBox6.Text)
            Dim VatStatus As String = If(Guna2CustomRadioButton1.Checked, "VAT", "NonVAT")

            ' From selected row
            Dim OldQuantity As Integer = Convert.ToInt32(selectedRow.Cells("Quantity").Value)
            Dim OldExpense As Decimal = Convert.ToDecimal(selectedRow.Cells("Expense").Value)
            Dim OldStatus As String = selectedRow.Cells("ProductStatus").Value.ToString()
            Dim OldCapitalPrice As Decimal = Convert.ToDecimal(selectedRow.Cells("CapitalPrice").Value)

            ' Compute updated values
            Dim QuantityDifference As Integer = NewQuantity - OldQuantity
            Dim NewExpense As Decimal = NewQuantity * CapitalPrice
            Dim TotalExpense As Decimal = OldExpense + (CapitalPrice * QuantityDifference)

            ' Determine if we need to adjust InventoryRecord
            Dim QuantityAdjust As Integer = 0
            Dim ExpenseAdjust As Decimal = 0D

            If (OldStatus = "LOSS PRODUCT" OrElse OldStatus = "DAMAGED" OrElse OldStatus = "INACTIVE") AndAlso NewProductStatus = "ACTIVE" Then
                ' Status restored to ACTIVE → add back quantity/expense
                QuantityAdjust = OldQuantity
                ExpenseAdjust = OldQuantity * OldCapitalPrice

            ElseIf NewProductStatus = "LOSS PRODUCT" OrElse NewProductStatus = "DAMAGED" OrElse NewProductStatus = "INACTIVE" Then
                ' Status changed to loss/inactive/damaged → subtract current quantity
                QuantityAdjust = -OldQuantity
                ExpenseAdjust = -(OldQuantity * CapitalPrice)

            ElseIf NewProductStatus = "ACTIVE" AndAlso OldStatus = "ACTIVE" AndAlso NewQuantity <> OldQuantity Then
                ' Quantity changed while active → adjust difference
                QuantityAdjust = QuantityDifference
                ExpenseAdjust = QuantityDifference * CapitalPrice
            End If

            ' SQL Queries
            Dim queryUpdateInventory As String = "
            UPDATE InventoryTb 
            SET Price = @Price, Quantity = @NewQuantity, Expense = @TotalExpense, 
                CapitalPrice = @CapitalPrice, DateIn = GETDATE(), Supplier = @Supplier, 
                VatStatus = @VatStatus, ProductStatus = @ProductStatus
            WHERE Barcode = @Barcode"

            Dim queryUpdateInventoryRecord As String = "
            UPDATE InventoryRecord 
            SET Quantity = Quantity + @QuantityAdjust, 
                Expense = Expense + @ExpenseAdjust, 
                Price = @Price, CapitalPrice = @CapitalPrice, VatStatus = @VatStatus 
            WHERE Barcode = @Barcode"

            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Using transaction = con.BeginTransaction()
                    Try
                        ' Update InventoryTb
                        Using cmd As New SqlCommand(queryUpdateInventory, con, transaction)
                            cmd.Parameters.AddWithValue("@Barcode", Barcode)
                            cmd.Parameters.AddWithValue("@Category", Category)
                            cmd.Parameters.AddWithValue("@Brand", Brand)
                            cmd.Parameters.AddWithValue("@Model", Model)
                            cmd.Parameters.AddWithValue("@ProductName", ProductName)
                            cmd.Parameters.AddWithValue("@NewQuantity", NewQuantity)
                            cmd.Parameters.AddWithValue("@TotalExpense", TotalExpense)
                            cmd.Parameters.AddWithValue("@Price", Price)
                            cmd.Parameters.AddWithValue("@CapitalPrice", CapitalPrice)
                            cmd.Parameters.AddWithValue("@Supplier", Supplier)
                            cmd.Parameters.AddWithValue("@VatStatus", VatStatus)
                            cmd.Parameters.AddWithValue("@ProductStatus", NewProductStatus)
                            cmd.ExecuteNonQuery()
                        End Using

                        ' Update InventoryRecord only if adjustment is not zero
                        If QuantityAdjust <> 0 OrElse ExpenseAdjust <> 0 Then
                            Using cmd As New SqlCommand(queryUpdateInventoryRecord, con, transaction)
                                cmd.Parameters.AddWithValue("@Barcode", Barcode)
                                cmd.Parameters.AddWithValue("@QuantityAdjust", QuantityAdjust)
                                cmd.Parameters.AddWithValue("@ExpenseAdjust", ExpenseAdjust)
                                cmd.Parameters.AddWithValue("@Price", Price)
                                cmd.Parameters.AddWithValue("@CapitalPrice", CapitalPrice)
                                cmd.Parameters.AddWithValue("@VatStatus", VatStatus)
                                cmd.ExecuteNonQuery()
                            End Using
                        End If

                        transaction.Commit()
                        Guna2CustomGradientPanel1.Visible = False
                        BindData()
                        MessageBox.Show("The product has been updated successfully.", "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Catch ex As SqlException
                        transaction.Rollback()
                        MessageBox.Show("SQL error occurred: " & ex.Message, "SQL Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Catch ex As Exception
                        transaction.Rollback()
                        MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End Try
                End Using
            End Using

            Clear()
        Else
            MessageBox.Show("Please select a row to update.", "Selection Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub




    Private Sub Guna2Button4_Click(sender As Object, e As EventArgs) Handles Guna2Button4.Click
        Try
            ' Create a new SaveFileDialog to choose file location and name
            Using saveFileDialog As New SaveFileDialog() With {
            .Filter = "PNG Image|*.png|JPEG Image|*.jpg",
            .Title = "Save Barcode",
            .FileName = "barcode" ' Default file name
        }

                ' Show the dialog and check if the user selected a file
                If saveFileDialog.ShowDialog() = DialogResult.OK Then
                    ' Create a new instance of the BarcodeEncoder
                    Dim encoder As New BarcodeEncoder With {
                    .IncludeLabel = True,
                    .CustomLabel = Guna2TextBox2.Text ' Set the barcode label from Guna2TextBox2
                }

                    ' Generate the barcode image (Code128 format)
                    Dim barcodeImage As Image = encoder.Encode(BarcodeFormat.Code128, Guna2TextBox2.Text)

                    ' Save the barcode as PNG or JPEG depending on the file extension selected
                    Dim fileExtension As String = System.IO.Path.GetExtension(saveFileDialog.FileName).ToLower()

                    If fileExtension = ".png" Then
                        barcodeImage.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Png)
                    ElseIf fileExtension = ".jpg" Then
                        barcodeImage.Save(saveFileDialog.FileName, System.Drawing.Imaging.ImageFormat.Jpeg)
                    End If

                    ' Confirmation message
                    MessageBox.Show("Barcode saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            ' Handle any errors that may occur during the barcode generation or saving
            MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Guna2TextBox4_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox4.TextChanged
        MultiplyValues()
    End Sub

    Private Sub Guna2TextBox6_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox6.TextChanged
        MultiplyValues()
    End Sub

    Private Sub Guna2TextBox3_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles Guna2TextBox3.KeyPress
        ' Allow only numbers and control keys (like Backspace)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub Guna2TextBox4_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles Guna2TextBox4.KeyPress
        ' Allow only numbers and control keys (like Backspace)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub Guna2TextBox5_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles Guna2TextBox5.KeyPress
        ' Allow only numbers and control keys (like Backspace)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub Guna2TextBox6_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles Guna2TextBox6.KeyPress
        ' Allow only numbers and control keys (like Backspace)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Guna2ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Guna2ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub Guna2ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Guna2ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub Guna2CustomGradientPanel1_Paint(sender As Object, e As PaintEventArgs) Handles Guna2CustomGradientPanel1.Paint

    End Sub

    Private Sub Guna2CircleButton1_Click(sender As Object, e As EventArgs) Handles Guna2CircleButton1.Click
        Dim rnd As New Random()
        Dim barcode As String = rnd.Next(1000, 9999).ToString() & DateTime.Now.ToString("yyyyMMddHHmmss")


        Guna2TextBox2.Text = barcode
    End Sub

    Private Sub Guna2CircleButton2_Click(sender As Object, e As EventArgs) Handles Guna2CircleButton2.Click
        Guna2CustomGradientPanel1.Visible = Not Guna2CustomGradientPanel1.Visible
        DataGridView1.ClearSelection()
        Clear()
    End Sub

    Private Sub Guna2TextBox3_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox3.TextChanged

    End Sub
End Class