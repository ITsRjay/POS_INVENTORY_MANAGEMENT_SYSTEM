Imports System.Data.SqlClient

Public Class Form11
    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InventoryRecord()
    End Sub
    Public Sub InventoryRecord()
        Dim query As String = "SELECT Barcode, Category, ProductName, Brand, Model, CapitalPrice, Price, Quantity, Expense, VatStatus, DateIn, Supplier FROM InventoryRecord ORDER BY DateIn DESC"
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Dim dt As New DataTable()
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
                DataGridView1.DataSource = dt
            End Using
        End Using

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs)
        InventoryRecord()
    End Sub

    Private Sub SearchByBarcode(barcode As String)
        ' Check if the barcode is empty
        If String.IsNullOrWhiteSpace(barcode) Then
            MessageBox.Show("Please enter a barcode before searching.", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim query As String = "SELECT Barcode, Category, ProductName, Brand, Model, CapitalPrice, Price, Quantity, Expense, VatStatus, DateIn, Supplier FROM InventoryRecord WHERE Barcode = @Barcode"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Barcode", barcode)

                Dim dt As New DataTable()
                Using da As New SqlDataAdapter(cmd)
                    Try
                        con.Open()
                        da.Fill(dt)
                        InventoryRecord()
                    Catch ex As Exception
                        MessageBox.Show("An error occurred while searching: " & ex.Message, "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return
                    End Try
                End Using

                ' Check if any records were found
                If dt.Rows.Count > 0 Then
                    DataGridView1.DataSource = dt
                Else
                    MessageBox.Show("No records found for the given barcode.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim barcodeInput As String = Guna2TextBox1.Text.Trim() ' Assuming you have a TextBox for input
        If Not String.IsNullOrEmpty(barcodeInput) Then
            SearchByBarcode(barcodeInput)
        Else
            MessageBox.Show("Please enter a barcode to search.")
        End If
    End Sub

    Public Sub FilterByDateRange()
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()

                ' Define the SQL query to filter by date range
                Dim query As String = "SELECT Barcode, Category, ProductName, Brand, Model, CapitalPrice, Price, Quantity, Expense, VatStatus, DateIn, Supplier FROM InventoryRecord WHERE DateIn BETWEEN @StartDate AND @EndDate"
                Using cmd As New SqlCommand(query, con)
                    ' Add parameters for the start and end dates
                    cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Value.Date)
                    cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Value.Date.AddDays(0).AddTicks(-1)) ' End of the day

                    Using da As New SqlDataAdapter(cmd)
                        Using dt As New DataTable()
                            da.Fill(dt)
                            DataGridView1.DataSource = dt
                        End Using
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message, "Filter Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        FilterByDateRange()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        FilterByDateRange()
    End Sub
    Private Sub Guna2TextBox1_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox1.TextChanged
        Module1.FilterDataGridView1(Guna2TextBox1.Text, DataGridView1, ConnectionString)
    End Sub

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

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        InventoryRecord()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        SaveDataGridViewToCSV()
    End Sub
    ' Function to Save DataGridView to CSV File
    Private Sub SaveDataGridViewToCSV()
        ' Initialize and configure the SaveFileDialog
        Dim sfd As New SaveFileDialog With {
        .Filter = "CSV files (*.csv)|*.csv", ' Only allow CSV files
        .Title = "Save a CSV File",
        .FileName = "PrintedDataGridView.csv" ' Default file name
    }

        ' Show the dialog and continue only if the user clicks Save
        If sfd.ShowDialog() = DialogResult.OK Then
            Try
                ' Create a StreamWriter to write to the selected file path
                Using writer As New System.IO.StreamWriter(sfd.FileName)
                    ' Write column headers
                    For col As Integer = 0 To DataGridView1.Columns.Count - 1
                        writer.Write(DataGridView1.Columns(col).HeaderText)
                        If col < DataGridView1.Columns.Count - 1 Then writer.Write(",")
                    Next
                    writer.WriteLine()

                    ' Write rows
                    For row As Integer = 0 To DataGridView1.Rows.Count - 1
                        For col As Integer = 0 To DataGridView1.Columns.Count - 1
                            writer.Write(DataGridView1.Rows(row).Cells(col).Value?.ToString())
                            If col < DataGridView1.Columns.Count - 1 Then writer.Write(",")
                        Next
                        writer.WriteLine()
                    Next
                End Using

                ' Notify success
                MessageBox.Show("Data saved to CSV file successfully!", "Save Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Catch ex As Exception
                MessageBox.Show("An error occurred while saving the data: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        InventoryRecord()
        Guna2TextBox1.Clear()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub
End Class