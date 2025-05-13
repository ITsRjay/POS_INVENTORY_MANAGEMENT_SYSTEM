Imports System.Data.SqlClient
Public Class Form13
    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Refund()

    End Sub
    Public Sub Refund()
        Dim query As String = "
SELECT 
    r.Id,
    r.PaymentID,
    i.ProductName,
    i.Brand,
    r.RefundCause,
    r.RefundAmount,
    r.Quantity,
    r.HandledBy,
    r.RefundDate
FROM 
    RefundTb r
INNER JOIN 
    PaymentTb p ON r.PaymentID = p.Id
INNER JOIN 
    InventoryTb i ON p.Barcode = i.Barcode
ORDER BY 
    r.Id DESC"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Dim dt As New DataTable()
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using

                DataGridView1.DataSource = dt

                If DataGridView1.Columns.Contains("Id") Then
                    DataGridView1.Columns("Id").Visible = False
                End If

                If DataGridView1.Columns.Contains("PaymentID") Then
                    DataGridView1.Columns("PaymentID").Visible = False
                End If
            End Using
        End Using
    End Sub


    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Refund()
    End Sub

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

    Public Sub FilterByDateRange()
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()

                Dim query As String = "
            SELECT 
                r.Id,
                r.PaymentID, 
                i.ProductName,
                i.Brand,
                r.RefundCause,
                r.RefundAmount,
                r.Quantity,
                r.RefundDate
            FROM 
                RefundTb r
            INNER JOIN 
                PaymentTb p ON r.PaymentID = p.Id
            INNER JOIN 
                InventoryTb i ON p.Barcode = i.Barcode
            WHERE 
                r.RefundDate BETWEEN @StartDate AND @EndDate
            ORDER BY 
                r.Id DESC"

                Using cmd As New SqlCommand(query, con)
                    cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Value.Date)
                    cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Value.Date.AddDays(1).AddTicks(-1)) ' End of selected day

                    Using da As New SqlDataAdapter(cmd)
                        Using dt As New DataTable()
                            da.Fill(dt)
                            DataGridView1.DataSource = dt

                            If DataGridView1.Columns.Contains("Id") Then
                                DataGridView1.Columns("Id").Visible = False
                            End If

                            If DataGridView1.Columns.Contains("PaymentID") Then
                                DataGridView1.Columns("PaymentID").Visible = False
                            End If
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


    Private Sub SearchByRefundBarcode(RefundBarcodeNo As String)
        ' Check if the barcode is empty
        If String.IsNullOrWhiteSpace(RefundBarcodeNo) Then
            MessageBox.Show("Please enter a barcode before searching.", "Search Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim query As String = "SELECT  Id, PaymentID, RefundCause, RefundAmount, Quantity, RefundDate FROM RefundTb WHERE RefundBarcodeNo = @RefundBarcodeNo"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@RefundBarcodeNo", RefundBarcodeNo)

                Dim dt As New DataTable()
                Using da As New SqlDataAdapter(cmd)

                    con.Open()
                    da.Fill(dt)
                    Refund()

                    If DataGridView1.Columns.Contains("Id") Then
                        DataGridView1.Columns("Id").Visible = False
                    End If

                    If DataGridView1.Columns.Contains("PaymentID") Then
                        DataGridView1.Columns("PaymentID").Visible = False
                    End If

                End Using

                If dt.Rows.Count > 0 Then
                    DataGridView1.DataSource = dt
                Else
                    MessageBox.Show("No records found for the given barcode.", "Search Result", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        End Using
    End Sub

End Class