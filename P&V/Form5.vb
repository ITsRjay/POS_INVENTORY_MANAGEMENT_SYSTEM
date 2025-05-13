Imports System.Data.SqlClient
Imports System.Drawing.Printing
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form5

    Public Sub Form5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadData()
        PopulateComboBox()
        InitializeListView()
        DataGridView1.Columns("DateIssued").DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss"
        Panel1.Visible = False
    End Sub
    Public Sub LoadData(Optional mopFilter As String = "", Optional startDate As Date = Nothing, Optional endDate As Date = Nothing)
        Dim query As String = "SELECT i.Barcode, i.Category, i.Brand, i.ProductName, i.Model, p.CustomerName, p.MOP, p.Amount, p.Discount, p.Change, p.ReferenceNo, p.Quantity, i.Price, p.TotalPrice, p.HandledBy, p.ReceiptBarcodeNo, p.DateIssued " &
                          "FROM InventoryTb AS i " &
                          "JOIN PaymentTb AS p ON i.Barcode = p.Barcode"

        Dim whereConditions As New List(Of String)()

        ' Add Mode of Payment filter if provided and not "All"
        If Not String.IsNullOrEmpty(mopFilter) AndAlso mopFilter <> "All" Then
            whereConditions.Add("p.MOP = @MOP")
        End If

        ' Add Date Range filter if both dates are provided
        If startDate <> Date.MinValue AndAlso endDate <> Date.MinValue Then
            whereConditions.Add("p.DateIssued BETWEEN @StartDate AND @EndDate")
        End If

        ' Append WHERE conditions if any exist
        If whereConditions.Count > 0 Then
            query &= " WHERE " & String.Join(" AND ", whereConditions)
        End If

        ' Append ORDER BY clause
        query &= " ORDER BY p.Id DESC"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                ' Add parameters if filters are applied
                If Not String.IsNullOrEmpty(mopFilter) AndAlso mopFilter <> "All" Then
                    cmd.Parameters.AddWithValue("@MOP", mopFilter)
                End If

                If startDate <> Date.MinValue AndAlso endDate <> Date.MinValue Then
                    cmd.Parameters.AddWithValue("@StartDate", startDate.Date)
                    cmd.Parameters.AddWithValue("@EndDate", endDate.Date.AddDays(1).AddTicks(-1)) ' End of the selected end date
                End If

                ' Execute the query and fill the DataGridView
                Using da As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt ' Bind the DataGridView to the DataTable
                    End Using
                End Using
            End Using
        End Using
    End Sub


    Public Sub PopulateComboBox()
        Using con As New SqlConnection(ConnectionString)
            con.Open()

            ' Query to get distinct MOP values from the PaymentTb
            Dim query As String = "SELECT DISTINCT MOP FROM PaymentTb"
            Using cmd As New SqlCommand(query, con)
                Using da As New SqlDataAdapter(cmd)
                    Dim dt As New DataTable()
                    da.Fill(dt)

                    ComboBox1.Items.Clear() ' Clear existing items before adding new ones

                    ' Add MOP values to ComboBox1
                    For Each row As DataRow In dt.Rows
                        ComboBox1.Items.Add(row("MOP").ToString())
                    Next

                    ' Optionally, add a default item for all MOP values (e.g., "All")
                    ComboBox1.Items.Insert(0, "All")
                    ComboBox1.SelectedIndex = 0 ' Select the first item by default
                End Using
            End Using
        End Using
    End Sub

    Private Sub Guna2TextBox1_TextChanged(sender As Object, e As EventArgs) Handles Guna2TextBox1.TextChanged
        Module1.FilterTransact(Guna2TextBox1.Text, DataGridView1, ConnectionString)
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            PopulateListView(row)
            Panel1.Visible = True
        End If
    End Sub

    Private Sub PopulateListView(row As DataGridViewRow)
        ListView1.Items.Clear()
        AddListViewItem("Barcode:", row.Cells("Barcode").Value.ToString())
        AddListViewItem("Category:", row.Cells("Category").Value.ToString())
        AddListViewItem("ProductName:", row.Cells("ProductName").Value.ToString())
        AddListViewItem("Brand:", row.Cells("Brand").Value.ToString())
        AddListViewItem("Model:", row.Cells("Model").Value.ToString())
        AddListViewItem("Cash Amount:", row.Cells("Amount").Value.ToString())
        AddListViewItem("Mode of Payment:", row.Cells("MOP").Value.ToString())
        AddListViewItem("Reference No.:", row.Cells("ReferenceNo").Value.ToString())
        AddListViewItem("Discount%:", row.Cells("Discount").Value.ToString())
        AddListViewItem("Customer Name:", row.Cells("CustomerName").Value.ToString())
        AddListViewItem("Quantity Purchase:", row.Cells("Quantity").Value.ToString())
        AddListViewItem("Item Price:", row.Cells("Price").Value.ToString())
        AddListViewItem("Total Price:", row.Cells("TotalPrice").Value.ToString())
        AddListViewItem("Change:", row.Cells("Change").Value.ToString())
        AddListViewItem("Cashier:", row.Cells("HandledBy").Value.ToString())
        AddListViewItem("Receipt Barcode No:", row.Cells("ReceiptBarcodeNo").Value.ToString())
        AddListViewItem("Date Issued:", row.Cells("DateIssued").Value.ToString())
    End Sub

    Private Sub AddListViewItem(label As String, value As String)
        Dim item As New ListViewItem(label)
        item.SubItems.Add(value)
        ListView1.Items.Add(item)
    End Sub

    Private Sub InitializeListView()

        ListView1.View = View.Details
        ListView1.Columns.Clear()
        ListView1.Columns.Add("Label", 150, HorizontalAlignment.Left)
        ListView1.Columns.Add("Value", 250, HorizontalAlignment.Left)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Panel1.Visible = False
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        ApplyFilters()
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Restrict to only 1-day export
        Dim startDate As Date = DateTimePicker1.Value.Date
        Dim endDate As Date = DateTimePicker2.Value.Date

        If (endDate - startDate).Days > 0 Then
            MessageBox.Show("You can only export one day at a time.", "Date Restriction", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' === Optional Preview ===
        If MessageBox.Show("Do you want to preview the sales summary before saving as PDF?", "Preview", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.WindowState = FormWindowState.Maximized
            PrintPreviewDialog1.ShowDialog()
        End If

        ' === Save as PDF ===
        Dim saveDialog As New SaveFileDialog()
        saveDialog.Filter = "PDF Files|*.pdf"
        saveDialog.Title = "Save PDF Report"
        saveDialog.FileName = "SalesSummary_" & startDate.ToString("yyyyMMdd") & ".pdf"

        If saveDialog.ShowDialog() = DialogResult.OK Then
            Dim pd As New PrintDocument()
            AddHandler pd.PrintPage, AddressOf PrintDocument1_PrintPage
            pd.DefaultPageSettings.PaperSize = New PaperSize("Receipt", 200, 1100)

            Dim ps As New PrinterSettings()
            ps.PrinterName = "Microsoft Print to PDF"
            ps.PrintToFile = True
            ps.PrintFileName = saveDialog.FileName
            pd.PrinterSettings = ps

            pd.Print()
            RemoveHandler pd.PrintPage, AddressOf PrintDocument1_PrintPage
        End If
    End Sub


    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim font As New Font("Arial", 6)
        Dim titleFont As New Font("Arial", 15, FontStyle.Bold)
        Dim headerFont As New Font("Arial", 8, FontStyle.Bold)
        Dim brush As New SolidBrush(Color.Black)

        Dim selectedDate As Date = DateTimePicker1.Value.Date
        Dim yPos As Integer = 5
        Dim paperWidth As Integer = e.PageBounds.Width

        ' Shrink total content width to center content more closely
        Dim contentWidth As Integer = paperWidth \ 1.5 ' Reduce overall width to about 66% of page
        Dim contentStartX As Integer = (paperWidth - contentWidth) \ 2

        ' Divide into 3 columns
        Dim columnWidth As Integer = contentWidth \ 3
        Dim col1X As Integer = contentStartX
        Dim col2X As Integer = col1X + columnWidth
        Dim col3X As Integer = col2X + columnWidth

        ' Helper to center text within a column
        Dim drawCenteredInColumn As Action(Of String, Font, Integer, Integer) = Sub(text, fnt, colX, width)
                                                                                    Dim textWidth As Single = e.Graphics.MeasureString(text, fnt).Width
                                                                                    Dim drawX As Single = colX + (width - textWidth) / 2
                                                                                    e.Graphics.DrawString(text, fnt, brush, drawX, yPos)
                                                                                End Sub

        ' Draw title
        yPos += 20
        Dim titleWidth As Single = e.Graphics.MeasureString("Sales Summary", titleFont).Width
        e.Graphics.DrawString("Sales Summary", titleFont, brush, (paperWidth - titleWidth) / 2, yPos)

        ' Draw centered date
        yPos += 30
        Dim dateText As String = "Date: " & selectedDate.ToShortDateString()
        Dim dateWidth As Single = e.Graphics.MeasureString(dateText, font).Width
        e.Graphics.DrawString(dateText, font, brush, (paperWidth - dateWidth) / 2, yPos)

        ' Separator line
        yPos += 30
        e.Graphics.DrawString(New String("-"c, 150), font, brush, contentStartX, yPos)

        ' Column headers
        yPos += 20
        drawCenteredInColumn("Product", headerFont, col1X, columnWidth)
        drawCenteredInColumn("Customer", headerFont, col2X, columnWidth)
        drawCenteredInColumn("Amount", headerFont, col3X, columnWidth)

        ' Separator line
        yPos += 20
        e.Graphics.DrawString(New String("-"c, 150), font, brush, contentStartX, yPos)

        ' Data rows
        Dim total As Decimal = 0
        Dim rowPrinted As Boolean = False

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim rowDate As Date = Convert.ToDateTime(row.Cells("DateIssued").Value).Date
                If rowDate = selectedDate Then
                    yPos += 20
                    drawCenteredInColumn(row.Cells("ProductName").Value.ToString(), font, col1X, columnWidth)
                    drawCenteredInColumn(row.Cells("CustomerName").Value.ToString(), font, col2X, columnWidth)
                    drawCenteredInColumn(Convert.ToDecimal(row.Cells("TotalPrice").Value).ToString("C"), font, col3X, columnWidth)

                    total += Convert.ToDecimal(row.Cells("TotalPrice").Value)
                    rowPrinted = True
                End If
            End If
        Next

        ' Footer
        yPos += 20
        e.Graphics.DrawString(New String("-"c, 150), font, brush, contentStartX, yPos)
        yPos += 25

        If rowPrinted Then
            drawCenteredInColumn("Grand Total:", headerFont, col2X, columnWidth)
            drawCenteredInColumn(total.ToString("C"), headerFont, col3X, columnWidth)
        Else
            Dim noDataMsg As String = "No sales data available for this date."
            Dim msgWidth As Single = e.Graphics.MeasureString(noDataMsg, headerFont).Width
            e.Graphics.DrawString(noDataMsg, headerFont, brush, (paperWidth - msgWidth) / 2, yPos)
        End If
    End Sub



    Public Sub FilterByDateRange()
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()

                ' Corrected SQL query to include the date filter
                Dim query As String = "SELECT i.Barcode, i.Category, i.Brand, i.ProductName, i.Model, p.CustomerName, p.MOP, p.Amount, p.Discount, p.Change, p.ReferenceNo, p.Quantity, i.Price, p.TotalPrice, p.HandledBy, p.ReceiptBarcodeNo, p.DateIssued " &
                                  "FROM InventoryTb AS i " &
                                  "JOIN PaymentTb AS p ON i.Barcode = p.Barcode " &
                                  "WHERE p.DateIssued BETWEEN @StartDate AND @EndDate " & ' Added space before ORDER BY
                                  "ORDER BY p.Id DESC"

                Using cmd As New SqlCommand(query, con)
                    ' Add parameters for the start and end dates
                    cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Value.Date)
                    cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Value.Date.AddDays(1).AddTicks(-1)) ' End of the day for the selected end date

                    Using da As New SqlDataAdapter(cmd)
                        Using dt As New DataTable()
                            da.Fill(dt)
                            ' Bind the data to DataGridView
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
        ApplyFilters() ' Call the function to filter both MOP and Date Range
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        ApplyFilters() ' Call the function to filter both MOP and Date Range
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ApplyFilters() ' Call the function to filter both MOP and Date Range
    End Sub

    ' Unified function to filter by both Mode of Payment (MOP) and Date Range
    Private Sub ApplyFilters()
        Dim selectedMOP As String = ComboBox1.SelectedItem.ToString()
        Dim startDate As Date = DateTimePicker1.Value
        Dim endDate As Date = DateTimePicker2.Value

        ' Call LoadData with both filters
        If selectedMOP = "All" Then
            LoadData(startDate:=startDate, endDate:=endDate) ' No MOP filter, only date range
        Else
            LoadData(mopFilter:=selectedMOP, startDate:=startDate, endDate:=endDate) ' Filter by MOP and Date Range
        End If
    End Sub


End Class
