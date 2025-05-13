Imports System.Data.SqlClient
Imports System.Drawing.Printing
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form15

    Private userRole As String
    Private currentCashierName As String

    ' Constructor to receive role and fullName
    Public Sub New(role As String, fullName As String)
        InitializeComponent()
        userRole = role
        currentCashierName = fullName
    End Sub

    Private Sub Form15_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateComboBox()
        InitializeListView()
        ApplyFilters()
        DataGridView1.Columns("DateIssued").DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss"
        Panel1.Visible = False


        ' Use the name passed in constructor
        SetFullName(currentCashierName)
    End Sub

    Public Sub SetFullName(fullName As String)
        Label16.Text = fullName
        LoadData(handledBy:=fullName)
    End Sub



    Public Sub LoadData(Optional mopFilter As String = "", Optional startDate As Date = Nothing, Optional endDate As Date = Nothing, Optional handledBy As String = "")
        Dim query As String = "SELECT i.Barcode, i.Category, i.Brand, i.ProductName, i.Model, p.CustomerName, p.Quantity, p.TotalPrice, p.MOP, p.ReceiptBarcodeNo, p.HandledBy, p.DateIssued " &
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

        ' Add HandledBy filter if provided
        If Not String.IsNullOrEmpty(handledBy) Then
            whereConditions.Add("p.HandledBy = @HandledBy")
        End If

        ' Combine conditions
        If whereConditions.Count > 0 Then
            query &= " WHERE " & String.Join(" AND ", whereConditions)
        End If

        query &= " ORDER BY p.Id DESC"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                If Not String.IsNullOrEmpty(mopFilter) AndAlso mopFilter <> "All" Then
                    cmd.Parameters.AddWithValue("@MOP", mopFilter)
                End If

                If startDate <> Date.MinValue AndAlso endDate <> Date.MinValue Then
                    cmd.Parameters.AddWithValue("@StartDate", startDate.Date)
                    cmd.Parameters.AddWithValue("@EndDate", endDate.Date.AddDays(1).AddTicks(-1))
                End If

                If Not String.IsNullOrEmpty(handledBy) Then
                    cmd.Parameters.AddWithValue("@HandledBy", handledBy)
                End If

                Using da As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
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
        AddListViewItem("Customer Name:", row.Cells("CustomerName").Value.ToString())
        AddListViewItem("Quantity Purchase:", row.Cells("Quantity").Value.ToString())
        AddListViewItem("Total Price:", row.Cells("TotalPrice").Value.ToString())
        AddListViewItem("Receipt Barcode No:", row.Cells("ReceiptBarcodeNo").Value.ToString())
        AddListViewItem("Date Issued:", row.Cells("DateIssued").Value.ToString())
        AddListViewItem("Cashier:", row.Cells("HandledBy").Value.ToString())
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
        LoadData(handledBy:=Label16.Text)
    End Sub



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
            If String.IsNullOrWhiteSpace(currentCashierName) Then
                MessageBox.Show("Cashier is not set.", "Missing Cashier", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Using con As New SqlConnection(ConnectionString)
                con.Open()

                Dim query As String = "SELECT i.Barcode, i.Category, i.Brand, i.ProductName, i.Model, p.CustomerName, " &
                                  "p.Quantity, p.TotalPrice, p.ReceiptBarcodeNo, p.HandledBy, p.DateIssued " &
                                  "FROM InventoryTb AS i " &
                                  "JOIN PaymentTb AS p ON i.Barcode = p.Barcode " &
                                  "WHERE p.DateIssued BETWEEN @StartDate AND @EndDate " &
                                  "AND p.HandledBy = @HandledBy " &
                                  "ORDER BY p.Id DESC"

                Using cmd As New SqlCommand(query, con)
                    cmd.Parameters.AddWithValue("@StartDate", DateTimePicker1.Value.Date)
                    cmd.Parameters.AddWithValue("@EndDate", DateTimePicker2.Value.Date.AddDays(1).AddTicks(-1))
                    cmd.Parameters.AddWithValue("@HandledBy", currentCashierName)

                    Using da As New SqlDataAdapter(cmd)
                        Dim dt As New DataTable()
                        da.Fill(dt)
                        DataGridView1.DataSource = dt
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message, "Filter Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ApplyFilters()
        If ComboBox1.SelectedItem Is Nothing Then Exit Sub

        Dim selectedMOP As String = ComboBox1.SelectedItem.ToString()
        Dim startDate As Date = DateTimePicker1.Value
        Dim endDate As Date = DateTimePicker2.Value

        If selectedMOP = "All" Then
            LoadData(startDate:=startDate, endDate:=endDate, handledBy:=currentCashierName)
        Else
            LoadData(mopFilter:=selectedMOP, startDate:=startDate, endDate:=endDate, handledBy:=currentCashierName)
        End If
    End Sub

    ' Trigger ApplyFilters on date changes
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        ApplyFilters()
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        ApplyFilters()
    End Sub

    ' Trigger MOP filter change
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        ApplyFilters()
    End Sub


End Class
