Imports System.Data.SqlClient

Public Class Form7


    Private Sub Form7_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        GenerateDailySalesReport()
        LoadSalesReport()
        InitializeListView()
        Panel1.Visible = False

    End Sub

    Private Sub GenerateDailySalesReport()
        ' Check if today's report already exists
        Dim checkQuery As String = "SELECT COUNT(*) FROM SalesReportTb WHERE ReportDate = CAST(GETDATE() AS DATE)"
        Dim reportExists As Boolean = False

        Using con As New SqlConnection(ConnectionString)
            Using checkCmd As New SqlCommand(checkQuery, con)
                Try
                    con.Open()
                    reportExists = Convert.ToInt32(checkCmd.ExecuteScalar()) > 0
                Catch ex As Exception
                    MessageBox.Show("An error occurred while checking for existing reports: " & ex.Message)
                    Return
                End Try
            End Using
        End Using

        Dim query As String

        If reportExists Then
            ' Update the existing report for today
            query = "
        UPDATE SalesReportTb
        SET 
            TotalSales = COALESCE((SELECT SUM(TotalPrice) FROM PaymentTb WHERE CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE)), 0),
            TotalItemSold = COALESCE((SELECT SUM(Quantity) FROM PaymentTb WHERE CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE)), 0),
            MostSoldProduct = (
                SELECT TOP 1 Barcode 
                FROM PaymentTb 
                WHERE CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE) 
                GROUP BY Barcode 
                ORDER BY SUM(Quantity) DESC
            ),
            HighestRevenueProduct = (
                SELECT TOP 1 Barcode 
                FROM PaymentTb 
                WHERE CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE) 
                GROUP BY Barcode 
                ORDER BY SUM(TotalPrice * Quantity) DESC
            )
        WHERE 
            ReportDate = CAST(GETDATE() AS DATE);
        "
        Else
            ' Insert a new report for today 
            query = "
        INSERT INTO SalesReportTb (ReportDate, TotalSales, TotalItemSold, MostSoldProduct, HighestRevenueProduct)
        SELECT 
            CAST(GETDATE() AS DATE) AS ReportDate,
            COALESCE(SUM(TotalPrice), 0) AS TotalSales,
            COALESCE(SUM(Quantity), 0) AS TotalItemSold,
            (
                SELECT TOP 1 Barcode 
                FROM PaymentTb 
                WHERE CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE) 
                GROUP BY Barcode 
                ORDER BY SUM(Quantity) DESC
            ) AS MostSoldProduct,
            (
                SELECT TOP 1 Barcode 
                FROM PaymentTb 
                WHERE CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE) 
                GROUP BY Barcode 
                ORDER BY SUM(TotalPrice * Quantity) DESC
            ) AS HighestRevenueProduct
        FROM 
            PaymentTb
        WHERE 
            CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE);
        "
        End If

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    LoadSalesReport()
                    UpdateSalesAndProfitTotals()
                Catch ex As SqlException
                    MessageBox.Show("SQL error occurred: " & ex.Message)
                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub


    Private Sub UpdateSalesAndProfitTotals()
        ' Queries for total sales
        Dim todaySalesQuery As String = "SELECT COALESCE(SUM(TotalPrice), 0) FROM PaymentTb WHERE CAST(DateIssued AS DATE) = CAST(GETDATE() AS DATE)"
        Dim monthlySalesQuery As String = "SELECT COALESCE(SUM(TotalPrice), 0) FROM PaymentTb WHERE YEAR(DateIssued) = YEAR(GETDATE()) AND MONTH(DateIssued) = MONTH(GETDATE())"
        Dim yearlySalesQuery As String = "SELECT COALESCE(SUM(TotalPrice), 0) FROM PaymentTb WHERE YEAR(DateIssued) = YEAR(GETDATE())"
        Dim salesQuery As String = "SELECT COALESCE(SUM(TotalPrice), 0) FROM PaymentTb"

        ' Queries for total capital cost
        Dim todayCapitalQuery As String = "SELECT COALESCE(SUM(IR.CapitalPrice * PT.Quantity), 0) " &
                                   "FROM PaymentTb PT " &
                                   "INNER JOIN InventoryRecord IR ON PT.Barcode = IR.Barcode " &
                                   "WHERE CAST(PT.DateIssued AS DATE) = CAST(GETDATE() AS DATE)"
        Dim monthlyCapitalQuery As String = "SELECT COALESCE(SUM(IR.CapitalPrice * PT.Quantity), 0) " &
                                     "FROM PaymentTb PT " &
                                     "INNER JOIN InventoryRecord IR ON PT.Barcode = IR.Barcode " &
                                     "WHERE YEAR(PT.DateIssued) = YEAR(GETDATE()) AND MONTH(PT.DateIssued) = MONTH(GETDATE())"
        Dim yearlyCapitalQuery As String = "SELECT COALESCE(SUM(IR.CapitalPrice * PT.Quantity), 0) " &
                                    "FROM PaymentTb PT " &
                                    "INNER JOIN InventoryRecord IR ON PT.Barcode = IR.Barcode " &
                                    "WHERE YEAR(PT.DateIssued) = YEAR(GETDATE())"
        Dim salesCapitalQuery As String = "SELECT COALESCE(SUM(IR.CapitalPrice * PT.Quantity), 0) " &
                                    "FROM PaymentTb PT " &
                                    "INNER JOIN InventoryRecord IR ON PT.Barcode = IR.Barcode"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand()
                cmd.Connection = con
                con.Open()

                ' Retrieve total sales
                cmd.CommandText = todaySalesQuery
                Dim todaySales As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                cmd.CommandText = monthlySalesQuery
                Dim monthlySales As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                cmd.CommandText = yearlySalesQuery
                Dim yearlySales As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                cmd.CommandText = salesQuery
                Dim totalSales As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                ' Retrieve total capital costs
                cmd.CommandText = todayCapitalQuery
                Dim todayCapital As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                cmd.CommandText = monthlyCapitalQuery
                Dim monthlyCapital As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                cmd.CommandText = yearlyCapitalQuery
                Dim yearlyCapital As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                cmd.CommandText = salesCapitalQuery
                Dim totalCapital As Decimal = Convert.ToDecimal(cmd.ExecuteScalar())

                ' Calculate profits
                Dim todayProfit As Decimal = todaySales - todayCapital
                Dim monthlyProfit As Decimal = monthlySales - monthlyCapital
                Dim yearlyProfit As Decimal = yearlySales - yearlyCapital
                Dim totalProfit As Decimal = totalSales - totalCapital

                ' Update labels
                Label7.Text = todaySales.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Today's sales
                Label8.Text = monthlySales.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Monthly sales
                Label9.Text = yearlySales.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Yearly sales
                Label17.Text = totalSales.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Total sales

                ' Update profit labels
                Label11.Text = todayProfit.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Today's profit
                Label12.Text = monthlyProfit.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Monthly profit
                Label13.Text = yearlyProfit.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Yearly profit
                Label18.Text = totalProfit.ToString("C", Globalization.CultureInfo.CreateSpecificCulture("fil-PH")) ' Total profit

                con.Close()
            End Using
        End Using
    End Sub




    Private Sub LoadSalesReport()
        Dim query As String = "SELECT ReportDate, TotalSales, TotalItemSold, MostSoldProduct, HighestRevenueProduct FROM SalesReportTb ORDER BY Id DESC"
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

    Public Sub FilterByDateRange()
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()

                ' Define the SQL query to filter by date range
                Dim query As String = "SELECT  ReportDate, TotalSales, TotalItemSold, MostSoldProduct, HighestRevenueProduct FROM SalesReportTb WHERE ReportDate BETWEEN @StartDate AND @EndDate"
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

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        LoadSalesReport()
        GenerateDailySalesReport()
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
    Private Sub InitializeListView()
        ' Configure ListView columns for vertical display
        ListView1.View = View.Details
        ListView1.Columns.Clear()

        ' First column for labels (e.g., "Brand", "Model", etc.)
        ListView1.Columns.Add("Label", 150, HorizontalAlignment.Left)

        ' Second column for details of the first barcode (MostSoldProduct)
        ListView1.Columns.Add("Most Sold Product", 250, HorizontalAlignment.Left)

        ' Third column for details of the second barcode (HighestRevenueProduct)
        ListView1.Columns.Add("Highest Revenue Product", 250, HorizontalAlignment.Left)
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Ensure the click is on a valid row
        If e.RowIndex >= 0 Then
            ' Reinitialize ListView columns
            InitializeListView()

            ' Retrieve the clicked row
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            Dim selectedBarcode As String = row.Cells("MostSoldProduct").Value.ToString()
            Dim selectedBarcode1 As String = row.Cells("HighestRevenueProduct").Value.ToString()



            ' Clear ListView before adding new details
            ListView1.Items.Clear()



            ' Fetch details for each barcode
            Dim detailsBarcode1 As Dictionary(Of String, String) = FetchBarcodeDetails(selectedBarcode)
            Dim detailsBarcode2 As Dictionary(Of String, String) = FetchBarcodeDetails(selectedBarcode1)

            ' Ensure both dictionaries have data
            If detailsBarcode1.Count = 0 OrElse detailsBarcode2.Count = 0 Then
                MessageBox.Show("No details found for one of the barcodes.")
            End If

            ' Add the details to the ListView side by side
            AddDetailsToListView(detailsBarcode1, detailsBarcode2)

            ' Show the panel with the ListView
            Panel1.Visible = True
        End If
    End Sub

    Private Function FetchBarcodeDetails(barcode As String) As Dictionary(Of String, String)
        ' Create a dictionary to store the details
        Dim details As New Dictionary(Of String, String)

        ' SQL query to get details for the selected Barcode
        Dim query As String = "
    SELECT i.Barcode, i.Category, i.ProductName, i.Model, i.Brand
    FROM InventoryTb AS i
    WHERE i.Barcode = @Barcode"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Barcode", barcode)

                Try
                    con.Open()
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        If reader.HasRows Then
                            ' Read the data and store it in the dictionary
                            While reader.Read()
                                details("Barcode") = reader("Barcode").ToString()
                                details("Category") = reader("Category").ToString()
                                details("ProductName") = reader("ProductName").ToString()
                                details("Brand") = reader("Brand").ToString()
                                details("Model") = reader("Model").ToString()
                            End While
                        End If
                    End Using
                Catch ex As Exception
                    MessageBox.Show("An error occurred while fetching barcode details: " & ex.Message)
                End Try
            End Using
        End Using

        Return details
    End Function

    Private Sub AddDetailsToListView(details1 As Dictionary(Of String, String), details2 As Dictionary(Of String, String))
        ' Ensure both dictionaries are populated
        Dim fields As String() = {"Barcode", "Category", "ProductName", "Brand", "Model"}

        ' Loop through each field and add it to the ListView
        For Each field In fields
            Dim item As New ListViewItem(field)

            ' Add the first barcode details to the second column
            If details1.ContainsKey(field) Then
                item.SubItems.Add(details1(field))
            Else
                item.SubItems.Add("") ' Add an empty value if no data
            End If

            ' Add the second barcode details to the third column
            If details2.ContainsKey(field) Then
                item.SubItems.Add(details2(field))
            Else
                item.SubItems.Add("") ' Add an empty value if no data
            End If

            ' Add the item to the ListView
            ListView1.Items.Add(item)
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        ' Hide the panel when the button is clicked
        Panel1.Visible = False
    End Sub

End Class
