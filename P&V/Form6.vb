Imports System.Data.SqlClient
Imports System.Drawing.Printing


Public Class Form6
    Private WithEvents PrintPreviewDialog1 As New PrintPreviewDialog
    Private barcodeClearTimer As Timer
    Private scanDelayTimer As Timer
    Private receiptBarcodeNo As String
    Private pendingBarcode As String


    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Focus()
        InventoryData()
        Panel1.Visible = False
        Panel4.Visible = False
        Panel5.Visible = False
        DataGridView1.Columns.Clear()
        DataGridView1.Columns.Add("Barcode", "Barcode")
        DataGridView1.Columns("Barcode").Visible = False
        DataGridView1.Columns.Add("Brand", "Brand")
        DataGridView1.Columns.Add("Model", "Model")
        DataGridView1.Columns.Add("ProductName", "ProductName")
        DataGridView1.Columns.Add("Quantity", "Quantity")
        DataGridView1.Columns.Add("Price", "Price")
        DataGridView1.Columns.Add("TotalPrice", "Total Price")
        DataGridView1.Columns.Add("VatStatus", "VatStatus")
        DataGridView1.Columns("VatStatus").Visible = False
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.WindowState = FormWindowState.Maximized

        ' Initialize timers
        barcodeClearTimer = New Timer() With {.Interval = 1000}
        AddHandler barcodeClearTimer.Tick, AddressOf BarcodeClearTimer_Tick

        scanDelayTimer = New Timer() With {.Interval = 200}
        AddHandler scanDelayTimer.Tick, AddressOf ScanDelayTimer_Tick

        Autocomplete()
        LoadCategories()
        TextBox6.Text = "Search by ProductName"
        TextBox6.ForeColor = Color.DarkGray
        TextBox5.Text = "QTY"
        TextBox5.ForeColor = Color.DarkGray
        Me.KeyPreview = True
        AddRemoveButtonColumn()
        Module1.LoadStoreDetails()
        SetFullName(userFullName)
    End Sub



    Public isPanel1Visible As Boolean = False

    Public Sub SetFullName(fullName As String)
        Label16.Text = fullName
    End Sub
    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        isPanel1Visible = Not isPanel1Visible

        If isPanel1Visible Then
            Panel1.Visible = True
            DataGridView1.Visible = False
        Else
            Panel1.Visible = False
            DataGridView1.Visible = True
        End If
    End Sub
    Private Sub Form6_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' Check if the Esc key is pressed

        If e.Control AndAlso e.KeyCode = Keys.L Then
            Button2.PerformClick()
        End If
        If e.Control AndAlso e.KeyCode = Keys.D Then
            Button10.PerformClick()
        End If
        If e.Control AndAlso e.KeyCode = Keys.S Then
            Button3.PerformClick()
        End If
        If e.Control AndAlso e.KeyCode = Keys.P Then
            Button9.PerformClick()
        End If
        If e.Control AndAlso e.KeyCode = Keys.X Then
            Button5.PerformClick()
        End If
        If e.Control AndAlso e.KeyCode = Keys.Q Then
            Button1.PerformClick()
        End If
    End Sub

    Public Sub SetRoleVisibility1(role As String)

        If role = "STAFF" Then
            DataGridView1.Height = 500
            Panel1.Height = 500

        End If
    End Sub
    Public Sub BarcodeClearTimer_Tick(sender As Object, e As EventArgs)
        barcodeClearTimer.Stop()
        TextBox1.Clear()
        TextBox1.Focus()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ResetFormFields()

    End Sub
    Private Sub ResetFormFields()
        DataGridView1.Rows.Clear()
        Label8.Text = ""
        Label9.Text = ""
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.SelectedIndex = -1
        RadioButton1.Checked = False
        RadioButton2.Checked = False
        RadioButton3.Checked = False
        ' RadioButton4.Checked = False
    End Sub
    Dim formattedDate As String = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss")


    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.ColumnIndex = DataGridView1.Columns("Quantity").Index OrElse e.ColumnIndex = DataGridView1.Columns("Price").Index Then
            UpdateTotalPrice()
        End If

    End Sub

    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If e.ColumnIndex = DataGridView1.Columns("Remove").Index AndAlso e.RowIndex >= 0 Then
            ' Customize the button cell appearance
            e.Paint(e.CellBounds, DataGridViewPaintParts.All)

            ' Define a smaller rectangle within the cell bounds for the background color (optional)
            Dim customRect As New Rectangle(e.CellBounds.X + 5, e.CellBounds.Y + 5, e.CellBounds.Width - 10, e.CellBounds.Height - 10)

            ' Draw the trash image from Resources (replace "TrashIcon" with the actual resource name)
            Dim trashImage As Image = My.Resources.trash ' Ensure the image name is correct in your resources

            ' Calculate position to center the image in the button cell
            Dim imgX As Integer = e.CellBounds.X + (e.CellBounds.Width - trashImage.Width) \ 2
            Dim imgY As Integer = e.CellBounds.Y + (e.CellBounds.Height - trashImage.Height) \ 2

            ' Draw the image inside the cell
            e.Graphics.DrawImage(trashImage, New Rectangle(imgX, imgY, trashImage.Width, trashImage.Height))

            ' Prevent the default painting to avoid showing text
            e.Handled = True
        End If
    End Sub



    Private Sub AddRemoveButtonColumn()
        ' Create a new button column for removing items
        Dim removeButtonColumn As New DataGridViewButtonColumn With {
        .HeaderText = "Action",
        .Name = "Remove",
        .Text = "",
        .UseColumnTextForButtonValue = True ' Use the same text for all rows
    }
        DataGridView1.Columns.Add(removeButtonColumn)
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If e.ColumnIndex = DataGridView1.Columns("Remove").Index AndAlso e.RowIndex >= 0 Then
            ' Set the background color and text color for the button
            With DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Style
                .ForeColor = Color.White ' Foreground color
                ' You can also set a custom background color here if needed
            End With
        End If
    End Sub

    Public Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        isPanel1Visible = Not isPanel1Visible

        If isPanel1Visible Then
            Panel5.Visible = True
            Panel4.Visible = True
            Label1.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Label5.Visible = False
            ComboBox1.Visible = False
            TextBox2.Visible = False
            TextBox3.Visible = False
            TextBox4.Visible = False
        Else
            Panel5.Visible = False
            Panel4.Visible = False
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label5.Visible = True
            ComboBox1.Visible = True
            TextBox2.Visible = True
            TextBox3.Visible = True
            TextBox4.Visible = True
        End If
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        isPanel1Visible = Not isPanel1Visible

        If isPanel1Visible Then
            Panel4.Visible = True
            Label1.Visible = False
            Label2.Visible = False
            Label3.Visible = False
            Label5.Visible = False
            ComboBox1.Visible = False
            TextBox2.Visible = False
            TextBox3.Visible = False
            TextBox4.Visible = False


        Else
            Panel4.Visible = False
            Label1.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Label5.Visible = True
            ComboBox1.Visible = True
            TextBox2.Visible = True
            TextBox3.Visible = True
            TextBox4.Visible = True


        End If
    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Panel4.Visible = False
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim newQuantity As Integer
        Dim selectedRowIndex As Integer = DataGridView1.CurrentCell.RowIndex
        Dim barcode As String = DataGridView1.Rows(selectedRowIndex).Cells("Barcode").Value.ToString()
        Dim availableQuantity As Integer = GetInventoryQuantity(barcode)


        If Integer.TryParse(TextBox5.Text, newQuantity) Then

            If newQuantity > 0 AndAlso newQuantity <= availableQuantity Then

                DataGridView1.Rows(selectedRowIndex).Cells("Quantity").Value = newQuantity
                Dim pricePerUnit As Decimal = Convert.ToDecimal(DataGridView1.Rows(selectedRowIndex).Cells("Price").Value)
                Dim totalPrice As Decimal = newQuantity * pricePerUnit
                DataGridView1.Rows(selectedRowIndex).Cells("TotalPrice").Value = totalPrice
                UpdateTotalPrice()

                Panel4.Visible = False
                Label1.Visible = True
                Label2.Visible = True
                Label3.Visible = True
                Label5.Visible = True
                ComboBox1.Visible = True
                TextBox2.Visible = True
                TextBox3.Visible = True
                TextBox4.Visible = True
                TextBox5.Clear()
            Else
                If newQuantity <= 0 Then
                    MessageBox.Show("Quantity must be greater than 0.")
                ElseIf newQuantity > availableQuantity Then
                    MessageBox.Show($"Only {availableQuantity} units available in stock.")
                End If
            End If
        Else
            MessageBox.Show("Please enter a valid quantity.")
        End If
    End Sub


    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        If e.ColumnIndex = DataGridView1.Columns("Remove").Index AndAlso e.RowIndex >= 0 Then
            ' Confirm with the user before removing the row
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to remove this item?", "Confirm Removal", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                ' Remove the row
                DataGridView1.Rows.RemoveAt(e.RowIndex)
                ' Update the total price after the row is removed
                UpdateTotalPrice()
            End If
        ElseIf e.ColumnIndex = DataGridView1.Columns("Quantity").Index OrElse e.ColumnIndex = DataGridView1.Columns("Price").Index Then
            UpdateTotalPrice() ' Update total price when Quantity or Price is clicked
        End If
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        pendingBarcode = TextBox1.Text.Trim()
        scanDelayTimer.Stop()
        scanDelayTimer.Start()
    End Sub
    Private Sub ScanDelayTimer_Tick(sender As Object, e As EventArgs)
        scanDelayTimer.Stop()

        Dim scannedBarcode As String = pendingBarcode
        If String.IsNullOrEmpty(scannedBarcode) Then Exit Sub

        TextBox1.Focus()

        Dim product As Product = Module1.GetProductDetails(scannedBarcode)

        If product IsNot Nothing Then
            If Not product.ProductStatus.Equals("ACTIVE", StringComparison.OrdinalIgnoreCase) Then
                Dim statusMessage As String

                Select Case product.ProductStatus.ToUpper()
                    Case "INACTIVE"
                        statusMessage = "This product is marked as INACTIVE and cannot be processed."
                    Case "DAMAGED"
                        statusMessage = "This product is DAMAGED and is not available for sale."
                    Case "LOSS PRODUCT"
                        statusMessage = "This product is marked as LOST and cannot be sold."
                    Case Else
                        statusMessage = "This product is not available for sale due to its current status: " & product.ProductStatus
                End Select

                MessageBox.Show(statusMessage, "Unavailable Product", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TextBox1.Clear()
                Exit Sub
            End If

            Dim inventoryQuantity As Integer = product.Quantity
            If inventoryQuantity > 0 Then
                Dim existingRow As DataGridViewRow = Nothing
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If row.IsNewRow Then Continue For
                    If row.Cells("Barcode").Value.ToString().Equals(product.Barcode, StringComparison.OrdinalIgnoreCase) Then
                        existingRow = row
                        Exit For
                    End If
                Next

                If existingRow IsNot Nothing Then
                    Dim currentQuantity As Integer = Convert.ToInt32(existingRow.Cells("Quantity").Value)
                    If currentQuantity < inventoryQuantity Then
                        existingRow.Cells("Quantity").Value = currentQuantity + 1
                        Dim pricePerUnit As Decimal = Convert.ToDecimal(existingRow.Cells("Price").Value)
                        existingRow.Cells("TotalPrice").Value = (currentQuantity + 1) * pricePerUnit
                        UpdateTotalPrice()
                    Else
                        MessageBox.Show("This item is currently out of stock and cannot be processed. Please restock to fulfill customer orders.",
                             "Out of Stock", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                Else
                    DataGridView1.Rows.Add(product.Barcode, product.Brand, product.Model, product.ProductName, 1, product.Price, product.Price, product.VatStatus)
                    UpdateTotalPrice()
                End If
            Else
                MessageBox.Show("This item is currently out of stock and cannot be processed. Please restock to fulfill customer orders.",
                     "Out of Stock", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Else
            MessageBox.Show("Barcode not found. Please check and try again.", "Invalid Barcode", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        TextBox1.Focus()
        barcodeClearTimer.Start()
    End Sub





    Private Function GetInventoryQuantity(barcode As String) As Integer
        Dim query As String = "SELECT Quantity FROM InventoryTb WHERE Barcode = @Barcode"
        Dim quantity As Integer = 0

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Barcode", barcode)

                Try
                    con.Open()
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing Then
                        quantity = Convert.ToInt32(result)
                    End If
                Catch ex As Exception
                    MessageBox.Show("Error fetching inventory quantity: " & ex.Message)
                Finally
                    con.Close()
                End Try
            End Using
        End Using

        Return quantity
    End Function

    Private Function GetTotalPrice() As Decimal
        Dim total As Decimal = 0D


        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim quantity As Integer = Convert.ToInt32(row.Cells("Quantity").Value)
                Dim pricePerUnit As Decimal = Convert.ToDecimal(row.Cells("Price").Value)
                total += quantity * pricePerUnit
            End If
        Next

        Return total
    End Function
    Private Sub UpdateTotalPrice()
        Dim total As Decimal = 0D
        Dim discount As Decimal = 0

        If RadioButton5.Checked Then
            discount = 0 ' 0% discount
        ElseIf RadioButton1.Checked Then
            discount = 0.2 ' 20% discount
        ElseIf RadioButton2.Checked Then
            discount = 0.5 ' 50% discount
        ElseIf RadioButton3.Checked Then
            discount = 0.7 ' 70% discount
            '  ElseIf RadioButton4.Checked Then
            ' discount = 1 ' 100% discount
        End If

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim quantity As Integer = Convert.ToInt32(row.Cells("Quantity").Value)
                Dim pricePerUnit As Decimal = Convert.ToDecimal(row.Cells("Price").Value)
                Dim discountedPrice As Decimal = pricePerUnit - (pricePerUnit * discount)
                Dim totalPrice As Decimal = quantity * discountedPrice

                row.Cells("TotalPrice").Value = totalPrice
                total += totalPrice
            End If
        Next
        Label8.Text = total.ToString("N2")
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim cashEntered As Decimal
        Dim Total As Decimal

        ' Parse the entered cash amount
        If Decimal.TryParse(TextBox2.Text, cashEntered) Then
            ' Use the discounted total from Label8
            Total = Convert.ToDecimal(Label8.Text.Replace("₹", "").Replace(",", ""))

            ' Calculate the change
            Dim change As Decimal = cashEntered - Total

            If change < 0 Then
                Label9.Text = "Not enough amount"
            Else
                Label9.Text = $"{change:N2}"
            End If
        Else
            Label9.Text = "Change: Invalid amount"
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Ensure both Cash Amount and Method of Payment are provided
        If TextBox2.Text <> "" AndAlso ComboBox1.SelectedItem IsNot Nothing Then
            ' Fetch the current date from SQL Server
            Dim currentDate As DateTime = GetSQLServerDate()
            If currentDate = DateTime.MinValue Then
                MessageBox.Show("Failed to retrieve server date. Please check the connection.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Initialize variables
            Dim CustomerName As String = TextBox4.Text
            Dim ReferenceNo As String = TextBox3.Text
            Dim Amount As Decimal
            Dim Discount As Decimal = 0
            Dim MOP As String = If(ComboBox1.SelectedItem Is Nothing, String.Empty, ComboBox1.SelectedItem.ToString())
            Dim Total As Decimal = Convert.ToDecimal(Label8.Text.Replace("₹", "").Replace(",", ""))
            Dim Change As Decimal

            ' Validate the cash amount
            If Not Decimal.TryParse(TextBox2.Text, Amount) Then
                MessageBox.Show("Invalid cash amount.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Validate the change amount
            Dim changeText As String = Label9.Text.Replace("₹", "").Replace(",", "")
            If Not Decimal.TryParse(changeText, Change) Then
                MessageBox.Show("Invalid change amount.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' Apply discount based on the selected RadioButton
            If RadioButton5.Checked Then
                Discount = 0 ' 0% discount
            ElseIf RadioButton1.Checked Then
                Discount = 0.2 ' 20% discount
            ElseIf RadioButton2.Checked Then
                Discount = 0.5 ' 50% discount
            ElseIf RadioButton3.Checked Then
                Discount = 0.7 ' 70% discount
                ' ElseIf RadioButton4.Checked Then
                ' Discount = 1 ' 100% discount
            End If

            ' Generate a unique barcode for the receipt
            receiptBarcodeNo = GenerateReceiptBarcodeNo()

            ' Process each row in the DataGridView
            For Each row As DataGridViewRow In DataGridView1.Rows
                If Not row.IsNewRow Then
                    Dim Barcode As String = row.Cells("Barcode").Value.ToString()
                    Dim Quantity As Integer = Convert.ToInt32(row.Cells("Quantity").Value)
                    Dim Price As Decimal = Convert.ToDecimal(row.Cells("Price").Value)
                    Dim TotalPrice As Decimal = Convert.ToDecimal(row.Cells("TotalPrice").Value)
                    Dim vatStatus As String = row.Cells("VatStatus").Value.ToString()

                    ' Tax-related calculations
                    Dim VAT As Decimal = 0D
                    Dim VATable As Decimal = 0D
                    Dim VATExemptSales As Decimal = 0D
                    Dim VATZeroRatedSales As Decimal = 0D
                    Dim NonVAT As Decimal = 0D

                    If vatStatus = "VAT" Then
                        VATable = (Price * Quantity) / 1.12D ' 12% tax
                        VAT = (Price * Quantity) - VATable
                    ElseIf vatStatus = "NonVAT" Then
                        NonVAT = (Price * Quantity) / 1.03D ' 3% tax
                    End If

                    Dim query As String = "INSERT INTO PaymentTb (Barcode, CustomerName, ReferenceNo, Amount, Quantity, Discount, MOP, TotalPrice, Total, Change, DateIssued, RefundExpirationDate, ReceiptNo, ReceiptBarcodeNo, VAT, VATable, NonVAT, VATExemptSales, VATZeroRatedSales, HandledBy) " &
    "VALUES (@Barcode, @CustomerName, @ReferenceNo, @Amount, @Quantity, @Discount, @MOP, @TotalPrice, @Total, @Change, GETDATE(), DATEADD(YEAR, 1, GETDATE()), @ReceiptNo, @ReceiptBarcodeNo, @VAT, @VATable, @NonVAT, @VATExemptSales, @VATZeroRatedSales, @HandledBy)"

                    Using con As New SqlConnection(ConnectionString)
                        con.Open()
                        Using cmd As New SqlCommand(query, con)
                            ' Add parameters
                            cmd.Parameters.AddWithValue("@Barcode", Barcode)
                            cmd.Parameters.AddWithValue("@CustomerName", CustomerName)
                            cmd.Parameters.AddWithValue("@ReferenceNo", ReferenceNo)
                            cmd.Parameters.AddWithValue("@Amount", Amount)
                            cmd.Parameters.AddWithValue("@Quantity", Quantity)
                            cmd.Parameters.AddWithValue("@Discount", Discount * 100 & "%")
                            cmd.Parameters.AddWithValue("@MOP", MOP)
                            cmd.Parameters.AddWithValue("@TotalPrice", TotalPrice)
                            cmd.Parameters.AddWithValue("@Total", Total)
                            cmd.Parameters.AddWithValue("@Change", Change)
                            cmd.Parameters.AddWithValue("@ReceiptNo", GenerateReceiptNo())
                            cmd.Parameters.AddWithValue("@ReceiptBarcodeNo", receiptBarcodeNo)
                            cmd.Parameters.AddWithValue("@VAT", VAT)
                            cmd.Parameters.AddWithValue("@VATable", VATable)
                            cmd.Parameters.AddWithValue("@NonVAT", NonVAT)
                            cmd.Parameters.AddWithValue("@VATExemptSales", VATExemptSales)
                            cmd.Parameters.AddWithValue("@VATZeroRatedSales", VATZeroRatedSales)
                            cmd.Parameters.AddWithValue("@HandledBy", Label16.Text)


                            Try
                                ' Execute the SQL command
                                cmd.ExecuteNonQuery()

                                ' Update inventory quantity
                                UpdateInventoryQuantityForItem(Barcode, Quantity)

                            Catch ex As SqlException
                                MessageBox.Show("SQL error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Catch ex As Exception
                                MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            End Try
                        End Using
                        con.Close()
                    End Using
                End If
            Next

            ' Show confirmation message
            MessageBox.Show("Transaction completed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            ' Show an error message if required fields are missing
            MessageBox.Show("Please fill in all required fields.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub





    Private Sub UpdateInventoryQuantityForItem(barcode As String, quantity As Integer)
        Dim query As String = "UPDATE InventoryTb SET Quantity = Quantity - @Quantity WHERE Barcode = @Barcode"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Quantity", quantity)
                cmd.Parameters.AddWithValue("@Barcode", barcode)

                Try
                    con.Open()
                    cmd.ExecuteNonQuery()
                    Dim remainingQuantity As Integer = GetInventoryQuantity(barcode)
                    If remainingQuantity <= 5 Then
                        If remainingQuantity = 0 Then
                            MessageBox.Show("This item is completely out of stock. Please restock to continue selling.",
                        "Out of Stock", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Else
                            MessageBox.Show("Warning: Low stock for this item. Only " & remainingQuantity & " units remaining. Consider restocking soon.",
                        "Low Stock Alert", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        End If
                    End If

                Catch ex As SqlException
                    MessageBox.Show("SQL error occurred while updating inventory: " & ex.Message)
                Catch ex As Exception
                    MessageBox.Show("An error occurred while updating inventory: " & ex.Message)
                Finally
                    con.Close()
                End Try
            End Using
        End Using

    End Sub
    Public Sub InventoryData()
        Dim query As String = "SELECT Category, ProductName, Brand, Model, Quantity, Barcode FROM InventoryTb"
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Dim dt As New DataTable()
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
                DataGridView2.DataSource = dt
            End Using
        End Using
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' Validate required fields
        If String.IsNullOrWhiteSpace(TextBox2.Text) OrElse ComboBox1.SelectedIndex = -1 Then
            MessageBox.Show("Please fill in all fields before submitting.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' Set receipt paper size and show print preview
        PrintDocument1.DefaultPageSettings.PaperSize = New PaperSize("Receipt", 200, 1100) ' 4 inches by 11 inches
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
        TextBox1.Focus()
        ResetFormFields()

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Define reusable fonts and brush
        Dim font As New Font("Arial", 10)
        Dim Add As New Font("Arial", 5)
        Dim dates As New Font("Arial", 5)
        Dim Pd As New Font("Arial", 5)

        Dim Title As New Font("Arial", 8)
        Dim boldFont As New Font("Arial", 15, FontStyle.Bold)
        Dim addFont As New Font("Arial", 11)
        Dim msFont As New Font("Microsoft Sans Serif", 9)
        Dim brush As New SolidBrush(Color.Black)

        ' Initial Y position for content
        Dim yPos As Integer = 5
        Dim paperWidth As Integer = e.PageBounds.Width

        Dim currentDate As DateTime = GetSQLServerDate()
        Dim expirationDate As DateTime = currentDate.AddDays(1)

        ' Helper function to draw centered text
        Dim drawCenteredText As Action(Of String, Font, Integer) = Sub(text As String, fnt As Font, offset As Integer)
                                                                       Dim width As Single = e.Graphics.MeasureString(text, fnt).Width
                                                                       e.Graphics.DrawString(text, fnt, brush, (paperWidth - width) / 2, yPos + offset)
                                                                   End Sub

        ' Print header
        yPos += 20
        drawCenteredText(StoreName, boldFont, 0)
        yPos += 25
        drawCenteredText(StoreType, addFont, 0)
        yPos += 17
        drawCenteredText(StoreAddress, Add, 0)
        yPos += 12
        drawCenteredText("Date Issued:  " & GetSQLServerDate().ToString(), dates, 0)
        yPos += 12
        drawCenteredText("Refund Exp:  " & expirationDate, dates, 0)
        yPos += 12
        drawCenteredText("Cashier:" & Label16.Text, dates, 0)
        yPos += 12
        drawCenteredText("Non Vat Reg. Tin: 632-757-013-00000", Pd, 0)


        ' Item list header
        yPos += 20
        e.Graphics.DrawString("----------------------------------", font, brush, 20, yPos)
        yPos += 30
        e.Graphics.DrawString("ITEM", Title, brush, 50, yPos)
        e.Graphics.DrawString("QTY", Title, brush, 20, yPos)
        e.Graphics.DrawString("UNIT PRICE", Title, brush, 115, yPos)
        yPos += 30

        Dim totalVATable As Decimal = 0
        Dim totalVAT As Decimal = 0
        Dim totalNonVAT As Decimal = 0
        'Dim totalVATExemptSales As Decimal = 0
        ' Dim totalVATZeroRatedSales As Decimal = 0

        For Each row As DataGridViewRow In DataGridView1.Rows
            If Not row.IsNewRow Then
                Dim productInfo As String = row.Cells("ProductName").Value.ToString()
                Dim price As Decimal = Convert.ToDecimal(row.Cells("Price").Value)
                Dim quantity As Integer = Convert.ToInt32(row.Cells("Quantity").Value)
                Dim vatStatus As String = row.Cells("VatStatus").Value.ToString()

                ' Calculate totals based on VAT status
                If vatStatus = "VAT" Then
                    Dim VATable As Decimal = (price * quantity) / 1.12D
                    Dim VAT As Decimal = (price * quantity) - VATable
                    totalVATable += VATable
                    totalVAT += VAT
                ElseIf vatStatus = "NonVAT" Then
                    totalNonVAT += (price * quantity) / 1.03D
                End If

                ' Print item details
                yPos += 20
                e.Graphics.DrawString(productInfo, Pd, brush, 45, yPos)
                e.Graphics.DrawString(quantity.ToString(), Pd, brush, 20, yPos)
                e.Graphics.DrawString(price.ToString("F2"), Pd, brush, 140, yPos)
            End If
        Next


        yPos += 30
        e.Graphics.DrawString("----------------------------------", font, brush, 20, yPos)
        yPos += 30
        e.Graphics.DrawString("Total                                 " & Label8.Text, Add, brush, 20, yPos)
        yPos += 20
        e.Graphics.DrawString("Cash:                               " & TextBox2.Text, Add, brush, 20, yPos)
        yPos += 20
        e.Graphics.DrawString("Change:                           " & Label9.Text, Add, brush, 20, yPos)
        yPos += 40
        e.Graphics.DrawString("VATable:                           " & totalVATable.ToString("F2"), Add, brush, 20, yPos)
        yPos += 20
        e.Graphics.DrawString("VAT:                                 " & totalVAT.ToString("F2"), Add, brush, 20, yPos)
        yPos += 20
        e.Graphics.DrawString("Non-VAT:                          " & totalNonVAT.ToString("F2"), Add, brush, 20, yPos)
        yPos += 20
        e.Graphics.DrawString("VAT Exempt Sales:            0.00", Add, brush, 20, yPos) ' Corrected
        yPos += 20
        e.Graphics.DrawString("VAT Zero Rated Sales:       0.00", Add, brush, 20, yPos) ' Corrected
        yPos += 20
        e.Graphics.DrawString("----------------------------------", font, brush, 20, yPos)
        ' Customer and barcode section
        yPos += 30
        e.Graphics.DrawString("Customer:       " & TextBox4.Text, Add, brush, 20, yPos)
        yPos += 30
        Dim barcode As New MessagingToolkit.Barcode.BarcodeEncoder()
        Dim barcodeImage As Image = barcode.Encode(MessagingToolkit.Barcode.BarcodeFormat.Code128, receiptBarcodeNo)
        e.Graphics.DrawImage(barcodeImage, 0, yPos, 200, 80)
        yPos += 75
        e.Graphics.DrawString(receiptBarcodeNo, font, brush, 50, yPos)
        yPos += 30
        e.Graphics.DrawString("  We appreciate your visit!", msFont, brush, 20, yPos)
        yPos += 30
        drawCenteredText("20 bklts. (50 x 2) 1001-2000 Date of ATP 10/03/24", Pd, 0)
        yPos += 12
        drawCenteredText("BIR Permit No. OCN 024AU20240000007589", Pd, 0)
    End Sub




    Private Sub Button7_Click(sender As Object, e As EventArgs)
        InventoryData()
    End Sub
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)
        Module1.FilterDataGridView2(TextBox6.Text, DataGridView2, ConnectionString)
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem IsNot Nothing AndAlso ComboBox1.SelectedItem.ToString() = "CASH" Then

            TextBox3.Visible = False
            Label3.Visible = False
        Else

            TextBox3.Visible = True
            Label3.Visible = True
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        InventoryData()
    End Sub

    Private Sub TextBox6_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Module1.FilterDataGridView2(TextBox6.Text, DataGridView2, ConnectionString)
    End Sub
    Private Sub TextBox6_GotFocus(sender As Object, e As EventArgs) Handles TextBox6.GotFocus
        If TextBox6.Text = "Search by ProductName" Then
            TextBox6.Text = ""
            TextBox6.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox6_lostFocus(sender As Object, e As EventArgs) Handles TextBox6.LostFocus
        If TextBox6.Text = "" Then
            TextBox6.Text = "Search by ProductName"
            TextBox6.ForeColor = Color.DarkGray
        End If
    End Sub
    Sub Autocomplete()
        Try
            Using con As New SqlConnection(ConnectionString)
                con.Open()


                Dim cmd = New SqlCommand("SELECT Barcode FROM InventoryTb ", con)
                Dim da As New SqlDataAdapter(cmd)
                Dim dt As New DataTable
                da.Fill(dt)
                Dim column1 As New AutoCompleteStringCollection
                For Each row As DataRow In dt.Rows
                    Dim combinedValue As String = row("Barcode").ToString()
                    column1.Add(combinedValue)
                Next
                TextBox6.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                TextBox6.AutoCompleteSource = AutoCompleteSource.CustomSource
                TextBox6.AutoCompleteCustomSource = column1
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        End Try
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Dim selectedCategory As String = ComboBox3.SelectedItem.ToString()
        FilterByCategory(selectedCategory)
    End Sub
    Public Sub FilterByCategory(selectedCategory As String)
        Dim query As String = "SELECT  Category, ProductName, Brand, Model, Quantity, Barcode FROM InventoryTb WHERE Category = @Category"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@Category", selectedCategory)
                Using da As New SqlDataAdapter(cmd)
                    Using dt As New DataTable()
                        da.Fill(dt)
                        DataGridView2.DataSource = dt
                    End Using
                End Using
            End Using
        End Using
    End Sub
    Public Sub LoadCategories()
        ComboBox3.Items.Clear()

        Using con As New SqlConnection(ConnectionString)
            con.Open()

            Dim query As String = "SELECT DISTINCT Category FROM InventoryTb"
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

    Private Sub TextBox5_GotFocus(sender As Object, e As EventArgs) Handles TextBox5.GotFocus
        If TextBox5.Text = "QTY" Then
            TextBox5.Text = ""
            TextBox5.ForeColor = Color.Black
        End If
    End Sub

    Private Sub TextBox5_lostFocus(sender As Object, e As EventArgs) Handles TextBox5.LostFocus
        If TextBox5.Text = "" Then
            TextBox5.Text = "QTY"
            TextBox5.ForeColor = Color.DarkGray
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button3.PerformClick()
        End If
    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TextBox2.KeyPress
        ' Allow only numbers and control keys (like Backspace)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub


    Private Sub TextBox3_KeyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles TextBox3.KeyPress
        ' Allow only numbers and control keys (like Backspace)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
End Class