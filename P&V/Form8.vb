Imports System.Data.SqlClient
Imports System.Drawing.Printing
Public Class Form8

    Private RefundBarcodeNo As String
    Private WithEvents PrintPreviewDialog1 As New PrintPreviewDialog


    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Focus()
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.WindowState = FormWindowState.Maximized
        ' Panel1.Visible = False
        AddRemoveButtonColumn()
        Module1.LoadStoreDetails()

        SetFullName1(userFullName)


    End Sub
    Public Sub New(userRole As String, fullName As String)
        InitializeComponent()
        SetFullName1(fullName)
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub SetFullName1(fullName As String)
        Label16.Text = fullName
    End Sub



    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        FetchPaymentDetailsByBarcode(TextBox2.Text)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Timer1.Stop() ' Reset the timer on every keystroke
        Timer1.Start() ' Restart the timer
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop() ' Stop the timer to avoid repeated execution

        Dim searchText As String = TextBox1.Text.Trim()

        If String.IsNullOrWhiteSpace(searchText) Then
            DataGridView1.DataSource = Nothing
            DataGridView1.Rows.Clear()
        Else
            Module1.FilterRefund(searchText, DataGridView1, ConnectionString)
        End If
    End Sub


    Public Sub Payment()

        Dim query As String = "SELECT p.Id, i.ProductName, p.Quantity, i.Price, p.Discount, p.TotalPrice, p.ReceiptBarcodeNo " &
                          "FROM InventoryTb AS i " &
                          "JOIN PaymentTb AS p ON i.Barcode = p.Barcode"


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

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        ' Validate refund details
        If String.IsNullOrEmpty(TextBox8.Text) OrElse String.IsNullOrEmpty(TextBox3.Text) OrElse String.IsNullOrEmpty(TextBox4.Text) Then
            MessageBox.Show("Please complete all refund fields.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        ' Prepare refund details
        Dim PaymentId As Integer = Convert.ToInt32(DataGridView1.CurrentRow.Cells("Id").Value)
        Dim RefundCause As String = TextBox8.Text
        Dim RefundDate As Date = GetSQLServerDate()
        Dim RefundAmount As Decimal = Convert.ToDecimal(TextBox3.Text)
        Dim RefundQuantity As Integer = Convert.ToInt32(TextBox4.Text)
        Dim Barcode As String = DataGridView1.CurrentRow.Cells("Barcode").Value.ToString()

        RefundBarcodeNo = GenerateReceiptBarcodeNo()

        Dim AvailableQuantity As Integer = Convert.ToInt32(DataGridView1.CurrentRow.Cells("Quantity").Value)
        If RefundQuantity > AvailableQuantity Then
            MessageBox.Show("Refund quantity exceeds recorded quantity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ClearText()
            Return
        End If

        ' Get VatStatus from InventoryTb using the barcode
        Dim VatStatus As String = ""
        Using con As New SqlConnection(ConnectionString)
            con.Open()
            Using cmdCheckVat As New SqlCommand("SELECT VatStatus FROM InventoryTb WHERE Barcode = @Barcode", con)
                cmdCheckVat.Parameters.AddWithValue("@Barcode", Barcode)
                Dim result = cmdCheckVat.ExecuteScalar()
                If result IsNot Nothing Then
                    VatStatus = result.ToString()
                End If
            End Using
            con.Close()
        End Using

        ' Get unit breakdown for VATable, VAT, and NonVAT from PaymentTb
        Dim unitVATable As Decimal = 0D
        Dim unitVAT As Decimal = 0D
        Dim unitNonVAT As Decimal = 0D

        Using con As New SqlConnection(ConnectionString)
            con.Open()
            Using cmd As New SqlCommand("SELECT Quantity, VATable, VAT, NonVAT FROM PaymentTb WHERE Id = @PaymentId", con)
                cmd.Parameters.AddWithValue("@PaymentId", PaymentId)
                Using reader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim origQty As Integer = Convert.ToInt32(reader("Quantity"))
                        If origQty > 0 Then
                            unitVATable = Convert.ToDecimal(reader("VATable")) / origQty
                            unitVAT = Convert.ToDecimal(reader("VAT")) / origQty
                            unitNonVAT = Convert.ToDecimal(reader("NonVAT")) / origQty
                        End If
                    End If
                End Using
            End Using
            con.Close()
        End Using

        ' Calculate deduction based on refunded quantity
        Dim VATableDeduction As Decimal = unitVATable * RefundQuantity
        Dim VATDeduction As Decimal = unitVAT * RefundQuantity
        Dim NonVATDeduction As Decimal = unitNonVAT * RefundQuantity

        ' SQL queries
        Dim queryRefund As String = "INSERT INTO RefundTb (PaymentId, RefundCause, RefundDate, RefundAmount, Quantity, RefundBarcodeNo, HandledBy) " &
                                "VALUES (@PaymentId, @RefundCause, @RefundDate, @RefundAmount, @Quantity, @RefundBarcodeNo, @HandledBy)"

        Dim queryUpdatePayment As String = "UPDATE PaymentTb SET " &
                                       "Quantity = Quantity - @RefundQuantity, " &
                                       "TotalPrice = TotalPrice - @RefundAmount, " &
                                       "VATable = VATable - @VATable, " &
                                       "VAT = VAT - @VAT, " &
                                       "NonVAT = NonVAT - @NonVAT " &
                                       "WHERE Id = @PaymentId"

        ' Process refund and update payment
        Using con As New SqlConnection(ConnectionString)
            con.Open()
            Using cmdRefund As New SqlCommand(queryRefund, con)
                cmdRefund.Parameters.AddWithValue("@PaymentId", PaymentId)
                cmdRefund.Parameters.AddWithValue("@RefundCause", RefundCause)
                cmdRefund.Parameters.AddWithValue("@RefundDate", RefundDate)
                cmdRefund.Parameters.AddWithValue("@RefundAmount", RefundAmount)
                cmdRefund.Parameters.AddWithValue("@Quantity", RefundQuantity)
                cmdRefund.Parameters.AddWithValue("@RefundBarcodeNo", RefundBarcodeNo)
                cmdRefund.Parameters.AddWithValue("@HandledBy", Label16.Text)

                Try
                    cmdRefund.ExecuteNonQuery()

                    Using cmdUpdate As New SqlCommand(queryUpdatePayment, con)
                        cmdUpdate.Parameters.AddWithValue("@RefundQuantity", RefundQuantity)
                        cmdUpdate.Parameters.AddWithValue("@RefundAmount", RefundAmount)
                        cmdUpdate.Parameters.AddWithValue("@VATable", VATableDeduction)
                        cmdUpdate.Parameters.AddWithValue("@VAT", VATDeduction)
                        cmdUpdate.Parameters.AddWithValue("@NonVAT", NonVATDeduction)
                        cmdUpdate.Parameters.AddWithValue("@PaymentId", PaymentId)

                        cmdUpdate.ExecuteNonQuery()
                    End Using

                    MessageBox.Show("Refund processed and payment updated successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    ' Optional UI reset
                    TextBox3.Clear()
                    TextBox4.Clear()
                    TextBox8.Clear()

                Catch ex As Exception
                    MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using
    End Sub


    Private Sub FetchPaymentDetailsByBarcode(barcode As String)
        Dim query As String = "SELECT p.Id, i.ProductName, p.Quantity, i.Price, p.Discount, p.TotalPrice " &
                          "FROM InventoryTb AS i " &
                          "JOIN PaymentTb AS p ON i.Barcode = p.Barcode " &
                          "WHERE p.Id = @Barcode"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                ' Add parameter for the scanned barcode (ID in this case)
                cmd.Parameters.AddWithValue("@Barcode", barcode)

                con.Open()
                Using reader As SqlDataReader = cmd.ExecuteReader()
                    ' Check if any rows are returned
                    If reader.HasRows Then
                        ' Move to the first record
                        reader.Read()

                        ' Check if the Quantity is 0
                        Dim quantity As Integer = Convert.ToInt32(reader("Quantity"))
                        If quantity <= 0 Then
                            MessageBox.Show("Product already refunded.", "Quantity Unavailable", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Exit Sub
                        End If

                        ' Check if the same Id already exists in DataGridView2
                        Dim idExists As Boolean = False
                        For Each row As DataGridViewRow In DataGridView2.Rows
                            If row.Cells(0).Value IsNot Nothing AndAlso row.Cells(0).Value.ToString() = reader("Id").ToString() Then
                                idExists = True

                                ' If the row exists and is marked as edited, skip re-fetching the data
                                If row.Cells("IsEdited").Value IsNot Nothing AndAlso CBool(row.Cells("IsEdited").Value) Then
                                    ' Skip updating this row since it's already edited
                                    ' Update the Quantity and TotalPrice if they have been changed
                                    UpdateRowValues(row, reader)
                                    Exit Sub
                                End If

                                ' Update the current row values
                                UpdateRowValues(row, reader)

                                ' Set the current row as selected
                                DataGridView2.CurrentCell = row.Cells(0)
                                DataGridView2.Rows(row.Index).Selected = True

                                ' Set TextBox4 to the Quantity of the fetched record
                                TextBox4.Text = reader("Quantity").ToString()
                                Exit For
                            End If
                        Next

                        If Not idExists Then
                            ' Add the new data to TextBox and DataGridView2
                            TextBox4.Text = reader("Quantity").ToString()
                            TextBox3.Text = reader("TotalPrice").ToString()
                            Label3.Text = reader("Price").ToString()

                            Dim newRow As String() = New String() {
                            reader("Id").ToString(),        ' Payment Id
                            reader("ProductName").ToString(),     ' Brand from InventoryTb
                            reader("Quantity").ToString(),  ' Quantity from PaymentTb
                            reader("TotalPrice").ToString() ' TotalPrice from PaymentTb
                        }

                            ' Insert the new row at the top of DataGridView2 (index 0)
                            DataGridView2.Rows.Insert(0, newRow)
                            DataGridView2.Rows(0).Cells("IsEdited").Value = False

                            ' Select the new row and set TextBox4
                            DataGridView2.CurrentCell = DataGridView2.Rows(0).Cells(0)
                        Else
                            MessageBox.Show("This payment record has already been added.", "Duplicate Entry", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            ClearText()

                        End If
                    Else
                        ' No matching record found, show error and clear fields
                        ' MessageBox.Show("No matching payment record found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End Using
            End Using
        End Using
    End Sub



    Private Sub UpdateRowValues(row As DataGridViewRow, reader As SqlDataReader)
        row.Cells("Quantity").Value = reader("Quantity").ToString()
        row.Cells("TotalPrice").Value = reader("TotalPrice").ToString()
        row.Cells("IsEdited").Value = True ' Mark as edited
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If DataGridView2.CurrentRow IsNot Nothing Then
            If String.IsNullOrWhiteSpace(TextBox4.Text) Then
                ' If TextBox4 is empty (after backspace), do nothing
                Exit Sub
            End If

            If IsNumeric(TextBox4.Text) Then
                Dim quantity As Integer = Convert.ToInt32(TextBox4.Text)

                ' ❌ Prevent zero or negative input
                If quantity <= 0 Then
                    MessageBox.Show("Quantity cannot be zero or negative.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    ClearText()
                    Exit Sub ' Stop further execution
                End If

                Dim price As Decimal = Convert.ToDecimal(Label3.Text)
                Dim discount As Decimal = 0

                If DataGridView2.CurrentRow.Cells("IsEdited").Value IsNot Nothing AndAlso
               Not CBool(DataGridView2.CurrentRow.Cells("IsEdited").Value) Then

                    DataGridView2.CurrentRow.Cells("Quantity").Value = quantity

                    If DataGridView1.CurrentRow IsNot Nothing AndAlso
                   DataGridView1.CurrentRow.Cells("Discount").Value IsNot Nothing Then

                        Dim discountStr As String = DataGridView1.CurrentRow.Cells("Discount").Value.ToString().Replace("%", "")
                        Decimal.TryParse(discountStr, discount)
                    End If

                    ' Calculate discounted price
                    Dim discountedPrice As Decimal = price
                    Select Case discount
                        Case 0
                            discountedPrice = price
                        Case 20
                            discountedPrice *= 0.8D
                        Case 50
                            discountedPrice *= 0.5D
                        Case 70
                            discountedPrice *= 0.3D
                        Case 100
                            discountedPrice = 0
                    End Select

                    ' Calculate the total price based on the quantity
                    Dim totalPrice As Decimal = quantity * discountedPrice

                    ' Update TextBox3 (TotalPrice)
                    TextBox3.Text = totalPrice.ToString("F2")

                    ' Update the TotalPrice in DataGridView2
                    DataGridView2.CurrentRow.Cells("TotalPrice").Value = totalPrice.ToString("F2")

                    ' Mark the row as edited
                    DataGridView2.CurrentRow.Cells("IsEdited").Value = True
                End If
            Else
                ' Only show warning if user typed non-numeric input (NOT empty)
                MessageBox.Show("Please enter a valid number for Quantity.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                ClearText()
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Check if the clicked row is valid (i.e., not a header or out-of-range row)
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim selectedRow As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

            If selectedRow.Cells(0).Value IsNot Nothing Then
                Dim selectedId As String = selectedRow.Cells(0).Value.ToString()

                ' Check if the ID already exists in DataGridView2
                Dim alreadyExists As Boolean = False
                For Each row As DataGridViewRow In DataGridView2.Rows
                    If row.IsNewRow = False Then ' Avoid the last empty row
                        If row.Cells(0).Value IsNot Nothing AndAlso row.Cells(0).Value.ToString() = selectedId Then
                            alreadyExists = True
                            Exit For
                        End If
                    End If
                Next

                If alreadyExists Then
                    MessageBox.Show("This row is already selected.", "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    TextBox2.Text = selectedId
                End If
            Else
                TextBox2.Clear()
            End If
        End If
    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox8.Clear()
        Label3.Text = ""
        DataGridView2.Rows.Clear()

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim receiptSize As New PaperSize("Receipt", 200, 1100) ' 4 inches by 11 inches
        PrintDocument1.DefaultPageSettings.PaperSize = receiptSize
        PrintPreviewDialog1.Document = PrintDocument1
        PrintPreviewDialog1.ShowDialog()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox8.Clear()
        Label3.Text = ""
        DataGridView2.Rows.Clear()

    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Dim font As New Font("Arial", 10)
        Dim Add As New Font("Arial", 5)
        Dim Adds As New Font("Arial", 5)
        Dim Title As New Font("Arial", 8)
        Dim boldFont As New Font("Arial", 15, FontStyle.Bold)
        Dim addFont As New Font("Arial", 11)
        Dim msFont As New Font("Microsoft Sans Serif", 9)
        Dim msFonts As New Font("Microsoft Sans Serif", 7)
        Dim Pd As New Font("Arial", 5)
        Dim brush As New SolidBrush(Color.Black)

        ' Initial Y position for content
        Dim yPos As Integer = 5
        Dim paperWidth As Integer = e.PageBounds.Width ' Get the width of the paper

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
        drawCenteredText(GetSQLServerDate().ToString(), Add, 0)
        yPos += 12
        drawCenteredText("Cashier:" & Label16.Text, Add, 0)


        ' Item list header
        yPos += 40
        e.Graphics.DrawString("-------------------------------", font, brush, 20, yPos)
        yPos += 20
        e.Graphics.DrawString("ITEM", Title, brush, 50, yPos)
        e.Graphics.DrawString("QTY", Title, brush, 20, yPos)
        e.Graphics.DrawString("AMOUNT", Title, brush, 115, yPos)

        yPos += 20
        e.Graphics.DrawString("-------------------------------", font, brush, 20, yPos)

        For Each row As DataGridViewRow In DataGridView2.Rows
            yPos += 30
            e.Graphics.DrawString(row.Cells("ProductName").Value.ToString(), Adds, brush, 45, yPos)
            e.Graphics.DrawString(row.Cells("Quantity").Value.ToString(), Adds, brush, 20, yPos)
            e.Graphics.DrawString(row.Cells("TotalPrice").Value.ToString(), Adds, brush, 140, yPos)
        Next
        yPos += 20
        e.Graphics.DrawString("-------------------------------", font, brush, 20, yPos)

        ' Calculate and display the grand total
        Dim grandTotal As Decimal = CalculateGrandTotal()
        yPos += 40
        e.Graphics.DrawString("Total Refund:   " & grandTotal, Add, brush, 20, yPos)
        yPos += 30
        e.Graphics.DrawString("We appreciate your understanding.", msFonts, brush, 20, yPos)
        yPos += 30
        drawCenteredText("20 bklts. (50 x 2) 1001-2000 Date of ATP 10/03/24", pd, 0)
        yPos += 12
        drawCenteredText("BIR Permit No. OCN 024AU20240000007589", pd, 0)

    End Sub
    Private Function CalculateGrandTotal() As Decimal
        Dim grandTotal As Decimal = 0

        For Each row As DataGridViewRow In DataGridView2.Rows
            If row.Cells("TotalPrice").Value IsNot Nothing Then
                grandTotal += Convert.ToDecimal(row.Cells("TotalPrice").Value)
            End If
        Next

        Return grandTotal
    End Function
    Private Sub DataGridView2_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView2.CellPainting
        If e.ColumnIndex = DataGridView2.Columns("Remove").Index AndAlso e.RowIndex >= 0 Then
            ' Customize the button cell appearance
            e.Paint(e.CellBounds, DataGridViewPaintParts.All)

            ' Define a smaller rectangle within the cell bounds for the background color (optional)
            Dim customRect As New Rectangle(e.CellBounds.X + 5, e.CellBounds.Y + 5, e.CellBounds.Width - 10, e.CellBounds.Height - 10)

            Dim trashImage As Image = My.Resources.trash

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
        ' Check if the column already exists to avoid duplicates
        If DataGridView2.Columns.Contains("Remove") Then Exit Sub

        ' Initialize and configure the button column
        Dim removeButton As New DataGridViewButtonColumn With {
        .Name = "Remove",
        .HeaderText = "REMOVE",
        .UseColumnTextForButtonValue = True,
        .Width = 60
    }

        ' Add the column to DataGridView2
        DataGridView2.Columns.Add(removeButton)
    End Sub



    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        If e.ColumnIndex = DataGridView2.Columns("Remove").Index AndAlso e.RowIndex >= 0 Then
            ' Set the background color and text color for the button using simplified style
            Dim buttonStyle As New DataGridViewCellStyle With {
            .ForeColor = Color.White ' Foreground color
        }

            ' Apply the style to the cell
            With DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Style
                .BackColor = buttonStyle.BackColor
                .ForeColor = buttonStyle.ForeColor
            End With
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick
        ' Check if the clicked cell is the "Remove" button column
        If e.ColumnIndex = DataGridView2.Columns("Remove").Index AndAlso e.RowIndex >= 0 Then
            ' Confirm removal with the user
            Dim result As DialogResult = MessageBox.Show("Are you sure you want to remove this row?", "Confirm Removal", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                ' Remove the selected row from DataGridView2
                DataGridView2.Rows.RemoveAt(e.RowIndex)
                Label3.Text = ""
                ClearText()
            End If
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Panel1.Visible = False
    End Sub

    Public Sub ClearText()
        Label3.Text = ""
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox8.Clear()
    End Sub
End Class