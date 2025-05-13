Imports System.Data.SqlClient


Module Module1
    ' Public ConnectionString As String = "Data Source=MSOO\SQLEXPRESS;Initial Catalog=POS&INVT;Integrated Security=True"
    'Connection
    Public ConnectionString As String = "Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=|DataDirectory|\Database\POS.mdf;Database=POS&INVT; Integrated Security=True"
    Public con As SqlConnection
    Public isPanel1Visible As Boolean = False
    Public StoreName As String
    Public StoreType As String
    Public StoreAddress As String

    Public userRole As String
    Public userFullName As String
    'Fetch role and username to display specific name
    Public Function GetUserRole(username As String) As String
        Dim query As String = "SELECT Role FROM UserAccountTb WHERE UserName = @UserName"
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@UserName", username)
                con.Open()
                Dim role As String = Convert.ToString(cmd.ExecuteScalar())
                Return role
            End Using
        End Using
    End Function
    'Fetch Product details
    Public Class Product
        Public Property Barcode As String
        Public Property Brand As String
        Public Property Model As String
        Public Property ProductName As String
        Public Property Price As Decimal
        Public Property Quantity As Integer
        Public Property VatStatus As String
        Public Property ProductStatus As String
    End Class

    'Fetch Store Details
    Public Sub LoadStoreDetails()
        Using connection As New SqlConnection(ConnectionString)
            Dim query As String = "SELECT StoreName, StoreType, StoreAddress FROM StoreTable WHERE Id = 1"
            Dim command As New SqlCommand(query, connection)

            Try
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.Read() Then
                    StoreName = reader("StoreName").ToString()
                    StoreType = reader("StoreType").ToString()
                    StoreAddress = reader("StoreAddress").ToString()
                End If

                reader.Close()
            Catch ex As Exception
                MessageBox.Show("Error loading store details: " & ex.Message)
            End Try
        End Using
    End Sub



    'Fetch product details for transaction
    Public Function GetProductDetails(barcode As String) As Product
        Dim product As Product = Nothing ' Set to Nothing initially to handle not found cases

        Using connection As New SqlConnection(ConnectionString)
            Dim query As String = "SELECT Barcode, Brand, Model, ProductName, Price, Quantity, VatStatus, ProductStatus FROM InventoryTb WHERE Barcode = @Barcode"
            Dim command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@Barcode", barcode)

            Try
                connection.Open()
                Dim reader As SqlDataReader = command.ExecuteReader()

                If reader.Read() Then
                    ' Populate the Product object with data from the database
                    product = New Product With {
                    .Barcode = reader("Barcode").ToString(),
                    .Brand = reader("Brand").ToString(),
                    .Model = reader("Model").ToString(),
                    .ProductName = reader("ProductName").ToString(),
                    .Price = Convert.ToDecimal(reader("Price")),
                    .Quantity = Convert.ToInt32(reader("Quantity")),
                    .VatStatus = reader("VatStatus").ToString(),
                    .ProductStatus = reader("ProductStatus").ToString()
                }
                End If
            Catch ex As Exception
                MessageBox.Show("Error retrieving product details: " & ex.Message)
            End Try
        End Using

        Return product
    End Function


    'Function to foGenarate receipt number

    Public Function GenerateReceiptNo() As Integer
        Dim lastReceiptNo As Integer = 20240 ' Base receipt number
        Dim query As String = "SELECT ISNULL(MAX(ReceiptNo), 20240) FROM PaymentTb"

        ' Open database connection
        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                Try
                    con.Open()

                    ' Get the last receipt number from the PaymentTb table
                    Dim result As Object = cmd.ExecuteScalar()

                    If result IsNot Nothing Then
                        lastReceiptNo = Convert.ToInt32(result) ' Update with the latest receipt number
                    End If

                Catch ex As Exception
                    MessageBox.Show("Error fetching last receipt number: " & ex.Message)
                Finally
                    con.Close()
                End Try
            End Using
        End Using

        ' Increment the last receipt number
        lastReceiptNo += 1

        ' Return the new receipt number
        Return lastReceiptNo
    End Function

    'Generate receipt barcode number function
    Public Function GenerateReceiptBarcodeNo() As String
        Dim random As New Random()
        Dim barcode As String = ""
        For i As Integer = 1 To 11
            barcode &= random.Next(0, 10).ToString()
        Next
        Return barcode
    End Function



    'For searching function
    Public Sub FilterInventory(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()

                ' Define the SQL query to search in Brand, Model, or ProductName
                Dim query As String = "SELECT Barcode, Category, ProductName, Brand, Model, CapitalPrice, Price, Quantity, Expense, VatStatus, ProductStatus, DateIn, Supplier FROM InventoryTb WHERE " &
                                      "(Barcode LIKE @searchText)" &
                                      "ORDER BY DateIn DESC"

                Dim cmd As New SqlCommand(query, con)
                ' Append '%' wildcard directly to the parameter value
                cmd.Parameters.AddWithValue("@searchText", searchText & "%")

                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)


                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt

                    If dgv.Columns.Contains("Expense") Then
                        dgv.Columns("Expense").Visible = False
                    End If
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub

    'Search function for inventory record
    Public Sub FilterDataGridView1(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()


                Dim query As String = "SELECT Barcode, Category, ProductName, Brand, Model, CapitalPrice, Price, Quantity, Expense, VatStatus, DateIn, Supplier FROM InventoryRecord WHERE " &
                                      "(Barcode LIKE @searchText)"




                Dim cmd As New SqlCommand(query, con)

                cmd.Parameters.AddWithValue("@searchText", searchText & "%")

                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)

                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub

    'Search function for inventory list
    Public Sub FilterDataGridView2(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()


                Dim query As String = "SELECT Category, ProductName, Brand, Model, Quantity, Barcode FROM InventoryTb WHERE " &
                                      "(ProductName LIKE @searchText)"

                Dim cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@searchText", searchText & "%")
                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)

                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub
    'Search function for variety
    Public Sub FilterDataGridView3(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()


                Dim query As String = "SELECT Id, Category, ProductName, Brand, Model FROM VarietyTb WHERE " &
                                      "(Category LIKE @searchText OR " &
                                      "ProductName LIKE @searchText OR " &
                                      "Brand LIKE @searchText OR " &
                                      "Model LIKE @searchText)"

                Dim cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@searchText", searchText & "%")
                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)

                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub

    'Serach function for supplier
    Public Sub FilterSupplier(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()

                Dim query As String = "SELECT * FROM SupplierTb WHERE Name LIKE @searchText"
                Dim cmd As New SqlCommand(query, con)

                cmd.Parameters.AddWithValue("@searchText", "%" & searchText & "%")

                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)


                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub

    'Seach function for Account table
    Public Sub FilterAccount(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()

                Dim query As String = "SELECT Id, UserName, Password, Role, FullName, ContactNo, Email, Address, Status  FROM UserAccountTb WHERE FullName LIKE @searchText ORDER BY Id DESC"
                Dim cmd As New SqlCommand(query, con)

                cmd.Parameters.AddWithValue("@searchText", "%" & searchText & "%")

                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)


                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt

                    ' Mask Password and Role columns with asterisks (*)
                    For Each row As DataGridViewRow In dgv.Rows
                        If row.Cells("Password").Value IsNot Nothing Then
                            row.Cells("Password").Value = New String("*"c, row.Cells("Password").Value.ToString().Length)
                        End If

                        If row.Cells("Role").Value IsNot Nothing Then
                            row.Cells("Role").Value = New String("*"c, row.Cells("Role").Value.ToString().Length)
                        End If
                    Next
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub
    'Search function for transaction
    Public Sub FilterTransact(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()

                Dim query As String = "SELECT i.Barcode, i.Category, i.Brand, i.ProductName, i.Model, p.CustomerName, p.MOP, p.Amount, p.Discount, p.Change, p.ReferenceNo, p.Quantity, i.Price, p.TotalPrice, p.HandledBy, p.ReceiptBarcodeNo, p.DateIssued " &
                       "FROM InventoryTb AS i " &
                       "JOIN PaymentTb AS p ON i.Barcode = p.Barcode where p.ReceiptBarcodeNo LIKE @searchText " &
                         "ORDER BY p.Id DESC"

                Dim cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@searchText", "%" & searchText & "%")
                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)
                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub


    Public Sub FilterRefund(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()

                ' Include RefundExpirationDate in the query
                Dim query As String = "SELECT p.Id, p.Barcode, i.Brand, i.Model, i.ProductName, p.Quantity, i.Price, p.Discount, p.TotalPrice, p.VATable, p.VAT, p.NonVAT, i.VatStatus, p.ReceiptBarcodeNo, p.RefundExpirationDate " &
                                  "FROM InventoryTb AS i " &
                                  "JOIN PaymentTb AS p ON i.Barcode = p.Barcode " &
                                  "WHERE p.ReceiptBarcodeNo = @searchText"

                Dim cmd As New SqlCommand(query, con)
                cmd.Parameters.AddWithValue("@searchText", searchText)

                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)

                dgv.DataSource = Nothing
                dgv.Rows.Clear()

                If dt.Rows.Count > 0 Then
                    ' Filter out expired or today's refund expiration
                    Dim filteredRows = dt.AsEnumerable().Where(Function(row)
                                                                   Dim refundExpirationDate As DateTime
                                                                   If DateTime.TryParse(row.Field(Of Object)("RefundExpirationDate").ToString(), refundExpirationDate) Then
                                                                       Return refundExpirationDate > DateTime.Now.Date ' strictly greater than today
                                                                   Else
                                                                       Return False ' If can't parse date, treat as invalid
                                                                   End If
                                                               End Function)

                    If filteredRows.Any() Then
                        dgv.DataSource = filteredRows.CopyToDataTable()
                        ' Optionally hide RefundExpirationDate column if you don't want users to see it
                        If dgv.Columns.Contains("RefundExpirationDate") Then
                            dgv.Columns("RefundExpirationDate").Visible = False
                        End If
                    Else
                        MessageBox.Show("This item's refund period has expired.", "Refund Expired", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Else
                    MessageBox.Show("No matching records.", "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub


    Public Sub FilterRefundID(searchText As String, dgv As DataGridView, connectionString As String)
        Try
            Using con As New SqlConnection(connectionString)
                con.Open()

                Dim query As String = "SELECT p.Id,  i.Brand, i.Model, i.ProductName, p.Quantity, i.Price, p.TotalPrice, p.ReceiptBarcodeNo " &
                          "FROM InventoryTb AS i " &
                       "JOIN PaymentTb AS p ON i.Barcode = p.Barcode where p.Id LIKE @searchText "
                Dim cmd As New SqlCommand(query, con)
                ' Add wildcard for search text
                cmd.Parameters.AddWithValue("@searchText", "%" & searchText & "%")

                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                adapter.Fill(dt)

                ' If rows are found, bind the DataGridView
                If dt.Rows.Count > 0 Then
                    dgv.DataSource = dt
                End If
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while filtering data: " & ex.Message)
        End Try
    End Sub

    'Fetch Date using sqlserverdate
    Public Function GetSQLServerDate() As DateTime
        Try
            Dim query As String = "SELECT GETDATE()"
            Using con As New SqlConnection(ConnectionString)
                con.Open()
                Using cmd As New SqlCommand(query, con)
                    Return Convert.ToDateTime(cmd.ExecuteScalar())
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Error fetching date from server: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return DateTime.MinValue
        End Try
    End Function



End Module
