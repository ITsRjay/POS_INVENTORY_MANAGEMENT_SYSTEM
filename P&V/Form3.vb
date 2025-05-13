Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting


Public Class Form3


    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadChartData()
        Chart1.Invalidate()
        InventoryRecord()
        Purchase()
        Refund()
        Inventory()
        Timer1.Interval = 10000
        Timer1.Start()
        Me.KeyPreview = True
    End Sub
    Public Sub New()
        InitializeComponent()

    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        RefreshDashboard()
    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        RefreshDashboard()

    End Sub
    Private Sub RefreshDashboard()
        LoadChartData()
        Inventory()
        InventoryRecord()
        Purchase()
        Refund()
    End Sub

    Private Sub LoadChartData()
        Try
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                Dim query As String = "SELECT ReportDate, TotalSales FROM SalesReportTB ORDER BY ReportDate"
                Using cmd As New SqlCommand(query, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()

                        ' Clear previous chart data
                        Chart1.Series.Clear()

                        ' Create new series
                        Dim series As New Series("SalesData")
                        With series
                            .ChartType = SeriesChartType.Column
                            .BorderWidth = 2
                            .Color = Color.FromArgb(19, 50, 74)
                            .XValueType = ChartValueType.Date
                            .MarkerStyle = MarkerStyle.Circle
                            .MarkerSize = 5
                            .IsValueShownAsLabel = False
                        End With

                        Chart1.Series.Add(series)

                        ' Add data points from DB
                        While reader.Read()
                            Dim reportDate As DateTime = Convert.ToDateTime(reader("ReportDate"))
                            Dim totalSales As Decimal = Convert.ToDecimal(reader("TotalSales"))
                            series.Points.AddXY(reportDate.ToOADate(), totalSales)
                        End While
                    End Using
                End Using
            End Using

            ' Setup chart appearance
            With Chart1.ChartAreas(0)
                .AxisX.LabelStyle.Format = "MM/dd/yyyy"
                .AxisX.IntervalType = DateTimeIntervalType.Days
                .AxisX.MajorGrid.LineColor = Color.LightGray
                .AxisX.LabelStyle.Angle = -45
                .AxisY.MajorGrid.LineColor = Color.LightGray
            End With

            Chart1.ChartAreas(0).RecalculateAxesScale()
            Chart1.Invalidate()
            Chart1.Update()

        Catch ex As Exception
            MessageBox.Show("Error loading chart data: " & ex.Message)
        End Try
    End Sub



    Private Sub Form3_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown

        If e.Control AndAlso e.KeyCode = Keys.R Then
            Button17.PerformClick()
        End If
    End Sub
    Public Sub Inventory()
        Dim query As String = "SELECT SUM(Quantity) AS TotalQuantity FROM InventoryTb WHERE ProductStatus NOT IN ('DAMAGED', 'LOSS PRODUCT', 'INACTIVE')"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                con.Open()

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim totalQuantity As Integer = If(IsDBNull(reader("TotalQuantity")), 0, Convert.ToInt32(reader("TotalQuantity")))
                        Label15.Text = totalQuantity.ToString("N0")
                    End If
                End Using

                con.Close()
            End Using
        End Using
    End Sub

    Public Sub InventoryRecord()
        Dim query As String = "SELECT SUM(Quantity) AS TotalQuantity, SUM(Expense) AS TotalExpense, SUM(Price * Quantity) AS TotalRevenue FROM InventoryRecord"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                con.Open()

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then

                        Dim totalQuantity As Integer = If(IsDBNull(reader("TotalQuantity")), 0, Convert.ToInt32(reader("TotalQuantity")))
                        Dim totalExpense As Decimal = If(IsDBNull(reader("TotalExpense")), 0, Convert.ToDecimal(reader("TotalExpense")))
                        Dim totalRevenue As Decimal = If(IsDBNull(reader("TotalRevenue")), 0, Convert.ToDecimal(reader("TotalRevenue")))
                        Dim totalProfit As Decimal = totalRevenue - totalExpense


                        Label7.Text = totalQuantity.ToString("N0")
                        Label8.Text = totalExpense.ToString("N0")
                        Label9.Text = totalProfit.ToString("N0")
                        Label17.Text = totalRevenue.ToString("N0")

                    End If
                End Using
                con.Close()

            End Using
        End Using
    End Sub

    Public Sub Purchase()
        Dim query As String = "SELECT SUM(Quantity) AS TotalQuantity, SUM(TotalPrice) AS TotalPurchase, SUM(NonVat) AS TotalNonVat, SUM(Vat) AS TotalVat, SUM(VATable) AS TotalVATable FROM PaymentTb"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                con.Open()

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim totalQuantity As Integer = If(IsDBNull(reader("TotalQuantity")), 0, Convert.ToInt32(reader("TotalQuantity")))
                        Dim totalPurchase As Decimal = If(IsDBNull(reader("TotalPurchase")), 0, Convert.ToDecimal(reader("TotalPurchase")))
                        Dim totalNonVat As Decimal = If(IsDBNull(reader("TotalNonVat")), 0, Convert.ToDecimal(reader("TotalNonVat")))
                        Dim totalVat As Decimal = If(IsDBNull(reader("TotalVat")), 0, Convert.ToDecimal(reader("TotalVat")))
                        Dim totalVatable As Decimal = If(IsDBNull(reader("TotalVATable")), 0, Convert.ToDecimal(reader("TotalVATable")))

                        Dim totalSales As Decimal = totalPurchase



                        Label14.Text = totalQuantity.ToString("N0")
                        Label10.Text = totalSales.ToString("N0")
                    End If
                End Using
                con.Close()
            End Using
        End Using
    End Sub

    Public Sub Refund()


        Dim query As String = "SELECT SUM(Quantity) AS TotalQuantity, SUM(RefundAmount) AS TotalRefund FROM RefundTb"

        Using con As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand(query, con)
                con.Open() ' Open the connection

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        ' Get the total quantity and total expense from the query
                        Dim totalQuantity As Integer = If(IsDBNull(reader("TotalQuantity")), 0, Convert.ToInt32(reader("TotalQuantity")))
                        Dim totalRefund As Integer = If(IsDBNull(reader("TotalRefund")), 0, Convert.ToInt32(reader("TotalRefund")))
                        Label6.Text = totalQuantity.ToString("N0")
                        Label12.Text = totalRefund.ToString("N0")
                    End If
                End Using
                con.Close()

            End Using
        End Using
    End Sub



    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
        Purchase()
    End Sub
End Class