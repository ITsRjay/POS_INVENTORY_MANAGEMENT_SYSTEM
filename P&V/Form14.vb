Imports System.Data.SqlClient

Public Class Form14
    Private Sub Form14_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StoreDetails()

    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim storeName As String = TextBox3.Text   ' Assuming you have TextBoxes for input
        Dim storeType As String = TextBox4.Text
        Dim storeAddress As String = TextBox5.Text

        Using connection As New SqlConnection(ConnectionString)
            Dim checkQuery As String = "SELECT COUNT(*) FROM StoreTable"
            Dim checkCommand As New SqlCommand(checkQuery, connection)

            Try
                connection.Open()
                Dim recordCount As Integer = CInt(checkCommand.ExecuteScalar())

                ' Check if there is already one record in the table
                If recordCount > 0 Then
                    MessageBox.Show("Only one store detail entry is allowed. Please update the existing entry.")
                    Exit Sub
                End If

                Dim insertQuery As String = "INSERT INTO StoreTable (StoreName, StoreType, StoreAddress) VALUES (@StoreName, @StoreType, @StoreAddress)"
                Dim insertCommand As New SqlCommand(insertQuery, connection)
                insertCommand.Parameters.AddWithValue("@StoreName", storeName)
                insertCommand.Parameters.AddWithValue("@StoreType", storeType)
                insertCommand.Parameters.AddWithValue("@StoreAddress", storeAddress)

                insertCommand.ExecuteNonQuery()
                MessageBox.Show("Store details have been inserted successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
                MessageBox.Show("The application will restart automatically to apply the updates.", "Application Restart Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Application.Exit()
                StoreDetails()

            Catch ex As Exception
                MessageBox.Show("Error inserting store details: " & ex.Message)
            End Try
        End Using
    End Sub

    Private Sub StoreDetails()
        Using connection As New SqlConnection(ConnectionString)
            Dim query As String = "SELECT StoreName, StoreType, StoreAddress FROM StoreTable"
            Dim adapter As New SqlDataAdapter(query, connection)
            Dim table As New DataTable()

            Try
                connection.Open()
                adapter.Fill(table)
                DataGridView1.DataSource = table
            Catch ex As Exception
                MessageBox.Show("Error loading store details: " & ex.Message)
            End Try
        End Using
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        ' Check if the click is on a valid row
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            TextBox3.Text = row.Cells("StoreName").Value.ToString()
            TextBox4.Text = row.Cells("StoreType").Value.ToString()
            TextBox5.Text = row.Cells("StoreAddress").Value.ToString()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim storeName As String = TextBox3.Text
        Dim storeType As String = TextBox4.Text
        Dim storeAddress As String = TextBox5.Text

        Using connection As New SqlConnection(ConnectionString)
            Dim updateQuery As String = "UPDATE StoreTable SET StoreName = @StoreName, StoreType = @StoreType, StoreAddress = @StoreAddress"

            Dim command As New SqlCommand(updateQuery, connection)
            command.Parameters.AddWithValue("@StoreName", storeName)
            command.Parameters.AddWithValue("@StoreType", storeType)
            command.Parameters.AddWithValue("@StoreAddress", storeAddress)

            Try
                connection.Open()
                Dim rowsAffected As Integer = command.ExecuteNonQuery()

                If rowsAffected > 0 Then
                    MessageBox.Show("The store details have been updated successfully.", "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    MessageBox.Show("The application will restart automatically to apply the updates.", "Application Restart Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Application.Exit()
                Else
                    MessageBox.Show("No record found to update.", "No Record", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                StoreDetails()
            Catch ex As Exception
                MessageBox.Show("Error updating store details: " & ex.Message)
            End Try
        End Using
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox4.Clear()
        TextBox3.Clear()
        TextBox5.Clear()
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim message As String = "Please click a row first to update the address if it already exists." & vbCrLf & vbCrLf &
                                "Follow these guidelines for input:" & vbCrLf &
                                "- Store name must be below 12 characters." & vbCrLf &
                                "- Store type must be below 16 characters (e.g., Vape Shop, Sari-Sari Store)." & vbCrLf &
                                "- Address must be below 30 characters."
        MessageBox.Show(message, "Input Guidelines", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class