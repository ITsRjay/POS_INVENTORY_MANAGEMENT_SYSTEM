<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form8
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.TextBox8 = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button10 = New System.Windows.Forms.Button()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.Id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ProductName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Quantity = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TotalPrice = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IsEdited = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Location = New System.Drawing.Point(990, 97)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(193, 20)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "RECEIPT BARCODE NO:"
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ControlLight
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.TopCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.Location = New System.Drawing.Point(557, 200)
        Me.DataGridView1.MaximumSize = New System.Drawing.Size(1081, 818)
        Me.DataGridView1.MinimumSize = New System.Drawing.Size(1081, 818)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.RowTemplate.Height = 28
        Me.DataGridView1.Size = New System.Drawing.Size(1081, 818)
        Me.DataGridView1.TabIndex = 46
        '
        'TextBox1
        '
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.TextBox1.Location = New System.Drawing.Point(886, 120)
        Me.TextBox1.MaxLength = 11
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(400, 64)
        Me.TextBox1.TabIndex = 48
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PrintDocument1
        '
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'TextBox8
        '
        Me.TextBox8.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox8.Location = New System.Drawing.Point(193, 311)
        Me.TextBox8.Multiline = True
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(298, 111)
        Me.TextBox8.TabIndex = 0
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label11.Location = New System.Drawing.Point(31, 402)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(141, 20)
        Me.Label11.TabIndex = 87
        Me.Label11.Text = "REFUND CAUSE:"
        '
        'Button8
        '
        Me.Button8.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(67, Byte), Integer), CType(CType(67, Byte), Integer))
        Me.Button8.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button8.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button8.Location = New System.Drawing.Point(12, 663)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(175, 71)
        Me.Button8.TabIndex = 0
        Me.Button8.Text = "ENTER"
        Me.Button8.UseVisualStyleBackColor = False
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.FromArgb(CType(CType(107, Byte), Integer), CType(CType(61, Byte), Integer), CType(CType(40, Byte), Integer))
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button7.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button7.Location = New System.Drawing.Point(367, 663)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(175, 71)
        Me.Button7.TabIndex = 91
        Me.Button7.Text = "CLEAR"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label6.Location = New System.Drawing.Point(31, 220)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(153, 20)
        Me.Label6.TabIndex = 92
        Me.Label6.Text = "REFUND AMOUNT:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label5.Location = New System.Drawing.Point(59, 274)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(117, 20)
        Me.Label5.TabIndex = 93
        Me.Label5.Text = "REFUND QTY:"
        '
        'TextBox3
        '
        Me.TextBox3.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(193, 205)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(298, 35)
        Me.TextBox3.TabIndex = 0
        Me.TextBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox4
        '
        Me.TextBox4.BackColor = System.Drawing.SystemColors.ScrollBar
        Me.TextBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox4.Location = New System.Drawing.Point(193, 259)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(298, 35)
        Me.TextBox4.TabIndex = 1
        Me.TextBox4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(CType(19, Byte), Integer), CType(CType(50, Byte), Integer), CType(CType(74, Byte), Integer))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button1.Location = New System.Drawing.Point(190, 663)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(175, 71)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "PRINT"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button10
        '
        Me.Button10.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button10.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button10.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button10.Location = New System.Drawing.Point(1272, 3)
        Me.Button10.Name = "Button10"
        Me.Button10.Size = New System.Drawing.Size(35, 36)
        Me.Button10.TabIndex = 97
        Me.Button10.Text = "-"
        Me.Button10.UseVisualStyleBackColor = False
        '
        'TextBox2
        '
        Me.TextBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(25, 107)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(495, 77)
        Me.TextBox2.TabIndex = 0
        Me.TextBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TextBox2.Visible = False
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Id, Me.ProductName, Me.Quantity, Me.TotalPrice, Me.IsEdited})
        Me.DataGridView2.Location = New System.Drawing.Point(12, 740)
        Me.DataGridView2.MaximumSize = New System.Drawing.Size(530, 278)
        Me.DataGridView2.MinimumSize = New System.Drawing.Size(530, 278)
        Me.DataGridView2.MultiSelect = False
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.RowHeadersWidth = 62
        Me.DataGridView2.RowTemplate.Height = 28
        Me.DataGridView2.Size = New System.Drawing.Size(530, 278)
        Me.DataGridView2.TabIndex = 102
        Me.DataGridView2.TabStop = False
        '
        'Id
        '
        Me.Id.HeaderText = "ID"
        Me.Id.MinimumWidth = 8
        Me.Id.Name = "Id"
        Me.Id.ReadOnly = True
        '
        'ProductName
        '
        Me.ProductName.HeaderText = "PRODUCT"
        Me.ProductName.MinimumWidth = 8
        Me.ProductName.Name = "ProductName"
        Me.ProductName.ReadOnly = True
        '
        'Quantity
        '
        Me.Quantity.HeaderText = "QTY"
        Me.Quantity.MinimumWidth = 8
        Me.Quantity.Name = "Quantity"
        Me.Quantity.ReadOnly = True
        '
        'TotalPrice
        '
        Me.TotalPrice.HeaderText = "TOTAL CASH"
        Me.TotalPrice.MinimumWidth = 8
        Me.TotalPrice.Name = "TotalPrice"
        Me.TotalPrice.ReadOnly = True
        '
        'IsEdited
        '
        Me.IsEdited.HeaderText = "IsEdited"
        Me.IsEdited.MinimumWidth = 8
        Me.IsEdited.Name = "IsEdited"
        Me.IsEdited.ReadOnly = True
        Me.IsEdited.Visible = False
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(19, Byte), Integer), CType(CType(50, Byte), Integer), CType(CType(74, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label7)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Button10)
        Me.Panel1.Controls.Add(Me.TextBox4)
        Me.Panel1.Controls.Add(Me.TextBox3)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label11)
        Me.Panel1.Controls.Add(Me.TextBox8)
        Me.Panel1.Location = New System.Drawing.Point(12, 200)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(530, 457)
        Me.Panel1.TabIndex = 47
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label4.Location = New System.Drawing.Point(139, 59)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(247, 29)
        Me.Label4.TabIndex = 105
        Me.Label4.Text = "CLICK ID PER ITEM"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label7.Location = New System.Drawing.Point(80, 166)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(104, 20)
        Me.Label7.TabIndex = 104
        Me.Label7.Text = "ITEM PRICE:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label3.Location = New System.Drawing.Point(316, 159)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(55, 29)
        Me.Label3.TabIndex = 103
        Me.Label3.Text = "......."
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.Label15)
        Me.Panel6.Controls.Add(Me.Label16)
        Me.Panel6.Location = New System.Drawing.Point(1373, 13)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(265, 34)
        Me.Panel6.TabIndex = 60
        Me.Panel6.Visible = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Font = New System.Drawing.Font("Microsoft YaHei", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.Location = New System.Drawing.Point(13, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(98, 25)
        Me.Label15.TabIndex = 22
        Me.Label15.Text = "CASHIER:"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft YaHei", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(112, 8)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(92, 25)
        Me.Label16.TabIndex = 23
        Me.Label16.Text = "................"
        '
        'Timer1
        '
        '
        'Form8
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1649, 1099)
        Me.Controls.Add(Me.Panel6)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Form8"
        Me.Text = "RefundForm"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents PrintDialog1 As PrintDialog
    Friend WithEvents TextBox8 As TextBox
    Friend WithEvents Label11 As Label
    Friend WithEvents Button8 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Label6 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Button10 As Button
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label7 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Panel6 As Panel
    Friend WithEvents Label15 As Label
    Friend WithEvents Label16 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Id As DataGridViewTextBoxColumn
    Friend Shadows WithEvents ProductName As DataGridViewTextBoxColumn
    Friend WithEvents Quantity As DataGridViewTextBoxColumn
    Friend WithEvents TotalPrice As DataGridViewTextBoxColumn
    Friend WithEvents IsEdited As DataGridViewTextBoxColumn
    Friend WithEvents Timer1 As Timer
End Class
