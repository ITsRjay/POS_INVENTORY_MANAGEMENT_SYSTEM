<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form11
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Guna2TextBox1 = New Guna.UI2.WinForms.Guna2TextBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
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
        Me.DataGridView1.Location = New System.Drawing.Point(12, 201)
        Me.DataGridView1.MaximumSize = New System.Drawing.Size(1629, 895)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.RowTemplate.Height = 28
        Me.DataGridView1.Size = New System.Drawing.Size(1629, 895)
        Me.DataGridView1.TabIndex = 36
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker2.Location = New System.Drawing.Point(1443, 149)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(198, 37)
        Me.DateTimePicker2.TabIndex = 45
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1403, 163)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 20)
        Me.Label3.TabIndex = 44
        Me.Label3.Text = "TO:"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(1190, 149)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(198, 37)
        Me.DateTimePicker1.TabIndex = 43
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(1124, 163)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 20)
        Me.Label2.TabIndex = 42
        Me.Label2.Text = "FROM:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Location = New System.Drawing.Point(743, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(121, 20)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "ITEM RECORD"
        '
        'Button17
        '
        Me.Button17.Location = New System.Drawing.Point(0, 0)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(75, 23)
        Me.Button17.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(0, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(166, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button3.Image = Global.P_V.My.Resources.Resources.refresh__1_
        Me.Button3.Location = New System.Drawing.Point(1578, 84)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(62, 57)
        Me.Button3.TabIndex = 47
        Me.Button3.UseVisualStyleBackColor = False
        '
        'Guna2TextBox1
        '
        Me.Guna2TextBox1.BorderRadius = 20
        Me.Guna2TextBox1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Guna2TextBox1.DefaultText = ""
        Me.Guna2TextBox1.DisabledState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer), CType(CType(208, Byte), Integer))
        Me.Guna2TextBox1.DisabledState.FillColor = System.Drawing.Color.FromArgb(CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer), CType(CType(226, Byte), Integer))
        Me.Guna2TextBox1.DisabledState.ForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.Guna2TextBox1.DisabledState.PlaceholderForeColor = System.Drawing.Color.FromArgb(CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(138, Byte), Integer))
        Me.Guna2TextBox1.FillColor = System.Drawing.SystemColors.Control
        Me.Guna2TextBox1.FocusedState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Guna2TextBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Guna2TextBox1.ForeColor = System.Drawing.Color.Black
        Me.Guna2TextBox1.HoverState.BorderColor = System.Drawing.Color.FromArgb(CType(CType(94, Byte), Integer), CType(CType(148, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.Guna2TextBox1.IconLeft = Global.P_V.My.Resources.Resources.search__1_
        Me.Guna2TextBox1.Location = New System.Drawing.Point(12, 126)
        Me.Guna2TextBox1.Margin = New System.Windows.Forms.Padding(10, 5, 4, 5)
        Me.Guna2TextBox1.Name = "Guna2TextBox1"
        Me.Guna2TextBox1.PlaceholderForeColor = System.Drawing.Color.Gray
        Me.Guna2TextBox1.PlaceholderText = "Search by Barcode"
        Me.Guna2TextBox1.SelectedText = ""
        Me.Guna2TextBox1.Size = New System.Drawing.Size(374, 60)
        Me.Guna2TextBox1.TabIndex = 94
        Me.Guna2TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Form11
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1653, 1106)
        Me.Controls.Add(Me.Guna2TextBox1)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button17)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Form11"
        Me.Text = "ItemRecord"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents Button17 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Guna2TextBox1 As Guna.UI2.WinForms.Guna2TextBox
End Class
