<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form15
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form15))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button17 = New System.Windows.Forms.Button()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(19, Byte), Integer), CType(CType(50, Byte), Integer), CType(CType(74, Byte), Integer))
        Me.Button2.Image = Global.P_V.My.Resources.Resources.P123
        Me.Button2.Location = New System.Drawing.Point(1507, 82)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(62, 57)
        Me.Button2.TabIndex = 52
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Button17
        '
        Me.Button17.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(166, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button17.Image = Global.P_V.My.Resources.Resources.refresh__1_
        Me.Button17.Location = New System.Drawing.Point(1575, 82)
        Me.Button17.Name = "Button17"
        Me.Button17.Size = New System.Drawing.Size(62, 57)
        Me.Button17.TabIndex = 50
        Me.Button17.UseVisualStyleBackColor = False
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Visible = False
        '
        'PrintDocument1
        '
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.Label4.Location = New System.Drawing.Point(106, 10)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(251, 29)
        Me.Label4.TabIndex = 37
        Me.Label4.Text = "VIEW TRANSACTION"
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(19, Byte), Integer), CType(CType(50, Byte), Integer), CType(CType(74, Byte), Integer))
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.ListView1)
        Me.Panel1.Controls.Add(Me.Button7)
        Me.Panel1.Location = New System.Drawing.Point(558, 224)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(483, 557)
        Me.Panel1.TabIndex = 49
        '
        'ListView1
        '
        Me.ListView1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(3, 45)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(477, 508)
        Me.ListView1.TabIndex = 35
        Me.ListView1.UseCompatibleStateImageBehavior = False
        '
        'Button7
        '
        Me.Button7.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.Button7.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button7.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button7.Location = New System.Drawing.Point(437, 5)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(41, 36)
        Me.Button7.TabIndex = 36
        Me.Button7.Text = "X"
        Me.Button7.UseVisualStyleBackColor = False
        '
        'ComboBox1
        '
        Me.ComboBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(12, 145)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(242, 37)
        Me.ComboBox1.TabIndex = 48
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.Label1.Location = New System.Drawing.Point(753, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 20)
        Me.Label1.TabIndex = 47
        Me.Label1.Text = "SALES LOG"
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker2.Location = New System.Drawing.Point(1436, 145)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(198, 35)
        Me.DateTimePicker2.TabIndex = 46
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1396, 155)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 20)
        Me.Label3.TabIndex = 45
        Me.Label3.Text = "TO:"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateTimePicker1.Location = New System.Drawing.Point(1162, 143)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(210, 35)
        Me.DateTimePicker1.TabIndex = 44
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(1096, 155)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(60, 20)
        Me.Label2.TabIndex = 43
        Me.Label2.Text = "FROM:"
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
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ButtonHighlight
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
        Me.DataGridView1.Location = New System.Drawing.Point(12, 193)
        Me.DataGridView1.MaximumSize = New System.Drawing.Size(1625, 895)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidth = 62
        Me.DataGridView1.RowTemplate.Height = 28
        Me.DataGridView1.Size = New System.Drawing.Size(1625, 895)
        Me.DataGridView1.TabIndex = 41
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft YaHei", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(25, 19)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(92, 25)
        Me.Label16.TabIndex = 53
        Me.Label16.Text = "................"
        Me.Label16.Visible = False
        '
        'Form15
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1649, 1106)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button17)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "Form15"
        Me.Text = "Form15"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button2 As Button
    Friend WithEvents Button17 As Button
    Friend WithEvents PrintPreviewDialog1 As PrintPreviewDialog
    Friend WithEvents PrintDocument1 As Printing.PrintDocument
    Friend WithEvents Label4 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents ListView1 As ListView
    Friend WithEvents Button7 As Button
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents Label1 As Label
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents Label3 As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Label2 As Label
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents Label16 As Label
End Class
