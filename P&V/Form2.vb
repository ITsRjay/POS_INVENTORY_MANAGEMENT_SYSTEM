Imports System.IO
Imports System.Net
Imports System.Reflection.Emit
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class Form2

    Dim isCollapsed As Boolean = True
    Dim isCollapsed1 As Boolean = True
    Dim isCollapsed2 As Boolean = True

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetRoleVisibility(userRole)
        SetFullName(userFullName)

        Panel4.Size = Panel4.MinimumSize
        Panel5.Size = Panel5.MinimumSize
        Panel6.Size = Panel6.MinimumSize
        Me.KeyPreview = True
    End Sub

    Public Sub New()
        InitializeComponent()
    End Sub


    ' Constructor to accept role and full name
    Public Sub New(role As String, fullName As String)
        InitializeComponent()
        userRole = role
        userFullName = fullName
    End Sub
    Public Sub SetFullName(fullName As String)
        Label16.Text = fullName
    End Sub
    Public Sub SetRoleVisibility(role As String)
        If role = "CASHIER" Then
            Switchpanel(Form6)
            Panel2.Left = 90
            Panel7.Left = 100
            Panel1.Visible = False
        ElseIf role = "ADMIN" Then
            Switchpanel(Form3)
            Panel7.Visible = False
            Panel2.Height = 1050
        End If
    End Sub



    Sub Switchpanel(ByVal panel As Form)

        Panel2.Controls.Clear()
        panel.TopLevel = False
        Panel2.Controls.Add(panel)
        panel.Show()

    End Sub

    Private Sub Form2_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' Check if the Esc key is pressed
        If e.KeyCode = Keys.Escape Then
            Button4.PerformClick()
            Button7.PerformClick()
        End If
        If e.Control AndAlso e.KeyCode = Keys.T Then
            Button5.PerformClick()
        End If
        If e.Control AndAlso e.KeyCode = Keys.R Then
            Button6.PerformClick()
        End If
    End Sub


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        Switchpanel(New Form3())
    End Sub
    Private Sub Label1_MouseEnter(sender As Object, e As EventArgs) Handles Label1.MouseEnter
        Label1.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label1_MouseLeave(sender As Object, e As EventArgs) Handles Label1.MouseLeave
        Label1.ForeColor = Color.White ' Change this back to the original color
    End Sub


    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        Switchpanel(Form4)
    End Sub
    Private Sub Label2_MouseEnter(sender As Object, e As EventArgs) Handles Label2.MouseEnter
        Label2.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label2_MouseLeave(sender As Object, e As EventArgs) Handles Label2.MouseLeave
        Label2.ForeColor = Color.White ' Change this back to the original color
    End Sub


    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        Switchpanel(Form6)
    End Sub
    Private Sub Label6_MouseEnter(sender As Object, e As EventArgs) Handles Label6.MouseEnter
        Label6.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label6_MouseLeave(sender As Object, e As EventArgs) Handles Label6.MouseLeave
        Label6.ForeColor = Color.White ' Change this back to the original color
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Switchpanel(Form5)
    End Sub
    Private Sub Label3_MouseEnter(sender As Object, e As EventArgs) Handles Label3.MouseEnter
        Label3.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label3_MouseLeave(sender As Object, e As EventArgs) Handles Label3.MouseLeave
        Label3.ForeColor = Color.White ' Change this back to the original color
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        Switchpanel(Form7)

    End Sub
    Private Sub Label4_MouseEnter(sender As Object, e As EventArgs) Handles Label4.MouseEnter
        Label4.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label4_MouseLeave(sender As Object, e As EventArgs) Handles Label4.MouseLeave
        Label4.ForeColor = Color.White ' Change this back to the original color
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        Switchpanel(Form8)
    End Sub

    Private Sub Label5_MouseEnter(sender As Object, e As EventArgs) Handles Label5.MouseEnter
        Label5.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label5_MouseLeave(sender As Object, e As EventArgs) Handles Label5.MouseLeave
        Label5.ForeColor = Color.White ' Change this back to the original color
    End Sub
    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        Switchpanel(Form9)
    End Sub
    Private Sub Label7_MouseEnter(sender As Object, e As EventArgs) Handles Label7.MouseEnter
        Label7.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label7_MouseLeave(sender As Object, e As EventArgs) Handles Label7.MouseLeave
        Label7.ForeColor = Color.White ' Change this back to the original color
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        Switchpanel(Form10)
    End Sub
    Private Sub Label8_MouseEnter(sender As Object, e As EventArgs) Handles Label8.MouseEnter
        Label8.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label8_MouseLeave(sender As Object, e As EventArgs) Handles Label8.MouseLeave
        Label8.ForeColor = Color.White ' Change this back to the original color
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        Switchpanel(Form11)
    End Sub
    Private Sub Label12_MouseEnter(sender As Object, e As EventArgs) Handles Label12.MouseEnter
        Label12.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label12_MouseLeave(sender As Object, e As EventArgs) Handles Label12.MouseLeave
        Label12.ForeColor = Color.White ' Change this back to the original color
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        Switchpanel(Form12)
    End Sub
    Private Sub Label13_MouseEnter(sender As Object, e As EventArgs) Handles Label13.MouseEnter
        Label13.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label13_MouseLeave(sender As Object, e As EventArgs) Handles Label13.MouseLeave
        Label13.ForeColor = Color.White ' Change this back to the original color
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        Switchpanel(Form13)
    End Sub
    Private Sub Label10_MouseEnter(sender As Object, e As EventArgs) Handles Label10.MouseEnter
        Label10.ForeColor = Color.PowderBlue ' Change this to any color you want
    End Sub

    Private Sub Label10_MouseLeave(sender As Object, e As EventArgs) Handles Label10.MouseLeave
        Label10.ForeColor = Color.White ' Change this back to the original color
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If isCollapsed Then
            Panel4.Height += 10
            Button1.Image = My.Resources.Ad
            If Panel4.Size = Panel4.MaximumSize Then
                Timer1.Stop()
                isCollapsed = False
            End If
        Else
            Panel4.Height -= 10
            Button1.Image = My.Resources.Ar
            If Panel4.Size = Panel4.MinimumSize Then
                Timer1.Stop()
                isCollapsed = True
            End If
        End If
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If isCollapsed1 Then
            Panel5.Height += 10
            Button2.Image = My.Resources.Ad
            If Panel5.Size = Panel5.MaximumSize Then
                Timer2.Stop()
                isCollapsed1 = False
            End If
        Else
            Panel5.Height -= 10
            Button2.Image = My.Resources.Ar
            If Panel5.Size = Panel5.MinimumSize Then
                Timer2.Stop()
                isCollapsed1 = True
            End If
        End If
    End Sub

    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        If isCollapsed2 Then
            Panel6.Height += 10
            Button3.Image = My.Resources.Ad
            If Panel6.Size = Panel6.MaximumSize Then
                Timer3.Stop()
                isCollapsed2 = False
            End If
        Else
            Panel6.Height -= 10
            Button3.Image = My.Resources.Ar
            If Panel6.Size = Panel6.MinimumSize Then
                Timer3.Stop()
                isCollapsed2 = True
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Timer1.Start()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Timer2.Start()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Timer3.Start()


    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs)
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to log out?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Form1.Show()
            Me.Hide()
            Form8.Close()
            Form6.Close()
        End If
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Switchpanel(Form6)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim result As DialogResult = MessageBox.Show("Are you sure you want to log out?", "Confirm Logout", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Form1.Show()
            Me.Hide()
            Form8.Close()
            Form6.Close()
        End If
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Switchpanel(Form8)
    End Sub

    Private Sub Panel2_Paint(sender As Object, e As PaintEventArgs) Handles Panel2.Paint

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Switchpanel(Form14)
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs)
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim form15 As New Form15(userRole, userFullName)
        Switchpanel(form15)
    End Sub
End Class