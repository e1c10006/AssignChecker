Public Class DropDownTextBox : Inherits TextBox

    Private WithEvents ListBox1 As New ListBox
    Private ToolStripDropDown1 As New ToolStripDropDown
    Private ExSetting As String()

    Public Sub New()
        'InitializeComponent()

        Dim ToolStripControlHost1 As New ToolStripControlHost(ListBox1)
        ToolStripControlHost1.AutoSize = False
        ToolStripControlHost1.Margin = New Padding(0)

        ToolStripDropDown1.DropShadowEnabled = False
        ToolStripDropDown1.AutoSize = True
        ToolStripDropDown1.Padding = New Padding(0)
        ToolStripDropDown1.Items.Add(ToolStripControlHost1)
        ListBox1.Items.Add("=")
        ListBox1.Items.Add("<>")
        ListBox1.Items.Add("<=")
        ListBox1.Items.Add(">=")

        Me.TabStop = False
        Me.BackColor = Color.White
        Me.BorderStyle = Windows.Forms.BorderStyle.None
    End Sub

    Public Property AddListItems As String()
        Set(value As String())
            ExSetting = value
            If ExSetting IsNot Nothing AndAlso ExSetting.Length > 0 Then
                ListBox1.Items.Clear()
                For i As Integer = 0 To ExSetting.Length - 1
                    ListBox1.Items.Add(ExSetting(i))
                Next
            End If
        End Set
        Get
            Return ExSetting
        End Get
    End Property

    Public ReadOnly Property Items() As ListBox.ObjectCollection
        Get
            Return ListBox1.Items
        End Get
    End Property

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case e.KeyCode
            Case Keys.F1
                ListBox1.Height = ListBox1.Items.Count * (ListBox1.ItemHeight + 1)
                ToolStripDropDown1.Show(Me, -2, Me.Height - 3)
        End Select
        MyBase.OnKeyDown(e)
    End Sub

    Private Sub Text_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseClick
        Dim height As Decimal = ListBox1.Items.Count * (ListBox1.ItemHeight + 1)
        If height < 28 Then
            height = 28
        End If
        ListBox1.Height = height
        ListBox1.Width = 20
        ToolStripDropDown1.Show(Me, -2, Me.Height - 3)
    End Sub

    Private Sub ListBox1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListBox1.MouseClick
        If ListBox1.SelectedItem IsNot Nothing Then
            Me.Text = ListBox1.SelectedItem.ToString
            ToolStripDropDown1.Hide()
        End If
    End Sub
End Class
