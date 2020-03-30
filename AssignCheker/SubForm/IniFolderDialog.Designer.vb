<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IniFolderDialog
    Inherits MetroFramework.Forms.MetroForm

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
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

    'Windows フォーム デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナーで必要です。
    'Windows フォーム デザイナーを使用して変更できます。  
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.metroButton2 = New MetroFramework.Controls.MetroButton()
        Me.metroButton1 = New MetroFramework.Controls.MetroButton()
        Me.Folderpath = New MetroFramework.Controls.MetroTextBox()
        Me.metroLabel1 = New MetroFramework.Controls.MetroLabel()
        Me.SuspendLayout()
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 27)
        '
        'metroButton2
        '
        Me.metroButton2.Location = New System.Drawing.Point(166, 92)
        Me.metroButton2.Name = "metroButton2"
        Me.metroButton2.Size = New System.Drawing.Size(92, 23)
        Me.metroButton2.TabIndex = 5
        Me.metroButton2.Text = "Ok"
        Me.metroButton2.UseSelectable = True
        '
        'metroButton1
        '
        Me.metroButton1.Location = New System.Drawing.Point(333, 63)
        Me.metroButton1.Name = "metroButton1"
        Me.metroButton1.Size = New System.Drawing.Size(75, 23)
        Me.metroButton1.TabIndex = 4
        Me.metroButton1.Text = "参照"
        Me.metroButton1.UseSelectable = True
        '
        'Folderpath
        '
        '
        '
        '
        Me.Folderpath.CustomButton.Image = Nothing
        Me.Folderpath.CustomButton.Location = New System.Drawing.Point(282, 1)
        Me.Folderpath.CustomButton.Name = ""
        Me.Folderpath.CustomButton.Size = New System.Drawing.Size(21, 21)
        Me.Folderpath.CustomButton.Style = MetroFramework.MetroColorStyle.Blue
        Me.Folderpath.CustomButton.TabIndex = 1
        Me.Folderpath.CustomButton.Theme = MetroFramework.MetroThemeStyle.Light
        Me.Folderpath.CustomButton.UseSelectable = True
        Me.Folderpath.CustomButton.Visible = False
        Me.Folderpath.Lines = New String(-1) {}
        Me.Folderpath.Location = New System.Drawing.Point(22, 63)
        Me.Folderpath.MaxLength = 32767
        Me.Folderpath.Name = "Folderpath"
        Me.Folderpath.PasswordChar = Global.Microsoft.VisualBasic.ChrW(0)
        Me.Folderpath.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.Folderpath.SelectedText = ""
        Me.Folderpath.SelectionLength = 0
        Me.Folderpath.SelectionStart = 0
        Me.Folderpath.ShortcutsEnabled = True
        Me.Folderpath.Size = New System.Drawing.Size(304, 23)
        Me.Folderpath.TabIndex = 3
        Me.Folderpath.UseSelectable = True
        Me.Folderpath.WaterMarkColor = System.Drawing.Color.FromArgb(CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer), CType(CType(109, Byte), Integer))
        Me.Folderpath.WaterMarkFont = New System.Drawing.Font("Segoe UI", 12.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Pixel)
        '
        'metroLabel1
        '
        Me.metroLabel1.Location = New System.Drawing.Point(22, 32)
        Me.metroLabel1.Name = "metroLabel1"
        Me.metroLabel1.Size = New System.Drawing.Size(386, 23)
        Me.metroLabel1.TabIndex = 6
        Me.metroLabel1.Text = "INIフォルダ(販売フォルダ)を選択してください"
        '
        'IniFolderDialog
        '
        Me.AllowDrop = True
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange
        Me.ClientSize = New System.Drawing.Size(421, 136)
        Me.Controls.Add(Me.metroLabel1)
        Me.Controls.Add(Me.metroButton2)
        Me.Controls.Add(Me.metroButton1)
        Me.Controls.Add(Me.Folderpath)
        Me.KeyPreview = True
        Me.Name = "IniFolderDialog"
        Me.Padding = New System.Windows.Forms.Padding(4, 60, 4, 4)
        Me.ShadowType = MetroFramework.Forms.MetroFormShadowType.AeroShadow
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Style = PrimaryForm.Style
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Private WithEvents metroButton2 As MetroFramework.Controls.MetroButton
    Private WithEvents metroButton1 As MetroFramework.Controls.MetroButton
    Private WithEvents Folderpath As MetroFramework.Controls.MetroTextBox
    Private WithEvents metroLabel1 As MetroFramework.Controls.MetroLabel

End Class
