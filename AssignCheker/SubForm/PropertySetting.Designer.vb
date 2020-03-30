<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettingForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SettingForm))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.基本設定 = New System.Windows.Forms.GroupBox()
        Me.テーマ = New System.Windows.Forms.ComboBox()
        Me.MaxView = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.H_ESCKBB = New System.Windows.Forms.Label()
        Me.H_SIZEKBB = New System.Windows.Forms.Label()
        Me.H_HEIGHT = New System.Windows.Forms.TextBox()
        Me.H_WIDTH = New System.Windows.Forms.TextBox()
        Me.H_SUBKBE = New System.Windows.Forms.Label()
        Me.H_SUBKBD = New System.Windows.Forms.Label()
        Me.H_SUBKBC = New System.Windows.Forms.Label()
        Me.H_SUBKBB = New System.Windows.Forms.Label()
        Me.H_ESCKBA = New System.Windows.Forms.Label()
        Me.H_SIZEKBA = New System.Windows.Forms.Label()
        Me.H_SUBKBA = New System.Windows.Forms.Label()
        Me.ESCKB = New System.Windows.Forms.TextBox()
        Me.SIZEKB = New System.Windows.Forms.TextBox()
        Me.SUBKB = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.H_ICONKBA = New System.Windows.Forms.Label()
        Me.H_ICONKBB = New System.Windows.Forms.Label()
        Me.ICONKB = New System.Windows.Forms.TextBox()
        Me.基本設定.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Transparent
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.Name = "Button1"
        Me.Button1.UseVisualStyleBackColor = False
        '
        '基本設定
        '
        Me.基本設定.Controls.Add(Me.テーマ)
        Me.基本設定.Controls.Add(Me.MaxView)
        Me.基本設定.Controls.Add(Me.Label6)
        Me.基本設定.Controls.Add(Me.Label7)
        Me.基本設定.Controls.Add(Me.H_ESCKBB)
        Me.基本設定.Controls.Add(Me.H_ICONKBB)
        Me.基本設定.Controls.Add(Me.H_SIZEKBB)
        Me.基本設定.Controls.Add(Me.H_HEIGHT)
        Me.基本設定.Controls.Add(Me.H_WIDTH)
        Me.基本設定.Controls.Add(Me.H_SUBKBE)
        Me.基本設定.Controls.Add(Me.H_SUBKBD)
        Me.基本設定.Controls.Add(Me.H_SUBKBC)
        Me.基本設定.Controls.Add(Me.H_SUBKBB)
        Me.基本設定.Controls.Add(Me.H_ESCKBA)
        Me.基本設定.Controls.Add(Me.H_ICONKBA)
        Me.基本設定.Controls.Add(Me.H_SIZEKBA)
        Me.基本設定.Controls.Add(Me.H_SUBKBA)
        Me.基本設定.Controls.Add(Me.ESCKB)
        Me.基本設定.Controls.Add(Me.ICONKB)
        Me.基本設定.Controls.Add(Me.SIZEKB)
        Me.基本設定.Controls.Add(Me.SUBKB)
        Me.基本設定.Controls.Add(Me.Label1)
        Me.基本設定.Controls.Add(Me.Label2)
        Me.基本設定.Controls.Add(Me.Label8)
        Me.基本設定.Controls.Add(Me.Label9)
        Me.基本設定.Controls.Add(Me.Label5)
        Me.基本設定.Controls.Add(Me.Label3)
        Me.基本設定.Controls.Add(Me.Label4)
        resources.ApplyResources(Me.基本設定, "基本設定")
        Me.基本設定.Name = "基本設定"
        Me.基本設定.TabStop = False
        '
        'テーマ
        '
        Me.テーマ.BackColor = System.Drawing.Color.White
        Me.テーマ.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.テーマ, "テーマ")
        Me.テーマ.FormattingEnabled = True
        Me.テーマ.Items.AddRange(New Object() {resources.GetString("テーマ.Items"), resources.GetString("テーマ.Items1"), resources.GetString("テーマ.Items2"), resources.GetString("テーマ.Items3"), resources.GetString("テーマ.Items4"), resources.GetString("テーマ.Items5")})
        Me.テーマ.Name = "テーマ"
        Me.テーマ.TabStop = False
        '
        'MaxView
        '
        resources.ApplyResources(Me.MaxView, "MaxView")
        Me.MaxView.Name = "MaxView"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'Label7
        '
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.Name = "Label7"
        '
        'H_ESCKBB
        '
        Me.H_ESCKBB.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.H_ESCKBB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        resources.ApplyResources(Me.H_ESCKBB, "H_ESCKBB")
        Me.H_ESCKBB.Name = "H_ESCKBB"
        '
        'H_SIZEKBB
        '
        Me.H_SIZEKBB.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.H_SIZEKBB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        resources.ApplyResources(Me.H_SIZEKBB, "H_SIZEKBB")
        Me.H_SIZEKBB.Name = "H_SIZEKBB"
        '
        'H_HEIGHT
        '
        resources.ApplyResources(Me.H_HEIGHT, "H_HEIGHT")
        Me.H_HEIGHT.Name = "H_HEIGHT"
        '
        'H_WIDTH
        '
        resources.ApplyResources(Me.H_WIDTH, "H_WIDTH")
        Me.H_WIDTH.Name = "H_WIDTH"
        '
        'H_SUBKBE
        '
        Me.H_SUBKBE.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.H_SUBKBE.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        resources.ApplyResources(Me.H_SUBKBE, "H_SUBKBE")
        Me.H_SUBKBE.Name = "H_SUBKBE"
        '
        'H_SUBKBD
        '
        Me.H_SUBKBD.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.H_SUBKBD.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        resources.ApplyResources(Me.H_SUBKBD, "H_SUBKBD")
        Me.H_SUBKBD.Name = "H_SUBKBD"
        '
        'H_SUBKBC
        '
        Me.H_SUBKBC.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.H_SUBKBC.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        resources.ApplyResources(Me.H_SUBKBC, "H_SUBKBC")
        Me.H_SUBKBC.Name = "H_SUBKBC"
        '
        'H_SUBKBB
        '
        Me.H_SUBKBB.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.H_SUBKBB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        resources.ApplyResources(Me.H_SUBKBB, "H_SUBKBB")
        Me.H_SUBKBB.Name = "H_SUBKBB"
        '
        'H_ESCKBA
        '
        Me.H_ESCKBA.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.H_ESCKBA.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.H_ESCKBA.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        resources.ApplyResources(Me.H_ESCKBA, "H_ESCKBA")
        Me.H_ESCKBA.Name = "H_ESCKBA"
        '
        'H_SIZEKBA
        '
        Me.H_SIZEKBA.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.H_SIZEKBA.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.H_SIZEKBA.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        resources.ApplyResources(Me.H_SIZEKBA, "H_SIZEKBA")
        Me.H_SIZEKBA.Name = "H_SIZEKBA"
        '
        'H_SUBKBA
        '
        Me.H_SUBKBA.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.H_SUBKBA.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.H_SUBKBA.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        resources.ApplyResources(Me.H_SUBKBA, "H_SUBKBA")
        Me.H_SUBKBA.Name = "H_SUBKBA"
        '
        'ESCKB
        '
        resources.ApplyResources(Me.ESCKB, "ESCKB")
        Me.ESCKB.Name = "ESCKB"
        '
        'SIZEKB
        '
        resources.ApplyResources(Me.SIZEKB, "SIZEKB")
        Me.SIZEKB.Name = "SIZEKB"
        '
        'SUBKB
        '
        resources.ApplyResources(Me.SUBKB, "SUBKB")
        Me.SUBKB.Name = "SUBKB"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'Label8
        '
        resources.ApplyResources(Me.Label8, "Label8")
        Me.Label8.Name = "Label8"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Transparent
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'Label9
        '
        resources.ApplyResources(Me.Label9, "Label9")
        Me.Label9.Name = "Label9"
        '
        'H_ICONKBA
        '
        Me.H_ICONKBA.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.H_ICONKBA.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.H_ICONKBA.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        resources.ApplyResources(Me.H_ICONKBA, "H_ICONKBA")
        Me.H_ICONKBA.Name = "H_ICONKBA"
        '
        'H_ICONKBB
        '
        Me.H_ICONKBB.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.H_ICONKBB.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        resources.ApplyResources(Me.H_ICONKBB, "H_ICONKBB")
        Me.H_ICONKBB.Name = "H_ICONKBB"
        '
        'ICONKB
        '
        resources.ApplyResources(Me.ICONKB, "ICONKB")
        Me.ICONKB.Name = "ICONKB"
        '
        'SettingForm
        '
        Me.AllowDrop = True
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoValidate = System.Windows.Forms.AutoValidate.EnableAllowFocusChange
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.基本設定)
        Me.Controls.Add(Me.Button1)
        Me.Name = "SettingForm"
        Me.Style = MetroFramework.MetroColorStyle.[Default]
        Me.基本設定.ResumeLayout(False)
        Me.基本設定.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents 基本設定 As System.Windows.Forms.GroupBox
    Friend WithEvents MaxView As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents H_SIZEKBB As System.Windows.Forms.Label
    Friend WithEvents H_HEIGHT As System.Windows.Forms.TextBox
    Friend WithEvents H_WIDTH As System.Windows.Forms.TextBox
    Friend WithEvents H_SUBKBD As System.Windows.Forms.Label
    Friend WithEvents H_SUBKBC As System.Windows.Forms.Label
    Friend WithEvents H_SUBKBB As System.Windows.Forms.Label
    Friend WithEvents H_SIZEKBA As System.Windows.Forms.Label
    Friend WithEvents H_SUBKBA As System.Windows.Forms.Label
    Friend WithEvents SIZEKB As System.Windows.Forms.TextBox
    Friend WithEvents SUBKB As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents テーマ As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents H_SUBKBE As System.Windows.Forms.Label
    Friend WithEvents H_ESCKBB As System.Windows.Forms.Label
    Friend WithEvents H_ESCKBA As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents ESCKB As System.Windows.Forms.TextBox
    Friend WithEvents H_ICONKBB As System.Windows.Forms.Label
    Friend WithEvents H_ICONKBA As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ICONKB As System.Windows.Forms.TextBox
End Class
