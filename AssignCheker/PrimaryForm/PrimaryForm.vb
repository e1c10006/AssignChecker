Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms
Imports System.Text
Imports System.Reflection
Imports System.Threading
Imports System
Imports System.Security.Permissions
Imports System.Collections
Imports System.ComponentModel
Imports System.Media
Imports System.Runtime.InteropServices

Imports ILL.ERF.BaseLib
Imports ILL.ERF.G1Base
Imports ILL.ERF.G2Base
Imports ILL.ERF.AKBS1010
Imports System.Collections.Specialized

Public Class PrimaryForm

#Region "  共通変数"
    Dim DBPass As String = ""

    Dim dsDATA As DataSet
    Dim cn As New SqlConnection()
    Dim cmd As SqlCommand = cn.CreateCommand()

    Dim Cost_dsDATA As DataSet
    Dim Cost_cn As New SqlConnection()
    Dim Cost_cmd As SqlCommand = Cost_cn.CreateCommand()

    Dim Price_dsDATA As DataSet
    Dim Price_cn As New SqlConnection()
    Dim Price_cmd As SqlCommand = Price_cn.CreateCommand()

    Dim Tancd_dsDATA As DataSet
    Dim Tancd_cn As New SqlConnection()
    Dim Tancd_cmd As SqlCommand = Tancd_cn.CreateCommand()
    Dim TancdSub_dsDATA As DataSet
    Dim TancdSub_cn As New SqlConnection()
    Dim TancdSub_cmd As SqlCommand = TancdSub_cn.CreateCommand()

    Dim Selector_cn As New SqlConnection()
    Dim Selector_cmd As SqlCommand = Selector_cn.CreateCommand()

    Dim L_CurrentX As Integer = Nothing
    Dim L_CurrentY As Integer = Nothing
    Dim SubFormCanOpen As Boolean = True
    Public ColumnCount As Integer = 20
    Public LoadEnd As Boolean = False

    Private hotkeyActivate As HotKey
    Private hotkeyQuery As HotKey
#End Region

#Region "  サブ画面表示用変数"
    Public Jdnno As String
    Public Taidt As String
    Public Taiym As String
    Public Tancd As String
    Public Bmnnm As String
    Public Subkb As String
#End Region

#Region "  起動処理 / 更新履歴"
    '初期項目の設定
    Private Sub RegisterForm(sender As System.Object, e As System.EventArgs) Handles MyBase.Load


        If Application.ExecutablePath.Contains("\\kdc-hyv03") Then
            MetroFramework.MetroMessageBox.Show(Me, "ローカルに落としてから使用して下さい")
            Application.Exit()
            Return
        End If

        Try
            ' Ctrl+Shift+Z をID 0のホットキーとして登録
            Try
                hotkeyActivate = New HotKey(Me.Handle, 0, Keys.Alt Or Keys.Shift Or Keys.Z)
                hotkeyQuery = New HotKey(Me.Handle, 0, Keys.Alt Or Keys.Shift Or Keys.C)
            Catch ex As Exception
                'エラーにしない
            End Try

            Dim info As FileInfo = New System.IO.FileInfo(Application.ExecutablePath)
            If My.Settings("LastUpdate") = "" Then
                My.Settings("LastUpdate") = info.LastWriteTime.ToString("yyyyMMdd")
            End If

            '更新チェック
            Dim ServerPath As String = "\\kdc-hyv03\o-dc02\共有\部署フォルダ\サポート\PG\個人フォルダ\有本\ツール\負荷アサインチェックツール\AssignCheck\AssignChecker.exe"
            If File.Exists(ServerPath) Then
                Dim serverdt As String = File.GetLastWriteTime(ServerPath).ToString("yyyyMMdd")
                Dim lastupd As String = My.Settings("LastUpdate")
                If Integer.Parse(serverdt) > Integer.Parse(lastupd) Then

                    Using reader As StreamReader = New StreamReader("\\kdc-hyv03\o-dc02\共有\部署フォルダ\サポート\PG\個人フォルダ\有本\ツール\負荷アサインチェックツール\修正ログ.txt", Encoding.GetEncoding("Shift_JIS"))
                        Dim txt As StringBuilder
                        txt = New StringBuilder
                        txt.Append(" << 修正履歴 >>" & vbCrLf)
                        txt.Append(reader.ReadLine() & vbCrLf)
                        txt.Append(reader.ReadLine() & vbCrLf)
                        txt.Append(reader.ReadLine() & vbCrLf)
                        txt.Append(reader.ReadLine() & vbCrLf)
                        MetroFramework.MetroMessageBox.Show(Me, "サーバーに最新版があります（更新日" & serverdt & "'）" & vbCrLf & txt.ToString, "確認")
                    End Using
                    My.Settings("LastUpdate") = serverdt
                End If
            End If

            '画面レイアウト読み込み
            FN_LoadStyle()

            'DataViewを編集可能に
            LoadEnd = True
            dsDATA = New DataSet("取得")
            Cost_dsDATA = New DataSet("取得")
            Price_dsDATA = New DataSet("取得")
            Tancd_dsDATA = New DataSet("取得")
            TancdSub_dsDATA = New DataSet("取得")
            BODY.TabPages(1).Select()
            BODY.SelectedIndex = 1

            'DoubleBuffer を有効にする
            Me.DoubleBuffered = True
            EnableDoubleBuffering(BODY)
            EnableDoubleBuffering(AssignGrid)
            EnableDoubleBuffering(PGGrid)
            EnableDoubleBuffering(TancdGrid)

            'フォームのサイズ変更
            FN_SetFormSize()

            '行選択ボタンを非表示にする
            AssignGrid.RowHeadersVisible = False
            CostGrid.RowHeadersVisible = False
            PGGrid.RowHeadersVisible = False
            TancdGrid.RowHeadersVisible = False
            AssignGrid.RowHeadersDefaultCellStyle.BackColor = Color.Aqua

            'アイコン設定
            Me.Icon = My.Resources.ResourceManager.GetObject("アイコン")
            H_Order.Items.Clear()
            For Each col As String In DirectCast(My.Settings("VisibleColumns"), StringCollection)
                If col.ToString.TrimEnd <> "" Then
                    H_Order.Items.Add(col)
                End If
            Next
            H_Order.Text = H_Order.Items(0).ToString.TrimEnd
            H_Orderkb.Text = "降順"
            H_WHERE7.Text = My.Settings("ERPNAME").ToString.TrimEnd
            C_WHERE05.Text = My.Settings("ERPNAME").ToString.TrimEnd

            ' -- 他タブ初期値設定 --

            '事例検索初期化
            RegiterForm_Assign()
            RegiterForm_PG()

            imageList1.Images.Add("FOLDER", My.Resources.FOLDER)
            imageList1.Images.Add("Child", My.Resources.IcoNew)
            imageList1.Images.Add("AON12", My.Resources.AOMain)
            imageList1.Images.Add("AON13", My.Resources.AO_blue)
            imageList1.Images.Add("SVN", My.Resources.svn)
            imageList1.Images.Add("EXCEL", My.Resources.IcoExcel)
            imageList1.Images.Add("AON20", My.Resources.icon_AO_48)

            ShowIniList()
            Me.Activate()
            Me.Select()
            H_WHERE0.Select()
            アップロード抽出日.Value = DateTime.Today()
            C_WHERE03.Text = DateTime.Today().ToString("yyyyMM")
            If (DateTime.Today.Month > 8) Then
                T5_対象年月.Text = DateTime.Today().AddYears(-1).ToString("yyyy")
            Else
                T5_対象年月.Text = DateTime.Today().ToString("yyyy")
            End If
            T5_TAIDT.Text = "生産高確認"
            '----------------------

            Dim cmds As String() = System.Environment.GetCommandLineArgs()
            Select Case cmds.Length
                Case 2 '起動引数が1つの場合は負荷アサイン入力
                    BODY.SelectedIndex = 1
                    If IsNumeric(cmds(1)) Then
                        H_WHERE0.Text = cmds(1).ToString.TrimEnd
                    Else
                        H_WHERE2.Text = cmds(1).ToString.TrimEnd
                    End If
                    UpdateDataView()
                Case 3 '起動引数が3つの場合は初期表示タブを指定
                    If IsNumeric(cmds(1)) Then
                        BODY.SelectedIndex = Integer.Parse(cmds(1)) - 1
                        Select Case Integer.Parse(cmds(1))
                            Case 1
                                Filter.Text = cmds(2).ToString.TrimEnd
                                ShowIniList()
                            Case 2
                                If IsNumeric(cmds(1)) Then
                                    H_WHERE0.Text = cmds(1).ToString.TrimEnd
                                Else
                                    H_WHERE2.Text = cmds(1).ToString.TrimEnd
                                End If
                                UpdateDataView()
                        End Select
                    End If

            End Select
        Catch ex As Exception
            MessageBox.Show("起動に失敗しました。")
            'My.Settings.Reset()
        Finally
        End Try

    End Sub
#End Region

#Region "  各種処理"

    Private Sub FN_LoadStyle()
        Dim style As MetroFramework.MetroColorStyle
        Dim Col As Color

        If DirectCast(My.Settings("IconVisible"), Boolean) = True Then
            Me.Text = "    Assign Checker"
            HeaderIcon.Visible = True
        Else
            Me.Text = "Assign Checker"
            HeaderIcon.Visible = False
        End If
        Select Case My.Settings("DefaultStyle")
            Case "Silver"
                style = MetroFramework.MetroColorStyle.Silver
                Col = Color.FromArgb(255, 85, 85, 85)
                HeaderIcon.Image = My.Resources.icon_silver
            Case "Green"
                style = MetroFramework.MetroColorStyle.Green
                Col = Color.FromArgb(255, 0, 177, 89)
                HeaderIcon.Image = My.Resources.icon_green
            Case "Lime"
                style = MetroFramework.MetroColorStyle.Lime
                Col = Color.FromArgb(255, 175, 224, 61)
                HeaderIcon.Image = My.Resources.icon_lime
            Case "Orange"
                style = MetroFramework.MetroColorStyle.Orange
                Col = Color.FromArgb(255, 243, 119, 53)
                HeaderIcon.Image = My.Resources.icon_orange
            Case "Purple"
                style = MetroFramework.MetroColorStyle.Purple
                Col = Color.FromArgb(255, 124, 65, 153)
                HeaderIcon.Image = My.Resources.icon_purple
            Case "Blue", "Default"
                style = MetroFramework.MetroColorStyle.Default
                Col = Color.FromArgb(255, 0, 178, 219)
                HeaderIcon.Image = My.Resources.icon_blue
        End Select

        '各種スタイルを設定
        Me.Style = style

        Me.ValidateChildren()

        BODY.Style = style
        'MetroTabControl2.Style = style

        ボタン背景1.BackColor = Col
        ボタン背景2.BackColor = Col
        ボタン背景6.BackColor = Col
        ボタン背景7.BackColor = Col
        ボタン背景8.BackColor = Col
        ボタン背景9.BackColor = Col

        'テキストボックス一括設定
        For Each ctl In getControls(Me)
            'Metrotextbox
            If TypeOf (ctl) Is MetroFramework.Controls.MetroTextBox Then
                DirectCast(ctl, MetroFramework.Controls.MetroTextBox).Style = style
            End If

            'MetroButton
            If TypeOf (ctl) Is MetroFramework.Controls.MetroButton Then
                DirectCast(ctl, MetroFramework.Controls.MetroButton).Style = style
            End If

            'MetroCombobobx
            If TypeOf (ctl) Is MetroFramework.Controls.MetroComboBox Then
                DirectCast(ctl, MetroFramework.Controls.MetroComboBox).Style = style
            End If
        Next

        Me.Update()
    End Sub

    Private Function getControls(target As Control) As ArrayList
        Dim controls As ArrayList = New ArrayList
        For Each control As Control In target.Controls
            If TypeOf control Is MetroFramework.Controls.MetroTextBox Then controls.Add(control)
            If TypeOf control Is MetroFramework.Controls.MetroButton Then controls.Add(control)
            If TypeOf control Is MetroFramework.Controls.MetroComboBox Then controls.Add(control)
            If control.HasChildren Then controls.AddRange(getControls(control))
        Next
        Return controls
    End Function


    ''' <summary>
    ''' 
    '''     ''' フォームのサイズ変更
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FN_SetFormSize()
        Dim x As Integer
        Dim y As Integer
        If My.Settings("SizeKb") = "0" Then
            x = My.Settings("DefaultFormSizeX")
            y = My.Settings("DefaultFormSizeY")
        Else
            x = My.Settings("UserFormSizeX")
            y = My.Settings("UserFormSizeY")
        End If
        Me.Width = x
        Me.Height = y
        'BODY.Dock = DockStyle.Fill
    End Sub

    ''' <summary>
    ''' コントロールのDoubleBufferedプロパティをTrueにする
    ''' </summary>
    ''' <param name="control">対象のコントロール</param>
    Public Shared Sub EnableDoubleBuffering(control As Control)
        control.GetType().InvokeMember( _
        "DoubleBuffered", _
        BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, _
        Nothing, _
        control, _
        New Object() {True})
    End Sub


#End Region

#Region "  各種起動処理"
    ''' <summary>
    ''' 設定画面を開く
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenSettinfgform()
        Dim SubForm As New SettingForm()
        SubForm.ShowDialog()
        FN_LoadStyle()
        Me.Refresh()
    End Sub

    ''' <summary>
    ''' グリッド情報設定を開く
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenGridInfo()
        Dim SubForm As New GridInfo()
        SubForm.ShowDialog()
    End Sub

#End Region

#Region "  イベント "

#Region " [フォームイベント] "

    ''' <summary>
    ''' 画面終了処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FormClose()
        Me.Close()
        Application.Exit()
        MessageBox.Show("Close処理に失敗しました")
    End Sub

    Private Sub PrimaryForm_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If My.Settings("SizeKb") = "1" Then
            My.Settings("UserFormSizeX") = Me.Size.Width.ToString
            My.Settings("UserFormSizeY") = Me.Size.Height.ToString
        Else
        End If
    End Sub

    ' ホットキーの入力メッセージを処理する
    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_HOTKEY As Integer = &H312
        If m.Msg = WM_HOTKEY AndAlso m.LParam = hotkeyActivate.LParam Then
            ' フォームをアクティブにする
            If Me.WindowState = FormWindowState.Normal Then
                Me.WindowState = FormWindowState.Minimized
            Else
                Me.WindowState = FormWindowState.Normal
                Me.Activate()
            End If
        Else
            MyBase.WndProc(m)
        End If
    End Sub
#End Region

#Region " [ヘッダーToolStrip]"

    ' ''' <summary>
    ' ''' グリッド情報設定画面
    ' ''' </summary>
    ' ''' <param name="sender"></param>
    ' ''' <param name="e"></param>
    ' ''' <remarks></remarks>
    'Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles toolstrip
    '    Dim SubForm As New GridInfo()
    '    SubForm.Show()
    'End Sub

    ''' <summary>
    ''' 工数内訳画面を開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Sub1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 内訳工数ToolStripMenuItem.Click
        OpenSubformKousu()
    End Sub
    Private Sub プログラム一覧を表示ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles プログラム一覧を表示ToolStripMenuItem.Click
        OpenSubformPGList()
    End Sub

    ''' <summary>
    ''' 対応残画面を開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 内訳対応残ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 内訳対応残ToolStripMenuItem.Click
        OpenSubformTaiou()
    End Sub

#End Region

#Region " [KeyDownEvent]"

    ''' <summary>
    ''' DataGridView,H_Setting
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub D_DataGridView_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'BODY.KeyDown, MyBase.KeyDown
        '■Ctrl
        If e.Control Then
            Select Case e.KeyCode
                Case Keys.F
                    Select Case BODY.SelectedTab.Name
                        Case AssignTab.Name
                            H_WHERE2.Focus()
                        Case PgTab.Name
                            P_WHERE04.Focus()
                        Case SproTab.Name
                            C_WHERE02.Focus()
                    End Select
                Case Keys.D1
                    BODY.SelectTab(0)
                Case Keys.D2
                    BODY.SelectTab(1)
                Case Keys.D3
                    BODY.SelectTab(2)
                Case Keys.D4
                    BODY.SelectTab(3)
                Case Keys.D5
                    BODY.SelectTab(4)
                Case Keys.D6
                    BODY.SelectTab(5)
                Case Keys.D7
                    BODY.SelectTab(6)
            End Select
        End If

        '■Alt
        If e.Alt Then
            Select Case e.KeyCode
                Case Keys.F4
                    If My.Settings("EscClose").ToString.TrimEnd = "1" Then
                        Me.Close()
                    End If
            End Select
        End If

        'タブごとの処理
        Select Case BODY.SelectedTab.Name
            Case "SelectorTab"
                '【Ctrl+S】上書き保存
                If e.Control AndAlso e.KeyCode = Keys.S Then
                    SaveIni()
                End If

                If e.Control AndAlso e.KeyCode = Keys.F Then
                    Filter.Focus()
                End If

                If e.Control AndAlso e.KeyCode = Keys.D Then
                    If CurrentIniinfo.ExecDir <> "" AndAlso System.IO.Directory.Exists(CurrentIniinfo.ExecDir) Then
                        If CurrentIniinfo.ExecDir.EndsWith("\") Then
                            System.Diagnostics.Process.Start(Path.GetDirectoryName(Path.GetDirectoryName(CurrentIniinfo.ExecDir)) & "\")
                        Else
                            System.Diagnostics.Process.Start(Path.GetDirectoryName(CurrentIniinfo.ExecDir) & "\")
                        End If
                    End If
                End If
        End Select

        '共通処理
        Select Case e.KeyCode
            Case Keys.F1
            Case Keys.F2
            Case Keys.F5
                Select Case BODY.SelectedTab.Name
                    Case AssignTab.Name
                        UpdateDataView()
                    Case SproTab.Name
                        UpdateCostDataView()
                    Case PgTab.Name
                        FindPGData()
                    Case FreeTab.Name
                        FindTancdData()
                End Select
            Case Keys.Escape
                If My.Settings("EscClose").ToString.TrimEnd = "1" Then
                    Me.Close()
                End If
        End Select

    End Sub


#End Region

#Region " [条件設定欄]"

    ''' <summary>
    ''' タブ切り替え時の初期フォーカス位置設定
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub H_Setting_TabIndexChanged(sender As Object, e As EventArgs) Handles BODY.SelectedIndexChanged
        Select Case BODY.SelectedTab.Name
            Case "SelectorTab"
                If CurrentIniinfo.ExecDir = "" Then
                    Filter.Focus()
                    Filter.Select()
                End If
            Case "AssignTab"
                H_WHERE2.Select()
            Case "PgTab"
                P_WHERE04.Select()
            Case "CostTab"
                C_WHERE03.Select()
        End Select
    End Sub

#End Region


#End Region


    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        OpenSettinfgform()
    End Sub

    Private Sub MetroButton3_Click(sender As Object, e As EventArgs) Handles MetroButton3.Click
        OpenGridInfo()
    End Sub

    Private Sub MetroButton5_Click(sender As Object, e As EventArgs) Handles MetroButton8.Click
        UpdateDataView()
    End Sub

    Private Sub MetroButton10_Click(sender As Object, e As EventArgs) Handles MetroButton10.Click
        UpdateCostDataView()
    End Sub

    Private Sub BODY_SelectedIndexChanged(sender As Object, e As EventArgs) Handles BODY.SelectedIndexChanged
        StatusStrip1.Text = String.Empty
    End Sub

    Private Sub MetroButton5_Click_1(sender As Object, e As EventArgs) Handles MetroButton5.Click
        FindPGData()
    End Sub

    Private Sub MetroButton11_Click(sender As Object, e As EventArgs) Handles MetroButton11.Click
        FindTancdData()
    End Sub


End Class


''' <summary>ホットキーの登録・解除を行うためのクラス</summary>
Class HotKey
    <DllImport("user32", SetLastError:=True)> _
    Private Shared Function RegisterHotKey(ByVal hWnd As IntPtr, _
                                         ByVal id As Integer, _
                                         ByVal fsModifier As Integer, _
                                         ByVal vk As Integer) As Integer
    End Function

    <DllImport("user32", SetLastError:=True)> _
    Private Shared Function UnregisterHotKey(ByVal hWnd As IntPtr, _
                                           ByVal id As Integer) As Integer
    End Function

    Public Sub New(ByVal hWnd As IntPtr, ByVal id As Integer, ByVal key As Keys)
        Me.hWnd = hWnd
        Me.id = id

        ' Keys列挙体の値をWin32仮想キーコードと修飾キーに分離
        Dim keycode As Integer = CInt(key And Keys.KeyCode)
        Dim modifiers As Integer = CInt(key And Keys.Modifiers) >> 16

        Me._lParam = New IntPtr(modifiers Or keycode << 16)

        If RegisterHotKey(hWnd, id, modifiers, keycode) = 0 Then
            ' ホットキーの登録に失敗
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If
    End Sub

    Public Sub Unregister()
        If hWnd = IntPtr.Zero Then Return

        If UnregisterHotKey(hWnd, id) = 0 Then
            ' ホットキーの解除に失敗
            Throw New Win32Exception(Marshal.GetLastWin32Error())
        End If

        hWnd = IntPtr.Zero
    End Sub

    Public ReadOnly Property LParam As IntPtr
        Get
            Return _lParam
        End Get
    End Property

    Private hWnd As IntPtr ' ホットキーの入力メッセージを受信するウィンドウのhWnd
    Private ReadOnly id As Integer ' ホットキーのID(0x0000〜0xBFFF)
    Private ReadOnly _lParam As IntPtr ' WndProcメソッドで押下されたホットキーを識別するためのlParam値
End Class