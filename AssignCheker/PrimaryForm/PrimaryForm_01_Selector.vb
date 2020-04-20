Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms
Imports System.Text
Imports System.Reflection
Imports System.Threading
Imports ILL.ERF.BaseLib
Imports ILL.ERF.G1Base
Imports ILL.ERF.G2Base
Imports ILL.ERF.AKBS1010
Imports System.IO.StringReader
Imports System.Text.RegularExpressions
Imports System
Imports System.Security.Permissions
Imports System.Collections
Imports System.ComponentModel
Imports System.Media
Imports System.Runtime.InteropServices
Imports System.Collections.Specialized

Partial Class PrimaryForm

#Region "イベント"
    ''' <summary>
    ''' Inilistクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub IniView_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles IniView.AfterSelect
        Me.OpenIniFile()
        Me.ShowPRGList()
    End Sub

    ''' <summary>  FILTER キー押下イベント  </summary>
    Private Sub Filter_KeyDown(sender As Object, e As KeyEventArgs) Handles Filter.KeyDown
        ''【F1】キーが押されたか調べる
        If e.KeyData = Keys.Enter Then
            Me.ShowIniList()
        End If

        If e.KeyData = Keys.Down Then
            IniView.Select()
            IniView.Focus()
        End If

    End Sub

    Private Sub IniText_KeyDown(sender As Object, e As KeyEventArgs) Handles IniText.KeyDown
        If e.Control AndAlso e.KeyCode = Keys.A Then
            IniText.SelectAll()
        End If

    End Sub

    ''' <summary>  ダブルクリック</summary>
    Private Sub IniView_DoubleClick(sender As Object, e As EventArgs) Handles IniView.DoubleClick
        Me.Exec()
    End Sub

    Private Sub IniView_MouseDown(sender As Object, e As MouseEventArgs) Handles IniView.MouseDown
        '' 右クリックでもノードを選択させる
        If e.Button = Windows.Forms.MouseButtons.Right Then
            IniView.SelectedNode = IniView.GetNodeAt(e.X, e.Y)
        End If
    End Sub

    ''' <summary>
    ''' Saveボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles FNC01.Click
        SaveIni()
    End Sub

    ''' <summary>
    ''' タブ変更
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    'Private Sub PrgView_TabIndexChanged(sender As Object, e As EventArgs) Handles MetroTabControl2.SelectedIndexChanged
    'If MetroTabControl2.SelectedTab.Text = "Execフォルダ" Then
    '    ShowPRGList()
    'End If
    'End Sub

    ''' <summary>
    ''' KeyDown
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub iniView_KeyDown(sender As Object, e As KeyEventArgs) Handles IniView.KeyDown
        Select Case e.KeyCode
            Case Keys.Enter
                Me.Exec()
            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down
            Case Keys.F2
                NameEdit = True
                IniView.SelectedNode.BeginEdit()
                NameEdit = False
            Case Keys.F5

            Case Else
                If Filter.Focused = False Then
                    Filter.Focus()
                End If
        End Select

        '検索にフォーカスを当てる
        If e.KeyData = Keys.Up Then
            If IniView.SelectedNode.Index = 0 AndAlso IniView.SelectedNode.Parent Is Nothing Then
                Filter.Focus()
            End If
        End If


        'Me.SuspendLayout()
        '子を表示する
        If e.KeyData = Keys.Space Then
            If IniView.SelectedNode.Parent Is Nothing Then
                If IniView.SelectedNode.IsExpanded Then
                    IniView.SelectedNode.Toggle()
                Else
                    IniView.SelectedNode.Expand()
                End If
            Else
                IniView.SelectedNode.Parent.Toggle()
            End If
        End If
        'Me.ResumeLayout()
    End Sub

    Private Sub 名前を変更ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 名前を変更ToolStripMenuItem.Click
        NameEdit = True
        IniView.SelectedNode.BeginEdit()
        NameEdit = False
    End Sub

    Private Sub IniView_BeforeLabelEdit(sender As Object, e As NodeLabelEditEventArgs) Handles IniView.BeforeLabelEdit
        If IniView.SelectedNode.Parent Is Nothing Then
            e.CancelEdit = True
        End If
        If NameEdit = False Then
            e.CancelEdit = True
        End If
    End Sub

    Private Sub ListView_AfterLabelEdit(sender As Object, e As NodeLabelEditEventArgs) Handles IniView.AfterLabelEdit
        If e.Label IsNot Nothing Then
            Dim lv As TreeView = DirectCast(sender, TreeView)

            '同名アイテムチェック
            For Each lvi As TreeNode In lv.Nodes
                If lvi.Index <> e.Node.Index AndAlso lvi.Text = e.Label Then
                    MetroFramework.MetroMessageBox.Show(Me, "既に同じ名前のIniが存在しています。", "エラー")
                    e.CancelEdit = True
                    Return
                End If
            Next

            'ファイル名変更
            Dim fi As System.IO.FileInfo = New System.IO.FileInfo(CurrentIniinfo.FilePath)
            fi.MoveTo(Path.Combine(System.IO.Path.GetDirectoryName(CurrentIniinfo.FilePath), e.Label & ".ini"))

            Me.GetIniList()
            Me.ShowIniList()
        End If
    End Sub

    Private Sub INIを複製ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles INIを複製ToolStripMenuItem.Click
        CopyIni()
    End Sub

    Private Sub INIを削除ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles INIを削除ToolStripMenuItem.Click
        DeleteIni()
    End Sub

    Private Sub PG反映元にセットToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PG反映元にセットToolStripMenuItem.Click
        LocalExecPath.Text = CurrentIniinfo.ExecDir
    End Sub

    Private Sub PG反映先にセット_Click(sender As Object, e As EventArgs) Handles PG反映先にセットToolStripMenuItem.Click
        ServerExecPath.Text = CurrentIniinfo.ExecDir
    End Sub


    Private Sub InilistToolstrip_Opening(sender As Object, e As CancelEventArgs) Handles InilistToolstrip.Opening
        画面グループファイルを自動登録ToolStripMenuItem.Visible = CurrentIniinfo.Version.StartsWith("2")
    End Sub

    Private Sub セキュリティマスタを自動登録ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles セキュリティマスタを自動登録ToolStripMenuItem.Click

        '常にDBに接続済の時は接続を切る 
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        Using cn As SqlConnection = New SqlConnection()

            '確認ダイアログ表示
            If Not MetroFramework.MetroMessageBox.Show(Me, "メニューマスタのみ登録されているメニューをセキュリティマスタへ反映します(セキュリティグループ:9999)" & vbCrLf _
                                                         & "よろしいですか？ " & vbCrLf _
                                                         & "Server = " & CurrentIniinfo.Server & " " & vbCrLf _
                                                         & "Database = " & CurrentIniinfo.Database, "確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                Return
            End If


            'データベースを選択
            Try
                cn.ConnectionString = "Data Source=" & CurrentIniinfo.Server & ";" _
                                & "Trusted_Connection = Yes;" _
                                & "Initial Catalog=" & CurrentIniinfo.Database & ";"
                cn.Open()
                Dim queryString As String = "INSERT INTO セキュリティマスタ ("
                queryString &= "                FORMCHK "
                queryString &= "	,セキュリティグループ "
                queryString &= "	,使用区分 "
                queryString &= "	,登録区分 "
                queryString &= "	,削除区分 "
                queryString &= "	,印刷区分 "
                queryString &= "	,更新者 "
                queryString &= "	,登録者 "
                queryString &= "	) ( "
                queryString &= "	SELECT DISTINCT a.FORMCHK "
                queryString &= "	,'9999' "
                queryString &= "	,'1' "
                queryString &= "	,'1' "
                queryString &= "	,'1' "
                queryString &= "	,'1' "
                queryString &= "	,'9999' "
                queryString &= "	,'9999' FROM メニューマスタ a LEFT OUTER JOIN セキュリティマスタ b ON a.FORMCHK = b.FORMCHK WHERE b.FORMCHK IS NULL AND a.EXE名 NOT LIKE 'MENU%'"
                queryString &= "	)"
                Dim command As SqlCommand = New SqlCommand(queryString, cn)
                Dim ret As Integer = command.ExecuteNonQuery()
                ret = ret
                MetroFramework.MetroMessageBox.Show(Me, ret.ToString & "件更新しました。", "確認")
            Catch ex As Exception
                ex = ex
            End Try
        End Using
    End Sub


    Private Sub 画面グループファイルを自動登録ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 画面グループファイルを自動登録ToolStripMenuItem.Click
        '常にDBに接続済の時は接続を切る 
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        Using cn As SqlConnection = New SqlConnection()

            '確認ダイアログ表示
            If Not MetroFramework.MetroMessageBox.Show(Me, "作成済みのデザイナから未整備のグループ名ファイルを一括修正します。" & vbCrLf _
                                                         & "よろしいですか？ " & vbCrLf _
                                                         & "Server = " & CurrentIniinfo.Server & " " & vbCrLf _
                                                         & "Database = " & CurrentIniinfo.Database, "確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

                Return
            End If


            'データベースを選択
            Try
                cn.ConnectionString = "Data Source=" & CurrentIniinfo.Server & ";" _
                                & "Trusted_Connection = Yes;" _
                                & "Initial Catalog=" & CurrentIniinfo.Database & ";"
                cn.Open()
                Dim queryString As String = " "
                queryString &= " UPDATE ADF"
                queryString &= "    SET システムID = SUB.システムID"
                queryString &= "   FROM AD画面グループファイル ADF"
                queryString &= "   INNER JOIN"
                queryString &= "   (SELECT システムID=SUBSTRING(システムID,1,8),システム名"
                queryString &= "      FROM AD画面構成ファイル"
                queryString &= "     WHERE AD画面構成区分 = '10' "
                queryString &= " 	  AND 作成区分 = '0' "
                queryString &= " 	  AND SUBSTRING(システムID,5,1) = '1'"
                queryString &= " 	  AND SUBSTRING(システムID,5,1) NOT LIKE '%[^0123456789]%'"
                queryString &= " 	  AND LEN(システムID) = 8 --念の為"
                queryString &= "  ) SUB ON (SUBSTRING(ADF.システムID,1,4) + '1' + SUBSTRING(ADF.システムID,6,3)) = SUB.システムID"
                queryString &= " "
                queryString &= " UPDATE ADF"
                queryString &= "    SET グループID = (SUBSTRING(ADF.グループID,1,4) + '1' + SUBSTRING(ADF.グループID,6,3))"
                queryString &= "   FROM AD画面グループファイル ADF "
                queryString &= "  WHERE SUBSTRING(システムID,5,1) = '1' "
                queryString &= "    AND SUBSTRING(システムID,5,1) NOT LIKE '%[^0123456789]%'"
                queryString &= "    AND SUBSTRING(システムID,8,1) = '0'"
                queryString &= "    AND LEN(システムID) = 8 --念の為"
                queryString &= " "
                queryString &= " UPDATE ADF"
                queryString &= "    SET グループID = SUB.グループID"
                queryString &= "   FROM AD画面グループ名ファイル ADF"
                queryString &= "        INNER JOIN AD画面グループファイル SUB ON (SUBSTRING(ADF.グループID,1,4) + '1' + SUBSTRING(ADF.グループID,6,3)) = SUB.グループID AND LEN(SUB.グループID) = 8"
                queryString &= " "
                queryString &= " UPDATE ADF2"
                queryString &= "    SET グループID = ADF.グループID"
                queryString &= "   FROM AD画面グループファイル ADF"
                queryString &= "        LEFT OUTER JOIN AD画面グループファイル ADF2"
                queryString &= "     ON SUBSTRING(ADF.グループID,1,4) = SUBSTRING(ADF2.グループID,1,4)"
                queryString &= "    AND SUBSTRING(ADF.グループID,6,3) = SUBSTRING(ADF2.グループID,6,3)"
                queryString &= "  WHERE SUBSTRING(ADF.グループID,5,1) = '1' "
                queryString &= "    AND SUBSTRING(ADF.グループID,5,1) NOT LIKE '%[^0123456789]%'"
                queryString &= "    AND SUBSTRING(ADF.グループID,8,1) = '0'"
                queryString &= "    AND SUBSTRING(ADF2.グループID,5,1) = '0'"
                queryString &= "    AND SUBSTRING(ADF2.グループID,5,1) NOT LIKE '%[^0123456789]%'"
                queryString &= "    AND LEN(ADF.グループID) = 8"
                Dim command As SqlCommand = New SqlCommand(queryString, cn)
                Dim ret As Integer = command.ExecuteNonQuery()
                ret = ret
                MetroFramework.MetroMessageBox.Show(Me, ret.ToString & "件更新しました。", "確認")
            Catch ex As Exception
                ex = ex
            End Try
        End Using
    End Sub

    Private Sub フォルダを指定してSQLを反映ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles フォルダを指定してSQLを反映ToolStripMenuItem.Click
        'Processオブジェクト
        Dim proc As New System.Diagnostics.Process()
        Dim FolderPath As String

        'フォルダ選択
        Using ofd = New OpenFileDialog()
            ofd.FileName = "SelectFolder"
            ofd.Filter = "Folder|."
            ofd.CheckFileExists = False
            ofd.InitialDirectory = Path.GetDirectoryName(Path.GetDirectoryName(CurrentIniinfo.ExecDir))
            If ofd.ShowDialog() = DialogResult.OK Then
                FolderPath = Path.GetDirectoryName(ofd.FileName)
            End If
        End Using

        '確認ダイアログ表示
        If Not MetroFramework.MetroMessageBox.Show(Me, "フォルダ内のsqlファイルを全て実行します。" & vbCrLf _
                                                     & "よろしいですか？ " & vbCrLf _
                                                     & "Server = " & CurrentIniinfo.Server & " " & vbCrLf _
                                                     & "Database = " & CurrentIniinfo.Database, "確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then

            Return
        End If

        With proc.StartInfo
            'ComSpecパスの取得
            .FileName = System.Environment.GetEnvironmentVariable("ComSpec")
            '出力読取り可能化
            .WorkingDirectory = FolderPath
            .UseShellExecute = False
            .RedirectStandardOutput = True
            .RedirectStandardInput = True
            'コマンドプロンプトのウィンドウ非表示
            '.CreateNoWindow = True
            'コマンド(/c:コマンド実行後、コマンドを実行したcmd.exeを終了)
            .Arguments = "/c for %f in (*.sql) do sqlcmd /S " & CurrentIniinfo.Server & " /d " & CurrentIniinfo.Database & " /E /i ""%f"" "
        End With

        'コマンド実行
        proc.Start()

        '実行結果取得
        Dim result As String = proc.StandardOutput.ReadToEnd()
        proc.WaitForExit()
        proc.Close()

        MessageBox.Show(Me, result)
        'MetroFramework.MetroMessageBox.Show(Me, result)

    End Sub

    Private Sub iniview_DragEnter(sender As Object, e As DragEventArgs) Handles IniView.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub iniview_DragDrop(sender As Object, e As DragEventArgs) Handles IniView.DragDrop
        Dim SourceFilename As String = e.Data.GetData(DataFormats.FileDrop)(0)
        Dim DestFilename As String = Path.Combine(My.Settings("IniFolder"), Path.GetFileName(e.Data.GetData(DataFormats.FileDrop)(0)))

        If File.Exists(SourceFilename) Then
            '' ファイルが存在しない場合のみコピー
            If Not File.Exists(DestFilename) Then
                File.Copy(SourceFilename, DestFilename)
            End If

            Filter.Text = Path.GetFileNameWithoutExtension(SourceFilename)
            Me.ShowIniList()
        End If
    End Sub
#End Region

#Region "共通メソッド"
    ''' <summary>  メニュー起動  </summary>
    Public Sub Exec()
        If CurrentIniinfo.ExecDir = "" OrElse Not System.IO.File.Exists(Path.Combine(CurrentIniinfo.ExecDir, "AONMENU.exe")) Then
            Return
        End If

        IniCopy(False)

        Try
            Dim Arg As String = ""
            If CurrentIniinfo.Version = "12" AndAlso CurrentIniinfo.Server <> "" AndAlso CurrentIniinfo.Database <> "" Then
                Arg = String.Format("/server {0} /db {1}", CurrentIniinfo.Server, CurrentIniinfo.Database)
            End If

            If CurrentIniinfo.Version = "13" AndAlso CurrentIniinfo.Server <> "" AndAlso CurrentIniinfo.Database <> "" Then
                Arg = String.Format("/server {0} /db {1}", CurrentIniinfo.Server, CurrentIniinfo.Database)
            End If

            If CurrentIniinfo.Version.StartsWith("2") AndAlso CurrentIniinfo.Server <> "" AndAlso CurrentIniinfo.Database <> "" Then
                Arg = String.Format("/ini {0}", CurrentIniinfo.FilePath)
            End If

            If Arg = "" Then
                Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(Path.Combine(CurrentIniinfo.ExecDir, "AONMENU.EXE"))
            Else
                Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(Path.Combine(CurrentIniinfo.ExecDir, "AONMENU.EXE"), Arg)
            End If
        Catch ex As Exception

        End Try
    End Sub
#End Region

#Region "Iniリスト"
    ''' <summary>  INIファイルの保存  </summary>
    Public Sub SaveIni()
        ' ノード選択時のみ
        If IniList.Keys.Contains(IniView.SelectedNode.Text) Then
            '存在するファイル
            Dim filepath As String = IniList(IniView.SelectedNode.Text).FilePath
            If Not Directory.Exists(My.Settings("IniFolder")) Then
                Return
            End If

            Dim text As String = IniText.Text
            'INIを反映

            Using sw As StreamWriter = New StreamWriter(filepath, False, Encoding.GetEncoding("Shift_JIS"))
                sw.Write(text)
                sw.Close()
                MetroFramework.MetroMessageBox.Show(Me, "iniファイルを保存しました。", "確認")
            End Using
            ReadIniFile(filepath, CurrentIniinfo.FileName)
            OpenIniFile()
            'Me.GetIniList()
            'Me.ShowIniList()
        End If
    End Sub

    ''' <summary>  INIリストの取得  </summary>
    Public Sub GetIniList()
        IniList.Clear()
        Dim files As String() = System.IO.Directory.GetFiles(My.Settings("IniFolder"), "*", System.IO.SearchOption.AllDirectories)
        For Each filepath In files
            ReadIniFile(filepath, "")
        Next
    End Sub

    ''' <summary>  iniファイルを開きます  </summary>
    Public Sub OpenIniFile()
        Dim str As String = ""
        'Me.SuspendLayout()
        'MetroPanel1.SuspendLayout()

        If IniView.SelectedNode Is Nothing OrElse Not IniList.ContainsKey(IniView.SelectedNode.Text) Then
            IniText.Text = ""
            CurrentIniinfo.Server = ""
            CurrentIniinfo.Database = ""
            CurrentIniinfo.ExecDir = ""
            CurrentIniinfo.FileName = ""
            CurrentIniinfo.FilePath = ""
            CurrentIniinfo.Version = ""
            CurrentIniinfo.visible = True
            Return
        End If

        Using reader As StreamReader = New StreamReader(IniList(IniView.SelectedNode.Text).FilePath, Encoding.GetEncoding("Shift_JIS"))
            IniText.Text = reader.ReadToEnd()
            CurrentIniinfo.Server = IniList(IniView.SelectedNode.Text).Server
            CurrentIniinfo.Database = IniList(IniView.SelectedNode.Text).Database
            CurrentIniinfo.ExecDir = IniList(IniView.SelectedNode.Text).ExecDir
            If Not CurrentIniinfo.ExecDir.EndsWith("\") Then
                CurrentIniinfo.ExecDir &= "\"
            End If
            CurrentIniinfo.FileName = IniList(IniView.SelectedNode.Text).FileName
            CurrentIniinfo.FilePath = IniList(IniView.SelectedNode.Text).FilePath
            CurrentIniinfo.Version = IniList(IniView.SelectedNode.Text).Version
            CurrentIniinfo.visible = IniList(IniView.SelectedNode.Text).visible
        End Using

        For i As Integer = 1 To 20
            If FunctionsPanel.Controls.ContainsKey("FNC" & i.ToString.PadLeft(2, "0"c)) Then
                With FunctionsPanel.Controls("FNC" & i.ToString.PadLeft(2, "0"c))
                    .Text = ""
                    .Enabled = False
                End With
            End If
        Next

        FNC01.Enabled = True
        FNC01.Text = "上書き保存"""
        If CurrentIniinfo.Version.StartsWith("2") Then
            FNC01.Enabled = True
            FNC02.Enabled = True
            FNC03.Enabled = True
            FNC04.Enabled = True
            FNC05.Enabled = True
            FNC06.Enabled = True
            FNC07.Enabled = True
            FNC08.Enabled = True
            FNC09.Enabled = True
            FNC10.Enabled = True
            FNC11.Enabled = True
            FNC12.Enabled = True

            FNC02.Text = "IniCopy"
            FNC03.Text = "Debug" & vbCrLf & "Executer"
            FNC04.Text = "帳票D"
            FNC05.Text = "マスタD"
            FNC06.Text = "検索D"
            FNC07.Text = "伝発D"
            FNC08.Text = "伝票D"
            FNC09.Text = "ベース"
            FNC10.Text = "連携"
            FNC11.Text = "DBD"
            FNC12.Text = "保存先" & vbCrLf & "登録"

        ElseIf CurrentIniinfo.Version.StartsWith("13") Then
            FNC01.Enabled = True
            FNC02.Enabled = True
            FNC03.Enabled = True


            FNC02.Text = "IniCopy"
            FNC03.Text = "db.ini" & vbCrLf & "Copy"
        Else
            FNC01.Enabled = True
            FNC02.Enabled = True
            FNC02.Text = "IniCopy"
        End If
        'MetroPanel1.ResumeLayout()
        'Me.ResumeLayout()
    End Sub

    ''' <summary>  INIリストの取得  </summary>
    Public Sub IniCopy(ByVal ShowMsg As Boolean)
        '' ノード選択時のみ
        If IniView.SelectedNode IsNot Nothing AndAlso IniList.Keys.Contains(IniView.SelectedNode.Text) Then

            ' 存在するファイルのみ
            If (Not Directory.Exists(My.Settings("IniFolder"))) Then
                MetroFramework.MetroMessageBox.Show(Me, "iniファイルの取得に失敗しました。")
                Return
            End If

            Dim filePath As String = "C:\ProgramData\AONET.ini"
            Dim text As String = IniText.Text


            Dim rs As New System.IO.StringReader(text)
            Dim Resulttext As String = ""
            While rs.Peek() > -1
                Dim line As String = rs.ReadLine

                'Ver13の場合、ServerとDatabaseの情報はdb.iniに持つため不要
                If CurrentIniinfo.Version.StartsWith("13") Then
                    If line.StartsWith("Server") Then
                        Continue While
                    End If
                    If line.StartsWith("Database") Then
                        Continue While
                    End If
                End If
                Resulttext &= line & vbCrLf
            End While

            Using sw As StreamWriter = New StreamWriter(filePath, False, Encoding.GetEncoding("Shift_JIS"))
                sw.Write(Resulttext)
                sw.Close()
                If (ShowMsg) Then
                    MetroFramework.MetroMessageBox.Show(Me, "AONET.iniを更新しました。")
                End If
            End Using
        End If

    End Sub

    ''' <summary>  INIリストの取得  </summary>
    Public Sub DbIniCopy(ByVal ShowMsg As Boolean)
        '' ノード選択時のみ
        If IniView.SelectedNode IsNot Nothing AndAlso IniList.Keys.Contains(IniView.SelectedNode.Text) Then

            ' 存在するファイルのみ
            If (Not Directory.Exists(My.Settings("IniFolder"))) Then
                MetroFramework.MetroMessageBox.Show(Me, "iniファイルの取得に失敗しました。")
                Return
            End If

            'db.ini のパスを取得
            Dim dbini As String = Path.Combine(System.IO.Directory.GetParent(CurrentIniinfo.ExecDir).Parent.FullName, "db.ini")
            If Not File.Exists(dbini) Then
                MetroFramework.MetroMessageBox.Show(Me, "db.ini が存在しません", "エラー")
                Return
            End If

            'ini ファイルのみ読み込む
            Dim ResultText As String = ""
            Using reader As StreamReader = New StreamReader(dbini, Encoding.GetEncoding("Shift_JIS"))
                Dim str As String = ""
                While (True)
                    Dim line = reader.ReadLine()
                    If line Is Nothing Then
                        Exit While
                    End If

                    If ResultText <> "" Then ResultText &= vbCrLf
                    If line.StartsWith("Server") Then
                        line = "Server=" & CurrentIniinfo.Server
                    ElseIf line.StartsWith("Database") Then
                        line = "Database=" & CurrentIniinfo.Database
                    End If
                    ResultText &= line.ToString()
                End While
            End Using

            Try
                Using sw As StreamWriter = New StreamWriter(dbini, False, Encoding.GetEncoding("Shift_JIS"))
                    sw.Write(ResultText)
                    sw.Close()
                    MetroFramework.MetroMessageBox.Show(Me, "db.iniを保存しました。", "確認")
                End Using
            Catch ex As Exception
                MetroFramework.MetroMessageBox.Show(Me, ex.Message, "エラー")
            End Try
        End If

    End Sub

    ''' <summary>  INIファイルの削除 </summary>
    Public Sub DeleteIni()
        '' ノード選択時のみ
        If IniView.SelectedNode IsNot Nothing AndAlso IniList.Keys.Contains(IniView.SelectedNode.Text) Then
            If MetroFramework.MetroMessageBox.Show(Me, "INIファイルを削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                System.IO.File.Delete(CurrentIniinfo.FilePath)
                IniView.SelectedNode.Remove()
                Me.GetIniList()
                Me.ShowIniList()
            End If
        End If
    End Sub

    ''' <summary>  INIファイルの複製 </summary>
    Public Sub CopyIni()
        '' ノード選択時のみ
        If IniView.SelectedNode IsNot Nothing AndAlso IniList.Keys.Contains(IniView.SelectedNode.Text) Then

            Dim lv As TreeView = IniView
            Dim CopyNode As TreeNode = IniView.SelectedNode

            '同名アイテムチェック
            For Each lvi As TreeNode In lv.Nodes
                If lvi.Index <> CopyNode.Index AndAlso lvi.Text = CopyNode.Text.TrimEnd Then
                    MetroFramework.MetroMessageBox.Show(Me, "既に同じ名前のIniが存在しています。", "エラー")
                    Return
                End If
            Next

            'ファイル名変更
            Dim afterFileName As String = CopyNode.Text & "_コピー"
            Dim afterName As String = Path.Combine(System.IO.Path.GetDirectoryName(CurrentIniinfo.FilePath), afterFileName & ".ini")
            Dim fi As System.IO.FileInfo = New System.IO.FileInfo(CurrentIniinfo.FilePath)
            Try
                fi.CopyTo(afterName)
            Catch ex As Exception
                MetroFramework.MetroMessageBox.Show(Me, ex.Message)
            End Try
            ReadIniFile(afterName, afterFileName)
            Dim add As TreeNode = IniView.SelectedNode.Clone()
            add.Text = afterFileName
            IniView.SelectedNode.Parent.Nodes.Insert(IniView.SelectedNode.Index + 1, add)
        End If
    End Sub

    ''' <summary>
    ''' INIのファンクションクリック
    ''' </summary> 
    Private Sub FNC_Click(sender As Object, e As EventArgs) Handles FNC01.Click, FNC02.Click, FNC03.Click, FNC04.Click, FNC05.Click, FNC06.Click, FNC07.Click, FNC08.Click, FNC09.Click, FNC10.Click, FNC11.Click, FNC12.Click, FNC13.Click, FNC14.Click, FNC15.Click, FNC16.Click, FNC17.Click, FNC18.Click, FNC19.Click, FNC20.Click
        ExuecuteDbe(sender)
    End Sub

    Public Sub ExuecuteDbe(sender As Object)

        'iniを選択していなければ抜ける
        If CurrentIniinfo.ExecDir = "" Then
            Return
        End If

        '' ノード選択時のみ
        If IniView.SelectedNode IsNot Nothing AndAlso IniList.Keys.Contains(IniView.SelectedNode.Text) Then
            Dim name As String = ""
            Select Case CurrentIniinfo.Version
                Case "12"
                    Select Case DirectCast(sender, MetroFramework.Controls.MetroButton).Name.Substring(3, 2)
                        Case "02"
                            IniCopy(True)
                        Case Else
                    End Select

                Case "13"
                    Select Case DirectCast(sender, MetroFramework.Controls.MetroButton).Name.Substring(3, 2)
                        Case "02"
                            IniCopy(True)
                        Case "03"
                            DbIniCopy(True)
                        Case Else
                    End Select

                Case Else
                    Select Case DirectCast(sender, MetroFramework.Controls.MetroButton).Name.Substring(3, 2)
                        Case "02"
                            IniCopy(True)
                        Case "03"
                            Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(Path.Combine(CurrentIniinfo.ExecDir, "DebugExecutor.exe"), "/ini " & CurrentIniinfo.FilePath & " /seqoff /logoff")
                            Return
                        Case "04"
                            name = "ADSYS010"
                        Case "05"
                            name = "ADSYS020"
                        Case "06"
                            name = "ADSYS030"
                        Case "07"
                            name = "ADSYS040"
                        Case "08"
                            name = "ADSYS140"
                        Case "09"
                            name = "ADSYS110"
                        Case "10"
                            name = "ADSYS900"
                        Case "11"
                            name = "DataBaseDesigner"
                        Case "12"
                            If MetroFramework.MetroMessageBox.Show(Me, "INI情報から各デザイナ保存先を登録します。DataBase.xmlのパスを選択してください", "確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                                Dim insertquery As String = ""
                                Dim machinename As String = Environment.MachineName
                                Dim username As String = Environment.UserName
                                Dim rootpath As String = Path.GetDirectoryName(Path.GetDirectoryName(CurrentIniinfo.ExecDir))
                                insertquery &= "IF EXISTS (SELECT * FROM ユーザー別設定ファイル WHERE キー = 'REPORT')"
                                insertquery &= "    DELETE ユーザー別設定ファイル WHERE 端末名 = '" & machinename & "' AND Windowsアカウント = '" & username & "' AND キー = 'REPORT'"
                                insertquery &= "INSERT INTO ユーザー別設定ファイル VALUES('" & machinename & "','" & username & "','DESIGNER','REPORT','" & Path.Combine(rootpath, "Report") & "','" & username & "',GETDATE())"
                                insertquery &= "IF EXISTS (SELECT * FROM ユーザー別設定ファイル WHERE キー = 'XMLPATH')"
                                insertquery &= "    DELETE ユーザー別設定ファイル WHERE 端末名 = '" & machinename & "' AND Windowsアカウント = '" & username & "' AND キー = 'XMLPATH'"
                                insertquery &= "INSERT INTO ユーザー別設定ファイル VALUES('" & machinename & "','" & username & "','DESIGNER','XMLPATH','" & Path.Combine(rootpath, "Designer") & "','" & username & "',GETDATE())"
                                insertquery &= "IF EXISTS (SELECT * FROM ユーザー別設定ファイル WHERE キー = 'BASE')"
                                insertquery &= "    DELETE ユーザー別設定ファイル WHERE 端末名 = '" & machinename & "' AND Windowsアカウント = '" & username & "' AND キー = 'BASE'"
                                insertquery &= "INSERT INTO ユーザー別設定ファイル VALUES('" & machinename & "','" & username & "','DESIGNER','BASE','" & Path.Combine(rootpath, "Base") & "','" & username & "',GETDATE())"
                                insertquery &= "    DELETE ユーザー別設定ファイル WHERE 端末名 = '" & machinename & "' AND Windowsアカウント = '" & username & "' AND キー = 'MSTBASE'"
                                insertquery &= "INSERT INTO ユーザー別設定ファイル VALUES('" & machinename & "','" & username & "','DESIGNER','MSTBASE','" & Path.Combine(Path.Combine(rootpath, "MasterIO"), "Base") & "','" & username & "',GETDATE())"
                                insertquery &= "    DELETE ユーザー別設定ファイル WHERE 端末名 = '" & machinename & "' AND Windowsアカウント = '" & username & "' AND キー = 'MST'"
                                insertquery &= "INSERT INTO ユーザー別設定ファイル VALUES('" & machinename & "','" & username & "','DESIGNER','MST','" & Path.Combine(rootpath, "MasterIO") & "','" & username & "',GETDATE())"

                                'OpenFileDialogクラスのインスタンスを作成
                                Dim ofd As New OpenFileDialog()
                                'はじめに「ファイル名」で表示される文字列を指定する
                                ofd.FileName = "DataBase.xml"
                                'はじめに表示されるフォルダを指定する
                                '指定しない（空の文字列）の時は、現在のディレクトリが表示される
                                ofd.InitialDirectory = Path.Combine(Path.Combine(rootpath, "sql"), "環境作成")
                                '[ファイルの種類]に表示される選択肢を指定する
                                '指定しないとすべてのファイルが表示される
                                ofd.Filter = "DataBase.Xml(*.xml)|*.xml"
                                '[ファイルの種類]ではじめに選択されるものを指定する
                                ofd.Title = "開くファイルを選択してください"
                                'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
                                ofd.RestoreDirectory = True
                                'ダイアログを表示する
                                If ofd.ShowDialog() = DialogResult.OK Then
                                    'OKボタンがクリックされたとき、選択されたファイル名を表示する
                                    insertquery &= "    DELETE ユーザー別設定ファイル WHERE 端末名 = '" & machinename & "' AND Windowsアカウント = '" & username & "' AND キー = 'DBXMLFILE'"
                                    insertquery &= "INSERT INTO ユーザー別設定ファイル VALUES('" & machinename & "','" & username & "','DESIGNER','DBXMLFILE','" & ofd.FileName & "','" & username & "',GETDATE())"
                                End If


                                Try
                                    '常にDBに接続済の時は接続を切る 
                                    Selector_cmd.Dispose()
                                    Selector_cn.Close()
                                    Selector_cn.Dispose()
                                    Selector_cn.ConnectionString = "Data Source=" & CurrentIniinfo.Server & ";" _
                                                              & "Trusted_Connection = Yes;" _
                                                              & "Initial Catalog=" & CurrentIniinfo.Database & ";"
                                    Selector_cn.Open()
                                    Selector_cmd.CommandText = insertquery
                                    Selector_cmd.ExecuteNonQuery()
                                    Selector_cmd.Dispose()
                                    Selector_cn.Close()
                                    Selector_cn.Dispose()
                                Catch ex As Exception
                                End Try
                            End If


                            MetroFramework.MetroMessageBox.Show(Me, "登録完了しました。", "確認")
                            Return
                        Case "13"
                            Return
                    End Select

                    '' ノード選択時のみ
                    If IniView.SelectedNode IsNot Nothing AndAlso IniList.Keys.Contains(IniView.SelectedNode.Text) Then
                        Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(Path.Combine(CurrentIniinfo.ExecDir, "DebugExecutor.exe"), "/ini " & CurrentIniinfo.FilePath & " /seqoff /-v /Exe " & name)
                    End If
            End Select
        End If

    End Sub

    Public Sub ShowDebugExecuter()
        '' ノード選択時のみ
        If IniView.SelectedNode IsNot Nothing AndAlso IniList.Keys.Contains(IniView.SelectedNode.Text) Then

            If CurrentIniinfo.ExecDir = "" OrElse Not System.IO.File.Exists(Path.Combine(CurrentIniinfo.ExecDir, "AONMENU.exe")) Then
                Return
            End If

            Try
                If CurrentIniinfo.Version = "12" AndAlso CurrentIniinfo.Server <> "" AndAlso CurrentIniinfo.Database <> "" Then
                    Return
                End If

                If CurrentIniinfo.Version = "13" AndAlso CurrentIniinfo.Server <> "" AndAlso CurrentIniinfo.Database <> "" Then
                    Return
                End If

                Dim p As System.Diagnostics.Process = System.Diagnostics.Process.Start(Path.Combine(CurrentIniinfo.ExecDir, "DebugExecutor.exe"), "/ini " & CurrentIniinfo.FilePath & " /seqoff /logoff")
            Catch ex As Exception

            End Try
        End If
    End Sub

#End Region

#Region "PRGリスト"

    ''' <summary>  PRGフォルダを表示 </summary>
    Private Sub ShowPRGList()

        PRGList.Clear()
    End Sub

    ''' <summary>  販売フォルダからini一覧を取得します  </summary>
    Public Sub ShowPRGListMain()
        'PRGView.Nodes.Clear()
        'PRGView2.Items.Clear()
        For Each info As KeyValuePair(Of String, PRGInfo) In PRGList
            Dim item As ListViewItem = New ListViewItem()
            Select Case info.Value.FileType
                Case PRGInfo.DirectoryType.Folder
                    item.ImageKey = "FOLDER"
                Case PRGInfo.DirectoryType.SVN
                    item.ImageKey = "SVN"
                Case PRGInfo.DirectoryType.Excel
                    item.ImageKey = "EXCEL"
                Case PRGInfo.DirectoryType.File
                    item.ImageKey = "Child"
            End Select
            item.Text = info.Value.FileName
            item.SubItems.Add(System.IO.File.GetLastWriteTime(info.Value.FilePath).ToString("yyyy/MM/dd hh:mm:ss"))
        Next
    End Sub


#End Region

#Region "定義一覧"
    Public Class IniInfo
        Public FileName As String
        Public FilePath As String
        Public Version As String
        Public Database As String
        Public Server As String
        Public ExecDir As String
        Public visible As Boolean = True

        Public Sub New()
            FilePath = ""
            FileName = ""
            Version = ""
            ExecDir = ""
            Database = ""
            Server = ""
            visible = True
        End Sub

        Public Sub New(ByVal argFilepath As String, ByVal argFileName As String, ByVal argVersion As String, ByVal argServer As String, ByVal argDatabase As String, ByVal argExecDir As String)
            FilePath = argFilepath
            FileName = argFileName
            Version = argVersion
            Server = argServer
            Database = argDatabase
            ExecDir = argExecDir
            visible = True
        End Sub

    End Class

    Public Class PRGInfo
        Public FileName As String
        Public FilePath As String
        Public FileType As DirectoryType
        Public Enum DirectoryType
            Folder = 1
            Excel = 2
            SVN = 3
            File = 4
        End Enum

        Public Sub New()
            FilePath = ""
            FileName = ""
        End Sub

        Public Sub New(argFilepath As String, argFileName As String, argFileType As DirectoryType)
            FilePath = argFilepath
            FileName = argFileName
            FileType = argFileType
        End Sub

    End Class

    ''' <summary>  iniファイルの一覧</summary>
    Public IniList As Dictionary(Of String, IniInfo)

    ''' <summary>  iniファイルの一覧</summary>
    Public PRGList As Dictionary(Of String, PRGInfo)

    ''' <summary>  選択中のini情報</summary>
    Public CurrentIniinfo As IniInfo

    ''' <summary>  ini名編集中フラグ </summary>
    Public NameEdit As Boolean

#End Region

#Region "起動処理"
    ''' <summary>  画面起動時処理</summary>
    Private Sub RegistrFormIni()
        ''必要な変数の初期化
        CurrentIniinfo = New IniInfo()

        '"C:\販売"以下のファイルをすべて取得する
        'ワイルドカード"*"は、すべてのファイルを意味する
        If (Not System.IO.Directory.Exists(My.Settings("IniFolder"))) Then
            Dim iniDialog As IniFolderDialog = New IniFolderDialog()
            iniDialog.ShowDialog(Me)
            If (Not System.IO.Directory.Exists(My.Settings("IniFolder"))) Then
                MetroFramework.MetroMessageBox.Show(Me, "Iniフォルダの取得に失敗しました。", "エラー", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, 100)
                Application.Exit()
                Return
            End If
        End If

        IniList = New Dictionary(Of String, IniInfo)
        PRGList = New Dictionary(Of String, PRGInfo)

        'IniListの取得
        Me.GetIniList()

        '初期フォーカス位置設定
        Filter.Focus()
        Filter.Select()
    End Sub

    ''' <summary>  販売フォルダからini一覧を取得します  </summary>
    Public Sub ShowIniList()
        RegistrFormIni()
        IniView.Nodes.Clear()
        ' Ver
        '[VER12]
        Dim Ver12 As TreeNode = New TreeNode()
        Ver12.Text = "  AladdinOffice 1.2"
        Ver12.ImageKey = "AON12"
        Ver12.SelectedImageKey = Ver12.ImageKey
        IniView.Nodes.Add(Ver12)


        '[VER13]
        Dim Ver13 As TreeNode = New TreeNode()
        Ver13.Text = "  AladdinOffice 1.3"
        Ver13.ImageKey = "AON13"
        Ver13.SelectedImageKey = Ver13.ImageKey
        IniView.Nodes.Add(Ver13)

        '[VER20]
        Dim Ver20 As TreeNode = New TreeNode()
        Ver20.ImageKey = "AON20"
        Ver20.SelectedImageKey = Ver20.ImageKey
        Ver20.Text = "  AladdinOffice 2.X"
        IniView.Nodes.Add(Ver20)


        For Each Info As KeyValuePair(Of String, IniInfo) In IniList
            If Not Info.Key.Contains(Filter.Text.TrimEnd()) Then
                Continue For
            End If

            Dim Child As TreeNode = New TreeNode()
            Child.Text = Info.Value.FileName
            Child.ImageKey = "Child"
            Child.SelectedImageKey = "Child"
            Child.NodeFont = New Font("Yu Gothic UI", 9)


            Dim addflg As Boolean = False
            Select Case Info.Value.Version
                Case "12"
                    Ver12.Nodes.Add(Child)
                    addflg = True
                Case "13"
                    Ver13.Nodes.Add(Child)
                    addflg = True
            End Select
            'Ver20
            If Info.Value.Version.ToString.StartsWith("2") Then
                Ver20.Nodes.Add(Child)
                addflg = True
            End If

            If addflg = False Then

            End If
        Next

        If Filter.Text.TrimEnd.Length > 0 Then
            Ver12.ExpandAll()
            Ver13.ExpandAll()
            Ver20.ExpandAll()
        End If

        If (IniView.GetNodeCount(True) > 0) Then
            IniView.Select()
            'IniView.SelectedNode = IniView.Nodes(1)
            IniView.Focus()
        End If
    End Sub


    Public Sub ReadIniFile(filepath As String, key As String)
        Dim L_Version As String = ""
        Dim L_Server As String = ""
        Dim L_Database As String = ""
        Dim L_ExecDir As String = ""
        Dim L_FileName As String = ""

        'ini以外は読まない
        If Not filepath.Contains(".ini") AndAlso Not filepath.Contains(".INI") Then
            Return
        End If

        'ini ファイルのみ読み込む
        Using reader As StreamReader = New StreamReader(filepath, Encoding.GetEncoding("Shift_JIS"))
            Dim sections As Dictionary(Of String, Dictionary(Of String, String)) = New Dictionary(Of String, Dictionary(Of String, String))
            Dim regexSection As Regex = New Regex("^\s*\[(?<section>[^\]]+)\].*$", RegexOptions.Singleline Or RegexOptions.CultureInvariant)
            Dim regexNameValue = New Regex("^\s*(?<name>[^=]+)=(?<value>.*?)(\s+;(?<comment>.*))?$", RegexOptions.Singleline Or RegexOptions.CultureInvariant)
            Dim currentSection = String.Empty

            While (True)
                Dim line = reader.ReadLine()
                If line Is Nothing Then
                    Exit While
                End If

                '空行
                If line.Length = 0 Then
                    Continue While
                End If

                'コメント
                ' VER情報
                If line.StartsWith(";", StringComparison.Ordinal) Then
                    Continue While
                End If

                '読み込み
                Dim matchNameValue = regexNameValue.Match(line)
                If (matchNameValue.Success) Then
                    Select Case matchNameValue.Groups("name").Value.TrimEnd()
                        Case "Version"
                            Dim val As Decimal = 0
                            '1.2
                            If matchNameValue.Groups("value").Value.TrimEnd().StartsWith("1.2") Then
                                L_Version = "12"
                            ElseIf matchNameValue.Groups("value").Value.TrimEnd().StartsWith("1.3") Then
                                L_Version = "13"
                            ElseIf (Decimal.TryParse(matchNameValue.Groups("value").Value.TrimEnd(), val)) Then
                                L_Version = (Integer.Parse((val * 10).ToString(0))).ToString()
                            End If

                        Case "ExecDir"
                            L_ExecDir = matchNameValue.Groups("value").Value.TrimEnd()
                        Case "Server"
                            L_Server = matchNameValue.Groups("value").Value.TrimEnd()
                        Case "Database"
                            L_Database = matchNameValue.Groups("value").Value.TrimEnd()
                    End Select
                End If
                L_FileName = System.IO.Path.GetFileNameWithoutExtension(filepath)
            End While

            '既に存在する場合は上書き
            If IniList.ContainsKey(L_FileName) Then
                IniList.Remove(L_FileName)
                IniList.Add(L_FileName, New IniInfo(filepath, L_FileName, L_Version, L_Server, L_Database, L_ExecDir))
            Else
                IniList.Add(L_FileName, New IniInfo(filepath, L_FileName, L_Version, L_Server, L_Database, L_ExecDir))
            End If
        End Using
    End Sub
#End Region
End Class
