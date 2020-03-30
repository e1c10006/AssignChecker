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


Partial Class PrimaryForm

    Private Sub ExecOpen_Click(sender As Object, e As EventArgs) Handles ExecOpen.Click
        アップロードファイル.Clear()
        If Not System.IO.Directory.Exists(LocalExecPath.Text.TrimEnd) AndAlso Not System.IO.File.Exists(ServerExecPath.Text.TrimEnd) Then
            MetroFramework.MetroMessageBox.Show(Me, "フォルダが存在しません", "エラー")
            Return
        End If

        Dim files As String() = System.IO.Directory.GetFiles(LocalExecPath.Text, "*", System.IO.SearchOption.TopDirectoryOnly)
        For Each filepath In files
            If filepath.Contains("vshost") Then
                Continue For
            End If
            Dim ext As String = Path.GetExtension(filepath)
            If (ext = ".exe" AndAlso CHK_EXE.Checked) OrElse _
               (ext = ".dll" AndAlso CHK_DLL.Checked) OrElse
               (ext = ".xml" AndAlso CHK_XML.Checked) OrElse
               (ext = ".rpx" AndAlso CHK_RPX.Checked) OrElse
               (ext = ".sql" AndAlso CHK_SQL.Checked) Then

                If File.GetLastWriteTime(filepath).Date >= アップロード抽出日.Value Then
                    '前1つきのみを抽出する場合
                    If ext = ".exe" AndAlso CHK_前1.Checked Then
                        If Not (Path.GetFileName(filepath).Length >= 8 AndAlso Path.GetFileName(filepath).Substring(4, 1) = "1") Then
                            Continue For
                        End If
                    End If

                    アップロードファイル.Text &= filepath & vbCrLf
                End If
            End If
        Next
    End Sub

    Private Sub アップロードファイル_DragEnter(sender As Object, e As DragEventArgs) Handles アップロードファイル.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        End If


    End Sub

    Private Sub アップロードファイル_DragDrop(sender As Object, e As DragEventArgs) Handles アップロードファイル.DragDrop
        アップロードファイル.Clear()
        For Each filenames As String In e.Data.GetData(DataFormats.FileDrop)
            アップロードファイル.Text &= filenames & vbCrLf
        Next
    End Sub



    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        If MetroFramework.MetroMessageBox.Show(Me, "ファイルをアップロードしますか？", "確認", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
            Return
        End If

        Dim ReplaceCnt As Integer = 0
        Dim CopyCnt As Integer = 0

        If Not System.IO.Directory.Exists(ServerExecPath.Text) Then
            MetroFramework.MetroMessageBox.Show(Me, "アップロード先のフォルダが見つかりません。", "エラー")
            Return
        End If

        'ファイル一覧を取得
        'Dim UploadFilepaths As String() = System.IO.Directory.GetFiles(LocalExecPath.Text, "*", System.IO.SearchOption.TopDirectoryOnly)
        Dim ServerFilepaths As String() = System.IO.Directory.GetFiles(ServerExecPath.Text, "*", System.IO.SearchOption.TopDirectoryOnly)

        Dim UploadFiles As List(Of String) = New List(Of String)
        Dim ServerFiles As List(Of String) = New List(Of String)

        'ファイル名に変換
        Dim rs As New System.IO.StringReader(アップロードファイル.Text)
        While rs.Peek() > -1
            Dim fileName As String = rs.ReadLine()
            If System.IO.File.Exists(fileName) Then
                Dim str As String = ""
                UploadFiles.Add(Path.GetFileName(fileName))
            End If
        End While


        For Each filename As String In ServerFilepaths
            ServerFiles.Add(System.IO.Path.GetFileName(filename))
        Next

        'ファイル移動処理
        For Each filename As String In UploadFiles

            'リネーム処理(上書き時は不要)
            If Not CHK_上書き.Checked Then
                If ServerFiles.Contains(filename) Then
                    Dim RenameName As String = Path.Combine(ServerExecPath.Text, Path.GetFileNameWithoutExtension(filename) & System.DateTime.Now.ToString("_yyyyMMdd"))
                    Dim cnt As Integer = 1
                    If File.Exists(RenameName & Path.GetExtension(filename)) Then
                        While (True)
                            If Not File.Exists(RenameName & "_" & cnt.ToString & Path.GetExtension(filename)) Then
                                File.Move(Path.Combine(ServerExecPath.Text, filename), RenameName & "_" & cnt.ToString & Path.GetExtension(filename))
                                Exit While
                            End If
                            cnt += 1
                        End While
                    Else
                        File.Move(Path.Combine(ServerExecPath.Text, filename), RenameName & Path.GetExtension(filename))
                    End If
                    ReplaceCnt += 1
                End If
            End If

            Try
                File.Copy(Path.Combine(LocalExecPath.Text, filename), Path.Combine(ServerExecPath.Text, filename))
            Catch ex As Exception

            End Try
            CopyCnt += 1
        Next

        MetroFramework.MetroMessageBox.Show(Me, "対象のファイルをアップロードしました。" & CopyCnt.ToString & "件（更新 " & ReplaceCnt.ToString() & "）", "完了")
    End Sub

    Private Sub CHK_EXE_CheckedChanged(sender As Object, e As EventArgs) Handles CHK_EXE.CheckedChanged
        CHK_前1.Enabled = CHK_EXE.Checked
        CHK_前1.Checked = False
    End Sub
End Class
