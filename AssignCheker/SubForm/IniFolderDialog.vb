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
Imports MetroFramework.Forms

Partial Public Class IniFolderDialog

    Public Sub IniFolderDialog()
        InitializeComponent()
    End Sub

    Private Sub metroButton1_Click(sender As Object, e As EventArgs) Handles metroButton1.Click
        'FolderBrowserDialogクラスのインスタンスを作成
        Dim fbd As FolderBrowserDialog = New FolderBrowserDialog()

        '上部に表示する説明テキストを指定する
        fbd.Description = "フォルダを指定してください。"

        'ルートフォルダを指定する
        'デフォルトでDesktop
        fbd.RootFolder = Environment.SpecialFolder.Desktop

        '最初に選択するフォルダを指定する
        'RootFolder以下にあるフォルダである必要がある
        fbd.SelectedPath = "C:\"
        'ユーザーが新しいフォルダを作成できるようにする
        'デフォルトでTrue
        fbd.ShowNewFolderButton = True

        'ダイアログを表示する
        If fbd.ShowDialog(Me) = DialogResult.OK Then
            ''選択されたフォルダを表示する
            Folderpath.Text = fbd.SelectedPath
        End If
    End Sub

    Private Sub metroButton2_Click(sender As Object, e As EventArgs) Handles metroButton2.Click
        My.Settings("IniFolder") = Folderpath.Text
        My.Settings.Save()
        Me.Close()
    End Sub
End Class

