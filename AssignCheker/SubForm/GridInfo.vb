Imports System.Collections.Specialized
Public Class GridInfo

    Public cols As StringCollection = New StringCollection()


    Private Sub GridInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cols.Add("案件NO")
        cols.Add("受注NO")
        cols.Add("枝番")
        cols.Add("得意先コード")
        cols.Add("得意先名")
        cols.Add("エディション")
        cols.Add("物件工数")
        cols.Add("PG工数")
        cols.Add("開発障害")
        cols.Add("設計障害")
        cols.Add("QA件数")
        cols.Add("仕様書")
        cols.Add("PG完了")
        cols.Add("SEテスト")
        cols.Add("納品日")
        cols.Add("検収月")
        cols.Add("対応残")
        cols.Add("開発主管者名")
        cols.Add("設計主管者名")
        cols.Add("進捗表更新日")
        cols.Add("料金表パス")
        cols.Add("料金表を開く")
        cols.Add("進捗表パス")
        cols.Add("進捗表を開く")

        FN_LoadGridInfo()
        CheckText1.Select()
    End Sub

    Private Sub FN_LoadGridInfo()

        'テキストセット
        For i As Integer = 1 To cols.Count
            If GroupBox1.Controls.ContainsKey("CheckText" & i.ToString) Then
                GroupBox1.Controls("CheckText" & i.ToString).Text = My.Settings(cols(i - 1).ToString).ToString
            End If
        Next

        Dim viscols As StringCollection = DirectCast(My.Settings("VisibleColumns"), StringCollection)
        'チェックボックスオン
        For i As Integer = 0 To viscols.Count - 1
            If GroupBox1.Controls.ContainsKey("CheckVisible" & (i + 1).ToString) Then
                DirectCast(GroupBox1.Controls("CheckVisible" & (cols.IndexOf(viscols(i).ToString.TrimEnd) + 1).ToString()), CheckBox).Checked = True
            End If
        Next
    End Sub

    Private Sub ResetSetting(sender As Object, e As EventArgs) Handles Button3.Click
        If MsgBox("設定を初期化しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            My.Settings.Reset()
            FN_LoadGridInfo()
        End If
    End Sub

    Private Sub CommitSetting(sender As Object, e As EventArgs) Handles Button1.Click
        Dim viscols As StringCollection = New StringCollection
        Dim hidcols As StringCollection = New StringCollection
        For i As Integer = 1 To cols.Count
            If GroupBox1.Controls.ContainsKey("CheckText" & i.ToString) Then
                My.Settings(cols(i - 1).ToString) = GroupBox1.Controls("CheckText" & i.ToString).Text
                If DirectCast(GroupBox1.Controls("CheckVisible" & i.ToString), CheckBox).Checked Then
                    viscols.Add(cols(i - 1).ToString)
                Else
                    hidcols.Add(cols(i - 1).ToString)
                End If
            End If
        Next
        '案件NOは必須
        If Not viscols.Contains("案件NO") Then
            viscols.Add("案件NO")
        End If
        My.Settings("VisibleColumns") = viscols
        My.Settings("HiddenColumns") = hidcols
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub GridInfo_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Escape
                If My.Settings("EscClose").ToString.TrimEnd = "1" Then
                    Me.Close()
                End If
        End Select
    End Sub
End Class