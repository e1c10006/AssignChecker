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
Imports System.Collections.Generic
'Imports System.Windows.Forms

Partial Class PrimaryForm

    Private Sub H6_Edit_Click(sender As Object, e As EventArgs) Handles H6_Edit.Click
        QueryEdit()
    End Sub

    Private Originaltxt As ArrayList
    Private CommitTxt As ArrayList


    Private Enum nextlist
        currentright = 0 '同じ行の次のセクション
        fromfirst = 1
        fromjoin = 2
        where = 3
    End Enum


    Private TaskList As List(Of Integer)


    Private Enum Align
        Left = 0
        Right = 1
    End Enum

    Private Sub QueryEdit()
        CommitTxt = New ArrayList()
        Originaltxt = New ArrayList()
        TaskList = New List(Of Integer)

        Dim rs As New System.IO.StringReader(Clipboard.GetText)
        While rs.Peek() > -1
            Originaltxt.Add(rs.ReadLine())
        End While


        'select * from 商品マスタ

        '読み込み行をループ
        While 1 = 1

            '解析中の行をセット
            'CurrentRow = Originaltxt(CurrentIndent)
            'TaskList.Add(CurrentIndex)


            ''単語を検索
            'ReadText = ""
            'For i As Integer = 0 To CurrentRow.Length - 1

            '    '単語確定
            '    If ReadText <> "" AndAlso CurrentRow.Substring(i, 1) <> " " Then
            '        If ReadText = "select" OrElse ReadText = "SELECT" Then
            '            CommitRow += ReadText
            '        End If

            '        If NextType = nextlist.currentright Then
            '            CommitRow
            '        End If
            '    End If

            '    '単語が終わるまで次の文字を取得する
            '    If CurrentRow.Substring(i, 1) <> " " Then
            '        ReadText += CurrentRow.Substring(i, 1)
            '    End If
            '    ReadIndex = i
            'Next
            'CommitRow += ReadText

            'CurrentIndex += 1

        End While

    End Sub

End Class
