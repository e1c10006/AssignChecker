Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Collections.Specialized
Imports System.Data.SqlClient


Partial Class PrimaryForm

    Private Sub RegiterForm_PG()
        P_WHERE01.ImeMode = ImeMode.Off
        P_WHERE02.ImeMode = ImeMode.On
        P_WHERE03.ImeMode = ImeMode.Off
        P_WHERE04.ImeMode = ImeMode.On
        P_WHERE05.ImeMode = ImeMode.On
        P_WHERE06.ImeMode = ImeMode.On
        P_WHERE07.ImeMode = ImeMode.On
        P_WHERE08.ImeMode = ImeMode.Off
        P_WHERE09.ImeMode = ImeMode.On
        P_WHERE10.ImeMode = ImeMode.Off
        P_WHERE11.ImeMode = ImeMode.Off
    End Sub

    Private Sub FindPGData()

        'メッセージ初期化
        StatusStrip1.Text = ""

        '常にDBに接続済の時は接続を切る 
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        L_CurrentX = Nothing
        L_CurrentY = Nothing
        PGGrid.ClearSelection()
        For i As Integer = 0 To dsDATA.Tables.Count - 1
            dsDATA.Tables(i).DefaultView.Sort = String.Empty
            dsDATA.Tables(i).Clear()
            dsDATA.Tables(i).Constraints.Clear()
            For j As Integer = dsDATA.Tables(i).Columns.Count - 1 To 0 Step -1
                dsDATA.Tables(i).Columns.RemoveAt(j)
            Next
            PGGrid.Columns.Clear()
            If i = dsDATA.Tables.Count - 1 Then
                dsDATA.Tables.Clear()
                PGGrid.DataSource = Nothing
            End If
        Next

        Try
            'データベースを選択
            cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                & "Trusted_Connection = Yes;" _
                                & "Initial Catalog=S開発品質管理;"
            cn.Open()
        Catch ex As Exception
            MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
            Return
        End Try

        Dim L_Query As String = ""
        L_Query = FN_PGLogQuery()
        Try
            'Columnのサイズは固定にしてから列を設定
            PGGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'DB接続
            Dim daAuthors As New SqlDataAdapter(L_Query, cn)
            daAuthors.FillSchema(dsDATA, SchemaType.Source)
            dsDATA.Tables("Table").PrimaryKey = Nothing
            If PGGrid.Columns.Count > 0 Then
                PGGrid.Sort(PGGrid.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            End If
            daAuthors.Fill(dsDATA)
            PGGrid.DataMember = dsDATA.Tables("table").TableName
            PGGrid.DataSource = dsDATA.Tables(0)
            'D_DataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            PGGrid.Columns(5).DefaultCellStyle.Format = "###,#0.0"
            PGGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            FN_CellSettingPG()
            SubFormCanOpen = False
        Catch ex As Exception
            Microsoft.VisualBasic.MsgBox(ex.Message, MsgBoxStyle.OkCancel, "クエリエラー")
        End Try
    End Sub

    ''' <summary>
    ''' 事例検索取得クエリ作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FN_PGLogQuery() As String
        Dim L_SQL As String = ""
        Dim L_ResultSQL As String = ""
        Dim L_ResultWhereSQL As String = ""


        'SELECT句
        L_SQL &= " SELECT TOP " & My.Settings("MaxView").ToString.TrimEnd & " PM.受注NO" & vbCrLf
        L_SQL &= " 	     ,プログラムID=RTRIM(PM.プログラムID)" & vbCrLf
        L_SQL &= " 	     ,プログラム名=RTRIM(PM.プログラム名)" & vbCrLf
        L_SQL &= " 	     ,開発担当=RTRIM(PM.開発担当者名)" & vbCrLf
        L_SQL &= " 	     ,開発完了日=RTRIM(PM.開発完了日)" & vbCrLf
        L_SQL &= " 	     ,割振工数=CONVERT(MONEY,RTRIM(PM.割振工数))" & vbCrLf
        L_SQL &= " 	     ,得意先名=RTRIM(AM.案件名)" & vbCrLf
        L_SQL &= " 	     ,設計主管者=RTRIM(AM.設計主管者名)" & vbCrLf
        L_SQL &= " 	     ,ソフト種類=RTRIM(AM.ソフト種類)" & vbCrLf
        L_SQL &= " 	     ,エディション=RTRIM(AM.業種エディション)" & vbCrLf
        L_SQL &= "   FROM T_プログラムマスタ PM LEFT OUTER JOIN" & vbCrLf
        L_SQL &= "   	  T_案件マスタ AM ON PM.受注NO = AM.受注NO" & vbCrLf
        L_SQL &= "  WHERE PM.受注NO           LIKE '%" & P_WHERE01.Text.TrimEnd & "%'" & vbCrLf
        L_SQL &= "  　AND AM.案件名           LIKE '%" & P_WHERE02.Text.TrimEnd & "%'" & vbCrLf
        L_SQL &= "    AND PM.プログラムID     LIKE '%" & P_WHERE03.Text.TrimEnd & "%'" & vbCrLf

        Dim param() As String = Split(P_WHERE04.Text.TrimEnd, " ")
        If param.Length = 1 Then
            L_SQL &= "    AND PM.プログラム名     LIKE '%" & P_WHERE04.Text.TrimEnd & "%'" & vbCrLf
        ElseIf param.Length > 1 Then
            L_SQL &= "    AND (" & vbCrLf
            For i As Integer = 0 To param.Length - 1
                If i > 0 Then
                    L_SQL &= " OR "
                End If
                L_SQL &= " PM.プログラム名 LIKE '" & param(i).ToString.TrimEnd & "%'" & vbCrLf
            Next
            L_SQL &= "        )" & vbCrLf
        End If

        Dim param2() As String = Split(P_WHERE05.Text.TrimEnd, " ")
        If param2.Length = 1 Then
            L_SQL &= "    AND PM.プログラム名     LIKE '%" & P_WHERE05.Text.TrimEnd & "%'" & vbCrLf
        ElseIf param2.Length > 1 Then
            L_SQL &= "    AND (" & vbCrLf
            For i As Integer = 0 To param2.Length - 1
                If i > 0 Then
                    L_SQL &= " OR "
                End If
                L_SQL &= " PM.プログラム名 LIKE '" & param2(i).ToString.TrimEnd & "%'" & vbCrLf
            Next
            L_SQL &= "        )" & vbCrLf
        End If
        L_SQL &= "    AND PM.開発担当者名     LIKE '%" & P_WHERE06.Text.TrimEnd & "%'" & vbCrLf
        L_SQL &= "    AND AM.設計主管者名     LIKE '%" & P_WHERE07.Text.TrimEnd & "%'" & vbCrLf
        L_SQL &= "    AND AM.ソフト種類       LIKE '%" & P_WHERE08.Text.TrimEnd & "%'" & vbCrLf
        L_SQL &= "    AND AM.業種エディション LIKE '%" & P_WHERE09.Text.TrimEnd & "%'" & vbCrLf
        If P_WHERE10.Text.TrimEnd <> "" AndAlso P_WHERE11.Text.TrimEnd <> "" Then
            L_SQL &= " AND PM.割振工数 BETWEEN " & P_WHERE10.Text.TrimEnd & " AND " & P_WHERE11.Text.TrimEnd & " " & vbCrLf
        ElseIf P_WHERE10.Text.TrimEnd <> "" Then
            L_SQL &= " AND PM.割振工数 >= " & P_WHERE10.Text.TrimEnd & " " & vbCrLf
        ElseIf P_WHERE11.Text.TrimEnd <> "" Then
            L_SQL &= " AND PM.割振工数 <= " & P_WHERE11.Text.TrimEnd & " " & vbCrLf
        End If

        Return L_SQL
    End Function

    ''' <summary>
    ''' [事例検索]カラム設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FN_CellSettingPG()
        'Return
        For i As Integer = 0 To PGGrid.Columns.Count - 1
            With PGGrid.Columns(i)
                Select Case i
                    Case 0  '受注NO
                        .Width = 70
                    Case 1  'プログラムID
                        .Width = 80
                    Case 2  'プログラム名
                        .Width = 120
                    Case 3  '開発担当者名
                        .Width = 120
                    Case 4  '開発完了日
                        .Width = 80
                    Case 5  '割振工数
                        .DefaultCellStyle.Format = "###,#0.0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 70
                    Case 6  '案件名
                        .Width = 150
                    Case 7  '設計主管者名
                        .Width = 80
                    Case 8  'ソフト種類
                        .Width = 70
                    Case 9  '業種エディション
                        .Width = 70
                End Select
            End With
        Next
    End Sub

    ''' <summary>
    ''' 事例検索カーソル移動（KeyDown）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TAB3_SetFocusEvent(sender As Object, e As EventArgs) Handles P_WHERE01.KeyDown, P_WHERE02.KeyDown, P_WHERE03.KeyDown, P_WHERE04.KeyDown, P_WHERE05.KeyDown, P_WHERE06.KeyDown, P_WHERE07.KeyDown, P_WHERE08.KeyDown, P_WHERE09.KeyDown, P_WHERE10.KeyDown, P_WHERE11.KeyDown
        'エンターか↓キーで移動
        Select Case DirectCast(e, System.Windows.Forms.KeyEventArgs).KeyValue
            Case Keys.Enter
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 2, 2))
                    If TabPage2_Panel.Controls("P_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)) Is Nothing Then
                        Index = 0
                    End If
                    TabPage2_Panel.Controls("P_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)).Select()
                End If
            Case Keys.Down
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 2, 2))
                    If TabPage2_Panel.Controls("P_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)) Is Nothing Then
                        Index = 0
                    End If
                    TabPage2_Panel.Controls("P_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)).Select()
                End If
            Case Keys.Up
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 2, 2))
                    If TabPage2_Panel.Controls("P_WHERE" & (Index - 1).ToString.PadLeft(2, "0"c)) IsNot Nothing Then
                        TabPage2_Panel.Controls("P_WHERE" & (Index - 1).ToString.PadLeft(2, "0"c)).Select()
                    End If
                End If
        End Select
    End Sub


End Class
