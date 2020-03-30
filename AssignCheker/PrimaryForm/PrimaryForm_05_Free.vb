Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Collections.Specialized
Imports System.Data.SqlClient


Partial Class PrimaryForm
    ''' <summary>
    ''' 事例取得メイン処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FindTancdData()

        'メッセージ初期化
        StatusStrip1.Text = ""

        Dim a As Integer = 4
        'Dim str As String() = {}
        Dim str(a) As String
        ReDim Preserve str(a)

        '常にDBに接続済の時は接続を切る 
        Tancd_cmd.Dispose()
        Tancd_cn.Close()
        Tancd_cn.Dispose()

        TancdSub_cmd.Dispose()
        TancdSub_cn.Close()
        TancdSub_cn.Dispose()

        L_CurrentX = Nothing
        L_CurrentY = Nothing

        TancdGrid.ClearSelection()
        'データをクリアする
        For i As Integer = 0 To Tancd_dsDATA.Tables.Count - 1
            Tancd_dsDATA.Tables(i).DefaultView.Sort = String.Empty
            Tancd_dsDATA.Tables(i).Clear()
            Tancd_dsDATA.Tables(i).Constraints.Clear()
            For j As Integer = Tancd_dsDATA.Tables(i).Columns.Count - 1 To 0 Step -1
                Tancd_dsDATA.Tables(i).Columns.RemoveAt(j)
            Next
            TancdGrid.Columns.Clear()
            If i = Tancd_dsDATA.Tables.Count - 1 Then
                Tancd_dsDATA.Tables.Clear()
                TancdGrid.DataSource = Nothing
            End If
        Next
        'サブデータをクリア
        For i As Integer = 0 To TancdSub_dsDATA.Tables.Count - 1
            TancdSub_dsDATA.Tables(i).DefaultView.Sort = String.Empty
            TancdSub_dsDATA.Tables(i).Clear()
            TancdSub_dsDATA.Tables(i).Constraints.Clear()
            For j As Integer = TancdSub_dsDATA.Tables(i).Columns.Count - 1 To 0 Step -1
                TancdSub_dsDATA.Tables(i).Columns.RemoveAt(j)
            Next
            If i = TancdSub_dsDATA.Tables.Count - 1 Then
                TancdSub_dsDATA.Tables.Clear()
            End If
        Next

        Dim L_Query As String = ""
        L_Query = FN_TancdQuery()
        If L_Query = "" Then
            Return
        End If
        Try
            'Columnのサイズは固定にしてから列を設定
            TancdGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            TancdGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
            TancdGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            'Me.SuspendLayout()

            'DB接続
            Dim daAuthors As New SqlDataAdapter(L_Query, Tancd_cn)
            daAuthors.FillSchema(Tancd_dsDATA, SchemaType.Source)
            Tancd_dsDATA.Tables("Table").PrimaryKey = Nothing
            If TancdGrid.Columns.Count > 0 Then
                TancdGrid.Sort(TancdGrid.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            End If
            daAuthors.Fill(Tancd_dsDATA)
            TancdGrid.DataMember = Tancd_dsDATA.Tables("table").TableName
            TancdGrid.DataSource = Tancd_dsDATA.Tables(0)

            'データ編集
            FN_ModifyData()

            FN_CellSettingTancd()
            SubFormCanOpen = True

            'Me.ResumeLayout()
        Catch ex As Exception
            Microsoft.VisualBasic.MsgBox(ex.Message, MsgBoxStyle.OkCancel, "クエリエラー")
        End Try
    End Sub

    ''' <summary>
    ''' 担当者検索取得クエリ作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FN_TancdQuery() As String
        Dim L_SQL As String = ""
        Dim L_ResultSQL As String = ""
        Dim L_ResultWhereSQL As String = ""
        If Not IsNumeric(T5_対象年月.Text.ToString) Then
            MetroFramework.MetroMessageBox.Show(Me, "対象年度が無効です", "エラー")
            Return ""
        End If


        If T5_対象年月.Text.TrimEnd.Length < 4 OrElse Integer.TryParse(Not T5_対象年月.Text.TrimEnd, 0) = False Then
            MetroFramework.MetroMessageBox.Show(Me, "対象年月は数値4桁以上で入力して下さい。")
            Return ""
        End If

        Dim dt As DateTime = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), 1, 1)
        Dim L_BfYear As String = dt.ToString("yyyy")
        Dim L_NwYear As String = dt.AddYears(+1).ToString("yyyy")
        Dim L_PGdt As String = ""
        Dim L_SEdt As String = ""
        Dim L_DCdt As String = ""

        L_SQL = ""
        Select Case T5_TAIDT.Text
            Case "生産高確認"
                Try
                    'データベースを選択
                    Tancd_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                        & "Trusted_Connection = Yes;" _
                                        & "Initial Catalog=S開発品質管理;"
                    Tancd_cn.Open()
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                    Return ""
                End Try

                If T5_予定を含む.Checked Then
                    L_PGdt = "ISNULL(NULLIF(P.開発完了日,''),P.開発予定日)"
                    L_SEdt = "ISNULL(NULLIF(P.SE引渡日,''),P.SE引渡予定日)"
                    L_DCdt = "ISNULL(NULLIF(P.仕様書作成完了日,''),P.仕様書作成予定日)"
                Else
                    L_PGdt = "P.開発完了日"
                    L_SEdt = "P.SE引渡日"
                    L_DCdt = "P.仕様書作成完了日"
                End If

                'SELECT句
                If T5_四半期表示.Checked Then
                    If T5_チーム別表示.Checked Then
                        L_SQL &= " SELECT 部署名" & vbCrLf
                        L_SQL &= "       ,[第一四半期] = [08月] + [09月] + [10月]" & vbCrLf
                        L_SQL &= "       ,[第二四半期] = [11月] + [12月] + [01月]" & vbCrLf
                        L_SQL &= "       ,[第三四半期] = [02月] + [03月] + [04月]" & vbCrLf
                        L_SQL &= "       ,[第四四半期] = [05月] + [06月] + [07月]" & vbCrLf
                    Else
                        L_SQL &= " SELECT 担当者" & vbCrLf
                        L_SQL &= "       ,名" & vbCrLf
                        L_SQL &= "       ,[第一四半期] = [08月] + [09月] + [10月]" & vbCrLf
                        L_SQL &= "       ,[第二四半期] = [11月] + [12月] + [01月]" & vbCrLf
                        L_SQL &= "       ,[第三四半期] = [02月] + [03月] + [04月]" & vbCrLf
                        L_SQL &= "       ,[第四四半期] = [05月] + [06月] + [07月]" & vbCrLf
                    End If
                Else
                    L_SQL &= " SELECT *" & vbCrLf
                    L_SQL &= "       ,順位   = ROW_NUMBER() OVER (ORDER BY 累計 DESC)" & vbCrLf
                End If
                L_SQL &= "   FROM " & vbCrLf
                L_SQL &= " (" & vbCrLf
                If T5_チーム別表示.Checked Then
                    L_SQL &= " SELECT 部署名=BMN.部署名" & vbCrLf
                    L_SQL &= " 	  ,[08月] = SUM([08月_PG]) + SUM([08月_CR]) + SUM([08月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[09月] = SUM([09月_PG]) + SUM([09月_CR]) + SUM([09月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[10月] = SUM([10月_PG]) + SUM([10月_CR]) + SUM([10月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[11月] = SUM([11月_PG]) + SUM([11月_CR]) + SUM([11月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[12月] = SUM([12月_PG]) + SUM([12月_CR]) + SUM([12月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[01月] = SUM([01月_PG]) + SUM([01月_CR]) + SUM([01月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[02月] = SUM([02月_PG]) + SUM([02月_CR]) + SUM([02月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[03月] = SUM([03月_PG]) + SUM([03月_CR]) + SUM([03月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[04月] = SUM([04月_PG]) + SUM([04月_CR]) + SUM([04月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[05月] = SUM([05月_PG]) + SUM([05月_CR]) + SUM([05月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[06月] = SUM([06月_PG]) + SUM([06月_CR]) + SUM([06月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,[07月] = SUM([07月_PG]) + SUM([07月_CR]) + SUM([07月_仕様書])" & vbCrLf
                    L_SQL &= " 	  ,累計   = SUM([08月_PG]) + SUM([08月_CR]) + SUM([08月_仕様書])" & vbCrLf
                    L_SQL &= " 	          + SUM([09月_PG]) + SUM([09月_CR]) + SUM([09月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([10月_PG]) + SUM([10月_CR]) + SUM([10月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([11月_PG]) + SUM([11月_CR]) + SUM([11月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([12月_PG]) + SUM([12月_CR]) + SUM([12月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([01月_PG]) + SUM([01月_CR]) + SUM([01月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([02月_PG]) + SUM([02月_CR]) + SUM([02月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([03月_PG]) + SUM([03月_CR]) + SUM([03月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([04月_PG]) + SUM([04月_CR]) + SUM([04月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([05月_PG]) + SUM([05月_CR]) + SUM([05月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([06月_PG]) + SUM([06月_CR]) + SUM([06月_仕様書])" & vbCrLf
                    L_SQL &= " 			  + SUM([07月_PG]) + SUM([07月_CR]) + SUM([07月_仕様書])" & vbCrLf
                Else
                    L_SQL &= " SELECT 担当者=MAIN.担当者コード" & vbCrLf
                    L_SQL &= "       ,名=TNM.担当者名" & vbCrLf
                    L_SQL &= " 	  ,[08月] = [08月_PG] + [08月_CR] + [08月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[09月] = [09月_PG] + [09月_CR] + [09月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[10月] = [10月_PG] + [10月_CR] + [10月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[11月] = [11月_PG] + [11月_CR] + [11月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[12月] = [12月_PG] + [12月_CR] + [12月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[01月] = [01月_PG] + [01月_CR] + [01月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[02月] = [02月_PG] + [02月_CR] + [02月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[03月] = [03月_PG] + [03月_CR] + [03月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[04月] = [04月_PG] + [04月_CR] + [04月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[05月] = [05月_PG] + [05月_CR] + [05月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[06月] = [06月_PG] + [06月_CR] + [06月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,[07月] = [07月_PG] + [07月_CR] + [07月_仕様書]" & vbCrLf
                    L_SQL &= " 	  ,累計   = [08月_PG] + [08月_CR] + [08月_仕様書] " & vbCrLf
                    L_SQL &= " 	          + [09月_PG] + [09月_CR] + [09月_仕様書]" & vbCrLf
                    L_SQL &= " 			  + [10月_PG] + [10月_CR] + [10月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [11月_PG] + [11月_CR] + [11月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [12月_PG] + [12月_CR] + [12月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [01月_PG] + [01月_CR] + [01月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [02月_PG] + [02月_CR] + [02月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [03月_PG] + [03月_CR] + [03月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [04月_PG] + [04月_CR] + [04月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [05月_PG] + [05月_CR] + [05月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [06月_PG] + [06月_CR] + [06月_仕様書] " & vbCrLf
                    L_SQL &= " 			  + [07月_PG] + [07月_CR] + [07月_仕様書]" & vbCrLf
                End If
                L_SQL &= "   FROM ("
                L_SQL &= " 	   SELECT "
                L_SQL &= " 	    担当者コード = REVERSE(CONVERT(CHAR(6),REVERSE('000000' + RTRIM((CASE WHEN T.BP区分 = '0' THEN T.担当者コード ELSE T2.担当者コード END)))))" & vbCrLf
                If Not T5_BP実績を含む.Checked Then
                    L_SQL &= "     ,[08月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[08月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[09月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[09月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[10月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[10月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[11月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[11月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[12月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[12月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[01月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[01月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[02月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[02月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[03月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[03月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[04月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[04月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[05月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[05月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[06月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[06月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[07月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[07月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 15000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[08月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[09月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[10月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[11月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[12月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[01月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[02月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[03月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[04月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[05月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[06月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[07月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                Else
                    L_SQL &= "     ,[08月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[08月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[09月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[09月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[10月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[10月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[11月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[11月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[12月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[12月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[01月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[01月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[02月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[02月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[03月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[03月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[04月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[04月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[05月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[05月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[06月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[06月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[07月_PG]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "     ,[07月_CR]    = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 50000 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 37500 END,0))" & vbCrLf
                    L_SQL &= "                   + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 12500 END,0))" & vbCrLf
                    L_SQL &= "     ,[08月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '08' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[09月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '09' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[10月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '10' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[11月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '11' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[12月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_BfYear & "' + '12' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[01月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '01' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[02月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '02' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[03月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '03' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[04月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '04' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[05月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '05' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[06月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '06' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf
                    L_SQL &= "     ,[07月_仕様書]= SUM(ISNULL(CASE WHEN P.仕様書担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_NwYear & "' + '07' THEN P.割振工数 * 50000 * 0.3 END,0))" & vbCrLf

                End If


                L_SQL &= "     ,存在区分   = CASE WHEN SUM(割振工数)>0 THEN '1' ELSE '0'　END" & vbCrLf
                L_SQL &= " 	   FROM T_プログラムマスタ P" & vbCrLf
                L_SQL &= " 	   INNER JOIN T_案件マスタ A    ON A.受注NO        = P.受注NO AND A.枝番=P.枝番" & vbCrLf
                L_SQL &= " 	   INNER JOIN T_担当者マスタ T  ON (P.仕様書作成担当者コード = T.担当者コード OR T.担当者コード  = P.開発担当者コード) AND T.部署コード <> ''" & vbCrLf
                L_SQL &= " 	   INNER JOIN T_担当者マスタ T2 ON T2.担当者コード = A.開発主管者コード" & vbCrLf
                L_SQL &= " 	   WHERE REPLACE(P.開発完了日,'/','')   BETWEEN '" & L_BfYear & "' + '0801' AND '" & L_NwYear & "' + '0799'" & vbCrLf
                L_SQL &= " 	      OR REPLACE(P.開発予定日,'/','')   BETWEEN '" & L_BfYear & "' + '0801' AND '" & L_NwYear & "' + '0799'" & vbCrLf
                L_SQL &= " 	      OR REPLACE(P.SE引渡日,'/','')     BETWEEN '" & L_BfYear & "' + '0801' AND '" & L_NwYear & "' + '0799'" & vbCrLf
                L_SQL &= "           OR REPLACE(P.SE引渡予定日,'/','') BETWEEN '" & L_BfYear & "' + '0801' AND '" & L_NwYear & "' + '0799'" & vbCrLf
                L_SQL &= " 	   GROUP BY (CASE WHEN T.BP区分 = '0' THEN T.担当者コード ELSE T2.担当者コード END)" & vbCrLf
                L_SQL &= " 	   ) MAIN" & vbCrLf
                L_SQL &= " 	   INNER JOIN T_担当者マスタ TNM  ON TNM.担当者コード  = MAIN.担当者コード" & vbCrLf
                L_SQL &= "     INNER JOIN T_部署マスタ BMN ON TNM.部署コード = BMN.部署コード " & vbCrLf
                L_SQL &= " 	   WHERE 存在区分 = '1'" & vbCrLf
                If T5_BMNLIST.Text.TrimEnd <> "" Then
                    If T5_BMNLIST.Text.TrimEnd = "関西" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                    ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                    Else
                        L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                    End If
                End If
                If T5_社員コード.Text.TrimEnd <> "" Then
                    L_SQL &= " AND TNM.担当者コード LIKE '%" & T5_社員コード.Text.TrimEnd & "%'" & vbCrLf
                End If
                If T5_チーム別表示.Checked Then
                    L_SQL &= " GROUP BY BMN.部署名" & vbCrLf
                End If
                L_SQL &= " ) MAIN " & vbCrLf
                L_SQL &= " WHERE 累計 > 0" & vbCrLf
                If T5_チーム別表示.Checked Then
                    L_SQL &= " ORDER BY 部署名" & vbCrLf
                Else
                    L_SQL &= " ORDER BY 累計 DESC" & vbCrLf
                End If

            Case "月別割振工数"

                Try
                    'データベースを選択
                    Tancd_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                        & "Trusted_Connection = Yes;" _
                                        & "Initial Catalog=S開発アサイン管理;"
                    Tancd_cn.Open()
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                    Return ""
                End Try

                Dim FromCol As String = "仕様書提出日"
                If T5_SE渡しを含む.Checked Then
                    FromCol = "PG完了希望日"
                End If

                'SELECT句
                If T5_チーム別表示.Checked = False Then
                    L_SQL &= "SELECT 社員NO=RIGHT('000000' + LTRIM(MAIN.担当者コード),6)"
                    L_SQL &= "      ,社員名=TNM.担当者名"
                Else
                    L_SQL &= "SELECT 部署名=BMN.部署名"
                End If

                L_SQL &= "	  ,[08月] = [08月_PG] " & vbCrLf
                L_SQL &= "	  ,[09月] = [09月_PG] " & vbCrLf
                L_SQL &= "	  ,[10月] = [10月_PG] " & vbCrLf
                L_SQL &= "	  ,[11月] = [11月_PG] " & vbCrLf
                L_SQL &= "	  ,[12月] = [12月_PG] " & vbCrLf
                L_SQL &= "	  ,[01月] = [01月_PG] " & vbCrLf
                L_SQL &= "	  ,[02月] = [02月_PG] " & vbCrLf
                L_SQL &= "	  ,[03月] = [03月_PG] " & vbCrLf
                L_SQL &= "	  ,[04月] = [04月_PG] " & vbCrLf
                L_SQL &= "	  ,[05月] = [05月_PG] " & vbCrLf
                L_SQL &= "	  ,[06月] = [06月_PG] " & vbCrLf
                L_SQL &= "	  ,[07月] = [07月_PG] " & vbCrLf
                L_SQL &= "	  ,累計   = [08月_PG] " & vbCrLf
                L_SQL &= "	          + [09月_PG] " & vbCrLf
                L_SQL &= "			  + [10月_PG] " & vbCrLf
                L_SQL &= "			  + [11月_PG] " & vbCrLf
                L_SQL &= "			  + [12月_PG] " & vbCrLf
                L_SQL &= "			  + [01月_PG] " & vbCrLf
                L_SQL &= "			  + [02月_PG] " & vbCrLf
                L_SQL &= "			  + [03月_PG] " & vbCrLf
                L_SQL &= "			  + [04月_PG] " & vbCrLf
                L_SQL &= "			  + [05月_PG] " & vbCrLf
                L_SQL &= "			  + [06月_PG] " & vbCrLf
                L_SQL &= "			  + [07月_PG] " & vbCrLf
                If T5_チーム別表示.Checked = False Then
                    L_SQL &= "  FROM (SELECT ATM.担当者コード" & vbCrLf
                Else
                    L_SQL &= "  FROM (SELECT BMN.部署コード" & vbCrLf
                End If
                L_SQL &= "              ,[08月_PG]    = SUM(ISNULL(CASE WHEN '" & L_BfYear & "'+'08' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[09月_PG]    = SUM(ISNULL(CASE WHEN '" & L_BfYear & "'+'09' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[10月_PG]    = SUM(ISNULL(CASE WHEN '" & L_BfYear & "'+'10' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[11月_PG]    = SUM(ISNULL(CASE WHEN '" & L_BfYear & "'+'11' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[12月_PG]    = SUM(ISNULL(CASE WHEN '" & L_BfYear & "'+'12' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[01月_PG]    = SUM(ISNULL(CASE WHEN '" & L_NwYear & "'+'01' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[02月_PG]    = SUM(ISNULL(CASE WHEN '" & L_NwYear & "'+'02' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[03月_PG]    = SUM(ISNULL(CASE WHEN '" & L_NwYear & "'+'03' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[04月_PG]    = SUM(ISNULL(CASE WHEN '" & L_NwYear & "'+'04' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[05月_PG]    = SUM(ISNULL(CASE WHEN '" & L_NwYear & "'+'05' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[06月_PG]    = SUM(ISNULL(CASE WHEN '" & L_NwYear & "'+'06' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,[07月_PG]    = SUM(ISNULL(CASE WHEN '" & L_NwYear & "'+'07' BETWEEN CONVERT(CHAR(6),仕様書提出日) AND CONVERT(CHAR(6)," & FromCol & ") THEN PG依頼工数 END,0))" & vbCrLf
                L_SQL &= "              ,存在区分   = CASE WHEN SUM(PG依頼工数)>0 THEN '1' ELSE '0'　END" & vbCrLf
                L_SQL &= "	   FROM 基本案件情報トラン TRN" & vbCrLf
                L_SQL &= "          LEFT OUTER JOIN アサイン一覧案件トラン Sub ON TRN.案件NO = Sub.案件NO"
                L_SQL &= "          LEFT OUTER JOIN 担当者マスタ ATM ON Sub.CR > '' AND Sub.CR = ATM.略称"
                If T5_チーム別表示.Checked = False Then
                    L_SQL &= "	   GROUP BY ATM.担当者コード" & vbCrLf
                Else
                    L_SQL &= " inner join S開発品質管理.dbo.T_担当者マスタ TNM ON TNM.略称  = ATM.担当者コード" & vbCrLf
                    L_SQL &= " inner join S開発品質管理.dbo.T_部署マスタ BMN on TNM.部署コード = BMN.部署コード " & vbCrLf
                    If T5_BMNLIST.Text.ToString.TrimEnd <> "" Then
                        If T5_BMNLIST.Text.TrimEnd = "関西" Then
                            L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                        ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                            L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                        Else
                            L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                        End If
                    End If
                    L_SQL &= "	   GROUP BY BMN.部署コード" & vbCrLf
                End If
                L_SQL &= "	   ) MAIN" & vbCrLf
                If T5_チーム別表示.Checked = False Then
                    L_SQL &= "	   inner join S開発品質管理.dbo.T_担当者マスタ TNM  on TNM.担当者コード  = MAIN.担当者コード" & vbCrLf
                    L_SQL &= "	   inner join S開発品質管理.dbo.T_部署マスタ BMN on TNM.部署コード = BMN.部署コード " & vbCrLf
                    L_SQL &= "	   WHERE 存在区分 = '1'" & vbCrLf
                    If T5_社員コード.Text.ToString.TrimEnd <> "" Then
                        L_SQL &= "   AND TNM.担当者コード = '" & T5_社員コード.Text.TrimEnd & "' " & vbCrLf
                    End If
                    If T5_社員コード.Text.ToString.TrimEnd <> "" Then
                        L_SQL &= "   AND TNM.担当者名 = '" & T5_社員名.Text.TrimEnd & "' " & vbCrLf
                    End If
                    If T5_BMNLIST.Text.ToString.TrimEnd <> "" Then
                        If T5_BMNLIST.Text.TrimEnd = "関西" Then
                            L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                        ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                            L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                        Else
                            L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                        End If
                    End If
                Else
                    L_SQL &= " inner join S開発品質管理.dbo.T_部署マスタ BMN on MAIN.部署コード = BMN.部署コード " & vbCrLf
                End If
                L_SQL &= "    ORDER BY 累計 DESC" & vbCrLf
            Case "原価入力リスト"
                '    Try
                '        'データベースを選択
                '        Tancd_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                '                            & "Trusted_Connection = Yes;" _
                '                            & "Initial Catalog=S開発品質管理;"
                '        Tancd_cn.Open()
                '    Catch ex As Exception
                '        MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                '        Return ""
                '    End Try

                '    dt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 1)

                '    L_SQL &= " SELECT 社員NO=RIGHT('000' + LTRIM(TNM.担当者コード),6) "
                '    L_SQL &= "       ,社員名=TNM.担当者名 "
                '    L_SQL &= "       ,累計=0.0"
                '    L_SQL &= "       ,残業=0.0"
                '    L_SQL &= "   FROM T_担当者マスタ TNM"
                '    L_SQL &= "        INNER JOIN T_部署マスタ BMN on TNM.部署コード = BMN.部署コード"
                '    L_SQL &= "  WHERE TNM.削除区分 = '0'"
                '    L_SQL &= "    AND BP区分 = '0'"
                '    L_SQL &= "    AND 部署種別 = '2'"
                '    If T5_社員コード.Text <> "" Then
                '        L_SQL &= "    AND TNM.担当者コード LIKE '%" & Integer.Parse(T5_社員コード.Text.TrimEnd) & "%'"
                '    End If
                '    If T5_社員名.Text <> "" Then
                '        L_SQL &= "    AND TNM.担当者名 LIKE '%" & T5_社員名.Text.TrimEnd & "%'"
                '    End If
                '    If T5_BMNLIST.Text <> "" Then
                '        If T5_BMNLIST.Text.TrimEnd = "関西" Then
                '            L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                '        ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                '            L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                '        Else
                '            L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                '        End If
                '    End If

                '    If T5_BMNLIST.Text.ToString.TrimEnd = "SE" Then
                '        L_SQL = "    SELECT 社員NO = RIGHT(('000000' + LTRIM(T.担当者コード)),6) "
                '        L_SQL &= "          ,社員名 = T.担当者名 "
                '        L_SQL &= "          ,累計   = 0.0"
                '        L_SQL &= "          ,残業   =0.0"
                '        L_SQL &= "      FROM SE担当一覧 SE LEFT OUTER JOIN S開発品質管理.dbo.T_担当者マスタ T ON T.担当者コード = SE.担当者コード  "
                '    End If
                dt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 1)

                Dim sttdt As DateTime
                Dim enddt As DateTime
                If T5_15日締表示.Checked Then
                    sttdt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 16).AddMonths(-1)
                    enddt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 15)
                Else
                    sttdt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 1)
                    enddt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), sttdt.AddMonths(1).AddDays(-1).Day)
                End If
                Dim sttdays As Integer = DateTime.DaysInMonth(sttdt.Year, sttdt.Month)
                Dim enddays As Integer = DateTime.DaysInMonth(enddt.Year, enddt.Month)

                Try
                    'データベースを選択
                    Tancd_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                        & "Trusted_Connection = Yes;" _
                                        & "Initial Catalog=S開発品質管理;"
                    Tancd_cn.Open()
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                    Return ""
                End Try


                L_SQL = ""
                L_SQL &= "  WITH 日付情報 (ID,日付)"
                L_SQL &= "  AS"
                L_SQL &= "  ("
                L_SQL &= "       SELECT	1,CONVERT(DATE,'" & sttdt.ToString("yyyyMMdd") & "')"
                L_SQL &= "       UNION ALL"
                L_SQL &= "       SELECT	ID+1,DATEADD(dd, 1, 日付)"
                L_SQL &= "       FROM   日付情報"
                L_SQL &= "       WHERE  日付 < CONVERT(DATE,'" & enddt.ToString("yyyyMMdd") & "') "
                L_SQL &= "  )"
                L_SQL &= " SELECT "
                L_SQL &= "  社員NO"
                L_SQL &= " ,社員名"
                L_SQL &= " ,累計 = SUM(GNK.作業時間)"
                L_SQL &= " ,残業 = SUM(CASE  WHEN GNK.作業時間 > 7.5 THEN GNK.作業時間 - 7.5 ELSE 0 END)"
                If T5_15日締表示.Checked Then
                    For i As Integer = 16 To sttdays
                        L_SQL &= " ,[" & sttdt.ToString("MM") & "/" & i.ToString.PadLeft(2, "0"c) & "] =NULLIF(SUM( CASE WHEN MAIN.日付 = '" & i.ToString & "' THEN GNK.作業時間 ELSE 0 END),0)"
                    Next
                    For i As Integer = 1 To 15
                        L_SQL &= " ,[" & enddt.ToString("MM") & "/" & i.ToString.PadLeft(2, "0"c) & "] =NULLIF(SUM( CASE WHEN MAIN.日付 = '" & i.ToString & "' THEN GNK.作業時間 ELSE 0 END),0)"
                    Next
                Else
                    For i As Integer = 1 To sttdays
                        L_SQL &= " ,[" & sttdt.ToString("MM") & "/" & i.ToString.PadLeft(2, "0"c) & "] =NULLIF(SUM( CASE WHEN MAIN.日付 = '" & i.ToString & "' THEN GNK.作業時間 ELSE 0 END),0)"
                    Next
                End If
                L_SQL &= " FROM ( SELECT 社員NO=RIGHT('000' + LTRIM(TNM.担当者コード),6)"
                L_SQL &= " 				       ,社員名=TNM.担当者名"
                L_SQL &= " 					   ,DAY(日付) 日付"
                L_SQL &= " 				   FROM 日付情報 DT "
                L_SQL &= " 				   LEFT JOIN T_担当者マスタ TNM ON 1= 1"
                L_SQL &= " 				   INNER JOIN T_部署マスタ BMN on TNM.部署コード = BMN.部署コード"
                L_SQL &= " 				WHERE TNM.削除区分 = '0' AND BP区分 = '0' AND 部署種別 = '2' "
                If T5_社員コード.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者コード LIKE '%" & Integer.Parse(T5_社員コード.Text.TrimEnd) & "%'"
                End If
                If T5_社員名.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者名 LIKE '%" & T5_社員名.Text.TrimEnd & "%'"
                End If
                If T5_BMNLIST.Text <> "" Then
                    If T5_BMNLIST.Text.TrimEnd = "関西" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                    ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                    Else
                        L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                    End If
                End If

                'If T5_BMNLIST.Text.ToString.TrimEnd = "SE" Then
                '    L_SQL = "    SELECT 社員NO = RIGHT(('000000' + LTRIM(T.担当者コード)),6) "
                '    L_SQL &= "          ,社員名 = T.担当者名 "
                '    L_SQL &= "          ,累計   = 0.0"
                '    L_SQL &= "          ,残業   =0.0"
                '    L_SQL &= "      FROM SE担当一覧 SE LEFT OUTER JOIN S開発品質管理.dbo.T_担当者マスタ T ON T.担当者コード = SE.担当者コード  "
                'End If
                L_SQL &= " 				) MAIN"
                L_SQL &= " LEFT OUTER JOIN OPENQUERY([IC01\SHARE],"
                L_SQL &= "    'SELECT 担当者コード,DAY(作業日付) AS 作業日付,SUM(作業時間) AS 作業時間"
                L_SQL &= "   FROM 原価管理link.dbo.原価トランビュー "
                L_SQL &= "  WHERE 作業日付 BETWEEN ''" & sttdt.ToString("yyyy/MM/dd") & "'' AND ''" & enddt.ToString("yyyy/MM/dd") & "''"
                L_SQL &= "  GROUP BY 担当者コード,DAY(作業日付)') GNK ON GNK.担当者コード = 社員NO AND  GNK.作業日付 = MAIN.日付"
                L_SQL &= "  GROUP BY 社員NO,社員名"
                L_SQL &= " ORDER BY 累計 DESC"


            Case "仕様書提出リスト"
                Try
                    'データベースを選択
                    Tancd_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                        & "Trusted_Connection = Yes;" _
                                        & "Initial Catalog=S開発アサイン管理;"
                    Tancd_cn.Open()
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                    Return ""
                End Try

                dt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 1)

                L_SQL &= " WITH 日付情報 (ID,日付)"
                L_SQL &= " AS"
                L_SQL &= " ("
                L_SQL &= "      SELECT	1,CONVERT(DATE,'" & dt.ToString("yyyy-MM-dd") & "')"
                L_SQL &= "      UNION ALL"
                L_SQL &= "      SELECT	ID+1,DATEADD(dd, 1, 日付)"
                L_SQL &= "      FROM   日付情報"
                L_SQL &= "      WHERE  日付 < CONVERT(DATE,'" & dt.AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd") & "') "
                L_SQL &= " )"
                L_SQL &= " SELECT ADT.案件NO,"
                L_SQL &= "        ADT.受注コード,"
                L_SQL &= "        得意先名=LTRIM(REPLACE(TKM.得意先名,'株式会社','')),"
                L_SQL &= " 	      工数=ADT.PG依頼工数,"
                L_SQL &= " 	      PG担当=ISNULL(TNM.担当者名,''),"
                L_SQL &= " 	      SE担当=ISNULL(STM.担当者名,''),"
                L_SQL &= "        [01] =MAIN.[1日] , "
                L_SQL &= "        [02] =MAIN.[2日] , "
                L_SQL &= "        [03] =MAIN.[3日] , "
                L_SQL &= "        [04] =MAIN.[4日] , "
                L_SQL &= "        [05] =MAIN.[5日] , "
                L_SQL &= "        [06] =MAIN.[6日] , "
                L_SQL &= "        [07] =MAIN.[7日] , "
                L_SQL &= "        [08] =MAIN.[8日] , "
                L_SQL &= "        [09] =MAIN.[9日] , "
                L_SQL &= "        [10]=MAIN.[10日],"
                L_SQL &= "        [11]=MAIN.[11日],"
                L_SQL &= "        [12]=MAIN.[12日],"
                L_SQL &= "        [13]=MAIN.[13日],"
                L_SQL &= "        [14]=MAIN.[14日],"
                L_SQL &= "        [15]=MAIN.[15日],"
                L_SQL &= "        [16]=MAIN.[16日],"
                L_SQL &= "        [17]=MAIN.[17日],"
                L_SQL &= "        [18]=MAIN.[18日],"
                L_SQL &= "        [19]=MAIN.[19日],"
                L_SQL &= "        [20]=MAIN.[20日],"
                L_SQL &= "        [21]=MAIN.[21日],"
                L_SQL &= "        [22]=MAIN.[22日],"
                L_SQL &= "        [23]=MAIN.[23日],"
                L_SQL &= "        [24]=MAIN.[24日],"
                L_SQL &= "        [25]=MAIN.[25日],"
                L_SQL &= "        [26]=MAIN.[26日],"
                L_SQL &= "        [27]=MAIN.[27日],"
                L_SQL &= "        [28]=MAIN.[28日],"
                L_SQL &= "        [29]=MAIN.[29日],"
                L_SQL &= "        [30]=MAIN.[30日],"
                L_SQL &= "        [31]=MAIN.[31日]"
                L_SQL &= "   FROM (SELECT 案件NO  = ISNULL(SDT.案件NO,PDT.案件NO)"
                L_SQL &= " 	            , [1日]  = MAX(CASE WHEN ID = 1    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 1    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [2日]  = MAX(CASE WHEN ID = 2    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 2    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [3日]  = MAX(CASE WHEN ID = 3    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 3    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [4日]  = MAX(CASE WHEN ID = 4    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 4    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [5日]  = MAX(CASE WHEN ID = 5    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 5    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [6日]  = MAX(CASE WHEN ID = 6    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 6    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [7日]  = MAX(CASE WHEN ID = 7    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 7    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [8日]  = MAX(CASE WHEN ID = 8    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 8    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [9日]  = MAX(CASE WHEN ID = 9    AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 9    AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [10日] = MAX(CASE WHEN ID = 10   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 10   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [11日] = MAX(CASE WHEN ID = 11   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 11   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [12日] = MAX(CASE WHEN ID = 12   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 12   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [13日] = MAX(CASE WHEN ID = 13   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 13   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [14日] = MAX(CASE WHEN ID = 14   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 14   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [15日] = MAX(CASE WHEN ID = 15   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 15   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [16日] = MAX(CASE WHEN ID = 16   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 16   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [17日] = MAX(CASE WHEN ID = 17   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 17   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [18日] = MAX(CASE WHEN ID = 18   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 18   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [19日] = MAX(CASE WHEN ID = 19   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 19   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [20日] = MAX(CASE WHEN ID = 20   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 20   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [21日] = MAX(CASE WHEN ID = 21   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 21   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [22日] = MAX(CASE WHEN ID = 22   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 22   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [23日] = MAX(CASE WHEN ID = 23   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 23   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [24日] = MAX(CASE WHEN ID = 24   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 24   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [25日] = MAX(CASE WHEN ID = 25   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 25   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [26日] = MAX(CASE WHEN ID = 26   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 26   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [27日] = MAX(CASE WHEN ID = 27   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 27   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [28日] = MAX(CASE WHEN ID = 28   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 28   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [29日] = MAX(CASE WHEN ID = 29   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 29   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [30日] = MAX(CASE WHEN ID = 30   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 30   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= " 	            , [31日] = MAX(CASE WHEN ID = 31   AND SDT.受注コード IS NOT NULL THEN '●'  WHEN ID = 31   AND PDT.受注コード IS NOT NULL THEN '★' ELSE '' END)"
                L_SQL &= "   FROM 日付情報 DT "
                L_SQL &= "        LEFT OUTER JOIN 基本案件情報トラン SDT ON SDT.仕様書提出日 = DT.日付"
                L_SQL &= " 	      LEFT OUTER JOIN 基本案件情報トラン PDT ON PDT.PG完了希望日 = DT.日付"
                L_SQL &= "  GROUP BY ISNULL(SDT.案件NO,PDT.案件NO)"
                L_SQL &= "   ) MAIN"
                L_SQL &= "     INNER JOIN 基本案件情報トラン ADT ON ADT.案件NO = MAIN.案件NO"
                L_SQL &= "     LEFT OUTER JOIN 担当者マスタ STM ON ADT.担当者コード = STM.担当者コード"
                L_SQL &= "     LEFT OUTER JOIN 得意先マスタ TKM ON ADT.得意先コード = TKM.得意先コード"
                L_SQL &= "     LEFT OUTER JOIN アサイン一覧案件トラン Sub ON ADT.案件NO = Sub.案件NO"
                L_SQL &= "     LEFT OUTER JOIN 担当者マスタ ATM ON Sub.CR > '' AND ATM.略称 = Sub.CR"
                L_SQL &= " 	   LEFT OUTER JOIN (SELECT 受注NO,担当者コード=MAX(開発主管者コード) FROM S開発品質管理.dbo.T_案件マスタ GROUP BY 受注NO) ANM ON ADT.受注コード =ANM.受注NO"
                L_SQL &= " 	   LEFT OUTER JOIN S開発品質管理.dbo.T_担当者マスタ TNM ON TNM.担当者コード = ISNULL(ANM.担当者コード,ATM.担当者コード)"
                L_SQL &= " 	   LEFT OUTER JOIN S開発品質管理.dbo.T_部署マスタ BMN on TNM.部署コード = BMN.部署コード"
                L_SQL &= "   WHERE 1=1"
                If T5_社員コード.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者コード LIKE '%" & Integer.Parse(T5_社員コード.Text.TrimEnd) & "%'"
                End If
                If T5_社員名.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者名 LIKE '%" & T5_社員名.Text.TrimEnd & "%'"
                End If
                If T5_BMNLIST.Text <> "" Then
                    If T5_BMNLIST.Text.TrimEnd = "関西" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                    ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                    Else
                        L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                    End If
                End If
                L_SQL &= "   ORDER BY ADT.仕様書提出日"
                L_SQL &= "   OPTION (MAXRECURSION 0)"
            Case "PG経験"
                Try
                    'データベースを選択
                    Tancd_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                        & "Trusted_Connection = Yes;" _
                                        & "Initial Catalog=S開発品質管理;"
                    Tancd_cn.Open()
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                    Return ""
                End Try

                L_SQL = ""
                L_SQL &= " SELECT 開発担当者コード AS 担当者"
                L_SQL &= " 	    , MAX(担当者名) AS 担当者名"
                L_SQL &= " 	    , SUM(割振工数) AS 総工数"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN ソフト種類 LIKE '%2.%'      THEN 割振工数 END),0) AS [2.X]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN ソフト種類 LIKE '%AONET%'   THEN 割振工数 END),0) AS [AONET]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN 業種エディション LIKE '%ファッション%'  THEN 割振工数 END),0) AS [ﾌｧｯｼｮﾝ]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN 業種エディション LIKE '%食品%'          THEN 割振工数 END),0) AS [食品]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN 業種エディション LIKE '%鋼材%'          THEN 割振工数 END),0) AS [鋼材]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN ソフト種類 LIKE '2.%' AND 業種エディション LIKE '%小売%'          THEN 割振工数 END),0) AS [21小売]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN プログラム名 LIKE '%EC%'                                   THEN 割振工数 END),0) AS [EC]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN プログラム名 LIKE '%CM%' OR プログラム名 LIKE '%CROSS%'    THEN 割振工数 END),0) AS [CM]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN プログラム名 LIKE '%取込%'                                 THEN 割振工数 END),0) AS [取込]"
                L_SQL &= " 	    , ISNULL(SUM(CASE  WHEN プログラム名 LIKE '%ハンディ%' OR プログラム名 LIKE '%検品入力%' THEN 割振工数 END),0) AS [ハンディ]"
                L_SQL &= "   FROM T_プログラムマスタ PM"
                L_SQL &= "        INNER JOIN T_担当者マスタ TNM ON PM.開発担当者コード = TNM.担当者コード"
                L_SQL &= " 	      INNER JOIN T_案件マスタ AM ON PM.受注NO = AM.受注NO"
                L_SQL &= " 	      LEFT OUTER JOIN S開発品質管理.dbo.T_部署マスタ BMN on TNM.部署コード = BMN.部署コード"
                L_SQL &= "  WHERE TNM.部署コード > ''"
                L_SQL &= "    AND TNM.BP区分 = '0'"
                L_SQL &= "    AND TNM.部署コード LIKE '%PG%'"
                L_SQL &= "    AND PM.開発完了日 LIKE '" & T5_対象年月.Text & "%'"


                If T5_社員コード.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者コード LIKE '%" & Integer.Parse(T5_社員コード.Text.TrimEnd) & "%'"
                End If
                If T5_社員名.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者名 LIKE '%" & T5_社員名.Text.TrimEnd & "%'"
                End If
                If T5_BMNLIST.Text <> "" Then
                    If T5_BMNLIST.Text.TrimEnd = "関西" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                    ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                    Else
                        L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                    End If
                End If

                L_SQL &= "  GROUP BY "
                L_SQL &= "        PM.開発担当者コード"
                L_SQL &= " ORDER BY 総工数 DESC"

            Case "UPRO未入力"
                Try
                    'データベースを選択
                    Tancd_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                        & "Trusted_Connection = Yes;" _
                                        & "Initial Catalog=S開発アサイン管理;"
                    Tancd_cn.Open()
                Catch ex As Exception
                    MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                    Return ""
                End Try

                L_SQL = ""
                L_SQL &= " SELECT "
                L_SQL &= "    JBT.受注NO"
                L_SQL &= "  , MAX(JBT.得意先略称)AS 得意先略称"
                L_SQL &= "  , ISNULL(MAX(TNM.担当者名),'') AS 最終ｶｽﾀﾏｲｽﾞ者"
                L_SQL &= "   FROM OPENQUERY([IC01\SHARE],"
                L_SQL &= "        'SELECT MAX(受注伝票NO) AS 受注NO"
                L_SQL &= "              ,MAX(得意先略称) AS 得意先略称"
                L_SQL &= "          FROM 原価管理link.dbo.公開用受注ビュー "
                L_SQL &= "         WHERE UPRO_PGコード = 0 AND カスタマイズ料 > 0 AND 受注日 > ''" & T5_対象年月.Text & "%'' AND (部門名 LIKE ''%関西%'' OR 部門名 LIKE ''%名古屋%'')"
                L_SQL &= "         GROUP BY 得意先コード"
                L_SQL &= "         ORDER BY 受注NO '"
                L_SQL &= "         ) JBT"
                L_SQL &= "         INNER JOIN 基本案件情報トラン TRN ON JBT.受注NO = TRN.受注コード"
                L_SQL &= "         LEFT OUTER JOIN S開発品質管理.dbo.T_案件マスタ ANM ON JBT.受注NO = ANM.受注NO"
                L_SQL &= "         LEFT OUTER JOIN アサイン一覧案件トラン Sub ON TRN.案件NO = Sub.案件NO"
                L_SQL &= "         LEFT OUTER JOIN 担当者マスタ ATM ON Sub.CR > '' AND Sub.CR = ATM.略称"
                L_SQL &= " 		   LEFT OUTER JOIN S開発品質管理.dbo.T_担当者マスタ TNM ON TNM.担当者コード = ISNULL(ATM.担当者コード,ANM.開発主管者コード)"
                L_SQL &= " 		   LEFT OUTER JOIN S開発品質管理.dbo.T_部署マスタ BMN ON TNM.部署コード = BMN.部署コード"
                L_SQL &= "  WHERE 1 = 1 "
                If T5_社員コード.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者コード LIKE '%" & Integer.Parse(T5_社員コード.Text.TrimEnd) & "%'"
                End If
                If T5_社員名.Text <> "" Then
                    L_SQL &= "    AND TNM.担当者名 LIKE '%" & T5_社員名.Text.TrimEnd & "%'"
                End If
                If T5_BMNLIST.Text <> "" Then
                    If T5_BMNLIST.Text.TrimEnd = "関西" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
                    ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                        L_SQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
                    Else
                        L_SQL &= "         AND BMN.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                    End If
                End If

                L_SQL &= "  GROUP BY JBT.受注NO "
                L_SQL &= " ORDER BY 最終ｶｽﾀﾏｲｽﾞ者 DESC"

            Case Else
                If T5_チーム別表示.Checked = False Then

                    L_SQL &= "    SELECT 社員NO = RIGHT(('000000' + LTRIM(担当者コード)),6) "
                    L_SQL &= "          ,社員名 = 担当者名 "
                    L_SQL &= "      FROM S開発品質管理.dbo.T_担当者マスタ T   "
                    L_SQL &= "           LEFT OUTER JOIN S開発品質管理.dbo.T_部署マスタ B ON T.部署コード = B.部署コード"
                    L_SQL &= " 	   WHERE T.部署コード <> '' AND BP区分 ='0'"
                    If T5_社員コード.Text.ToString.TrimEnd <> "" Then
                        L_SQL &= "   AND T.担当者コード = '%" & T5_社員コード.Text.TrimEnd & "%' "
                    End If
                    If T5_BMNLIST.Text.ToString.TrimEnd <> "" Then
                        If T5_BMNLIST.Text.TrimEnd = "関西" Then
                            L_SQL &= "         AND B.拠点区分 IN('1','3') " & vbCrLf
                        ElseIf T5_BMNLIST.Text.TrimEnd = "首都圏" Then
                            L_SQL &= "         AND B.拠点区分 IN('2') " & vbCrLf
                        Else
                            L_SQL &= "         AND B.部署名 LIKE '%" & T5_BMNLIST.Text.TrimEnd & "%' " & vbCrLf
                        End If
                    End If
                    L_SQL &= "     ORDER BY 担当者コード"

                Else
                    L_SQL &= "SELECT * FROM S開発品質管理.dbo.T_部署マスタ WHERE 部署種別 = '2' AND 削除区分 = '0'"
                End If
        End Select

        Return L_SQL
    End Function

    Private Sub FN_ModifyData()
        Dim L_SQL As String = ""

        Select Case T5_TAIDT.Text
            Case "原価入力リスト"
                ''テーブルを編集可能にする
                'For Each col As DataColumn In Tancd_dsDATA.Tables(0).Columns
                '    col.ReadOnly = False
                'Next

                Dim taidt As DateTime = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 1)
                Dim sttdt As DateTime
                Dim enddt As DateTime
                If T5_15日締表示.Checked Then
                    sttdt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 16).AddMonths(-1)
                    enddt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 15)
                Else
                    sttdt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), 1)
                    enddt = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), sttdt.AddMonths(1).AddDays(-1).Day)
                End If
                Dim sttdays As Integer = DateTime.DaysInMonth(sttdt.Year, sttdt.Month)
                Dim enddays As Integer = DateTime.DaysInMonth(enddt.Year, enddt.Month)
                If Not T5_15日締表示.Checked Then
                    For i As Integer = 1 To enddays
                        Dim dt As DateTime = New DateTime(Integer.Parse(T5_対象年月.Text.ToString.Substring(0, 4)), Integer.Parse(T5_対象年月.Text.ToString.Substring(4, 2)), i)
                        If dt.DayOfWeek = DayOfWeek.Sunday Or dt.DayOfWeek = DayOfWeek.Saturday Then
                            TancdGrid.Columns(3 + i).DefaultCellStyle.BackColor = Color.DarkGray
                        End If
                    Next
                Else
                    For i As Integer = 0 To sttdays - 1
                        Dim dt As DateTime = New DateTime(sttdt.AddDays(i).Year, sttdt.AddDays(i).Month, sttdt.AddDays(i).Day)
                        If dt.DayOfWeek = DayOfWeek.Sunday Or dt.DayOfWeek = DayOfWeek.Saturday Then
                            TancdGrid.Columns(4 + i).DefaultCellStyle.BackColor = Color.DarkGray
                        End If
                    Next
                End If

                'Try
                '    'データベースを選択
                '    TancdSub_cn.ConnectionString = "Data Source=ic01\share;" _
                '                        & "Trusted_Connection = Yes;" _
                '                        & "Initial Catalog=原価管理link;"
                '    TancdSub_cn.Open()
                'Catch
                '    MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
                '    Return
                'End Try

                'For Each row As DataGridViewRow In TancdGrid.Rows
                '    L_SQL = " SELECT 作業日付=REPLACE(作業日付,'/',''),作業時間=FORMAT(SUM(作業時間),'#0.0') "
                '    L_SQL &= "  FROM 原価トランビュー TRN "
                '    L_SQL &= " WHERE TRN.担当者コード = '" & row.Cells("社員NO").Value.ToString().PadLeft(6, "0"c) & "' "
                '    If T5_15日締表示.Checked Then
                '        L_SQL &= "   AND TRN.作業日付 BETWEEN '" & sttdt.ToString("yyyy/MM/dd").TrimEnd() & "' AND '" & enddt.ToString("yyyy/MM/dd").TrimEnd() & "'"
                '    Else
                '        L_SQL &= "   AND TRN.作業日付 LIKE '%" & taidt.ToString("yyyy/MM").TrimEnd() & "%' "
                '    End If
                '    L_SQL &= " GROUP BY 作業日付 "

                '    Dim daAuthors As New SqlDataAdapter(L_SQL, TancdSub_cn)
                '    daAuthors.FillSchema(TancdSub_dsDATA, SchemaType.Source)
                '    TancdSub_dsDATA.Tables("Table").PrimaryKey = Nothing
                '    TancdSub_dsDATA.Tables(0).Clear()
                '    daAuthors.Fill(TancdSub_dsDATA)
                '    If TancdSub_dsDATA.Tables(0) IsNot Nothing AndAlso TancdSub_dsDATA.Tables(0).Rows.Count > 0 Then
                '        For Each subrow As DataRow In TancdSub_dsDATA.Tables(0).Rows
                '            If TancdGrid.Columns.Contains("日付" & subrow("作業日付").ToString) Then
                '                Dim val As Decimal = Decimal.Parse(subrow("作業時間").ToString)
                '                Tancd_dsDATA.Tables(0).Rows(row.Index).Item(TancdGrid.Columns("日付" & subrow("作業日付").ToString).Index) = subrow("作業時間").ToString
                '            End If
                '        Next
                '    End If
                'Next

                ''累計計算
                'For Each row As DataGridViewRow In TancdGrid.Rows
                '    Dim sum As Decimal = 0D
                '    Dim zan As Decimal = 0D
                '    For columnindex As Integer = 3 To TancdGrid.ColumnCount - 1
                '        If IsNumeric(row.Cells(columnindex).Value) Then
                '            sum += row.Cells(columnindex).Value
                '            If row.Cells(columnindex).Value > 7.5 Then
                '                zan += (row.Cells(columnindex).Value - 7.5)
                '            End If
                '        End If
                '    Next
                '    row.Cells("累計").Value = sum
                '    row.Cells("残業").Value = zan
                'Next

                'TancdGrid.Sort(TancdGrid.Columns("累計"), System.ComponentModel.ListSortDirection.Descending)
        End Select
        Return
    End Sub

    ''' <summary>
    ''' 色設定
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TancdGrid_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles TancdGrid.CellFormatting
        If T5_TAIDT.Text = "原価入力リスト" Then
            If e.ColumnIndex > 3 Then
                If e.Value.ToString() <> "" Then
                    Dim val As Decimal = Decimal.Parse(e.Value)
                    Dim col As Color = Color.Gainsboro
                    If val <= 7.5 Then
                        col = Color.FromArgb(247, 230, 210)
                    ElseIf val <= 8.5 Then
                        col = Color.FromArgb(247, 220, 160)
                    ElseIf val <= 9.5 Then
                        col = Color.FromArgb(247, 210, 110)
                    ElseIf val <= 10.5 Then
                        col = Color.FromArgb(247, 200, 60)
                    ElseIf val <= 11.5 Then
                        col = Color.FromArgb(247, 190, 0)
                    ElseIf val <= 12.5 Then
                        col = Color.FromArgb(255, 100, 100)
                    Else
                        col = Color.FromArgb(255, 45, 45)
                    End If
                    e.CellStyle.BackColor = col
                End If
            End If
        End If

        If T5_TAIDT.Text = "仕様書提出リスト" Then
            If IsNumeric(TancdGrid.Columns(e.ColumnIndex).Name) AndAlso (Today.ToString("yyyyMMdd") = (T5_対象年月.Text.TrimEnd() & TancdGrid.Columns(e.ColumnIndex).Name.ToString())) Then
                e.CellStyle.BackColor = Color.Red

            End If
        End If
    End Sub

    ''' <summary>
    ''' カラム設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FN_CellSettingTancd()

        For i As Integer = 0 To TancdGrid.Columns.Count - 1

            If T5_TAIDT.Text = "原価入力リスト" AndAlso i > 3 Then
                With TancdGrid.Columns(i)
                    .Width = 40
                    .DefaultCellStyle.Format = "#,0.0"
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    Continue For
                End With
            End If

            If T5_TAIDT.Text = "PG経験" Then
                If i > 1 Then
                    With TancdGrid.Columns(i)
                        .Width = 60
                        .DefaultCellStyle.Format = "#,#0.#"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    End With
                    Continue For
                End If
            End If

            With TancdGrid.Columns(i)
                If T5_TAIDT.Text = "生産高確認" AndAlso (.Name.Contains("月") OrElse .Name.Contains("累計")) Then
                    .Width = 70
                    .DefaultCellStyle.Format = "#,0"
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                ElseIf .Name.Contains("月") OrElse .Name.Contains("累計") OrElse .Name.Contains("残業") OrElse .Name.Contains("工数") Then
                    .Width = 50
                    .DefaultCellStyle.Format = "#,0.0"
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                ElseIf .Name.Contains("社員名") OrElse .Name.Contains("PG担当") OrElse .Name.Contains("SE担当") Then
                    .Width = 70
                    .Frozen = True
                ElseIf .Name.EndsWith("得意先名") OrElse .Name.EndsWith("得意先略称") OrElse .Name.Contains("最終ｶｽﾀﾏｲｽﾞ者") Then
                    .Width = 130
                ElseIf .Name.Contains("日付") Then
                    .Width = 40
                ElseIf .Name.Equals("名") Then
                    .Frozen = True
                ElseIf .Name.Contains("部署名") Then
                    .AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                    .Width = 180
                    .Frozen = True
                ElseIf .Name.Contains("四半期") Then
                    .Width = 100
                    .DefaultCellStyle.Format = "#,0"
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                ElseIf .Name.Contains("順位") Then
                    .Width = 35
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                ElseIf .Name.Contains("累計") Then
                    .Width = 60
                    .DefaultCellStyle.Format = "#,0"
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                ElseIf .Name.Contains("残業") Then
                    .Width = 60
                    .DefaultCellStyle.Format = "#,0"
                    .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                ElseIf .Name.Contains("社員NO") Then
                    .Width = 60
                ElseIf IsNumeric(.Name) Then
                    .Width = 20
                Else
                    .Width = 70
                End If
            End With
        Next
    End Sub

    ''' <summary>
    ''' サブ画面を開く(開く画面は設定による))
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenTancdSubform()
        If TancdGrid.CurrentRow Is Nothing Then Return
        If SubFormCanOpen = False Then Return

        Dim SubForm As New SubdataForm()
        If AssignGrid.SelectedCells.Count > 0 Then
            SubForm.posx = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.X + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Width / 2 + Me.Location.X + 200
            SubForm.posy = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.Y + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Height / 2 + Me.Location.Y + 109
        End If
        SubForm.Show()
    End Sub

    ''' <summary>
    ''' サブ画面を開く(開く画面は設定による))
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenTancdSubform_BMN()
        If TancdGrid.CurrentRow Is Nothing Then Return
        If SubFormCanOpen = False Then Return

        Dim SubForm As New SubdataForm()
        If AssignGrid.SelectedCells.Count > 0 Then
            SubForm.posx = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.X + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Width / 2 + Me.Location.X + 200
            SubForm.posy = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.Y + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Height / 2 + Me.Location.Y + 109
        End If
        SubForm.Show()
    End Sub


    ''' <summary>
    ''' 内訳サブ画面を開く
    ''' </summary>
    Private Sub TancdGrid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles TancdGrid.CellDoubleClick
        ' ヘッダ以外のセルか？
        If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
            '担当者毎に表示時
            If TancdGrid.Columns.Contains("担当者") Then
                '月毎表示時
                If TancdGrid.Columns(e.ColumnIndex).Name.Contains("月") Then
                    Tancd = TancdGrid.Rows(e.RowIndex).Cells("担当者").Value.ToString.TrimEnd
                    Taiym = T5_対象年月.Text.Substring(0, 4)
                    Taidt = TancdGrid.Columns(e.ColumnIndex).Name.Substring(0, 2)
                    Bmnnm = ""
                    Subkb = "生産高_月"
                    OpenTancdSubform()
                End If

                '四半期表示時
                If TancdGrid.Columns(e.ColumnIndex).Name.Contains("四半期") Then
                    Tancd = TancdGrid.Rows(e.RowIndex).Cells("担当者").Value.ToString.TrimEnd
                    Taiym = T5_対象年月.Text.Substring(0, 4)
                    Taidt = TancdGrid.Columns(e.ColumnIndex).Name.Substring(0, 2)
                    Bmnnm = ""
                    Subkb = "生産高_四半期"
                    OpenTancdSubform()
                End If
                'Else
                '    Subkb = "負荷状況"
                '    OpenTancdGridSubform()
            End If
        End If
    End Sub

    ''' <summary>
    ''' サブ画面を開く(開く画面は設定による))
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenTancdGridSubform()
        If TancdGrid.CurrentRow Is Nothing Then Return
        If SubFormCanOpen = False Then Return

        Dim SubForm As New SubdataForm()
        If AssignGrid.SelectedCells.Count > 0 Then
            SubForm.posx = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.X + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Width / 2 + Me.Location.X + 200
            SubForm.posy = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.Y + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Height / 2 + Me.Location.Y + 109
        End If
        SubForm.Show()
    End Sub

    Private Sub T5_TMCHK_CheckedChanged(sender As Object, e As EventArgs) Handles T5_チーム別表示.CheckedChanged
        If T5_チーム別表示.Checked Then
            T5_社員コード.Enabled = False
            T5_TANCD_LBL.Enabled = False
        Else
            T5_社員コード.Enabled = True
            T5_TANCD_LBL.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' 対象変更時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub T5_TAIDT_SelectedValueChanged(sender As Object, e As EventArgs) Handles T5_TAIDT.SelectedValueChanged

        T5_チーム別表示.Visible = False
        T5_SE渡しを含む.Visible = False
        T5_四半期表示.Visible = False
        T5_BP実績を含む.Visible = False
        T5_予定を含む.Visible = False
        T5_15日締表示.Visible = False
        T5_未入力時間を表示.Visible = False
        T5_ラベル説明.Visible = False
        T5_UPRO未入力説明.Visible = False

        If (DateTime.Today.Month < 8) Then
            T5_対象年月.Text = DateTime.Today().AddYears(-1).ToString("yyyy")
        Else
            T5_対象年月.Text = DateTime.Today().ToString("yyyy")
        End If


        Select Case T5_TAIDT.Text.TrimEnd
            Case "生産高確認"
                T5_チーム別表示.Visible = True
                T5_SE渡しを含む.Visible = True
                T5_四半期表示.Visible = True
                T5_BP実績を含む.Visible = True
                T5_予定を含む.Visible = True

            Case "月別割振工数"
                T5_チーム別表示.Visible = True

            Case "原価入力リスト"
                If Today.Day > 15 Then
                    T5_対象年月.Text = DateTime.Today.AddMonths(1).ToString("yyyyMM")
                Else
                    T5_対象年月.Text = DateTime.Today.ToString("yyyyMM")
                End If
                T5_15日締表示.Visible = True
                T5_未入力時間を表示.Visible = True

            Case "仕様書提出リスト"
                T5_ラベル説明.Visible = True
                T5_対象年月.Text = DateTime.Today().ToString("yyyyMM")
            Case "UPRO未入力"
                T5_対象年月.Text = DateTime.Today().AddYears(-3).ToString("yyyy")
                T5_UPRO未入力説明.Visible = True
        End Select
    End Sub
End Class

