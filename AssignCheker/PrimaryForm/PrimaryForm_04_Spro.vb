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
Imports System.Collections.Specialized

Partial Class PrimaryForm
    ''' <summary>
    ''' アサイン取得メイン処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateCostDataView()

        'メッセージ初期化
        StatusStrip1.Text = ""

        '常にDBに接続済の時は接続を切る 
        Cost_cmd.Dispose()
        Cost_cn.Close()
        Cost_cn.Dispose()
        L_CurrentX = Nothing
        L_CurrentY = Nothing
        CostGrid.ClearSelection()
        For i As Integer = 0 To Cost_dsDATA.Tables.Count - 1
            Cost_dsDATA.Tables(i).DefaultView.Sort = String.Empty
            Cost_dsDATA.Tables(i).Clear()
            Cost_dsDATA.Tables(i).Constraints.Clear()
            For j As Integer = Cost_dsDATA.Tables(i).Columns.Count - 1 To 0 Step -1
                Cost_dsDATA.Tables(i).Columns.RemoveAt(j)
            Next
            CostGrid.Columns.Clear()
            If i = Cost_dsDATA.Tables.Count - 1 Then
                Cost_dsDATA.Tables.Clear()
                CostGrid.DataSource = Nothing
            End If
        Next

        Try
            'データベースを選択
            Cost_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                & "Trusted_Connection = Yes;" _
                                & "Initial Catalog=S開発アサイン管理;"
            Cost_cn.Open()
        Catch ex As Exception
            MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
            Return
        End Try

        Dim L_Query As String = ""
        L_Query = FN_CreateCostQuery()

        Try
            'Columnのサイズは固定にしてから列を設定
            CostGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'DB接続
            Dim daAuthors As New SqlDataAdapter(L_Query, Cost_cn)
            daAuthors.FillSchema(Cost_dsDATA, SchemaType.Source)
            Cost_dsDATA.Tables("Table").PrimaryKey = Nothing
            daAuthors.Fill(Cost_dsDATA)
            CostGrid.DataMember = Cost_dsDATA.Tables("table").TableName
            CostGrid.DataSource = Cost_dsDATA.Tables(0)

            'セルの初期設定を行います。
            FN_CostCellSetting()

            SubFormCanOpen = True
        Catch ex As Exception
            Microsoft.VisualBasic.MsgBox(ex.Message, MsgBoxStyle.OkCancel, "クエリエラー")
        End Try
    End Sub

    ''' <summary>
    ''' アサイン情報取得クエリ作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FN_CreateCostQuery()
        Dim L_SQL As String = ""
        Dim L_ResultSQL As String = ""

        'SELECT句                                                                                                                                                        
        'L_ResultSQL = " "
        'L_ResultSQL &= " SELECT TOP " & My.Settings("MaxView").ToString.TrimEnd & vbCrLf
        'L_ResultSQL &= "       受注NO       = JBT.受注伝票NO " & vbCrLf
        'L_ResultSQL &= " 	  ,得意先コード = JBT.得意先コード " & vbCrLf
        'L_ResultSQL &= " 	  ,得意先名     = JBT.得意先略称 " & vbCrLf
        'L_ResultSQL &= " 	  ,進捗率       = ISNULL((SELECT MAX(進捗) FROM [IC01\SHARE].原価管理link.dbo.原価トランビュー GNK WHERE GNK.受注NO = JBT.受注伝票NO AND GNK.対応内訳コード IN ('U003','U00303') AND GNK.集計コード = '110'),0) " & vbCrLf
        'L_ResultSQL &= " 	  ,SE担当       = JBT.SPRO_設計名 " & vbCrLf
        'L_ResultSQL &= " 	  ,PG担当       = JBT.SPRO_PG名 " & vbCrLf
        'L_ResultSQL &= "      ,PG部署       = BMN.部署名 " & vbCrLf
        'L_ResultSQL &= " 	  ,検収月       = CONVERT(CHAR(6),JBT.売上予定日) " & vbCrLf
        'L_ResultSQL &= "      ,UPRO担当     = JBT.UPRO_PG名 " & vbCrLf
        'L_ResultSQL &= "  FROM [IC01\SHARE].原価管理link.dbo.公開用受注ビュー JBT" & vbCrLf
        'L_ResultSQL &= "       LEFT OUTER JOIN T_案件マスタ MAIN ON JBT.受注伝票NO = MAIN.受注NO" & vbCrLf
        'L_ResultSQL &= "       INNER JOIN T_担当者マスタ TNM  ON TNM.担当者コード  = MAIN.開発主管者コード" & vbCrLf
        'L_ResultSQL &= "       INNER JOIN T_部署マスタ BMN ON TNM.部署コード = BMN.部署コード" & vbCrLf

        'If C_WHERE_B.Text.TrimEnd <> "" Then
        '    If C_WHERE_B.Text.TrimEnd = "関西" Then
        '        L_ResultSQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
        '    ElseIf C_WHERE_B.Text.TrimEnd = "首都圏" Then
        '        L_ResultSQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
        '    Else
        '        L_ResultSQL &= "         AND BMN.部署名 LIKE '%" & C_WHERE_B.Text.TrimEnd & "%' " & vbCrLf
        '    End If
        'End If


        'L_ResultSQL &= "  WHERE 1=1" & vbCrLf
        'If C_WHERE01.Text.TrimEnd <> "" Then
        '    L_ResultSQL &= "    AND JBT.受注伝票NO " & C_OP01.Text & " '" & C_WHERE01.Text.TrimEnd & "'" & vbCrLf
        'End If
        'If C_WHERE02.Text.TrimEnd <> "" Then
        '    If C_OP02.Text.TrimEnd = "=" Then
        '        L_ResultSQL &= "    AND JBT.得意先略称 LIKE '%" & C_WHERE02.Text.TrimEnd & "%'" & vbCrLf
        '    Else
        '        L_ResultSQL &= "    AND JBT.得意先略称 " & C_OP02.Text.TrimEnd & " '" & C_WHERE02.Text.TrimEnd & "'" & vbCrLf
        '    End If
        'End If
        'If C_WHERE03.Text.TrimEnd <> "" Then
        '    If C_OP03.Text.TrimEnd = "=" Then
        '        L_ResultSQL &= "    AND JBT.売上予定日 LIKE '%" & C_WHERE03.Text.TrimEnd & "%'" & vbCrLf
        '    Else
        '        L_ResultSQL &= "    AND JBT.売上予定日 " & C_OP03.Text.TrimEnd & " '" & C_WHERE03.Text.TrimEnd & "'" & vbCrLf
        '    End If
        'End If
        'If C_WHERE04.Text.TrimEnd <> "" Then
        '    If C_OP04.Text.TrimEnd = "=" Then
        '        L_ResultSQL &= "    AND JBT.SPRO_設計名 LIKE '%" & C_WHERE04.Text.TrimEnd & "%'" & vbCrLf
        '    Else
        '        L_ResultSQL &= "    AND JBT.SPRO_設計名 " & C_OP04.Text.TrimEnd & " '" & C_WHERE04.Text.TrimEnd & "'" & vbCrLf
        '    End If
        'End If
        'If C_WHERE05.Text.TrimEnd <> "" Then
        '    If C_OP05.Text.TrimEnd = "=" Then
        '        L_ResultSQL &= "    AND JBT.SPRO_PG名 LIKE '%" & C_WHERE05.Text.TrimEnd & "%'" & vbCrLf
        '    Else
        '        L_ResultSQL &= "    AND JBT.SPRO_PG名 " & C_OP04.Text.TrimEnd & " '" & C_WHERE05.Text.TrimEnd & "'" & vbCrLf
        '    End If
        'End If


        L_ResultSQL &= " SELECT MAIN.受注NO" & vbCrLf
        L_ResultSQL &= "      , 得意先コード     = MAX(MAIN.得意先コード)" & vbCrLf
        L_ResultSQL &= "      , 得意先名         = MAX(MAIN.得意先略称)" & vbCrLf
        L_ResultSQL &= " 	  , 検収月           = MAX(MAIN.検収月)" & vbCrLf
        L_ResultSQL &= " 	  , 進捗率           = MAX(進捗率)" & vbCrLf
        L_ResultSQL &= "      , PG粗利           = ISNULL((SUM(TRN.PG依頼工数) * 40000 - MAX(MAIN.原価)) / NULLIF((SUM(TRN.PG依頼工数) * 40000),0),0)" & vbCrLf
        L_ResultSQL &= " 	  , SE担当           = MAX(MAIN.SE担当)" & vbCrLf
        L_ResultSQL &= " 	  , PG担当           = MAX(MAIN.PG担当)" & vbCrLf
        L_ResultSQL &= " 	  , UPRO担当         = MAX(MAIN.UPRO担当)" & vbCrLf
        L_ResultSQL &= " 	  , PG工数           = SUM(TRN.PG依頼工数)" & vbCrLf
        L_ResultSQL &= " 	  , PG原価           = MAX(MAIN.原価)" & vbCrLf
        L_ResultSQL &= "   FROM (SELECT * FROM openquery([IC01\SHARE],'SELECT 受注NO           = GNK.受注NO" & vbCrLf
        L_ResultSQL &= "              , 得意先コード     = MAX(JBT.得意先コード)" & vbCrLf
        L_ResultSQL &= "              , 得意先略称       = MAX(JBT.得意先略称)" & vbCrLf
        L_ResultSQL &= " 			  , 検収月           = MAX(JBT.売上予定日)" & vbCrLf
        L_ResultSQL &= "         	  , 原価             = CONVERT(INT,SUM(CASE WHEN 内容コード = ''015'' OR 内容コード = ''025'' OR 内容コード = ''131'' THEN 0 WHEN 集計コード = ''110'' THEN 実原価金額 ELSE 原価金額 END))" & vbCrLf
        L_ResultSQL &= " 			  , 進捗率           = MAX(GNK.進捗)" & vbCrLf
        L_ResultSQL &= " 			  , SE担当           = MAX(JBT.SPRO_設計名)" & vbCrLf
        L_ResultSQL &= " 			  , PG担当           = MAX(JBT.SPRO_PG名)" & vbCrLf
        L_ResultSQL &= " 			  , UPRO担当         = MAX(JBT.UPRO_PG名)" & vbCrLf
        L_ResultSQL &= "           FROM 原価管理link.dbo.原価トランビュー GNK" & vbCrLf
        L_ResultSQL &= "         	   INNER JOIN 原価管理link.dbo.公開用受注ビュー JBT ON GNK.受注NO = JBT.受注伝票NO" & vbCrLf
        L_ResultSQL &= "          WHERE 1 = 1 " & vbCrLf

        If C_WHERE03.Text.TrimEnd <> "" Then
            If C_OP03.Text.TrimEnd = "=" Then
                L_ResultSQL &= "    AND JBT.売上予定日 LIKE ''%" & C_WHERE03.Text.TrimEnd & "%''" & vbCrLf
            Else
                L_ResultSQL &= "    AND JBT.売上予定日 " & C_OP03.Text.TrimEnd & " ''" & C_WHERE03.Text.TrimEnd & "''" & vbCrLf
            End If
        End If
        If C_WHERE01.Text.TrimEnd <> "" Then
            L_ResultSQL &= "    AND JBT.受注伝票NO " & C_OP01.Text & " ''" & C_WHERE01.Text.TrimEnd & "''" & vbCrLf
        End If
        If C_WHERE02.Text.TrimEnd <> "" Then
            If C_OP02.Text.TrimEnd = "=" Then
                L_ResultSQL &= "    AND JBT.得意先略称 LIKE ''%" & C_WHERE02.Text.TrimEnd & "%''" & vbCrLf
            Else
                L_ResultSQL &= "    AND JBT.得意先略称 " & C_OP02.Text.TrimEnd & " ''" & C_WHERE02.Text.TrimEnd & "''" & vbCrLf
            End If
        End If
        If C_WHERE04.Text.TrimEnd <> "" Then
            If C_OP04.Text.TrimEnd = "=" Then
                L_ResultSQL &= "    AND JBT.SPRO_設計名 LIKE ''%" & C_WHERE04.Text.TrimEnd & "%''" & vbCrLf
            Else
                L_ResultSQL &= "    AND JBT.SPRO_設計名 " & C_OP04.Text.TrimEnd & " ''" & C_WHERE04.Text.TrimEnd & "''" & vbCrLf
            End If
        End If
        If C_WHERE05.Text.TrimEnd <> "" Then
            If C_OP05.Text.TrimEnd = "=" Then
                L_ResultSQL &= "    AND JBT.SPRO_PG名 LIKE ''%" & C_WHERE05.Text.TrimEnd & "%''" & vbCrLf
            Else
                L_ResultSQL &= "    AND JBT.SPRO_PG名 " & C_OP04.Text.TrimEnd & " ''" & C_WHERE05.Text.TrimEnd & "''" & vbCrLf
            End If
        End If
        L_ResultSQL &= "            AND GNK.対応内訳コード IN(''U003'',''U00303'') " & vbCrLf
        L_ResultSQL &= "            AND GNK.集計コード IN (''110'',''59'')" & vbCrLf
        L_ResultSQL &= "         GROUP BY GNK.受注NO "

        If C_WHERE06.Text.TrimEnd <> "" And C_WHERE07.Text <> "" Then
            L_ResultSQL &= "    HAVING MAX(GNK.進捗) BETWEEN ''" & C_WHERE06.Text.TrimEnd & "'' AND ''" & C_WHERE07.Text.TrimEnd & "'' " & vbCrLf
        ElseIf C_WHERE06.Text <> "" Then
            L_ResultSQL &= "    HAVING MAX(GNK.進捗) > ''" & C_WHERE06.Text.TrimEnd & "'' " & vbCrLf
        ElseIf C_WHERE07.Text <> "" Then
            L_ResultSQL &= "    HAVING MAX(GNK.進捗) < ''" & C_WHERE07.Text.TrimEnd & "'' " & vbCrLf
        End If

        L_ResultSQL &= " ')) MAIN"
        L_ResultSQL &= " INNER JOIN 基本案件情報トラン TRN ON TRN.受注コード = MAIN.受注NO"
        L_ResultSQL &= " INNER JOIN [S開発品質管理].dbo.T_案件マスタ ANM ON TRN.受注コード = ANM.受注NO"
        L_ResultSQL &= " INNER JOIN [S開発品質管理].dbo.T_担当者マスタ TNM ON TNM.担当者コード = ANM.開発主管者コード"
        L_ResultSQL &= " INNER JOIN [S開発品質管理].dbo.T_部署マスタ BMN ON TNM.部署コード = BMN.部署コード"
        L_ResultSQL &= " WHERE 1 = 1 "
        If C_WHERE_B.Text.TrimEnd <> "" Then
            If C_WHERE_B.Text.TrimEnd = "関西" Then
                L_ResultSQL &= "         AND BMN.拠点区分 IN('1','3') " & vbCrLf
            ElseIf C_WHERE_B.Text.TrimEnd = "首都圏" Then
                L_ResultSQL &= "         AND BMN.拠点区分 IN('2') " & vbCrLf
            Else
                L_ResultSQL &= "         AND BMN.部署名 LIKE '%" & C_WHERE_B.Text.TrimEnd & "%' " & vbCrLf
            End If
        End If
        L_ResultSQL &= " GROUP BY MAIN.受注NO"


        L_ResultSQL &= " ORDER BY 進捗率,PG担当" & vbCrLf
        Return L_ResultSQL
    End Function

    ''' <summary>
    ''' [原価情報]カラム設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FN_CostCellSetting()
        'Return

        For Each col As DataGridViewColumn In CostGrid.Columns
            'For i As Integer = 1 To 26
            With col
                Select Case .Name
                    Case "受注NO"
                        .Width = 70
                    Case "得意先コード"
                        .Width = 90
                    Case "得意先名"
                        .Width = 150
                    Case "エディション"
                    Case "PG原価"
                        .DefaultCellStyle.Format = "###,#0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case "PG工数"
                        .DefaultCellStyle.Format = "###,#0.0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case "進捗率"
                        .DefaultCellStyle.Format = "###,#0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case "PG粗利"
                        .DefaultCellStyle.Format = "###,#0.0%"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case "検収月"
                        .DefaultCellStyle.Format = "####/##"
                        .Width = 70
                    Case "SE担当"
                        .Width = 80
                    Case "PG担当"
                        .Width = 80
                    Case "PG部署"
                        .Width = 160
                    Case "UPRO担当"
                        .Width = 80
                End Select
            End With
        Next
    End Sub

    ''' <summary>
    ''' 原価情報カーソル移動（KeyDown）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CostTab_Enter(sender As Object, e As EventArgs) Handles C_WHERE01.KeyDown, C_WHERE02.KeyDown, C_WHERE03.KeyDown, C_WHERE04.KeyDown, C_WHERE05.KeyDown, C_WHERE06.KeyDown, C_WHERE07.KeyDown
        'エンターか↓キーで移動
        Select Case DirectCast(e, System.Windows.Forms.KeyEventArgs).KeyValue
            Case Keys.Enter
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 1, 1))
                    If CostTab_Panel.Controls("C_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)) Is Nothing Then
                        Index = 0
                    End If
                    CostTab_Panel.Controls("C_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)).Select()
                End If
            Case Keys.Down
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 1, 1))
                    If CostTab_Panel.Controls("C_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)) Is Nothing Then
                        Index = 0
                    End If
                    CostTab_Panel.Controls("C_WHERE" & (Index + 1).ToString.PadLeft(2, "0"c)).Select()
                End If
            Case Keys.Up
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 1, 1))
                    If CostTab_Panel.Controls("C_WHERE" & (Index - 1).ToString.PadLeft(2, "0"c)) IsNot Nothing Then
                        CostTab_Panel.Controls("C_WHERE" & (Index - 1).ToString.PadLeft(2, "0"c)).Select()
                    End If
                End If
        End Select
    End Sub

#Region "  右クリックメニュー "

    ''' <summary>
    ''' [右クリック] コピー
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 値をコピー_Click(sender As Object, e As EventArgs) Handles Cost_値をコピー.Click
        '選択されたセルをクリップボードにコピーする
        Clipboard.SetDataObject(CostGrid.GetClipboardContent())
    End Sub

    ''' <summary>
    ''' [右クリック] 値をヘッダ付きでコピーする
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 値をヘッダ付きでコピー_Click(sender As Object, e As EventArgs) Handles Cost_ヘッダ付きコピー.Click
        '選択されたセルをクリップボードにコピーする
        CostGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Clipboard.SetDataObject(CostGrid.GetClipboardContent())
        CostGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Subkb = "PG一覧"
        OpenCostGridSubform()
    End Sub
    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        Subkb = "工数"
        OpenCostGridSubform()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Subkb = "対応"
        OpenCostGridSubform()
    End Sub
    Private Sub 原価内訳ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 原価内訳ToolStripMenuItem.Click
        Subkb = "原価内訳"
        OpenCostGridSubform()
    End Sub
#End Region

#Region " [アサイン情報グリッド]"
    ''' <summary>
    ''' [CellMouseDown]クリック時にセルを選択状態にする
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CostGridView_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles CostGrid.CellMouseDown
        ' 右ボタンのクリックか？
        If e.Button = MouseButtons.Left Then
            'Me.SuspendLayout()
            ' ヘッダ以外のセルか？
            If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
                CostGrid.ClearSelection()
                Dim cell As DataGridViewCell = CostGrid(e.ColumnIndex, e.RowIndex)
                ' セルの選択状態を反転
                cell.Selected = True
            End If

            If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
                Dim dgv As DataGridView = CType(sender, DataGridView)
                If L_CurrentX <> Nothing Then
                    dgv(L_CurrentX, L_CurrentY).Style.BackColor = Color.Empty
                    dgv(L_CurrentX, L_CurrentY).Style.SelectionBackColor = Color.Empty
                End If

                L_CurrentX = e.ColumnIndex
                L_CurrentY = e.RowIndex
                Dim cell As DataGridViewCell = CostGrid(e.ColumnIndex, e.RowIndex)
                cell.Selected = True
            End If
            'Me.ResumeLayout()
        End If
    End Sub

    ''' <summary>
    ''' [CellMouseUp]右クリック時の処理
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub CostGridView_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles CostGrid.CellMouseUp

        If e.Button = MouseButtons.Right Then
            'Me.SuspendLayout()
            ' ヘッダ以外のセルか？
            If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
                If CostGrid.SelectedCells.Count <= 1 Then
                    CostGrid.ClearSelection()
                    ' 右クリックされたセル
                    Dim cell As DataGridViewCell = CostGrid(e.ColumnIndex, e.RowIndex)
                    ' セルの選択状態を反転
                    cell.Selected = True
                End If
            End If

            If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
                Dim dgv As DataGridView = CType(sender, DataGridView)
                If L_CurrentX <> Nothing Then
                    dgv(L_CurrentX, L_CurrentY).Style.BackColor = Color.Empty
                    dgv(L_CurrentX, L_CurrentY).Style.SelectionBackColor = Color.Empty
                End If
                L_CurrentX = e.ColumnIndex
                L_CurrentY = e.RowIndex
                Dim cell As DataGridViewCell = CostGrid(e.ColumnIndex, e.RowIndex)
                cell.Selected = True
            End If
            'Me.ResumeLayout()
        End If
    End Sub

    ''' <summary>
    ''' 内訳サブ画面を開く
    ''' </summary>
    Private Sub CosteGrid_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles CostGrid.CellDoubleClick
        ' ヘッダ以外のセルか？
        If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
            Subkb = "原価内訳"
            OpenCostGridSubform()
        End If

    End Sub

    ''' <summary>
    ''' サブ画面を開く(開く画面は設定による))
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenCostGridSubform()
        If CostGrid.CurrentRow Is Nothing Then Return
        If SubFormCanOpen = False Then Return
        Dim L_JdnColumnNM As String = "受注NO"
        Dim SubForm As New SubdataForm()
        If CostGrid.SelectedCells.Count > 0 Then
            Jdnno = CostGrid.Rows(CostGrid.SelectedCells(0).RowIndex).Cells(L_JdnColumnNM).Value
            SubForm.posx = CostGrid.GetCellDisplayRectangle(CostGrid.SelectedCells(0).ColumnIndex, CostGrid.SelectedCells(0).RowIndex, True).Location.X + CostGrid.GetCellDisplayRectangle(CostGrid.SelectedCells(0).ColumnIndex, CostGrid.SelectedCells(0).RowIndex, True).Width / 2 + Me.Location.X + 200
            SubForm.posy = CostGrid.GetCellDisplayRectangle(CostGrid.SelectedCells(0).ColumnIndex, CostGrid.SelectedCells(0).RowIndex, True).Location.Y + CostGrid.GetCellDisplayRectangle(CostGrid.SelectedCells(0).ColumnIndex, CostGrid.SelectedCells(0).RowIndex, True).Height / 2 + Me.Location.Y + 109
        End If
        SubForm.Show()
    End Sub
#End Region
End Class
