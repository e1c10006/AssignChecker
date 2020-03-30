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
            ''' 進捗表更新実行
            ''' </summary>
    Private Sub 進捗表を更新ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 進捗表を更新ToolStripMenuItem.Click

        Dim IWorkbook

        If MetroFramework.MetroMessageBox.Show(Me, "進捗表を更新しますか？", "確認", MessageBoxButtons.YesNo, Nothing, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            Return
        End If

        Dim updlist As List(Of String) = New List(Of String)
        Dim val As String = ""
        For Each cell As DataGridViewCell In AssignGrid.SelectedCells
            val = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("進捗表パス").ToString.TrimEnd).Value.ToString.TrimEnd
            If val <> "" AndAlso Not updlist.Contains(val) Then
                updlist.Add(val)
            End If
        Next


        Dim UpdCnt As Integer = 0
        Dim MaxCnt As Integer = 0

        '対象件数取得
        MaxCnt = updlist.Count

        '対象無し
        If MaxCnt = 0 Then
            MetroFramework.MetroMessageBox.Show(Me, "更新対象が選択されていません")
            Return
        End If

        'MetroProgressBar1.Value = 0
        For Each Item As String In updlist
            'プログレスバー更新
            UpdCnt += 1
            'MetroProgressBar1.Value = CInt((CDec(UpdCnt) / CDec(MaxCnt) * 100))
            Me.Refresh()
            Try

                If Not File.Exists(Item) Then
                    MetroFramework.MetroMessageBox.Show(Me, "進捗表が存在しません。" & Item)
                End If

                StatusStrip1.Text = "[" & Item & "]を更新中です..."

                Dim p = New ProcessStartInfo()
                If My.Settings("Batch").ToString.TrimEnd = "" Then
                    p.FileName = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) & "\upd.vbs"
                Else
                    p.FileName = My.Settings("Batch").ToString.TrimEnd
                End If
                p.Arguments = Item
                p.UseShellExecute = True
                Dim Proc As Process = Process.Start(p)

                Proc.WaitForExit()
            Catch ex As Exception
                MetroFramework.MetroMessageBox.Show(Me, "既に進捗表を開かれている可能性があります。閉じた後に再度実行して下さい。" & ex.Message)
                Return
            Finally
                StatusStrip1.Text = ""
            End Try
        Next
        'MetroProgressBar1.Value = MetroProgressBar1.Maximum
        MetroFramework.MetroMessageBox.Show(Me, "進捗表の更新が完了しました！", "確認")
    End Sub


    Private Sub RegiterForm_Assign()
        H_WHERE0.ImeMode = ImeMode.Off
        H_WHERE1.ImeMode = ImeMode.Off
        H_WHERE2.ImeMode = ImeMode.On
        H_WHERE3.ImeMode = ImeMode.Off
        H_WHERE4.ImeMode = ImeMode.Off
        H_WHERE5.ImeMode = ImeMode.Off
        H_WHERE6.ImeMode = ImeMode.On
        H_WHERE7.ImeMode = ImeMode.On
        H_WHERE8.ImeMode = ImeMode.Off
        H_WHERE9.ImeMode = ImeMode.Off
        H_WHERE10.ImeMode = ImeMode.Off
    End Sub

    ''' <summary>
    ''' アサイン取得メイン処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateDataView()

        'メッセージ初期化
        StatusStrip1.Text = ""

        '常にDBに接続済の時は接続を切る 
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        L_CurrentX = Nothing
        L_CurrentY = Nothing
        AssignGrid.ClearSelection()
        For i As Integer = 0 To dsDATA.Tables.Count - 1
            dsDATA.Tables(i).DefaultView.Sort = String.Empty
            dsDATA.Tables(i).Clear()
            dsDATA.Tables(i).Constraints.Clear()
            For j As Integer = dsDATA.Tables(i).Columns.Count - 1 To 0 Step -1
                dsDATA.Tables(i).Columns.RemoveAt(j)
            Next
            AssignGrid.Columns.Clear()
            If i = dsDATA.Tables.Count - 1 Then
                dsDATA.Tables.Clear()
                AssignGrid.DataSource = Nothing
            End If
        Next

        Try
            'データベースを選択
            cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                & "Trusted_Connection = Yes;" _
                                & "Initial Catalog=S開発アサイン管理;"
            cn.Open()
        Catch ex As Exception
            MetroFramework.MetroMessageBox.Show(Me, "DBへの接続に失敗しました。", "エラー", MessageBoxButtons.OK)
            Return
        End Try

        Dim L_Query As String = ""
        L_Query = FN_CreateQuery()

        If H_Order.Text.ToString.TrimEnd <> "" Then
            L_Query &= " ORDER BY " & H_Order.Text.ToString.TrimEnd
            If H_Orderkb.Text.ToString.TrimEnd = "降順" Then
                L_Query &= " DESC "
            End If
        End If

        Try
            'Columnのサイズは固定にしてから列を設定
            AssignGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
            'DB接続
            Dim daAuthors As New SqlDataAdapter(L_Query, cn)
            daAuthors.FillSchema(dsDATA, SchemaType.Source)
            dsDATA.Tables("Table").PrimaryKey = Nothing
            daAuthors.Fill(dsDATA)
            AssignGrid.DataMember = dsDATA.Tables("table").TableName
            AssignGrid.DataSource = dsDATA.Tables(0)

            Dim BColumn As New DataGridViewButtonColumn
            BColumn.Name = "料金表を開く"
            BColumn.HeaderText = My.Settings("料金表を開く").ToString.TrimEnd
            BColumn.UseColumnTextForButtonValue = True
            BColumn.Text = "開く"
            BColumn.DataPropertyName = My.Settings("料金表を開く").ToString.TrimEnd
            Dim idx As Integer = AssignGrid.Columns(My.Settings("料金表を開く").ToString.TrimEnd).Index
            AssignGrid.Columns.RemoveAt(idx)
            AssignGrid.Columns.Insert(idx, BColumn)
            AssignGrid.VirtualMode = True

            Dim ExcelColumn As New DataGridViewButtonColumn
            ExcelColumn.Name = "進捗表を開く"
            ExcelColumn.HeaderText = My.Settings("進捗表を開く").ToString.TrimEnd
            ExcelColumn.UseColumnTextForButtonValue = True
            ExcelColumn.Text = "進捗表を開く"
            ExcelColumn.DataPropertyName = My.Settings("進捗表を開く").ToString.TrimEnd
            idx = AssignGrid.Columns(My.Settings("進捗表を開く").ToString.TrimEnd).Index
            AssignGrid.Columns.RemoveAt(idx)
            AssignGrid.Columns.Insert(idx, ExcelColumn)
            AssignGrid.VirtualMode = True

            'セルの初期設定を行います。
            FN_CellSetting()
            For i As Integer = 0 To AssignGrid.Columns.Count - 1
                FN_ColumnVisibleSEttings(i)
            Next
            My.Settings("ERPNAME") = H_WHERE7.Text.ToString.TrimEnd
            SubFormCanOpen = True
        Catch ex As Exception
            Microsoft.VisualBasic.MsgBox(ex.Message, MsgBoxStyle.OkCancel, "クエリエラー")
            If Microsoft.VisualBasic.MsgBox("Columnの設定に誤りがあるようです。設定を初期化しますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                My.Settings.Reset()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' アサイン情報取得クエリ作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FN_CreateQuery()
        Dim L_SQL As String = ""
        Dim L_ResultSQL As String = ""
        Dim L_ResultWhereSQL As String = ""


        'SELECT句                                                                                                                                                        
        L_SQL &= "(SELECT " & vbCrLf & vbCrLf
        L_SQL &= "  案件NO			=TRN.案件NO " & vbCrLf
        L_SQL &= " ,受注NO			=TRN.受注コード" & vbCrLf
        L_SQL &= " ,枝番	        =SGT.枝番 " & vbCrLf
        L_SQL &= " ,得意先コード	=TRN.得意先コード" & vbCrLf
        L_SQL &= " ,得意先名		=RTRIM(TKM.得意先名)" & vbCrLf
        L_SQL &= " ,エディション	=CASE WHEN EDM.エディション名='通常' THEN '' ELSE EDM.エディション名 END" & vbCrLf
        L_SQL &= " ,物件工数		=(SELECT 工数 = SUM(CASE WHEN SUB.受注コード <> '' THEN SUB.PG依頼工数 ELSE 0 END) FROM S開発アサイン管理.dbo.基本案件情報トラン SUB WHERE SUB.受注コード = TRN.受注コード)" & vbCrLf
        L_SQL &= " ,PG工数			=CONVERT(MONEY,TRN.PG依頼工数)" & vbCrLf
        L_SQL &= " ,QA件数          =SGT.QA件数" & vbCrLf
        L_SQL &= " ,開発障害        =ISNULL(SGT2.開発障害,0)" & vbCrLf
        L_SQL &= " ,設計障害        =ISNULL(SGT2.設計障害,0)" & vbCrLf
        L_SQL &= " ,仕様書			=CONVERT(INT,TRN.仕様書提出日)" & vbCrLf
        L_SQL &= " ,PG完了			=CONVERT(INT,TRN.PG完了希望日)" & vbCrLf
        L_SQL &= " ,SEテスト  		=CONVERT(INT,TRN.SEテスト開始予定日)" & vbCrLf
        L_SQL &= " ,納品日          =CONVERT(INT,TRN.納品予定日)" & vbCrLf
        L_SQL &= " ,検収月          =CONVERT(INT,TRN.売上予定年月)" & vbCrLf
        L_SQL &= " ,対応残          =SGT.対応残" & vbCrLf
        L_SQL &= " ,開発主管者名    =ISNULL(NULLIF(RTRIM(SGT.開発主管者名),''),ISNULL(TNM.担当者名,''))" & vbCrLf
        L_SQL &= " ,設計主管者名    =ISNULL(NULLIF(RTRIM(SGT.設計主管者名),''),ISNULL(STM.担当者名,''))" & vbCrLf
        L_SQL &= " ,進捗表更新日    =CONVERT(INT,SGT.最終更新日)" & vbCrLf
        L_SQL &= " ,料金表パス      =RTRIM(TRN.カスタマイズ料金表)" & vbCrLf
        L_SQL &= " ,料金表を開く    =RTRIM(TRN.カスタマイズ料金表)" & vbCrLf
        L_SQL &= " ,進捗表パス      =RTRIM(TRN.進捗表)" & vbCrLf
        L_SQL &= " ,進捗表を開く      =RTRIM(TRN.進捗表)" & vbCrLf
        L_SQL &= " FROM S開発アサイン管理.dbo.基本案件情報トラン TRN" & vbCrLf
        L_SQL &= " LEFT OUTER JOIN S開発アサイン管理.dbo.エディションマスタ EDM ON TRN.エディションコード = EDM.エディションコード" & vbCrLf
        L_SQL &= " LEFT OUTER JOIN S開発アサイン管理.dbo.得意先マスタ TKM ON TRN.得意先コード = TKM.得意先コード" & vbCrLf
        L_SQL &= " LEFT OUTER JOIN S開発アサイン管理.dbo.担当者マスタ STM ON TRN.担当者コード = STM.担当者コード" & vbCrLf
        L_SQL &= " LEFT OUTER JOIN (SELECT 対応残=COUNT(状況)" & vbCrLf
        L_SQL &= " 					    ,受注NO=B_受注NO" & vbCrLf
        L_SQL &= " 					    ,枝番=B_枝番" & vbCrLf
        L_SQL &= " 					    ,案件名=MAX(案件名)" & vbCrLf
        L_SQL &= " 					    ,地区名=MAX(地区名)" & vbCrLf
        L_SQL &= " 					    ,開発主管者名=MAX(開発主管者名)" & vbCrLf
        L_SQL &= " 					    ,設計主管者名=MAX(設計主管者名)" & vbCrLf
        L_SQL &= " 					    ,最終更新日=CONVERT(CHAR(8),MAX(PM.更新日時),112)" & vbCrLf
        L_SQL &= " 					    ,納品日=MAX(納品日)" & vbCrLf
        L_SQL &= " 					    ,検収日=MAX(検収日)" & vbCrLf
        L_SQL &= "                      ,開発障害=0" & vbCrLf
        L_SQL &= "                      ,設計障害=0" & vbCrLf
        L_SQL &= "                      ,QA件数=(SELECT COUNT(*) FROM S開発品質管理.dbo.T_QAトラン QA WHERE QA.受注NO = B_受注NO AND QA.枝番 = B_枝番)" & vbCrLf
        L_SQL &= " 					FROM (SELECT 地区名,開発主管者名,設計主管者名,案件名,B_受注NO=受注NO,B_枝番=枝番,納品日,検収日 FROM S開発品質管理.dbo.T_案件マスタ ANM) JBT" & vbCrLf
        L_SQL &= " 						 LEFT OUTER JOIN (SELECT 受注NO=PM_SUB.受注NO " & vbCrLf
        L_SQL &= "                                              ,枝番=PM_SUB.枝番" & vbCrLf
        L_SQL &= "                                              ,更新日時=MAX(PM_SUB.更新日時)" & vbCrLf
        L_SQL &= "                                          FROM S開発品質管理.dbo.T_プログラムマスタ PM_SUB GROUP BY 受注NO,枝番) PM ON JBT.B_受注NO = PM.受注NO AND JBT.B_枝番=PM.枝番" & vbCrLf
        L_SQL &= " 						 LEFT OUTER JOIN S開発品質管理.dbo.T_障害トラン  SUB_SGT ON SUB_SGT.状況<>'完了' AND SUB_SGT.状況<>'問題なし' AND SUB_SGT.受注NO = JBT.B_受注NO AND SUB_SGT.枝番 = JBT.B_枝番" & vbCrLf
        L_SQL &= " 				GROUP BY B_受注NO,B_枝番) SGT ON TRN.受注コード = SGT.受注NO " & vbCrLf
        L_SQL &= " LEFT OUTER JOIN (SELECT 受注NO,枝番,開発障害=SUM(CASE WHEN 障害工程='開発' THEN 1 ELSE 0 END),設計障害=SUM(CASE WHEN 障害工程='設計' THEN 1 ELSE 0 END) FROM S開発品質管理.dbo.T_障害トラン GROUP BY 受注NO,枝番) SGT2 ON TRN.受注コード = SGT2.受注NO AND SGT.枝番 = SGT2.枝番" & vbCrLf
        'Edit t_Arimoto >>>>>>>>>>>>>>> 2018/12/20
        L_SQL &= " LEFT OUTER JOIN アサイン一覧案件トラン Sub ON TRN.案件NO = Sub.案件NO"
        L_SQL &= " LEFT OUTER JOIN 担当者マスタ TNM ON Sub.CR > '' AND Sub.CR = TNM.略称"
        'Edit t_Arimoto <<<<<<<<<<<<<<< 2018/12/20
        L_SQL &= "  )" & vbCrLf

        Dim AddComma As Boolean = False

        L_ResultSQL = ""
        L_ResultSQL &= " SELECT " & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("案件NO", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("受注NO", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("枝番", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("得意先コード", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("得意先名", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("エディション", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("物件工数", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("PG工数", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("開発障害", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("設計障害", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("QA件数", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("仕様書", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("PG完了", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("SEテスト", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("納品日", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("検収月", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("対応残", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("設計主管者名", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("開発主管者名", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("進捗表更新日", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("料金表パス", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("料金表を開く", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("進捗表パス", AddComma) & vbCrLf
        L_ResultSQL &= FN_CreateSelectQuery("進捗表を開く", AddComma) & vbCrLf

        L_ResultSQL &= " FROM " & L_SQL & " MAIN "

        If TAB1_Where_Group.Text.TrimEnd <> "" Then
            L_ResultSQL &= "       INNER JOIN S開発品質管理.dbo.T_担当者マスタ TNM ON MAIN.開発主管者名 = TNM.担当者名 "
            L_ResultSQL &= "       INNER JOIN S開発品質管理.dbo.T_部署マスタ   BMN ON TNM.部署コード = BMN.部署コード "
            If TAB1_Where_Group.Text.TrimEnd = "関西" Then
                L_ResultSQL &= "         AND BMN.拠点区分 IN('1','3') "
            ElseIf TAB1_Where_Group.Text.TrimEnd = "首都圏" Then
                L_ResultSQL &= "         AND BMN.拠点区分 IN('2') "
            Else
                L_ResultSQL &= "         AND BMN.部署名 LIKE '%" & TAB1_Where_Group.Text.TrimEnd & "%' "
            End If
        End If

        FN_AddWHERE2(L_ResultSQL)

        If H_WHERE9.Text.TrimEnd <> "" AndAlso H_WHERE10.Text.TrimEnd <> "" Then
            L_ResultSQL &= " AND " & LBL_WHERE9.Text & " BETWEEN " & H_WHERE9.Text.TrimEnd & " AND " & H_WHERE10.Text.TrimEnd & " " & vbCrLf
        ElseIf H_WHERE9.Text.TrimEnd <> "" Then
            L_ResultSQL &= " AND " & LBL_WHERE9.Text & " >= " & H_WHERE9.Text.TrimEnd & " " & vbCrLf
        ElseIf H_WHERE10.Text.TrimEnd <> "" Then
            L_ResultSQL &= " AND " & LBL_WHERE9.Text & " <= " & H_WHERE10.Text.TrimEnd & " " & vbCrLf
        End If

        L_ResultWhereSQL = "SELECT DISTINCT TOP " & My.Settings("MaxView").ToString.TrimEnd & " * FROM (" & L_ResultSQL & ") WSQL "

        Return L_ResultWhereSQL
    End Function

    ''' <summary>
    ''' 事例検索Whrere句作成
    ''' </summary>
    ''' <param name="L_SQL"></param>
    ''' <remarks></remarks>
    Private Sub FN_AddWHERE2(ByRef L_SQL As String)
        L_SQL &= " WHERE 1=1 "
        If BODY.SelectedTab.Name = "AssignTab" Then
            If H_WHERE0.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND 案件NO " & CreateWhere(0) & " "
            End If
            If H_WHERE1.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND 受注NO " & CreateWhere(1) & " "
            End If
            If H_WHERE2.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND 得意先名 " & CreateWhere(2) & " "
            End If
            If H_WHERE3.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND PG完了 " & CreateWhere(3) & " "
            End If
            If H_WHERE5.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND 検収月 " & CreateWhere(5) & " "
            End If
            If H_WHERE6.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND 設計主管者名 " & CreateWhere(6) & " "
            End If
            If H_WHERE7.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND 開発主管者名 " & CreateWhere(7) & " "
            End If
            If H_WHERE8.Text.ToString.TrimEnd <> "" Then
                L_SQL &= " AND 仕様書 " & CreateWhere(8) & " "
            End If

        End If
    End Sub

    ''' <summary>
    ''' アサイン情報取得Where句作成
    ''' </summary>
    ''' <param name="idx"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateWhere(ByVal idx As Integer) As String
        Dim op As String = MetroPanel10.Controls("H_OP" + idx.ToString()).Text.TrimEnd
        Dim val As String = MetroPanel10.Controls("H_WHERE" + idx.ToString()).Text.TrimEnd
        If op = "=" AndAlso Not IsNumeric(val) Then
            Return " COLLATE Japanese_CI_AS LIKE '%" & val & "%' "
        ElseIf (op = "=") Then
            Return " LIKE '%" & val & "%' "
        Else
            Return " " & op & " '" & val & "' "
        End If
    End Function

    ''' <summary>
    ''' 基本案件情報を開く
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ShowAKBS1010()
        Dim CellIndex As Integer = 0
        Dim denno As String = ""
        If AssignGrid.SelectedCells Is Nothing Then Return
        denno = AssignGrid.Rows(AssignGrid.SelectedCells(0).RowIndex).Cells(My.Settings("案件NO").ToString()).Value
        Dim arg As Object = New TYPMGR.typeTranLoad(SAIBAS.SAIKB.AKN, denno)
        EXEMGR.Main(arg)
    End Sub

    ''' <summary>
    ''' 負荷アサインシステム管理クラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EXEMGR
        Inherits ILL.ERF.G1Base.SYSMGR

        Shared Sub Main(ByVal arg As Object)
            Dim L_EXE As EXEMGR
            Dim L_INI As INIMGR
            Dim L_Form As BasicForm

            'Visualスタイル許可
            System.Windows.Forms.Application.EnableVisualStyles()
            Dim FilePath As String = "C:\ProgramData\AONET.ini"
            Dim Initxt As String = ""
            Dim Fukaini As StringBuilder = New StringBuilder

            Using reader As New StreamReader(FilePath, Encoding.GetEncoding("Shift_JIS"))
                Initxt = reader.ReadToEnd
                reader.Close()
            End Using
            Using writer As New StreamWriter(FilePath, False, Encoding.GetEncoding("Shift_JIS"))
                writer.WriteLine("### AONET ### ")
                writer.WriteLine("[AONET] ")
                writer.WriteLine("Version=0.01 ")
                writer.WriteLine("Server=KDC-O-SE01\s_kaihatsu ")
                writer.WriteLine("Database=S開発アサイン管理 ")
                writer.WriteLine("UserID=1 ")
                writer.WriteLine("LoginID=9501 ")
                writer.WriteLine("Password= ")
                writer.WriteLine("EXCELDir=D:\EXCEL出力\ ")
                writer.WriteLine("ExecDir=\\kdc-hyv03\o-dc02\共有\部署フォルダ\サポート\PG\負荷情報アサイン情報管理システム2\Exec\ ")
                writer.WriteLine("RptDir=\\kdc-hyv03\o-dc02\共有\部署フォルダ\サポート\PG\負荷情報アサイン情報管理システム2\Rpx\ ")
                writer.WriteLine("GrdxDir=\\kdc-hyv03\o-dc02\共有\部署フォルダ\サポート\PG\負荷情報アサイン情報管理システム2\Grdx\ ")
                writer.WriteLine("Updatable=1 ")
                writer.WriteLine("Timeout=300 ")
                writer.WriteLine("LockTimeout=5000 ")
                writer.WriteLine("UseSecurity=0 ")
                writer.WriteLine("DefaultPreviewSize=100  ")
                writer.WriteLine("UseFax=1 ")
                writer.WriteLine("Tier=2 ")
                writer.WriteLine("FullScreen=1 ")
                writer.WriteLine("LogOut=0 ")
                writer.Close()
            End Using

            'INI読み込み
            L_INI = New INIMGR
            If L_INI.LoadIniFile = False Then
                MetroFramework.MetroMessageBox.Show(Nothing, MSGMGR.Common.Alert.Load_IniFile)
                Exit Sub
            End If

            '自分のインスタンス作成
            L_EXE = New EXEMGR
            L_EXE.INI = L_INI

            '初期フォームのインスタンス作成するコードを記述する>>>
            L_EXE.PrimaryForm = New ILL.ERF.AKBS1010.PrimaryForm
            L_Form = L_EXE.ShowPrimaryForm("\\kdc-hyv03\o-dc02\共有\部署フォルダ\サポート\PG\負荷情報アサイン情報管理システム2\Exec\", Nothing, arg)

            Using writer As New StreamWriter(FilePath, False, Encoding.GetEncoding("Shift_JIS"))
                writer.Write(Initxt)
                writer.Close()
            End Using
        End Sub
    End Class

    ''' <summary>
    ''' [アサイン情報]カラム設定
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub FN_CellSetting()
        'Return

        For Each col As DataGridViewColumn In AssignGrid.Columns
            'For i As Integer = 1 To 26
            If DirectCast(My.Settings("HiddenColumns"), StringCollection).Contains(col.Name) Then Continue For
            With col
                Select Case .Name
                    Case My.Settings("案件NO").ToString()
                        .Width = 70
                    Case My.Settings("受注NO").ToString()
                        .Width = 70
                    Case My.Settings("枝番").ToString()
                        .Width = 50
                    Case My.Settings("得意先コード").ToString()
                        .Width = 70
                    Case My.Settings("得意先名").ToString()
                        .Width = 150
                    Case My.Settings("エディション").ToString()
                    Case My.Settings("物件工数").ToString()
                        .DefaultCellStyle.Format = "###,#0.0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case My.Settings("PG工数").ToString()
                        .DefaultCellStyle.Format = "###,#0.0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case My.Settings("開発障害").ToString()
                        .DefaultCellStyle.Format = "###,#0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case My.Settings("設計障害").ToString()
                        .DefaultCellStyle.Format = "###,#0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case My.Settings("QA件数").ToString()
                        .DefaultCellStyle.Format = "###,#0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 60
                    Case My.Settings("仕様書").ToString()
                        .DefaultCellStyle.Format = "####/##/##"
                        .Width = 70
                    Case My.Settings("PG完了").ToString()
                        .DefaultCellStyle.Format = "####/##/##"
                        .Width = 70
                    Case My.Settings("SEテスト").ToString()
                        .DefaultCellStyle.Format = "####/##/##"
                        .Width = 70
                    Case My.Settings("納品日").ToString()
                        .DefaultCellStyle.Format = "####/##/##"
                        .Width = 70
                    Case My.Settings("検収月").ToString()
                        .DefaultCellStyle.Format = "####/##"
                        .Width = 70
                    Case My.Settings("対応残").ToString()
                        .DefaultCellStyle.Format = "###,#0"
                        .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                        .Width = 70
                    Case My.Settings("設計主管者名").ToString()
                        .Width = 80
                    Case My.Settings("開発主管者名").ToString()
                        .Width = 80
                    Case My.Settings("進捗表更新日").ToString()
                        .DefaultCellStyle.Format = "####/##/##"
                        .Width = 90
                    Case My.Settings("料金表パス").ToString()
                        .Width = 150
                    Case My.Settings("料金表を開く").ToString()
                    Case My.Settings("進捗表パス").ToString()
                        .Width = 150
                    Case My.Settings("進捗表を開く").ToString()
                End Select
            End With
        Next
    End Sub

    Private Sub FN_ColumnVisibleSEttings(ByVal index As Integer)
        For Each col As String In DirectCast(My.Settings("HiddenColumns"), StringCollection)
            AssignGrid.Columns(My.Settings(col).ToString()).Visible = False
        Next
    End Sub

    Private Function FN_CreateSelectQuery(ByVal Target As String, ByRef AddComma As Boolean) As String
        If My.Settings(Target) = "" Then
            My.Settings(Target) = Target
            If AddComma = True Then
                AddComma = True
                Return "," & Target & "=" & Target
            Else
                AddComma = True
                Return Target & "=" & Target
            End If
            Return Target
        Else
            If AddComma = True Then
                AddComma = True
                Return "," & My.Settings(Target) & "=" & Target
            Else
                AddComma = True
                Return My.Settings(Target) & "=" & Target
            End If
        End If
    End Function

    ''' <summary>
    ''' 進捗表パス登録処理
    ''' </summary>
    ''' <param name="Denno"></param>
    ''' <param name="val"></param>
    ''' <remarks></remarks>
    Private Sub UpdateExcel(ByVal Denno As String, ByVal val As String)

        '常にDBに接続済の時は接続を切る 
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        Using cn As SqlConnection = New SqlConnection()
            'データベースを選択
            Try
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                & "Trusted_Connection = Yes;" _
                                & "Initial Catalog=S開発アサイン管理;"
                cn.Open()
                Dim queryString As String = "UPDATE 基本案件情報トラン" _
                                          & "   SET 進捗表 = '" & val & "' " _
                                          & " WHERE 案件NO = '" & Denno & "' "
                Dim command As SqlCommand = New SqlCommand(queryString, cn)
                Dim ret As Integer = command.ExecuteNonQuery()
            Catch ex As Exception
                ex = ex
            End Try
        End Using
    End Sub


    ''' <summary>
    ''' 工数割振りサブ画面を開く
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenSubformKousu()
        Subkb = "工数"
        OpenSubform()
    End Sub

    ''' <summary>
    ''' 対応内訳サブ画面を開く
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenSubformTaiou()
        Subkb = "対応"
        OpenSubform()
    End Sub

    ''' <summary>
    ''' プログラム一覧サブ画面を開く
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenSubformPGList()
        Subkb = "PG一覧"
        OpenSubform()
    End Sub

    ''' <summary>
    ''' 原価内訳サブ画面を開く
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenSubformCost()
        Subkb = "原価内訳"
        OpenSubform()
    End Sub

    ''' <summary>
    ''' 原価内訳サブ画面を開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 原価内訳ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 原価内訳ToolStripMenuItem1.Click
        OpenSubformCost()
    End Sub

    ''' <summary>
    ''' サブ画面を開く(開く画面は設定による))
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub OpenSubform()
        If AssignGrid.CurrentRow Is Nothing Then Return
        If SubFormCanOpen = False Then Return
        Dim L_JdnColumnNM As String = ""
        If Not DirectCast(My.Settings("HiddenColumns"), StringCollection).Contains("受注NO") Then
            L_JdnColumnNM = My.Settings("受注NO").ToString.TrimEnd
        End If
        Jdnno = AssignGrid.CurrentRow.Cells(L_JdnColumnNM).Value
        If Jdnno = "" Then
            MetroFramework.MetroMessageBox.Show(Me, "受注NOが設定されていません。", "確認")
            Me.Close()
            Return
        End If
        Dim SubForm As New SubdataForm()
        If AssignGrid.SelectedCells.Count > 0 Then
            SubForm.posx = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.X + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Width / 2 + Me.Location.X + 200
            SubForm.posy = AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Location.Y + AssignGrid.GetCellDisplayRectangle(AssignGrid.SelectedCells(0).ColumnIndex, AssignGrid.SelectedCells(0).RowIndex, True).Height / 2 + Me.Location.Y + 109
        End If
        SubForm.Show()
    End Sub

    ''' <summary>
    ''' アサイン情報カーソル移動（KeyDown）
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TAB1_SetFocusEvent(sender As Object, e As EventArgs) Handles H_WHERE0.KeyDown, H_WHERE1.KeyDown, H_WHERE2.KeyDown, H_WHERE3.KeyDown, H_WHERE4.KeyDown, H_WHERE5.KeyDown, H_WHERE6.KeyDown, H_WHERE7.KeyDown, H_WHERE8.KeyDown, H_WHERE9.KeyDown, H_WHERE10.KeyDown
        'エンターか↓キーで移動
        Select Case DirectCast(e, System.Windows.Forms.KeyEventArgs).KeyValue
            Case Keys.Enter
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 1, 1))
                    If DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name = "H_WHERE10" Then
                        Index = -1
                    End If
                    If MetroPanel10.Controls("H_WHERE" & (Index + 1)) IsNot Nothing Then
                        MetroPanel10.Controls("H_WHERE" & (Index + 1)).Select()
                    Else
                        UpdateDataView()
                    End If
                End If
            Case Keys.Down
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 1, 1))
                    If DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name = "H_WHERE10" Then
                        Index = -1
                    End If
                    If MetroPanel10.Controls("H_WHERE" & (Index + 1)) IsNot Nothing Then
                        MetroPanel10.Controls("H_WHERE" & (Index + 1)).Select()
                    End If
                End If
            Case Keys.Up
                If TypeOf (sender) Is MetroFramework.Controls.MetroTextBox Then
                    Dim Index As Integer = CInt(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Substring(DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name.Length - 1, 1))
                    If DirectCast(sender, MetroFramework.Controls.MetroTextBox).Name = "H_WHERE10" Then
                        Index = -1
                    End If
                    If MetroPanel10.Controls("H_WHERE" & (Index - 1)) IsNot Nothing Then
                        MetroPanel10.Controls("H_WHERE" & (Index - 1)).Select()
                    End If
                End If
        End Select
    End Sub

    ''' <summary>
    ''' 並び順変更時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub H_Order_SelectedIndexChanged(sender As Object, e As EventArgs) Handles H_Order.MouseDown
        '常に最新の情報を表示する
        H_Order.Items.Clear()
        For Each col As String In DirectCast(My.Settings("VisibleColumns"), StringCollection)
            Try
                H_Order.Items.Add(My.Settings(col))
            Catch ex As Exception
            End Try
        Next
    End Sub

    ''' <summary>
    ''' 今月仕様書一覧取得
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        H_WHERE0.Text = ""
        H_WHERE1.Text = ""
        H_WHERE2.Text = ""
        H_WHERE3.Text = ""
        H_WHERE4.Text = ""
        H_WHERE5.Text = ""
        H_WHERE6.Text = ""
        H_OP8.Text = "="
        H_WHERE8.Text = DateTime.Today.ToString("yyyyMM")
        H_WHERE9.Text = ""
        H_WHERE10.Text = ""
        H_Order.Text = My.Settings("仕様書").ToString.TrimEnd
        UpdateDataView()
    End Sub


    ''' <summary>
    ''' 仕掛一覧取得
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        H_WHERE0.Text = ""
        H_WHERE1.Text = ""
        H_WHERE2.Text = ""
        H_WHERE3.Text = ""
        H_WHERE4.Text = ""
        H_OP5.Text = ">="
        H_WHERE5.Text = DateTime.Today.ToString("yyyyMM")
        H_WHERE6.Text = ""
        H_OP8.Text = "<="
        H_WHERE8.Text = DateTime.Now.ToString("yyyyMMdd")
        H_WHERE9.Text = ""
        H_WHERE10.Text = ""

        H_Order.Text = My.Settings("検収月").ToString.TrimEnd
        UpdateDataView()

        If AssignGrid IsNot Nothing AndAlso AssignGrid.Rows.Count > 0 Then
            If AssignGrid.Columns.Contains(My.Settings("PG工数").ToString()) Then
                Dim sum As Decimal = Decimal.Zero

                For i As Integer = 0 To AssignGrid.Rows.Count - 1
                    sum += Decimal.Parse(AssignGrid(My.Settings("PG工数").ToString(), i).Value.ToString())
                Next

                StatusStrip1.Text = "総工数：" + sum.ToString("#.0")
            End If
        End If
    End Sub

    ''' <summary>
    ''' 仕掛一覧取得
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 仕様書提出前アサイン_Click(sender As Object, e As EventArgs) Handles 仕様書提出前アサイン.Click
        H_WHERE0.Text = ""
        H_WHERE1.Text = ""
        H_WHERE2.Text = ""
        H_WHERE3.Text = ""
        H_WHERE4.Text = ""
        H_OP5.Text = "="
        H_WHERE5.Text = ""
        H_WHERE6.Text = ""
        H_OP8.Text = ">="
        H_WHERE8.Text = DateTime.Now.ToString("yyyyMMdd")
        H_WHERE9.Text = ""
        H_WHERE10.Text = ""

        H_Order.Text = My.Settings("仕様書").ToString.TrimEnd
        UpdateDataView()

        If AssignGrid IsNot Nothing AndAlso AssignGrid.Rows.Count > 0 Then
            If AssignGrid.Columns.Contains(My.Settings("PG工数").ToString()) Then
                Dim sum As Decimal = Decimal.Zero

                For i As Integer = 0 To AssignGrid.Rows.Count - 1
                    sum += Decimal.Parse(AssignGrid(My.Settings("PG工数").ToString(), i).Value.ToString())
                Next

                StatusStrip1.Text = "総工数：" + sum.ToString("#.0")
            End If
        End If
    End Sub


    ''' <summary>
    ''' 条件をクリア
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 条件クリア_Click(sender As Object, e As EventArgs) Handles Button1.Click
        H_WHERE0.Text = ""
        H_WHERE1.Text = ""
        H_WHERE2.Text = ""
        H_WHERE3.Text = ""
        H_WHERE4.Text = ""
        H_WHERE5.Text = ""
        H_WHERE6.Text = ""
        H_WHERE7.Text = ""
        H_WHERE8.Text = ""
        H_WHERE9.Text = ""
        H_WHERE10.Text = ""

        H_OP1.Text = "="
        H_OP2.Text = "="
        H_OP3.Text = "="
        H_OP4.Text = "="
        H_OP5.Text = "="
        H_OP6.Text = "="
        H_OP7.Text = "="
        H_OP8.Text = "="

        H_Order.Text = My.Settings("案件NO").ToString.TrimEnd
    End Sub

#Region " [右クリック] "
    ''' <summary>
    ''' [右クリック] コピー
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub コピーToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles コピーToolStripMenuItem.Click
        '選択されたセルをクリップボードにコピーする
        Clipboard.SetDataObject(AssignGrid.GetClipboardContent())
    End Sub

    ''' <summary>
    ''' [右クリック] 値をヘッダ付きでコピーする
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 値をヘッダ付きでコピーToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 値をヘッダ付きでコピーToolStripMenuItem.Click
        '選択されたセルをクリップボードにコピーする
        AssignGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
        Clipboard.SetDataObject(AssignGrid.GetClipboardContent())
        AssignGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText
    End Sub

    ''' <summary>
    ''' [右クリック] 進捗表を登録
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 進捗表を登録ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 進捗表を登録ToolStripMenuItem.Click
        'OpenFileDialogクラスのインスタンスを作成
        Dim ofd As New OpenFileDialog()

        'はじめのファイル名を指定する
        'はじめに「ファイル名」で表示される文字列を指定する
        ofd.FileName = "進捗表を指定"
        'はじめに表示されるフォルダを指定する
        '指定しない（空の文字列）の時は、現在のディレクトリが表示される
        ofd.InitialDirectory = "\\192.168.0.20\users"
        '[ファイルの種類]に表示される選択肢を指定する
        '指定しないとすべてのファイルが表示される
        ofd.Filter = "(*.*)|*.*"
        '[ファイルの種類]ではじめに選択されるものを指定する
        '2番目の「すべてのファイル」が選択されているようにする
        ofd.FilterIndex = 2
        'タイトルを設定する
        ofd.Title = "開くファイルを選択してください"
        'ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
        ofd.RestoreDirectory = True
        '存在しないファイルの名前が指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckFileExists = True
        '存在しないパスが指定されたとき警告を表示する
        'デフォルトでTrueなので指定する必要はない
        ofd.CheckPathExists = True

        'ダイアログを表示する
        If ofd.ShowDialog() = DialogResult.OK Then
            'OKボタンがクリックされたとき、選択されたファイル名を表示する
            Dim CellIndex As Integer = 0
            Dim denno As String = ""

            For Each cell As DataGridViewCell In AssignGrid.SelectedCells
                denno = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("案件NO").ToString.TrimEnd).Value.ToString.TrimEnd
                CellIndex = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("進捗表パス").ToString()).ColumnIndex
                dsDATA.Tables(0).Rows(cell.RowIndex).BeginEdit()
                dsDATA.Tables(0).Columns(CellIndex).ReadOnly = False
                dsDATA.Tables(0).Rows(cell.RowIndex).Item(CellIndex) = ofd.FileName
                dsDATA.Tables(0).Columns(CellIndex).ReadOnly = True
                dsDATA.Tables(0).Rows(cell.RowIndex).EndEdit()

                ' 基本案件情報トランに登録
                UpdateExcel(denno, ofd.FileName)
                AssignGrid.Refresh()
            Next

            'AddPrgList()
        End If
    End Sub

    ''' <summary>
    ''' [右クリック] 登録を解除
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 登録を解除ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 登録を解除ToolStripMenuItem2.Click
        '常にDBに接続済の時は接続を切る 
        'OKボタンがクリックされたとき、選択されたファイル名を表示する
        Dim CellIndex As Integer = 0
        Dim denno As String = ""

        For Each cell As DataGridViewCell In AssignGrid.SelectedCells
            denno = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("案件NO").ToString.TrimEnd).Value.ToString.TrimEnd
            CellIndex = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("進捗表パス").ToString()).ColumnIndex
            dsDATA.Tables(0).Rows(cell.RowIndex).BeginEdit()
            dsDATA.Tables(0).Columns(CellIndex).ReadOnly = False
            dsDATA.Tables(0).Rows(cell.RowIndex).Item(CellIndex) = ""
            dsDATA.Tables(0).Columns(CellIndex).ReadOnly = True
            dsDATA.Tables(0).Rows(cell.RowIndex).EndEdit()

            ' 基本案件情報トランに登録
            UpdateExcel(denno, "")
            AssignGrid.Refresh()
        Next
    End Sub

    ''' <summary>
    ''' [右クリック] 進捗表を開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 料金表のパスを開くToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 料金表のパスを開くToolStripMenuItem.Click
        Dim CellIndex As Integer = 0
        Dim excelpath As String = ""
        For Each cell As DataGridViewCell In AssignGrid.SelectedCells
            excelpath = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("料金表パス").ToString()).Value
            If System.IO.Directory.Exists(excelpath) Then
                System.Diagnostics.Process.Start(excelpath)
            End If
        Next
    End Sub

    ''' <summary>
    ''' [右クリック] 進捗表を開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 進捗表を開くToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 進捗表を開くToolStripMenuItem.Click
        Dim CellIndex As Integer = 0
        Dim excelpath As String = ""
        For Each cell As DataGridViewCell In AssignGrid.SelectedCells
            excelpath = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("進捗表パス").ToString()).Value
            If System.IO.File.Exists(excelpath) Then
                System.Diagnostics.Process.Start(excelpath)
            End If
        Next
    End Sub

    ''' <summary>
    ''' [右クリック] フォルダを開く
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 進捗表フォルダを開くToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 進捗表フォルダを開くToolStripMenuItem.Click
        Dim CellIndex As Integer = 0
        Dim excelpath As String = ""
        For Each cell As DataGridViewCell In AssignGrid.SelectedCells
            excelpath = AssignGrid.Rows(cell.RowIndex).Cells(My.Settings("進捗表パス").ToString()).Value
            If System.IO.File.Exists(excelpath) Then
                System.Diagnostics.Process.Start(System.IO.Path.GetDirectoryName(excelpath))
            End If
        Next
    End Sub

    ''' <summary>
    ''' [右クリック] 基本案件情報を開くボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub 案件情報入力を開くToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 案件情報入力を開くToolStripMenuItem.Click
        ShowAKBS1010()
    End Sub

#End Region

#Region " [アサイン情報グリッド]"

    Public OwnBeginGrabRowindex As Integer = 0

    Private Sub AssignGrid_DragEnter(sender As Object, e As DragEventArgs) Handles AssignGrid.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        End If


    End Sub

    Private Sub AssignGrid_DragDrop(sender As Object, e As DragEventArgs) Handles AssignGrid.DragDrop
        Dim p As Point = AssignGrid.PointToClient(New Point(e.X, e.Y))
        Dim hit As DataGridView.HitTestInfo = AssignGrid.HitTest(p.X, p.Y)
        Dim denno As String = AssignGrid.Rows(hit.RowIndex).Cells(My.Settings("案件NO").ToString.TrimEnd).Value.ToString.TrimEnd
        Dim Cols As Integer = AssignGrid.Rows(hit.RowIndex).Cells(My.Settings("進捗表パス").ToString()).ColumnIndex

        If hit.RowIndex >= 0 Then
            dsDATA.Tables(0).Rows(hit.RowIndex).BeginEdit()
            dsDATA.Tables(0).Columns(Cols).ReadOnly = False
            dsDATA.Tables(0).Rows(hit.RowIndex).Item(Cols) = e.Data.GetData(DataFormats.FileDrop)(0)
            dsDATA.Tables(0).Columns(Cols).ReadOnly = True
            dsDATA.Tables(0).Rows(hit.RowIndex).EndEdit()
            AssignGrid.ClearSelection()
            AssignGrid.Rows(hit.RowIndex).Cells(Cols).Selected = True
            ' 基本案件情報トランに登録
            UpdateExcel(AssignGrid.Rows(hit.RowIndex).Cells("案件NO").Value.ToString(), e.Data.GetData(DataFormats.FileDrop)(0))
            AssignGrid.Refresh()

        End If
    End Sub

    Private Sub AssignGrid_DragOver(sender As Object, e As DragEventArgs) Handles AssignGrid.DragOver
        Dim p As Point = AssignGrid.PointToClient(New Point(e.X, e.Y))
        Dim hit As DataGridView.HitTestInfo = AssignGrid.HitTest(p.X, p.Y)
        If AssignGrid.Columns.Contains(My.Settings("進捗表パス").ToString()) Then
            Dim Cols As Integer = AssignGrid.Columns(My.Settings("進捗表パス").ToString()).Index
            If hit.RowIndex >= 0 Then
                AssignGrid.ClearSelection()
                AssignGrid.Rows(hit.RowIndex).Cells(Cols).Selected = True
            End If
        End If
    End Sub


    ''' <summary>
    ''' ダブルクリック(サブ画面を開く)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub D_DataGridView_CellDoubleClick(sender As Object, e As MouseEventArgs) Handles AssignGrid.DoubleClick
        ' 事例検索では実行しない
        If BODY.SelectedIndex = 1 Then
            ' ヘッダ以外のセルか？
            If DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.ColumnIndex >= 0 And DirectCast(sender, System.Windows.Forms.DataGridView).CurrentCell.RowIndex >= 0 Then
                If AssignGrid.SelectedCells.Count <= 1 Then
                    Select Case My.Settings("Taikb").ToString
                        Case "0"
                            ShowAKBS1010()
                        Case "1"
                            OpenSubformKousu()
                        Case "2"
                            OpenSubformTaiou()
                        Case "3"
                            OpenSubformPGList()
                        Case "4"
                            OpenSubformCost()
                    End Select
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' [CellMouseDown]クリック時にセルを選択状態にする
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub D_DataGridView_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles AssignGrid.CellMouseDown
        ' 右ボタンのクリックか？
        If e.Button = MouseButtons.Left Then
            'Me.SuspendLayout()
            ' ヘッダ以外のセルか？
            If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
                AssignGrid.ClearSelection()
                Dim cell As DataGridViewCell = AssignGrid(e.ColumnIndex, e.RowIndex)
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
                Dim cell As DataGridViewCell = AssignGrid(e.ColumnIndex, e.RowIndex)
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
    Private Sub D_DataGridView_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles AssignGrid.CellMouseUp

        If e.Button = MouseButtons.Right Then
            'Me.SuspendLayout()
            ' ヘッダ以外のセルか？
            If e.ColumnIndex >= 0 And e.RowIndex >= 0 Then
                If AssignGrid.SelectedCells.Count <= 1 Then
                    AssignGrid.ClearSelection()
                    ' 右クリックされたセル
                    Dim cell As DataGridViewCell = AssignGrid(e.ColumnIndex, e.RowIndex)
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
                Dim cell As DataGridViewCell = AssignGrid(e.ColumnIndex, e.RowIndex)
                cell.Selected = True
            End If
            'Me.ResumeLayout()
        End If
    End Sub

    ''' <summary>
    ''' 料金表を開くボタン
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub D_DataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles AssignGrid.CellContentClick
        If AssignGrid.Columns(e.ColumnIndex).Name = My.Settings("料金表を開く").ToString() Then

            If System.IO.Directory.Exists(AssignGrid.CurrentRow.Cells(My.Settings("料金表パス").ToString.TrimEnd).Value.ToString) Then
                Try
                    System.Diagnostics.Process.Start(AssignGrid.CurrentRow.Cells(My.Settings("料金表パス").ToString.TrimEnd).Value.ToString)
                Catch ex As Exception
                End Try
            Else
                MetroFramework.MetroMessageBox.Show(Me, "フォルダが存在しません", "エラー")
            End If
        End If

        If AssignGrid.Columns(e.ColumnIndex).Name = My.Settings("進捗表を開く").ToString() Then

            If System.IO.File.Exists(AssignGrid.CurrentRow.Cells(My.Settings("進捗表パス").ToString.TrimEnd).Value.ToString) Then
                Try
                    System.Diagnostics.Process.Start(AssignGrid.CurrentRow.Cells(My.Settings("進捗表パス").ToString.TrimEnd).Value.ToString)
                Catch ex As Exception
                End Try
            Else
                MetroFramework.MetroMessageBox.Show(Me, "フォルダが存在しません", "エラー")
            End If
        End If
    End Sub
#End Region
End Class
