Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms
Imports System.Text

Public Class SubdataForm
    Dim DBPass As String = ""
    Dim dsDATA As DataSet
    Dim dsSchima As DataSet
    Dim cn As New SqlConnection()
    Dim cmd As SqlCommand = cn.CreateCommand()
    Dim Sub_dsSchima As DataSet
    Dim Sub_dsDATA As DataSet
    Dim Sub_cn As New SqlConnection()
    Dim Sub_cmd As SqlCommand = cn.CreateCommand()
    Public Jdnno As String
    Public posx As Integer = 0
    Public posy As Integer = 0

    Private Sub SubdataForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim x As Integer = System.Windows.Forms.Cursor.Position.X
        Dim y As Integer = System.Windows.Forms.Cursor.Position.Y

        '設定されている場合は優先
        If posx <> 0 Then
            x = posx
        End If
        If posy <> 0 Then
            y = posy
        End If


        '各種初期設定
        Me.Location = New Point(x, y)
        D_DataGridView.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText


        'DataViewを編集可能に
        dsSchima = New DataSet("メインDB")
        dsDATA = New DataSet("取得")
        Sub_dsSchima = New DataSet("サブDB")
        Sub_dsDATA = New DataSet("サブ取得")

        Dim L_Jdnno As String
        L_Jdnno = My.Forms.m_PrimaryForm.Jdnno

        '常にDBに接続済の時は接続を切る 
        cmd.Dispose()
        cn.Close()
        cn.Dispose()
        D_DataGridView.ClearSelection()
        For i As Integer = 0 To dsDATA.Tables.Count - 1
            dsDATA.Tables(i).Clear()
            dsDATA.Tables(i).Constraints.Clear()
            For j As Integer = dsDATA.Tables(i).Columns.Count - 1 To 0 Step -1
                dsDATA.Tables(i).Columns.RemoveAt(j)
            Next
        Next

        Sub_cmd.Dispose()
        Sub_cn.Close()
        Sub_cn.Dispose()
        For i As Integer = 0 To Sub_dsDATA.Tables.Count - 1
            Sub_dsDATA.Tables(i).Clear()
            Sub_dsDATA.Tables(i).Constraints.Clear()
            For j As Integer = Sub_dsDATA.Tables(i).Columns.Count - 1 To 0 Step -1
                Sub_dsDATA.Tables(i).Columns.RemoveAt(j)
            Next
        Next

        'DBからマスタ一覧を取得
        '【table】マスタ名一覧
        '【table1】マスタ名と主キー一覧
        '【table2】履歴テーブル

        Dim L_Query As String = ""

        '起動元によってクエリを編集する
        Select Case My.Forms.PrimaryForm.Subkb
            Case "工数"
                'データベースを選択
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                            & "integrated security=SSPI;" _
                            & "Initial Catalog=S開発アサイン管理;"
                cn.Open()

                L_Query = ""
                L_Query &= " SELECT		工数=SUM(割振工数),担当者=(CASE WHEN 担当区分 = '1' THEN 開発担当者名 ELSE 開発担当者名 + '(SEPG)' END),更新日時=MAX(更新日時)"
                L_Query &= " FROM        S開発品質管理.dbo.T_プログラムマスタ PGM"
                L_Query &= "             LEFT OUTER JOIN 担当者マスタ TNM ON TNM.担当者コード = PGM.開発担当者コード"
                L_Query &= " WHERE		受注NO='" & L_Jdnno & "'"
                L_Query &= " GROUP BY	開発担当者コード,開発担当者名,担当区分"
            Case "対応"
                'データベースを選択
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                            & "integrated security=SSPI;" _
                            & "Initial Catalog=S開発アサイン管理;"
                cn.Open()
                L_Query = ""
                L_Query &= " SELECT"
                L_Query &= "        開発担当者名"
                L_Query &= "       ,障害工程"
                L_Query &= "       ,対応残=COUNT(障害工程)"
                L_Query &= " FROM   S開発品質管理.dbo.T_障害トラン"
                L_Query &= " WHERE  状況<>'完了' AND 状況<>'問題なし' AND 受注NO='" & L_Jdnno & "'"
                L_Query &= " GROUP BY 障害工程,開発担当者名"

            Case "PG一覧"
                'データベースを選択
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                            & "integrated security=SSPI;" _
                            & "Initial Catalog=S開発品質管理;"
                cn.Open()
                L_Query = ""
                L_Query &= "SELECT PGNO=プログラムNO"
                L_Query &= "      ,PGID=プログラムID"
                L_Query &= "	  ,PG名=プログラム名"
                L_Query &= "	  ,開発担当者名"
                L_Query &= "	  ,割振工数"
                L_Query &= "	  ,開発完了日"
                L_Query &= "	  ,SE引渡日 "
                L_Query &= "  FROM T_プログラムマスタ"
                L_Query &= " WHERE 受注NO = '" & L_Jdnno & "'"
                L_Query &= " ORDER BY プログラムNO"

            Case "生産高_月"
                'データベースを選択
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                            & "integrated security=SSPI;" _
                            & "Initial Catalog=S開発品質管理;"
                cn.Open()
                Dim L_Taidt As String
                Dim dt As DateTime = New DateTime(Integer.Parse(PrimaryForm.Taiym.ToString), 1, 1)
                If PrimaryForm.Taidt >= "08" Then
                    L_Taidt = dt.ToString("yyyy")
                Else
                    L_Taidt = dt.AddYears(1).ToString("yyyy")
                End If
                L_Taidt = L_Taidt + PrimaryForm.Taidt

                Dim L_PGdt As String = ""
                Dim L_SEdt As String = ""
                Dim L_DCdt As String = ""

                If PrimaryForm.T5_予定を含む.Checked Then
                    L_PGdt = "ISNULL(NULLIF(P.開発完了日,''),P.開発予定日)"
                    L_SEdt = "ISNULL(NULLIF(P.SE引渡日,''),P.SE引渡予定日)"
                    L_DCdt = "ISNULL(NULLIF(P.仕様書作成完了日,''),P.仕様書作成予定日)"
                Else
                    L_PGdt = "P.開発完了日"
                    L_SEdt = "P.SE引渡日"
                    L_DCdt = "P.仕様書作成完了日"
                End If

                L_Query = " SELECT 物件名 = RTRIM(LTRIM(物件名))"
                L_Query &= "      ,仕様書"
                L_Query &= "      ,PG完了 "
                L_Query &= "      ,SE渡し "
                L_Query &= "      ,担当 "
                L_Query &= "      ,担当者名 "
                L_Query &= " FROM ("
                L_Query &= " SELECT [物件名]     = REPLACE(MAX(A.案件名),'株式会社','')"
                L_Query &= " 	   ,[仕様書]     = SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 END,0)) "
                L_Query &= "       ,[PG完了]     = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND  T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 END,0)) "
                L_Query &= " 	                 + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND  T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 END,0))  "
                L_Query &= " 	                 + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND  T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 END,0)) "
                L_Query &= " 	   ,[SE渡し]     = SUM(ISNULL(CASE WHEN T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 END,0)) "
                L_Query &= " 	   ,[担当]       =  REVERSE(CONVERT(CHAR(6),REVERSE('000000' + RTRIM(T.担当者コード))))"
                L_Query &= " 	   ,[担当者名]   = MAX(T.担当者名) "
                L_Query &= " from T_プログラムマスタ P"
                L_Query &= " inner join T_案件マスタ A    on A.受注NO        = P.受注NO and A.枝番=P.枝番"
                L_Query &= " inner join T_担当者マスタ T  on (T.担当者コード  = P.開発担当者コード OR P.仕様書作成担当者コード = T.担当者コード) AND T.部署コード <> ''"
                L_Query &= " inner join T_担当者マスタ T2 on T2.担当者コード = A.開発主管者コード"
                L_Query &= " WHERE (REPLACE(P.開発完了日,'/','')  BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.開発予定日,'/','')   BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.仕様書作成完了日,'/','')   BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.仕様書作成予定日,'/','')   BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.SE引渡日,'/','')     BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.SE引渡予定日,'/','') BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99')"
                L_Query &= "    AND REVERSE(CONVERT(CHAR(6),REVERSE('000000' + RTRIM((CASE WHEN T.BP区分 = '0' THEN T.担当者コード ELSE T2.担当者コード END))))) LIKE '%" & PrimaryForm.Tancd & "%'"
                L_Query &= " GROUP BY A.受注NO,T.担当者コード"
                L_Query &= " ) MAIN "
                L_Query &= " WHERE (PG完了+SE渡し+仕様書) > 0"
                L_Query &= " ORDER BY 物件名"
            Case "生産高_四半期"
                'データベースを選択
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                            & "integrated security=SSPI;" _
                            & "Initial Catalog=S開発品質管理;"
                cn.Open()

                Dim Sttdt As String = ""
                Dim Enddt As String = ""

                Dim dt As DateTime = New DateTime(Integer.Parse(PrimaryForm.Taiym.ToString), 1, 1)
                Select Case My.Forms.PrimaryForm.Taidt
                    Case "第一"
                        Sttdt = dt.ToString("yyyy") + "0801"
                        Enddt = dt.ToString("yyyy") + "1099"
                    Case "第二"
                        Sttdt = dt.ToString("yyyy") + "1101"
                        Enddt = dt.AddYears(1).ToString("yyyy") + "0199"
                    Case "第三"
                        Sttdt = dt.AddYears(1).ToString("yyyy") + "0201"
                        Enddt = dt.AddYears(1).ToString("yyyy") + "0499"
                    Case "第四"
                        Sttdt = dt.AddYears(1).ToString("yyyy") + "0501"
                        Enddt = dt.AddYears(1).ToString("yyyy") + "0799"
                End Select


                Dim L_PGdt As String = ""
                Dim L_SEdt As String = ""
                Dim L_DCdt As String = ""

                If PrimaryForm.T5_予定を含む.Checked Then
                    L_PGdt = "ISNULL(NULLIF(P.開発完了日,''),P.開発予定日)"
                    L_SEdt = "ISNULL(NULLIF(P.SE引渡日,''),P.SE引渡予定日)"
                    L_DCdt = "ISNULL(NULLIF(P.仕様書作成完了日,''),P.仕様書作成予定日)"
                Else
                    L_PGdt = "P.開発完了日"
                    L_SEdt = "P.SE引渡日"
                    L_DCdt = "P.仕様書作成完了日"
                End If

                L_Query = "SELECT 物件名 = RTRIM(LTRIM(物件名))"
                L_Query &= "      ,仕様書 "
                L_Query &= "      ,PG完了 "
                L_Query &= "      ,SE渡し "
                L_Query &= "      ,担当 "
                L_Query &= "      ,担当者名 "
                L_Query &= " FROM ("
                L_Query &= " SELECT [物件名]     = REPLACE(MAX(A.案件名),'株式会社','')"
                L_Query &= " 	   ,[仕様書]     = SUM(ISNULL(CASE WHEN P.仕様書作成担当者コード = T.担当者コード AND CONVERT(CHAR(6),REPLACE(" & L_DCdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 END,0)) "
                L_Query &= "       ,[PG完了]     = SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 END,0)) "
                L_Query &= " 	                 + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 END,0))  "
                L_Query &= " 	                 + SUM(ISNULL(CASE WHEN P.開発担当者コード = T.担当者コード AND T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 END,0)) "
                L_Query &= " 	   ,[SE渡し]     = SUM(ISNULL(CASE WHEN T.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 END,0)) "
                L_Query &= " 	   ,[担当]       =  REVERSE(CONVERT(CHAR(6),REVERSE('000000' + RTRIM(T.担当者コード))))"
                L_Query &= " 	   ,[担当者名]   = MAX(T.担当者名) "
                L_Query &= " from T_プログラムマスタ P"
                L_Query &= " inner join T_案件マスタ A    on A.受注NO        = P.受注NO and A.枝番=P.枝番"
                L_Query &= " inner join (T_担当者マスタ T  on T.担当者コード  = P.開発担当者コード OR P.仕様書作成担当者コード = T.担当者コード) AND T.部署コード <> ''"
                L_Query &= " inner join T_担当者マスタ T2 on T2.担当者コード = A.開発主管者コード"
                L_Query &= " WHERE (REPLACE(P.開発完了日,'/','')  BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.開発予定日,'/','')   BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.仕様書作成完了日,'/','')   BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.仕様書作成予定日,'/','')   BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.SE引渡日,'/','')     BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.SE引渡予定日,'/','') BETWEEN '" & Sttdt & "' AND '" & Enddt & "')"
                L_Query &= "    AND REVERSE(CONVERT(CHAR(6),REVERSE('000000' + RTRIM((CASE WHEN T.BP区分 = '0' THEN T.担当者コード ELSE T2.担当者コード END))))) LIKE '%" & PrimaryForm.Tancd & "%'"
                L_Query &= " GROUP BY A.受注NO,T.担当者コード"
                L_Query &= " ) MAIN "
                L_Query &= " WHERE (PG完了+SE渡し+仕様書) > 0"
                L_Query &= " ORDER BY 物件名"
            Case "チーム生産高_月"
                'データベースを選択
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                            & "integrated security=SSPI;" _
                            & "Initial Catalog=S開発品質管理;"
                cn.Open()
                Dim L_Taidt As String
                Dim dt As DateTime = New DateTime(Integer.Parse(PrimaryForm.Taiym.ToString), 1, 1)
                If PrimaryForm.Taidt >= "08" Then
                    L_Taidt = dt.ToString("yyyy")
                Else
                    L_Taidt = dt.AddYears(1).ToString("yyyy")
                End If
                L_Taidt = L_Taidt + PrimaryForm.Taidt

                Dim L_PGdt As String = ""
                Dim L_SEdt As String = ""

                If PrimaryForm.T5_予定を含む.Checked Then
                    L_PGdt = "ISNULL(NULLIF(P.開発完了日,''),P.開発予定日)"
                    L_SEdt = "ISNULL(NULLIF(P.SE引渡日,''),P.SE引渡予定日)"
                Else
                    L_PGdt = "P.開発完了日"
                    L_SEdt = "P.SE引渡日"
                End If

                L_Query = " SELECT * FROM("
                L_Query &= "  SELECT PGM.担当者名"
                L_Query &= " 	    ,金額 = SUM(ISNULL(CASE WHEN PGM.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 * 50000 END,0))"
                L_Query &= " 	          + SUM(ISNULL(CASE WHEN PGM.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 * 15000 END,0))"
                L_Query &= " 	          + SUM(ISNULL(CASE WHEN PGM.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 * 12500 END,0))"
                L_Query &= " 	          + SUM(ISNULL(CASE WHEN PGM.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) ='" & L_Taidt & "' THEN P.割振工数 * 12500 END,0))"
                L_Query &= " 	from T_プログラムマスタ P"
                L_Query &= " 	INNER JOIN T_案件マスタ A    ON A.受注NO = P.受注NO and A.枝番=P.枝番"
                L_Query &= " 	LEFT OUTER JOIN T_担当者マスタ PGM ON PGM.担当者コード = P.開発担当者コード AND PGM.部署コード <> ''"
                L_Query &= " 	LEFT OUTER JOIN T_担当者マスタ ERP ON ERP.担当者コード = A.開発主管者コード"
                L_Query &= " 	LEFT OUTER JOIN T_部署マスタ BMN  ON PGM.部署コード = BMN.部署コード "
                L_Query &= " 	LEFT OUTER JOIN T_部署マスタ BMN2  ON ERP.部署コード = BMN2.部署コード "
                L_Query &= " WHERE (REPLACE(P.開発完了日,'/','')  BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.開発予定日,'/','')   BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.SE引渡日,'/','')     BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99' "
                L_Query &= "    OR REPLACE(P.SE引渡予定日,'/','') BETWEEN '" & L_Taidt & "01' AND '" & L_Taidt & "99')"
                L_Query &= "    AND BMN2.部署名 = '" & PrimaryForm.Bmnnm & "'"

                L_Query &= " GROUP BY PGM.担当者名"
                L_Query &= " ) MAIN WHERE MAIN.金額 > 0 ORDER BY 金額 DESC"
            Case "チーム生産高_四半期"
                'データベースを選択
                cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                            & "integrated security=SSPI;" _
                            & "Initial Catalog=S開発品質管理;"
                cn.Open()

                Dim Sttdt As String = ""
                Dim Enddt As String = ""
                Dim dt As DateTime = New DateTime(Integer.Parse(PrimaryForm.Taiym.ToString), 1, 1)
                Select Case My.Forms.PrimaryForm.Taidt
                    Case "第一"
                        Sttdt = dt.ToString("yyyy") + "0801"
                        Enddt = dt.ToString("yyyy") + "1099"
                    Case "第二"
                        Sttdt = dt.ToString("yyyy") + "1101"
                        Enddt = dt.AddYears(1).ToString("yyyy") + "0199"
                    Case "第三"
                        Sttdt = dt.AddYears(1).ToString("yyyy") + "0201"
                        Enddt = dt.AddYears(1).ToString("yyyy") + "0499"
                    Case "第四"
                        Sttdt = dt.AddYears(1).ToString("yyyy") + "0501"
                        Enddt = dt.AddYears(1).ToString("yyyy") + "0799"
                End Select


                Dim L_PGdt As String = ""
                Dim L_SEdt As String = ""

                If PrimaryForm.T5_予定を含む.Checked Then
                    L_PGdt = "ISNULL(NULLIF(P.開発完了日,''),P.開発予定日)"
                    L_SEdt = "ISNULL(NULLIF(P.SE引渡日,''),P.SE引渡予定日)"
                Else
                    L_PGdt = "P.開発完了日"
                    L_SEdt = "P.SE引渡日"
                End If

                L_Query = " SELECT * FROM("
                L_Query &= "  SELECT PGM.担当者名"
                L_Query &= " 	    ,金額 = SUM(ISNULL(CASE WHEN PGM.BP区分 = '0' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 * 50000 END,0))"
                L_Query &= " 	          + SUM(ISNULL(CASE WHEN PGM.BP区分 = '1' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 * 15000 END,0))"
                L_Query &= " 	          + SUM(ISNULL(CASE WHEN PGM.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_PGdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 * 12500 END,0))"
                L_Query &= " 	          + SUM(ISNULL(CASE WHEN PGM.BP区分 = '2' AND CONVERT(CHAR(6),REPLACE(" & L_SEdt & ",'/','')) BETWEEN '" & Sttdt.Substring(0, 6) & "' AND '" & Enddt.Substring(0, 6) & "' THEN P.割振工数 * 12500 END,0))"
                L_Query &= " 	from T_プログラムマスタ P"
                L_Query &= " 	INNER JOIN T_案件マスタ A    ON A.受注NO = P.受注NO and A.枝番=P.枝番"
                L_Query &= " 	LEFT OUTER JOIN T_担当者マスタ PGM ON PGM.担当者コード = P.開発担当者コード AND PGM.部署コード <> ''"
                L_Query &= " 	LEFT OUTER JOIN T_担当者マスタ ERP ON ERP.担当者コード = A.開発主管者コード"
                L_Query &= " 	LEFT OUTER JOIN T_部署マスタ BMN  ON PGM.部署コード = BMN.部署コード "
                L_Query &= " 	LEFT OUTER JOIN T_部署マスタ BMN2  ON ERP.部署コード = BMN2.部署コード "
                L_Query &= " WHERE (REPLACE(P.開発完了日,'/','')  BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.開発予定日,'/','')   BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.SE引渡日,'/','')     BETWEEN '" & Sttdt & "' AND '" & Enddt & "' "
                L_Query &= "    OR REPLACE(P.SE引渡予定日,'/','') BETWEEN '" & Sttdt & "' AND '" & Enddt & "')"
                L_Query &= "    AND BMN2.部署名 = '" & PrimaryForm.Bmnnm & "'"
                L_Query &= " GROUP BY PGM.担当者名"
                L_Query &= " ) MAIN "
                L_Query &= " WHERE 金額 > 0"
                L_Query &= " ORDER BY 金額 DESC"

            Case "原価内訳"
                'データベースを選択
                cn.ConnectionString = "Data Source=ic01\share;" _
                                    & "Trusted_Connection = Yes;" _
                                    & "Initial Catalog=原価管理link;"
                cn.Open()
                L_Query = ""
                L_Query &= "  SELECT" & vbCrLf
                L_Query &= "   担当者コード" & vbCrLf
                L_Query &= "  ,担当者名" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=ISNULL(SUM(作業時間),0)" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(実原価金額),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 集計コード IN ('110')" & vbCrLf
                L_Query &= "    AND 内容コード IN ('103','016','120','128','022')" & vbCrLf
                L_Query &= "    AND (担当者コード < '009592'" & vbCrLf
                L_Query &= "      OR 担当者コード > '009601') " & vbCrLf
                L_Query &= "    AND 担当者名 NOT LIKE '%大阪委託%' " & vbCrLf
                L_Query &= "  GROUP BY 担当者コード,担当者名" & vbCrLf
                L_Query &= "  UNION" & vbCrLf
                L_Query &= "  SELECT" & vbCrLf
                L_Query &= "   担当者コード='009068'" & vbCrLf
                L_Query &= "  ,担当者名='Wits大連'" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=ISNULL(SUM(作業時間),0)" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(実原価金額),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 集計コード IN ('110')" & vbCrLf
                L_Query &= "    AND 内容コード IN ('103','016','120','128','022')" & vbCrLf
                L_Query &= "    AND (担当者コード BETWEEN '009592' AND '009601' " & vbCrLf
                L_Query &= "     OR 担当者名 LIKE '%大阪委託%') " & vbCrLf
                L_Query &= "  GROUP BY 受注NO" & vbCrLf
                L_Query &= "  UNION" & vbCrLf
                L_Query &= " SELECT" & vbCrLf
                L_Query &= "   担当者コード = '@1'" & vbCrLf
                L_Query &= "  ,担当者名     = '仕様書作成'" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=ISNULL(SUM(作業時間),0)" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(実原価金額),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 部門コード = '000003'" & vbCrLf
                L_Query &= "    AND 内容コード IN ('102','015') " & vbCrLf
                L_Query &= "  UNION" & vbCrLf
                L_Query &= " SELECT" & vbCrLf
                L_Query &= "   担当者コード = '@2'" & vbCrLf
                L_Query &= "  ,担当者名     = '外注費'" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=0" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(原価金額),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 集計コード IN ('59')" & vbCrLf
                L_Query &= "    AND 内容コード IN ('901')" & vbCrLf
                L_Query &= "  UNION" & vbCrLf
                L_Query &= " SELECT" & vbCrLf
                L_Query &= "   担当者コード = '@3'" & vbCrLf
                L_Query &= "  ,担当者名     = 'コントロール'" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=ISNULL(SUM(作業時間),0)" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(実原価金額),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 集計コード IN ('110')" & vbCrLf
                L_Query &= "    AND 内容コード IN ('123','126','127','018','020')" & vbCrLf
                L_Query &= "  UNION" & vbCrLf
                L_Query &= "  SELECT" & vbCrLf
                L_Query &= "   担当者コード = '@4'" & vbCrLf
                L_Query &= "  ,担当者名     = '仕様変更'" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=ISNULL(SUM(作業時間),0)" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(実原価金額),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 集計コード IN ('110')" & vbCrLf
                L_Query &= "    AND 内容コード IN ('125','024')" & vbCrLf
                L_Query &= "  UNION" & vbCrLf
                L_Query &= " SELECT" & vbCrLf
                L_Query &= "   担当者コード = '@5'" & vbCrLf
                L_Query &= "  ,担当者名     = 'テスターテスト'" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=ISNULL(SUM(作業時間),0)" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(実原価金額),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 集計コード IN ('110')" & vbCrLf
                L_Query &= "    AND 内容コード IN ('121','017')" & vbCrLf
                L_Query &= "  UNION" & vbCrLf
                L_Query &= " SELECT" & vbCrLf
                L_Query &= "   担当者コード = '@9'" & vbCrLf
                L_Query &= "  ,担当者名     = '合計'" & vbCrLf
                L_Query &= "  ,割振工数=0.00" & vbCrLf
                L_Query &= "  ,作業時間=ISNULL(SUM(CASE 内容コード WHEN '015' THEN 0 WHEN '025' THEN 0 WHEN '131' THEN 0 ELSE 作業時間 END),0)" & vbCrLf
                L_Query &= "  ,原価金額=ISNULL(SUM(CASE 内容コード WHEN '901' THEN 原価金額 WHEN '015' THEN 0 WHEN '131' THEN 0 WHEN '025' THEN 0 ELSE 実原価金額 END),0)" & vbCrLf
                L_Query &= "  ,効率=0.00" & vbCrLf
                L_Query &= "  ,粗利=0.00" & vbCrLf
                L_Query &= "   FROM 原価トランビュー GNK" & vbCrLf
                L_Query &= "  WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                L_Query &= "    AND 集計コード IN ('110','59')" & vbCrLf
            Case "負荷状況"

        End Select
        Try
            Dim daAuthors As New SqlDataAdapter(L_Query, cn)
            daAuthors.FillSchema(dsDATA, SchemaType.Source)
            daAuthors.Fill(dsDATA)
            D_DataGridView.DataMember = dsDATA.Tables("table").TableName
            D_DataGridView.DataSource = dsDATA
            D_DataGridView.Height = (dsDATA.Tables("table").Rows.Count * 21) + 23
            D_DataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
            Select Case My.Forms.PrimaryForm.Subkb
                Case "工数"
                    D_DataGridView.Columns(0).Width = 70
                    D_DataGridView.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(0).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(1).Width = 150
                Case "対応"
                    D_DataGridView.Width = 298
                    D_DataGridView.Columns(0).Width = 150
                    D_DataGridView.Columns(1).Width = 80
                    D_DataGridView.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(2).DefaultCellStyle.Format = "##0"
                    D_DataGridView.Columns(2).Width = 68
                Case "PG一覧"
                    D_DataGridView.Width = 590
                    D_DataGridView.Columns(0).Width = 40
                    D_DataGridView.Columns(1).Width = 80
                    D_DataGridView.Columns(2).Width = 150
                    D_DataGridView.Columns(3).Width = 90
                    D_DataGridView.Columns(4).Width = 70
                    D_DataGridView.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(4).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(5).Width = 80
                    D_DataGridView.Columns(6).Width = 80
                Case "生産高_月", "生産高_四半期"
                    D_DataGridView.Width = 478
                    D_DataGridView.Columns(0).Width = 120
                    D_DataGridView.Columns(1).Width = 74
                    D_DataGridView.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(1).DefaultCellStyle.Format = "#,#0.0"
                    D_DataGridView.Columns(2).Width = 68
                    D_DataGridView.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(2).DefaultCellStyle.Format = "#,#0.0"
                    D_DataGridView.Columns(3).Width = 68
                    D_DataGridView.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(3).DefaultCellStyle.Format = "#,#0.0"
                    D_DataGridView.Columns(4).Width = 68
                    D_DataGridView.Columns(5).Width = 80
                Case "チーム生産高_月", "チーム生産高_四半期"
                    D_DataGridView.Width = 188
                    D_DataGridView.Columns(0).Width = 120
                    D_DataGridView.Columns(1).Width = 68
                    D_DataGridView.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(1).DefaultCellStyle.Format = "#,#0"
                Case "原価内訳"
                    D_DataGridView.Width = 500
                    D_DataGridView.Columns(0).Width = 80
                    D_DataGridView.Columns(1).Width = 70
                    D_DataGridView.Columns(2).Width = 70
                    D_DataGridView.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(2).DefaultCellStyle.Format = "#,#0.0"
                    D_DataGridView.Columns(3).Width = 70
                    D_DataGridView.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(3).DefaultCellStyle.Format = "#,#0.0"
                    D_DataGridView.Columns(4).Width = 70
                    D_DataGridView.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(4).DefaultCellStyle.Format = "#,#0"
                    D_DataGridView.Columns(5).Width = 70
                    D_DataGridView.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(5).DefaultCellStyle.Format = "#,#0.0%"
                    D_DataGridView.Columns(6).Width = 70
                    D_DataGridView.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(6).DefaultCellStyle.Format = "#,#0.0%"


                    'データベースを選択
                    Sub_cn.ConnectionString = "Data Source=KDC-O-SE01\s_kaihatsu;" _
                                & "integrated security=SSPI;" _
                                & "Initial Catalog=S開発品質管理;"
                    Sub_cn.Open()

                    L_Query = ""
                    L_Query &= "SELECT 担当者コード=REVERSE(CONVERT(VARCHAR(6),REVERSE('000000'+CONVERT(VARCHAR,開発担当者コード))))" & vbCrLf
                    L_Query &= "      ,担当者名=開発担当者名" & vbCrLf
                    L_Query &= "	  ,割振工数=SUM(割振工数)" & vbCrLf
                    L_Query &= "	  ,BP区分  = CASE WHEN TNM.BP区分 = '0' THEN '社員' "
                    L_Query &= "                      WHEN TNM.BP区分 = '1' AND 請負契約 = '0' THEN '社内' " & vbCrLf
                    L_Query &= "                      WHEN TNM.BP区分 = '1' AND 請負契約 = '1' THEN '社外' " & vbCrLf
                    L_Query &= "                      ELSE '社外' END " & vbCrLf
                    L_Query &= "  FROM T_プログラムマスタ PG" & vbCrLf
                    L_Query &= "	   LEFT OUTER JOIN T_担当者マスタ TNM ON PG.開発担当者コード = TNM.担当者コード" & vbCrLf
                    L_Query &= "   WHERE 受注NO = '" & L_Jdnno & "'" & vbCrLf
                    L_Query &= "   AND 割振工数 > 0" & vbCrLf
                    L_Query &= " GROUP BY 開発担当者コード,開発担当者名,TNM.BP区分,TNM.請負契約" & vbCrLf
                    Dim sub_daAuthors As New SqlDataAdapter(L_Query, Sub_cn)
                    sub_daAuthors.FillSchema(Sub_dsDATA, SchemaType.Source)
                    sub_daAuthors.Fill(Sub_dsDATA)

                    'テーブルを編集可能にする
                    For Each col As DataColumn In dsDATA.Tables(0).Columns
                        col.ReadOnly = False
                    Next


                    Dim kosu As Decimal = 0
                    Dim kingaku As Decimal = 0

                    If Sub_dsDATA.Tables(0) IsNot Nothing AndAlso Sub_dsDATA.Tables(0).Rows.Count > 0 Then
                        For Each mainrow As DataGridViewRow In D_DataGridView.Rows


                            kosu = 0
                            For Each subrow As DataRow In Sub_dsDATA.Tables(0).Rows
                                Select Case subrow("BP区分")
                                    Case "社員", "社内"
                                        If mainrow.Cells("担当者コード").Value.ToString() = subrow("担当者コード").ToString() Then
                                            mainrow.Cells("割振工数").Value = subrow("割振工数")
                                        End If
                                    Case "社外"
                                        'wits委託のみ例外
                                        If mainrow.Cells("担当者コード").Value.ToString() = "009068" AndAlso _
                                            mainrow.Cells("担当者コード").Value.ToString() = subrow("担当者コード").ToString() Then
                                            mainrow.Cells("割振工数").Value = subrow("割振工数")
                                        ElseIf subrow("担当者コード").ToString() <> "009068" AndAlso _
                                               mainrow.Cells("担当者名").Value.ToString() = "外注費" Then
                                            mainrow.Cells("割振工数").Value = subrow("割振工数")
                                        End If
                                End Select
                                If mainrow.Cells("割振工数").Value <> 0 Then
                                    If mainrow.Cells("作業時間").Value = 0 Then
                                        mainrow.Cells("効率").Value = 0
                                    Else
                                        mainrow.Cells("効率").Value = (mainrow.Cells("割振工数").Value * 7.5) / CDec(mainrow.Cells("作業時間").Value)
                                    End If
                                End If
                                If Not mainrow.Cells("割振工数").Value = 0 Then
                                    mainrow.Cells("粗利").Value = ((mainrow.Cells("割振工数").Value * 40000) - mainrow.Cells("原価金額").Value) / CDec((mainrow.Cells("割振工数").Value * 40000))
                                End If
                                kosu += subrow("割振工数")
                            Next

                            '集計レコード
                            If mainrow.Cells("担当者名").Value = "合計" Then
                                mainrow.Cells("割振工数").Value = kosu
                                'mainrow.Cells("原価金額").Value = kingaku
                                If mainrow.Cells("原価金額").Value <> 0 Then
                                    mainrow.Cells("効率").Value = (mainrow.Cells("割振工数").Value * 40000) / mainrow.Cells("原価金額").Value
                                End If
                                If mainrow.Cells("割振工数").Value <> 0 Then
                                    mainrow.Cells("粗利").Value = ((mainrow.Cells("割振工数").Value * 40000) - mainrow.Cells("原価金額").Value) / (mainrow.Cells("割振工数").Value * 40000)
                                End If
                            End If
                            If mainrow.Cells("担当者名").Value <> "仕様書作成" Then
                                kingaku += mainrow.Cells("原価金額").Value
                            End If
                        Next
                    End If

                    For i As Integer = D_DataGridView.Rows.Count - 1 To 0 Step -1
                        '表示不要な行を削除する
                        With D_DataGridView.Rows(i)
                            If .Cells("割振工数").Value = 0 _
                               AndAlso .Cells("原価金額").Value = 0 _
                               AndAlso .Cells("作業時間").Value = 0 Then
                                D_DataGridView.Rows.Remove(D_DataGridView.Rows(i))
                            End If
                        End With
                    Next

                    '最後に表示を整形する
                    For Each mainrow As DataGridViewRow In D_DataGridView.Rows
                        If mainrow.Cells("割振工数").Value = 0 OrElse mainrow.Cells("原価金額").Value = 0 OrElse mainrow.Cells("担当者コード").Value = "" Then
                            mainrow.Cells("粗利").Value = DBNull.Value
                            mainrow.Cells("効率").Value = DBNull.Value
                        End If

                        If mainrow.Cells("割振工数").Value = 0 Then
                            mainrow.Cells("割振工数").Value = DBNull.Value
                        End If

                        If mainrow.Cells("作業時間").Value = 0 Then
                            mainrow.Cells("作業時間").Value = DBNull.Value
                        End If

                        If Not IsDBNull(mainrow.Cells("効率").Value) AndAlso mainrow.Cells("効率").Value = 0 Then
                            mainrow.Cells("効率").Value = DBNull.Value
                        End If

                        If mainrow.Cells("担当者コード").Value.ToString.Contains("@") Then
                            mainrow.Cells("担当者コード").Value = ""
                        End If

                        '外注費の効率を求める
                        If mainrow.Cells("担当者名").Value = "外注費" Then
                            If Not IsDBNull(mainrow.Cells("割振工数").Value) AndAlso mainrow.Cells("割振工数").Value <> 0 Then
                                If Not IsDBNull(mainrow.Cells("原価金額").Value) AndAlso mainrow.Cells("原価金額").Value = 0 Then
                                    mainrow.Cells("効率").Value = 0
                                Else
                                    mainrow.Cells("効率").Value = (mainrow.Cells("割振工数").Value * 40000) / CDec(mainrow.Cells("原価金額").Value)
                                End If
                            End If
                            If Not IsDBNull(mainrow.Cells("割振工数").Value) AndAlso Not mainrow.Cells("割振工数").Value = 0 Then
                                mainrow.Cells("粗利").Value = ((mainrow.Cells("割振工数").Value * 40000) - mainrow.Cells("原価金額").Value) / CDec((mainrow.Cells("割振工数").Value * 40000))
                            End If
                        End If

                        'テスターテストの効率を求める
                        If mainrow.Cells("担当者名").Value = "テスターテスト" Then
                            If Not IsDBNull(mainrow.Cells("原価金額").Value) AndAlso mainrow.Cells("原価金額").Value <> 0 Then
                                mainrow.Cells("粗利").Value = ((kosu * 8000) - mainrow.Cells("原価金額").Value) / CDec(kosu * 8000)
                                mainrow.Cells("効率").Value = (kosu * 8000) / mainrow.Cells("原価金額").Value
                            End If
                        End If
                    Next
                Case "負荷状況"
                    D_DataGridView.Width = 790
                    D_DataGridView.Columns(0).Width = 70
                    D_DataGridView.Columns(1).Width = 70
                    D_DataGridView.Columns(2).Width = 50
                    D_DataGridView.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(2).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(3).Width = 50
                    D_DataGridView.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(3).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(4).Width = 50
                    D_DataGridView.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(4).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(4).Width = 50
                    D_DataGridView.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(5).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(5).Width = 50
                    D_DataGridView.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(6).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(6).Width = 50
                    D_DataGridView.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(7).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(7).Width = 50
                    D_DataGridView.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(8).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(8).Width = 50
                    D_DataGridView.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(9).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(9).Width = 50
                    D_DataGridView.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(10).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(10).Width = 50
                    D_DataGridView.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(11).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(11).Width = 50
                    D_DataGridView.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(12).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(12).Width = 50
                    D_DataGridView.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(13).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(13).Width = 50
                    D_DataGridView.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                    D_DataGridView.Columns(14).DefaultCellStyle.Format = "##0.0"
                    D_DataGridView.Columns(14).Width = 50


            End Select

            If (D_DataGridView.Rows.Count > 10) Then
                D_DataGridView.Height = (10 * 21) + 22
                D_DataGridView.ScrollBars = ScrollBars.Vertical
                D_DataGridView.Width += 20
            Else
                D_DataGridView.Height = (D_DataGridView.Rows.Count * 21) + 23
            End If

            Me.Height = D_DataGridView.Height
            Me.Width = D_DataGridView.Width
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkCancel, "クエリエラー")
        End Try
    End Sub

    Private Sub SubdataForm_Leave(sender As Object, e As EventArgs) Handles MyBase.Leave, D_DataGridView.Leave
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub SubdataForm_Deactivate(sender As Object, e As EventArgs) Handles Me.Deactivate
        Me.Close()
        Me.Dispose()
    End Sub

End Class