Dim excelApp : Set excelApp = CreateObject("Excel.Application")
' Excelを非表示にする
excelApp.Visible = False
Dim targetFile : targetFile = WScript.Arguments(0)

' Excelファイルを開く
excelApp.Workbooks.Open targetFile
excelApp.DisplayAlerts = false
excelApp.ActiveWorkBook.UnprotectSharing

excelApp.Worksheets("案件基本情報").Select
excelApp.Worksheets("案件基本情報").Cells(1, 1).Value = "1"

' マクロの実行
' PG進捗登録
excelApp.Worksheets("PG進捗").Select
excelApp.Run "Sheet5.cmdUpdateData_Click"

excelApp.Worksheets("案件基本情報").Select
excelApp.Worksheets("案件基本情報").Cells(1, 1).Value = ""

' Excelの終了
excelApp.Quit