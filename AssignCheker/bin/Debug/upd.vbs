Dim excelApp : Set excelApp = CreateObject("Excel.Application")
' Excel���\���ɂ���
excelApp.Visible = False
Dim targetFile : targetFile = WScript.Arguments(0)

' Excel�t�@�C�����J��
excelApp.Workbooks.Open targetFile
excelApp.DisplayAlerts = false
excelApp.ActiveWorkBook.UnprotectSharing

excelApp.Worksheets("�Č���{���").Select
excelApp.Worksheets("�Č���{���").Cells(1, 1).Value = "1"

' �}�N���̎��s
' PG�i���o�^
excelApp.Worksheets("PG�i��").Select
excelApp.Run "Sheet5.cmdUpdateData_Click"

excelApp.Worksheets("�Č���{���").Select
excelApp.Worksheets("�Č���{���").Cells(1, 1).Value = ""

' Excel�̏I��
excelApp.Quit