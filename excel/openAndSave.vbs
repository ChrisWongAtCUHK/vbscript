' Current directory
dim fso: set fso = CreateObject("Scripting.FileSystemObject")
currDir = fso.GetAbsolutePathName(".")

' Open and change
Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(currDir & "\testIn.xlsx")
objExcel.Application.DisplayAlerts = False
objExcel.Application.Visible = True
objExcel.Cells(2, 1).Value = "First value"
objExcel.Cells(2, 2).Value = "Second value"

'Save and close
objExcel.ActiveWorkbook.SaveAs currDir & "\testOut.xlsx"
objExcel.ActiveWorkbook.Close