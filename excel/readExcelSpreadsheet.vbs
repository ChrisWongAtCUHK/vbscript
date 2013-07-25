Set objExcel = CreateObject("Excel.Application")
' Must be full path...
Set objWorkbook = objExcel.Workbooks.Open _
    ("D:\github\vbscript\excel\server_ports.xlsx")

' Start from row 2
intRow = 2

Do Until objExcel.Cells(intRow,1).Value = ""
    Wscript.Echo "Server Name: " & objExcel.Cells(intRow, 1).Value
    Wscript.Echo "Server port: " & objExcel.Cells(intRow, 2).Value
    intRow = intRow + 1
Loop

objExcel.Quit