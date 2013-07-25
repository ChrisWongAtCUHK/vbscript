dim fso: set fso = CreateObject("Scripting.FileSystemObject")

Set objExcel = CreateObject("Excel.Application")
' Trick: directory in which this script is currently running
currentDirectory = fso.GetAbsolutePathName(".")

Set objWorkbook = objExcel.Workbooks.Open _
    (currentDirectory & "\server_ports.xlsx")

' Start from row 2
intRow = 2

Do Until objExcel.Cells(intRow,1).Value = ""
    Wscript.Echo "Server Name: " & objExcel.Cells(intRow, 1).Value
    Wscript.Echo "Server port: " & objExcel.Cells(intRow, 2).Value
    intRow = intRow + 1
Loop

objExcel.Quit