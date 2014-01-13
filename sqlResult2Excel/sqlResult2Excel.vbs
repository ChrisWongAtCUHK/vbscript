' http://www.visualbasicscript.com/Reading-a-txtfile-with-VBScript-m1510.aspx
' cscript /nologo sqlResult2Excel.vbs

' TODO: arguments
inputTxt   = "test.txt"
outputXlsx = "test.xlsx"
' Input text file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile (inputTxt, ForReading)
Const ForReading = 1

' Output excel file
currDir = objFSO.GetAbsolutePathName(".")
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add
'objExcel.Cells(1, 1).Value = "Test value"


' Headers of the sql query
Dim headers
' e.g. "--- --- ----"
Dim splitLine

' Get the headers, line 0
headers = objFile.ReadLine

' Get the splite line 1
splitLine = objFile.ReadLine

' count "-", colCnt
Dim wordsStr
wordsStr = Split(splitLine, " ")

' Get the lengthes of "-", colLengthes[colCnt]
Dim colLengthes()
startIndex = 1
For i = 0 to UBound(wordsStr)
	Redim Preserve colLengthes(i)
	colLengthes(i) = Len(wordsStr(i))
	' set the headers
	header = Mid(headers, startIndex, colLengthes(i))
	startIndex = startIndex + colLengthes(i) + 1 						' dynamic changing the start index of substring, 1 is for spliting space
	header = LTrim(header)												' remove leading spaces
	header = RTrim(header)												' remove trailing spaces
	objExcel.Cells(1, i + 1).Value = header
	' format the cell
	objExcel.Cells(1, i + 1).Font.Bold = TRUE
	objExcel.Cells(1, i + 1).HorizontalAlignment = -4108				' center alignmet
	' set the column widht of the spread sheet
	objExcel.ActiveSheet.columns(i + 1).columnwidth = colLengthes(i)
next

	
' Loop through the contents, i.e. the result from sql query, until reach a blank line
rowI = 1
Do Until objFile.AtEndOfStream
	row = objFile.ReadLine
	' Reach a blank lin
	If StrComp(row, "") = 0 Then Exit Do
	' Split each line to colCnt column, with colLengthes
	substringLen = 0
	startIndex = 1
	For colI = 0 to UBound(wordsStr)
		cell = Mid(row, startIndex, colLengthes(colI))						' get the cell content
		startIndex = startIndex + colLengthes(colI) + 1						' dynamic changing the start index of substring, 1 is for spliting space
		cell = LTrim(cell)													' remove leading spaces
		cell = RTrim(cell)													' remove trailing spaces
		objExcel.Cells(rowI + 1, colI + 1).Value = "" & CStr(cell) & ""		' TODO: how to format a cell, i.e. not automatic format
		WScript.Echo cell
	next
	WScript.Echo row
	rowI = rowI + 1
Loop

' Close all files

objExcel.ActiveWorkbook.SaveAs currDir & "\" & outputXlsx
'objExcel.ActiveWorkbook.Close
'objExcel.Quit
objFile.Close
