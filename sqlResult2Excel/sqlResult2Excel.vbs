' http://www.visualbasicscript.com/Reading-a-txtfile-with-VBScript-m1510.aspx
' cscript /nologo sqlResult2Excel.vbs

If WScript.Arguments.Count <> 2 Then
	WScript.Echo "Usage: cscript /nologo sqlResult2Excel.vbs inputTextFile outputExcelFile"
	WScript.Quit 0
End If

inputTxt   = WScript.Arguments(0)
outputXlsx = WScript.Arguments(1)

' Input text file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile (inputTxt, ForReading)
Const ForReading = 1

' Output excel file
' http://www.activexperts.com/activmonitor/windowsmanagement/scripts/msoffice/excel/
currDir = objFSO.GetAbsolutePathName(".")
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

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
wordsStr = Split(splitLine, " ")										' http://www.ezineasp.net/post/ASP-Vbscript-Split-String-into-Array-Examples.aspx

' Get the lengthes of "-", colLengthes[colCnt]
Dim colLengthes()
startIndex = 1
For i = 0 to UBound(wordsStr)
	Redim Preserve colLengthes(i)
	colLengthes(i) = Len(wordsStr(i))
	
	' set the headers
	header = Mid(headers, startIndex, colLengthes(i))
	startIndex = startIndex + colLengthes(i) + 1 						' dynamic changing the start index of substring, 1 is for spliting space
	header = LTrim(header)												' remove leading spaces, http://www.pctools.com/guides/scripting/detail/78/?act=reference
	header = RTrim(header)												' remove trailing spaces
	objExcel.Cells(1, i + 1).Value = header
	
	' format the cell
	objExcel.Cells(1, i + 1).Font.Bold = TRUE
	objExcel.Cells(1, i + 1).HorizontalAlignment = -4108				' center alignmet
	
	' set the column width of the spread sheet
	objExcel.ActiveSheet.columns(i + 1).columnwidth = colLengthes(i)
next

	
' Loop through the contents, i.e. the result from sql query, until reach a blank line
rowI = 1
Do Until objFile.AtEndOfStream
	row = objFile.ReadLine
	
	' Reach a blank line
	If StrComp(row, "") = 0 Then Exit Do
	' Split each line to colCnt column, with colLengthes
	substringLen = 0
	startIndex = 1
	For colI = 0 to UBound(wordsStr)
		cell = Mid(row, startIndex, colLengthes(colI))						' get the cell content
		startIndex = startIndex + colLengthes(colI) + 1						' dynamic changing the start index of substring, 1 is for spliting space
		cell = LTrim(cell)													' remove leading spaces
		cell = RTrim(cell)													' remove trailing spaces
		objExcel.Cells(rowI + 1, colI + 1).Value = "" & CStr(cell) & ""		' how to format a cell, i.e. not automatic format
		WScript.Echo cell
	next
	rowI = rowI + 1
Loop

' Save to excel file in current directory
objExcel.ActiveWorkbook.SaveAs currDir & "\" & outputXlsx

' Close all files
objExcel.ActiveWorkbook.Close
objExcel.Quit
objFile.Close
