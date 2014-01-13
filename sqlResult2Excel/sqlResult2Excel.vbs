' http://www.visualbasicscript.com/Reading-a-txtfile-with-VBScript-m1510.aspx
' cscript /nologo sqlResult2Excel.vbs
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile ("test.txt", ForReading)

Const ForReading = 1

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
For i = 0 to UBound(wordsStr)
	Redim Preserve colLengthes(i)
	colLengthes(i) = Len(wordsStr(i))
next

' Loop through the contents, i.e. the result from sql query
Do Until objFile.AtEndOfStream
	row = objFile.ReadLine
	' Split each line to colCnt column, with colLengthes
	substringLen = 0
	startIndex = 1
	For i = 0 to UBound(wordsStr)
		substringLen = substringLen + colLengthes(i)
		WScript.Echo Mid(row, startIndex, substringLen)
		startIndex = substringLen
	next
	WScript.Echo row
Loop
objFile.Close
