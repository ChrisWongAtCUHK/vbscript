Set objFS=CreateObject("Scripting.FileSystemObject")
count=0
strFile = "k-lunch.txt"
Set objFile = objFS.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
	strLine= objFile.ReadLine
	'Wscript.Echo strLine
	s=Split(strLine,",")
	For i=LBound(s) To UBound(s)
		'Remove leading spaces
		WScript.Echo LTrim(s(i))
		count = count + 1
	Next
Loop
WScript.Echo ""
WScript.Echo count & "pp will go to k-lunch"
objFile.Close
