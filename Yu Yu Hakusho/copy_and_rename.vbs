extension="TXT"
Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "."

Set objFolder = objFSO.GetFolder(objStartFolder)
Wscript.Echo objFolder.Path

Set colFiles = objFolder.Files


For Each objFile in colFiles
If UCase(objFSO.GetExtensionName(objFile.name)) = extension Then
    'Wscript.Echo objFile.Name & " " & a & ".txt"
	gd = grabDigit(objFile.Name)
	WScript.Echo gd
	'objFSO.CopyFile objFile.Name, gd & ".txt"
    End If
Next

'Grad the digits of first appearance
Function grabDigit(str)
	digitStr=""
	strLen = Len(str)
	for i=1 to strLen
		If IsNumeric(Mid(str, i, 1)) Then
			digitStr=digitStr & Mid(str, i, 1)
		End If
	Next

	grabDigit = digitStr
End Function




