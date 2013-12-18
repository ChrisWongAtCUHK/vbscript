extension="TXT"
Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "."

Set objFolder = objFSO.GetFolder(objStartFolder)
Wscript.Echo objFolder.Path

Set colFiles = objFolder.Files


For Each objFile in colFiles
If UCase(objFSO.GetExtensionName(objFile.name)) = extension Then
	gd = grabDigit(objFile.Name)
	If StrComp(gd, "") Then
		WScript.Echo gd & "." & extension
		'objFSO.CopyFile objFile.Name, gd & "." & extension
		objFSO.MoveFile objFile.Name, gd & "." & extension
	End If
    End If
Next

'Grad the digits of first appearance
Function grabDigit(str)
	digitStr=""
	strLen = Len(str)
	hasGrab = False
	For i=1 to strLen
		If IsNumeric(Mid(str, i, 1)) Then
			hasGrab = True
			digitStr=digitStr & Mid(str, i, 1)
		ElseIf hasGrab = True Then
			Exit For
		End If
	Next

	grabDigit = digitStr
End Function




