Option Explicit

'********************************************************************
'* SortDictionary
'* http://demon.tw/programming/vbs-scripting-dictionary-sort.html
'* Shell sort based on:
'* http://support.microsoft.com/support/kb/articles/q246/0/67.asp
'********************************************************************
Sub SortDictionary(objDict, intSort)

   Const dictKey  = 1
   Const dictItem = 2

   Dim strDict()
   Dim objKey
   Dim strKey, strItem
   Dim intCount, i, j

   intCount = objDict.Count

   If intCount > 1 Then

	  ReDim strDict(intCount, 2)

	  i = 0
	  For Each objKey In objDict
		 strDict(i,dictKey)  = CStr(objKey)
		 strDict(i,dictItem) = CStr(objDict(objKey))
		 i = i + 1
	  Next

	  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	  ' Perform a shell sort of the 2D string array
	  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	  For i = 0 To (intCount - 2)
		 For j = i To (intCount - 1)
			If StrComp(strDict(i,intSort), strDict(j,intSort), vbTextCompare) > 0 Then
			   strKey  = strDict(i,dictKey)
			   strItem = strDict(i,dictItem)
			   strDict(i,dictKey)  = strDict(j,dictKey)
			   strDict(i,dictItem) = strDict(j,dictItem)
			   strDict(j,dictKey)  = strKey
			   strDict(j,dictItem) = strItem
			End If
		 Next
	  Next

	  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	  ' Erase the contents of the dictionary object
	  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	  objDict.RemoveAll

	  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	  ' Repopulate the dictionary with the sorted information
	  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	  For i = 0 To (intCount - 1)
		 objDict.Add strDict(i,dictKey), strDict(i,dictItem)
	  Next

   End If

End Sub

Const HKEY_CLASSES_ROOT = &H80000000

Dim arrProgID, lstProgID, objReg, strProgID, strSubKey, subKey, subKeys(), objFSO, objFile, outFile, strWrite, i

Set lstProgID = CreateObject( "Scripting.Dictionary" )
Set objReg    = GetObject( "winmgmts://./root/default:StdRegProv" )
strWrite = ""
' List all subkeys of HKEY_CLASSES_ROOT\CLSID
objReg.EnumKey HKEY_CLASSES_ROOT, "CLSID", subKeys

' Loop through the list of subkeys
For Each subKey In subKeys
	' Check each subkey for the existence of a ProgID
	strSubKey = "CLSID\" & subKey & "\ProgID"
	objReg.GetStringValue HKEY_CLASSES_ROOT, strSubKey, "", strProgID
	' If a ProgID exists, add it to the list
	If Not IsNull( strProgID ) And Not lstProgID.exists(strProgID)  Then 
		lstProgID.Add strProgID, ""
	End If
Next

' Sort the list of ProgIDs
SortDictionary lstProgID, 1

' Copy the list to an array (this makes displaying it much easier)
For Each i in lstProgID
	strWrite = strWrite & i & vbCrLf
Next

' Write the entire array to file
Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile="listProgIDs.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write strWrite
objFile.Close