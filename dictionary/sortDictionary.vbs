'********************************************************************
'* SortDictionary
'* 
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


Dim oDic
Set oDic = CreateObject("scripting.dictionary")
oDic.Add "aaa", "demon"
oDic.Add "bbb", "blog"
oDic.Add "ccc", "http://demon.tw"

WScript.Echo "---- before sort ----"
For Each i In oDic
	WScript.Echo i, oDic(i)
Next

SortDictionary oDic, 2

WScript.Echo "---- after sort ----"
For Each i In oDic
	WScript.Echo i, oDic(i)
Next



