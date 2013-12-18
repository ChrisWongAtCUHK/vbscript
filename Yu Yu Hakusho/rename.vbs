Dim FSO
  Set FSO = CreateObject("Scripting.FileSystemObject")
  strFile = "test.txt"
  strRename = "rename.txt"
   If FSO.FileExists(strFile) Then
        FSO.MoveFile strFile, strRename
   End If

  Set FSO = Nothing