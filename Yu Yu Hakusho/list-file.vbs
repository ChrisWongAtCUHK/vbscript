Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "."

Set objFolder = objFSO.GetFolder(objStartFolder)
Wscript.Echo objFolder.Path

Set colFiles = objFolder.Files

For Each objFile in colFiles
If UCase(objFSO.GetExtensionName(objFile.name)) = "TXT" Then
    Wscript.Echo objFile.Name
    End If
Next

ShowSubfolders objFSO.GetFolder(objStartFolder)

Sub ShowSubFolders(Folder)
    For Each Subfolder in Folder.SubFolders
        Wscript.Echo Subfolder.Path
        Set objFolder = objFSO.GetFolder(Subfolder.Path)
        Set colFiles = objFolder.Files
        For Each objFile in colFiles
            Wscript.Echo objFile.Name
        Next
        ShowSubFolders Subfolder
    Next
End Sub