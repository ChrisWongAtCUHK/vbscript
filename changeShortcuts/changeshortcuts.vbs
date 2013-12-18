'D:\GoogleChromePortableMS\chrome-shortcuts-app


' Get the environment %CHROME_HOME%
oldPath = "D:\GoogleChromePortableMS"
newPath = "%CHROME_HOME%"

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = "D:\GoogleChromePortableMS\chrome-shortcuts-app"
Set objFolder = objFSO.GetFolder(objStartFolder)

Set files = objFolder.Files
For each folderIdx In files
	Set oWS = WScript.CreateObject("WScript.Shell")
	Set oLink = oWS.CreateShortcut(folderIdx.path)
	ChangeShortcut oLink, oldPath, newPath
Next
			
Showsubfolders objFolder

Sub Showsubfolders(folder)
    For Each subfolder in folder.subfolders
        ' folderDepth = (Length of current folder path) - (number if backslashes in current folder path) - (number of backslahes in path you have specified for objStartFolder)
        folderDepth = len(subfolder.Path) - len(replace(subfolder.Path,"\","")) - 1
        ' Specifying folderDepth > 0 will give everything inside your objStartFolder
        If folderDepth > 0 then
            ' Wscript.Echo subfolder.Path
            Set objFSO = CreateObject("Scripting.FileSystemObject")
			objStartFolder = subfolder.Path
			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set files = folder.Files
			For each folderIdx In files
    			'Wscript.Echo folderIdx.path
    			' Open the shortcut file
    			Set oWS = WScript.CreateObject("WScript.Shell")
				Set oLink = oWS.CreateShortcut(folderIdx.path)
				ChangeShortcut oLink, oldPath, newPath
    		Next
        End If
  
        Showsubfolders subfolder
    Next
End Sub

Sub ChangeShortcut(oLink, oldPath, newPath)
	' Target in
	targetPath = oLink.TargetPath
	targetPath = Replace(targetPath, oldPath, newPath)
	oLink.TargetPath = targetPath

	' Target in arguments
	arguments = oLink.Arguments
	arguments = Replace(arguments, oldPath, newPath)
	oLink.Arguments = arguments

	' Start in
	workingDirectory = oLink.WorkingDirectory
	workingDirectory = Replace(workingDirectory, oldPath, newPath)
	oLink.WorkingDirectory = workingDirectory

	' Change icon
	iconLocation = oLink.IconLocation
	iconLocation = Replace(iconLocation, oldPath, newPath)
	oLink.IconLocation = iconLocation

	'	other usages
	'	oLink.Description = "MyProgram"
	'	oLink.HotKey = "ALT+CTRL+F"
	'	oLink.WindowStyle = "1"

	' Save the change
	oLink.Save
End Sub