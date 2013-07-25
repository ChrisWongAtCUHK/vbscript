dim fso: set fso = CreateObject("Scripting.FileSystemObject")

' directory in which this script is currently running
currentDirectory = fso.GetAbsolutePathName(".")

Wscript.Echo currentDirectory