' How to write file
Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile="hello.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write "test string" & vbCrLf
objFile.Close