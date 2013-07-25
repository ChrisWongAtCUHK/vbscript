'How to read a file
Set objFS=CreateObject("Scripting.FileSystemObject")

strFile = "hello.txt"
Set objFile = objFS.OpenTextFile(strFile)
Do Until objFile.AtEndOfStream
    strLine= objFile.ReadLine
    Wscript.Echo strLine
Loop
objFile.Close