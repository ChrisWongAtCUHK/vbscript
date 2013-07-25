'http://stackoverflow.com/questions/1707058/how-to-split-a-string-in-a-windows-batch-file
'Usage: cscript /nologo test.vbs "AAA BBB CCC"
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
str1 = objArgs(0)
s=Split(str1," ")
For i=LBound(s) To UBound(s)
    WScript.Echo s(i)
Next