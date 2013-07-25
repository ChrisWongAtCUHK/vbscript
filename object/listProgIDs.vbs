'http://www.robvanderwoude.com/vbstech_objectbrowsers.php
Option Explicit

Const HKEY_CLASSES_ROOT = &H80000000

Dim arrProgID, lstProgID, objReg, strProgID, strSubKey, subKey, subKeys(), objFSO, objFile, outFile

Set lstProgID = CreateObject( "System.Collections.ArrayList" )
Set objReg    = GetObject( "winmgmts://./root/default:StdRegProv" )

' List all subkeys of HKEY_CLASSES_ROOT\CLSID
objReg.EnumKey HKEY_CLASSES_ROOT, "CLSID", subKeys

' Loop through the list of subkeys
For Each subKey In subKeys
	' Check each subkey for the existence of a ProgID
	strSubKey = "CLSID\" & subKey & "\ProgID"
	objReg.GetStringValue HKEY_CLASSES_ROOT, strSubKey, "", strProgID
	' If a ProgID exists, add it to the list
	If Not IsNull( strProgID ) Then lstProgID.Add strProgID
Next

' Sort the list of ProgIDs
lstProgID.Sort

' Copy the list to an array (this makes displaying it much easier)
arrProgID = lstProgID.ToArray

' WScript.Echo Join(arrProgID, vbCrLf)

' Write the entire array to file
Set objFSO=CreateObject("Scripting.FileSystemObject")
outFile="listProgIDs.txt"
Set objFile = objFSO.CreateTextFile(outFile,True)
objFile.Write Join(arrProgID, vbCrLf)
objFile.Close