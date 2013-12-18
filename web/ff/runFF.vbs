Set WshShell = WScript.CreateObject("WScript.Shell")
Return = WshShell.Run("firefox.exe http://www.google.com.hk", 1)