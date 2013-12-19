' A vbscript to turn on(speak) and turn off(mute) the sound
' cscript /nologo muteSpeak.vbs

' Open the shell
Set oShell = CreateObject("WScript.Shell")

' Open the volume control, (in winxp, it may be SndVol32)
oShell.run "Sndvol"

' Wait
WScript.Sleep 1500

' Send tab
oShell.SendKeys "{TAB}"

' Select enter
oShell.SendKeys " "

' Alt + F4
oShell.SendKeys "%{F4}"