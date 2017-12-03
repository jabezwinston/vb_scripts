' Author : Jabez Winston
'  
' Date : 3-12-2017

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad"
WScript.Sleep 1000
WshShell.AppActivate "Untitled - Notepad"
WshShell.SendKeys "% x"
WScript.Sleep 1000
WshShell.SendKeys "Hello{ENTER}"
WScript.Sleep 500

Dim x
For x = 1 To 50
    WshShell.SendKeys x & " "
	WScript.Sleep 100
Next