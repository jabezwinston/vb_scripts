' Author : Jabez Winston
'  
' Date : 3-12-2017

Set wshShell =wscript.CreateObject("WScript.Shell")

do while i < 50
wscript.sleep 100
wshshell.sendkeys "{CAPSLOCK}"
wshshell.sendkeys "{NUMLOCK}"
wshshell.sendkeys "{SCROLLLOCK}"
i = i + 1
loop