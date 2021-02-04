Set wshShell =wscript.CreateObject("WScript.Shell")
wshshell.sendkeys "{NUMLOCK}"
do
wshshell.sendkeys "{CAPSLOCK}"
wscript.sleep 120
wshshell.sendkeys "{SCROLLLOCK}"
wscript.sleep 50
wshshell.sendkeys "{CAPSLOCK}"
wscript.sleep 120
wshshell.sendkeys "{NUMLOCK}"
wscript.sleep 50
wshshell.sendkeys "{SCROLLLOCK}"
wscript.sleep 120
wshshell.sendkeys "{NUMLOCK}"
wscript.sleep 50
loop