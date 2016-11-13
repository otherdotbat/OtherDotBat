Set wshShell =wscript.CreateObject("WScript.Shell")
do
wscript.sleep 100
wshshell.sendkeys "{CAPSLOCK}"
wshshell.sendkeys "{NUMLOCK}"
wshshell.sendkeys "{i}"
wshshell.sendkeys "{d}"
wshshell.sendkeys "{i}"
wshshell.sendkeys "{o}"
wshshell.sendkeys "{t}"
wshshell.sendkeys "{ENTER}"
loop