Set WshShell = WScript.CreateObject("WScript.Shell")

'Open cmd Line'
WshShell.run "cmd"
'Wait for cmd line to open
wscript.sleep 500
'Shutdown code
WshShell.sendkeys "shutdown -a"
'Go
WshShell.sendkeys "{Enter}"
'Close window
wscript.sleep 500
WshShell.sendkeys "%{F4}"
