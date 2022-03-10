'NEEDS CHECKING

Set WshShell = WScript.CreateObject("WScript.Shell")

'User defines minutes to shut down
tminus=InputBox("In how many minutes do you want to shut down?","Shut Down in...","Whole Numbers Please")

If IsEmpty(tminus) Then
MsgBox "The Script has been cancelled"
Else
'Convert to milliseconds
tminus=tminus*60
'Open cmd Line'
WshShell.run "cmd"
'Wait for cmd line to open
wscript.sleep tminus
'Shutdown code
WshShell.sendkeys "rundll32.exe powrprof.dll, SetSuspendState Sleep"
'Go
WshShell.sendkeys "{Enter}"
'Close window
wscript.sleep 500
WshShell.sendkeys "%{F4}"
End If
