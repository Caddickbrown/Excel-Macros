'This script can be used to automatically shut down your PC in an amount of minutes specified by the user

set shellobj = CreateObject("WScript.Shell")

'User defines minutes to shut down
tminus=InputBox("In how many minutes do you want to shut down?","Shut Down in...","Whole Numbers Please")

If IsEmpty(tminus) Then
	'Cancel Box clicked
	MsgBox "The Script has been cancelled"
Else
	'Convert to milliseconds
	tminus=tminus*60
	'Open cmd Line'
	shellobj.run "cmd"
	'Wait for cmd line to open
	wscript.sleep 500
	'Shutdown code
	shellobj.sendkeys "shutdown -s -f -t "&tminus
	'Go
	shellobj.sendkeys "{Enter}"
	'Close window
	wscript.sleep 500
	shellobj.sendkeys "%{F4}"
End If
