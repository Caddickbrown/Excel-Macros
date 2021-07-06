set shellobj = CreateObject("WScript.Shell")

tminus=InputBox("In how many minutes do you want to shut down?","Shut Down in...")

If IsEmpty(tminus) Then
	'Cancel Box clicked
	MsgBox "The Script has been cancelled"
Else
	tminus=tminus*60000
	shellobj.run "cmd"
	wscript.sleep 1500
	shellobj.sendkeys "shutdown-s-f-t"
	Shellobj.sendkeys a
	WshShell.SendKeys "{Enter}"
End If
