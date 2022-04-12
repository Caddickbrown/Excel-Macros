'This script would be used in the IFS ERP System for filling out details to multiple pages one after another
Set WshShell=WScript.CreateObject("WScript.Shell")

MsgBox "Click in the box you wish the info to go into"
wscript.sleep 10000

'User defines the amount of orders that need to be copied down, this is specified as an integer
OrderCount=Int(InputBox("How many orders do you have?","# of Batches...","Whole Numbers Please"))

'If box is empty, cancelled, or has a qty of 0 then this will cancel the script
If OrderCount<1 Then
	MsgBox "Script Cancelled"

Else

	'User defines the value that needs to be inputted
	ConcVal=InputBox("What info do you need inputting?","Info...")

	'If box is empty, or cancelled this will cancel the script
	If IsEmpty(ConcVal) Then
	MsgBox "Script Cancelled"

	Else
		'Select window
		WshShell.AppActivate "Lot Batch Master - Daniel Caddick-Brown @ IFS Applications 8 SP 1 - Live Database - IFS Applications"
		'Wait - just in case
		wscript.sleep 200
		'define "Loop Counter" Variable
		a=0
		'Start Loop'
		Do While a<OrderCount

			wscript.sleep 500
			'Sends the inputted info as text
			WshShell.sendkeys ConcVal
			WScript.Sleep 500
			'Save Record
			WshShell.SendKeys "^{s}"
			WScript.Sleep 1000
			'Accept Dialog box
			WshShell.SendKeys "{Enter}"
			WScript.Sleep 1000
			'Load up next Record
			WshShell.SendKeys "^{DOWN}"

			a=a+1

		Loop
	End If
End If
