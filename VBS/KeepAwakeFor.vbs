'This script does not work entirely as required - seems to drift overtime, possibly due to 55000 - unsure
'This would be a script that asks you how long to keep your computer open for and then keeps it awake for that long

set shellobj = CreateObject("WScript.Shell")

'User defines minutes to shut down
tminus=Int(InputBox("How many minutes do you want keep your computer awake for?","Keep awake for...","Whole Numbers Please"))

If tminus<1 Then
	'Cancel Box clicked
	MsgBox "The Script has been cancelled"
Else
	'Convert to milliseconds
	tminus=(tminus*60000)/55000

  Set WshShell = WScript.CreateObject("WScript.Shell")
  a=0
  Do While a<tminus
          WshShell.SendKeys("{F15}")
          WScript.Sleep(55000)
  a=a+1
  Loop
	MsgBox "Your computer is no longer being kept awake."
End If
