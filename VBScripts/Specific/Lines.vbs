'This would fill out the Unit of Measure for batches in IFS - it would bring up menus and select the latest UoM that had been picked. Could likely be used for similar situations'

Set WshShell = WScript.CreateObject("WScript.Shell")

'The following part is looped, change the "Do until" to however many loops you want to do

LineCount=Int(InputBox("How many lines do you have?","# of Lines...","Whole Numbers Please"))

'If box is empty or has a qty of 0 then this will cancel the script
If LineCount<1 Then
	'Cancel Box clicked
	MsgBox "Script Cancelled"

  Else
    a=0
	MsgBox "WARNING: DO NOT TOUCH ANYTHING UNTIL SCRIPT HAS FINISHED"
'Start Loop'
    Do While a<LineCount

      WScript.Sleep 200
      WshShell.SendKeys "^{v}"
      WScript.Sleep 200
      WshShell.SendKeys "{DOWN}"

      a=a+1
    Loop

	WScript.Sleep 200
      	WshShell.SendKeys "^{s}"
      	WScript.Sleep 200
	MsgBox "Click Ok when IFS has Saved"

    a=0

    Do While a<LineCount

      	WScript.Sleep 200
	WshShell.SendKeys "{ENTER}"

      a=a+1
    Loop

End If
