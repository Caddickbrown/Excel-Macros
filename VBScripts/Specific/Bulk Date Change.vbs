'This would fill out multiple Lines of dates to quickly change details'

Set WshShell = WScript.CreateObject("WScript.Shell")

'The following part is looped, change the "Do until" to however many loops you want to do

LineCount=Int(InputBox("How many lines do you have?","# of Lines...","Whole Numbers Please"))

CalendarValue=InputBox("What date are you changing the lines to?","Date...","Use dd/mm/yyyy")

'If box is empty or has a qty of 0 then this will cancel the script
If LineCount<1 Then
	'Cancel Box clicked
	MsgBox "Script Cancelled"

  Else
    a=0
'Start Loop'
    Do While a<LineCount

      WScript.Sleep 200
      WshShell.SendKeys(CalendarValue)
      WScript.Sleep 200
      WshShell.SendKeys "{DOWN}"

      a=a+1
    Loop

End If

MsgBox "Script Complete"