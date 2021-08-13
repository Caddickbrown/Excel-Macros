'This is a guide to VBScripting for Windows. It will go into detail about how to set up a script and how to run it.
'You can setup "VBScripts" using Notepad and that's it! They can be incredibly useful for automating small tasks. They don't need installation or another program to run them. You only need a text editor to write them.
'Thoughout this guide, I am likely to refer to them as "VBScript", "VBS", "VBS Script", "Scripts" or some variation of the sort. Although technically "VBScript" is correct, as it stands for "Visual Basic Script", I will use them interchangably for ease. I'll possibly fix it where I can with time, but I'm not overly concerned with the syntax.

'To start with, you have to Create the "Environment" in which your code will run on your machine. This enables you to interact with Windows and run your scripts.
Set WshShell = WScript.CreateObject("WScript.Shell")
'Basically - If you input this line at the start of each script - you can code away

'The Best way to think about VBScripting to start with (although it can get a lot more complex) is a way to mimic the keyboard and mouse. It can be used to type things in, to copy and paste, and basically run small tasks that you would usually use a keyboard and mouse for.
'An easy example would be

'Sleep
'This simply pauses the script for a specific amount of milliseconds
WScript.Sleep(55000)

'Send Keys
'Send Keys imitates Keystrokes
WshShell.SendKeys("{F15}")

'Message Box
'This will pop up a box to alert the user of whatever is required. It has an Ok button which will continue the script. The Script is Paused until this OK button is pressed.
MsgBox "Pop Up Box"

'Input Box
'Input Boxes are used for defining values in variables
tminus=InputBox("In how many minutes do you want to shut down?","Shut Down in...","Whole Numbers Please")

'Loops
'For Loop
'Do Loop
'Do While Loop
Do While condition
Loop
'Do Until Loop

'If Else
If condition Then
  ' true
Else
  ' false
End if

'IsEmpty
'Checks if a variable is empty
IsEmpty(variable)

'Int
'This defines the value as an integer
Int(value)
