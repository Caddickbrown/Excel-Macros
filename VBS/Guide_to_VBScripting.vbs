'This is a guide to VBScripting for Windows. It will go into detail about how to set up a script and how to run it.
'You can setup "VBScripts" using Notepad and that's it! They can be incredibly useful for automating small tasks. They don't need installation or another program to run them. You only need a text editor to write them.
'Thoughout this guide, I am likely to refer to them as "VBScript", "VBS", "VBS Script", "Scripts" or some variation of the sort. Although technically "VBScript" is correct, as it stands for "Visual Basic Script", I will use them interchangably for ease. I'll possibly fix it where I can with time, but I'm not overly concerned with the syntax.

'To start with, you have to Create the "Environment" in which your code will run on your machine. This enables you to interact with Windows and run your scripts.
Set WshShell = WScript.CreateObject("WScript.Shell")
'Basically - If you input this line at the start of each script - you can code away

'The Best way to think about VBScripting to start with (although it can get a lot more complex) is a way to mimic the keyboard and mouse. It can be used to type things in, to copy and paste, and basically run small tasks that you would usually use a keyboard and mouse for.
'An easy example would be 
