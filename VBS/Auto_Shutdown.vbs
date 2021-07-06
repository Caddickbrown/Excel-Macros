set shellobj = CreateObject("WScript.Shell")

a=InputBox("In how many minutes do you want to shut down?","Shut Down in...")
a=a*60000

shellobj.run "cmd"
wscript.sleep 2000
shellobj.sendkeys "shutdown-s-f-t"
Shellobj.sendkeys a
