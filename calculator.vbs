msgbox "Press OK to open"
do
Set shell=CreateObject("wscript.shell")
Shell.Run("calc.exe")
loop