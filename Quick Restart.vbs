Set WshShell = WScript.CreateObject("WScript.Shell")

intAnswer = _
    Msgbox("Are you sure you want to restart?", _
        vbYesNo, "Are you sure?")
If intAnswer = vbYes Then
    Msgbox "Please Click (OK) to confirm"
Msgbox "Restarting..."
WshShell.run "cmd"
WScript.sleep 125
WshShell.sendkeys "s"
WScript.sleep 75
WshShell.sendkeys "h"
WScript.sleep 75
WshShell.sendkeys "u"
WScript.sleep 75
WshShell.sendkeys "t"
WScript.sleep 75
WshShell.sendkeys "d"
WScript.sleep 75
WshShell.sendkeys "o"
WScript.sleep 75
WshShell.sendkeys "w"
WScript.sleep 75
WshShell.sendkeys "n"
WScript.sleep 75
WshShell.sendkeys " "
WScript.sleep 75
WshShell.sendkeys "/"
WScript.sleep 75
WshShell.sendkeys "r"
WScript.sleep 75
WshShell.sendkeys "{enter}"
WScript.sleep 75

Else
    Msgbox "ok then ._."
End If
