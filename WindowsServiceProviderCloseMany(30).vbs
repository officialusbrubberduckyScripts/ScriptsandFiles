Set WshShell = CreateObject("WScript.Shell")

WScript.Sleep 500  ' focus the target window first

i = 0
While i < 9999
    WshShell.SendKeys "%{F4}"  ' Alt+F4
    WScript.Sleep 1000
    WshShell.SendKeys "%{F4}"  ' Alt+F4
    WScript.Sleep 1000
    WshShell.SendKeys "%{F4}"  ' Alt+F4
    WScript.Sleep 30000          
    i = i + 1
Wend
