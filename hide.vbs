Set oShell = CreateObject ("Wscript.Shell")
Dim strArgs
strArgs = "cmd /c ur program name"
oShell.Run strArgs, 0, false