Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c C:\Users\RautShyam\Downloads\test.bat"
oShell.Run strArgs, 0, false