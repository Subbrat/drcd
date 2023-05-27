Set objShell = CreateObject("WScript.Shell")

' Open Command Prompt
objShell.Run "cmd.exe"

' Wait for 4 seconds
WScript.Sleep 4000

' Send the "git pull" command to the Command Prompt
objShell.SendKeys "git pull{ENTER}"

' Wait for the command to execute and capture the output
WScript.Sleep 2000
strOutput = objShell.Exec("cmd /c echo %errorlevel%").StdOut.ReadAll

' Check if the response contains "Already up to date."
If InStr(strOutput, "Already up to date.") > 0 Then
    ' Wait for 2 seconds
    WScript.Sleep 2000
    
    ' Echo "updated"
    objShell.SendKeys "echo updated{ENTER}"
End If

' Exit the Command Prompt
objShell.SendKeys "exit{ENTER}"
