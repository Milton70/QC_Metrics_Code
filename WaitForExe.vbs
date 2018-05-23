On Error Resume Next
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Waits for an application to shut down. Usage:" & vbCrLf & Ucase(WScript.ScriptName) & " ""Program.exe""" & vbCrLf & "Where ""Program.exe"" is the executable name of the application you are waiting for."
Else
    blnRunning = True
    Do While blnRunning = True
        Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
        Set colItems = objWMIService.ExecQuery("Select Name from Win32_Process where Name='" & Wscript.Arguments(0) & "'",,48)
        blnRunning = False
        For Each objItem in colItems
            blnRunning = True
        Next
        WScript.Sleep 500
    Loop
End If
