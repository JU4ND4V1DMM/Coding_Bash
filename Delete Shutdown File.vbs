Dim fso, filePath
filePath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\w1x2y3z4\quality.bat")

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(filePath) Then
    fso.DeleteFile(filePath)
End If
