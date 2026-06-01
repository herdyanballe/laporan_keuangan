Set oShell = CreateObject("WScript.Shell")
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Cek Python
On Error Resume Next
oShell.Run "python --version", 0, True
If Err.Number <> 0 Then
    MsgBox "Python tidak ditemukan!" & vbCrLf & vbCrLf & _
           "Silakan install Python dari:" & vbCrLf & _
           "https://www.python.org/downloads/" & vbCrLf & vbCrLf & _
           "Pastikan centang 'Add Python to PATH' saat install.", _
           vbCritical, "Laporan Keuangan - Error"
    WScript.Quit
End If
On Error GoTo 0

' Jalankan aplikasi (tanpa jendela CMD)
oShell.Run "cmd /c cd /d """ & strPath & """ && python app.py", 0, False
