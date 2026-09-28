Attribute VB_Name = "PCD_Updater"
Option Explicit

Private Const VERSION_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/VERSION.json"
Private Const ARTIFACT_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-dev.xlsm"
Private Const UPDATE_TIMEOUT_SECONDS As Long = 30

Public Sub CheckForUpdate()
    Dim diagPath As String
    Dim diagFso As Object
    Dim diagTs As Object

    On Error Resume Next
    diagPath = Environ$("TEMP") & "\PCD-Excel-checkforupdate.log"
    Set diagFso = CreateObject("Scripting.FileSystemObject")
    Set diagTs = diagFso.CreateTextFile(diagPath, True, False)
    If Not diagTs Is Nothing Then
        diagTs.WriteLine Now & " | CHECKFORUPDATE START"
        diagTs.WriteLine "Workbook=" & ThisWorkbook.FullName
        diagTs.WriteLine "Version=" & CStr(ThisWorkbook.Worksheets("Intervention en cours").Range("H2").Value)
        diagTs.Close
    End If
    On Error GoTo UpdateError

    MsgBox "CHECKFORUPDATE IZ DEV-19" & vbCrLf & vbCrLf & _
           "Workbook: " & ThisWorkbook.FullName & vbCrLf & _
           "Version: " & ThisWorkbook.Worksheets("Intervention en cours").Range("H2").Value, _
           vbInformation, "PCD DIJAGNOSTIKA"

    Dim currentVersion As String
    Dim remoteVersion As String
    Dim remoteSha256 As String
    Dim tempPath As String
    Dim answer As VbMsgBoxResult

    On Error GoTo UpdateError

    currentVersion = Trim$(CStr(ThisWorkbook.Worksheets("Intervention en cours").Range("H2").Value))
    If Len(currentVersion) = 0 Then currentVersion = "0.0.0"

    Application.StatusBar = "PCD Excel: proveravam novu verziju..."
    remoteVersion = JsonValue(HttpGetText(VERSION_URL & "?t=" & CStr(Timer)), "version")
    remoteSha256 = LCase$(JsonValue(HttpGetText(VERSION_URL & "?t=" & CStr(Timer + 1)), "sha256"))

    If Len(remoteVersion) = 0 Then Err.Raise vbObjectError + 1001, , "GitHub nije vratio broj verzije."
    If Len(remoteSha256) <> 64 Then Err.Raise vbObjectError + 1002, , "GitHub nije vratio ispravan SHA-256."

    If CompareVersions(remoteVersion, currentVersion) <= 0 Then
        Application.StatusBar = False
        MsgBox "Koristis najnoviju dostupnu verziju: " & currentVersion, vbInformation, "PCD Excel"
        Exit Sub
    End If

    answer = MsgBox( _
        "Dostupna je nova verzija PCD Excel-a." & vbCrLf & vbCrLf & _
        "Trenutna: " & currentVersion & vbCrLf & _
        "Nova: " & remoteVersion & vbCrLf & vbCrLf & _
        "Da li zelis da je preuzmem i instaliram?", _
        vbQuestion + vbYesNo, "PCD Excel - azuriranje")

    If answer <> vbYes Then
        Application.StatusBar = False
        Exit Sub
    End If

    If Not ThisWorkbook.Saved Then
        answer = MsgBox( _
            "Postoje nesacuvane izmene u ovom Excel fajlu." & vbCrLf & vbCrLf & _
            "Sacuvaj ih pre azuriranja, pa ponovo pokreni proveru.", _
            vbExclamation + vbOKOnly, "PCD Excel - azuriranje")
        Application.StatusBar = False
        Exit Sub
    End If

    tempPath = DownloadUpdate(remoteVersion)
    If Len(tempPath) = 0 Then Err.Raise vbObjectError + 1003, , "Preuzimanje nove verzije nije uspelo."

    Application.StatusBar = "PCD Excel: proveravam integritet nove verzije..."
    Dim localSha256 As String
    localSha256 = LCase$(FileSha256(tempPath))
    If localSha256 <> remoteSha256 Then
        On Error Resume Next
        Kill tempPath
        On Error GoTo UpdateError
        Err.Raise vbObjectError + 1004, , "SHA-256 kontrola nije prosla." & vbCrLf & "Ocekivani: " & remoteSha256 & vbCrLf & "Dobijeni: " & localSha256
    End If

    If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        Err.Raise vbObjectError + 1005, , "Automatsko azuriranje je podrzano za .xlsm fajl."
    End If

    MsgBox "STIGAO SAM DO SCHEDULE REPLACEMENT" & vbCrLf & vbCrLf & _
           "Temp fajl: " & tempPath & vbCrLf & _
           "Ciljni fajl: " & ThisWorkbook.FullName, _
           vbInformation, "PCD DIJAGNOSTIKA"

    ScheduleReplacement tempPath, ThisWorkbook.FullName
    Application.StatusBar = False
    MsgBox "Nova verzija je preuzeta i proverena." & vbCrLf & vbCrLf & _
           "Excel ce sada zatvoriti staru verziju, zameniti je novom i ponovo je otvoriti.", _
           vbInformation, "PCD Excel - azuriranje"

    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    Application.Quit
    Exit Sub

UpdateError:
    Application.StatusBar = False
    MsgBox "Azuriranje nije izvrseno." & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Mozes nastaviti da koristis ovu verziju. Ako je korporativna mreza blokirala GitHub, koristi rucno preuzimanje najnovijeg XLSM fajla.", _
           vbExclamation, "PCD Excel - azuriranje"
End Sub

Private Function HttpGetText(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1010, , "GitHub HTTP greska: " & http.Status & " " & http.StatusText
    End If

    HttpGetText = CStr(http.responseText)
End Function

Private Function DownloadUpdate(ByVal remoteVersion As String) As String
    Dim http As Object
    Dim stream As Object
    Dim tempPath As String

    tempPath = Environ$("TEMP") & "PCD-Excel-update-" & Replace(remoteVersion, ".", "_") & ".xlsm"

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", ARTIFACT_URL & "?v=" & Replace(remoteVersion, " ", "%20") & "&t=" & CStr(Timer), False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1011, , "Preuzimanje XLSM fajla nije uspelo: HTTP " & http.Status
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile tempPath, 2
    stream.Close

    DownloadUpdate = tempPath
End Function

Private Function FileSha256(ByVal filePath As String) As String
    Dim outputPath As String
    Dim shell As Object
    Dim fso As Object
    Dim ts As Object
    Dim text As String
    Dim matches As Object
    Dim re As Object

    outputPath = Environ$("TEMP") & "PCD-sha256-" & Format$(Timer * 1000, "0") & ".txt"

    Set shell = CreateObject("WScript.Shell")
    shell.Run "cmd.exe /c certutil -hashfile " & QuoteArg(filePath) & " SHA256 > " & QuoteArg(outputPath), 0, True

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(outputPath) Then Err.Raise vbObjectError + 1012, , "Windows nije mogao da izracuna SHA-256."

    Set ts = fso.OpenTextFile(outputPath, 1, False)
    text = ts.ReadAll
    ts.Close
    On Error Resume Next
    fso.DeleteFile outputPath, True
    On Error GoTo 0

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "([0-9A-Fa-f]{64})"
    Set matches = re.Execute(text)

    If matches.Count = 0 Then Err.Raise vbObjectError + 1013, , "Windows nije vratio SHA-256 vrednost."
    FileSha256 = LCase$(matches(0).Value)
End Function

Private Function JsonValue(ByVal json As String, ByVal key As String) As String
    Dim re As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = """" & key & """" & "\s*:\s*""([^""]*)"""
    Set matches = re.Execute(json)

    If matches.Count > 0 Then JsonValue = matches(0).SubMatches(0)
End Function

Private Function CompareVersions(ByVal a As String, ByVal b As String) As Long
    Dim pa() As String, pb() As String
    Dim i As Long, na As Long, nb As Long
    Dim aCore As String, bCore As String

    aCore = Split(a, "-")(0)
    bCore = Split(b, "-")(0)
    pa = Split(aCore, ".")
    pb = Split(bCore, ".")

    For i = 0 To 2
        na = 0: nb = 0
        If i <= UBound(pa) And IsNumeric(pa(i)) Then na = CLng(pa(i))
        If i <= UBound(pb) And IsNumeric(pb(i)) Then nb = CLng(pb(i))
        If na > nb Then CompareVersions = 1: Exit Function
        If na < nb Then CompareVersions = -1: Exit Function
    Next i

    CompareVersions = CompareBuildSuffix(a, b)
End Function

Private Function CompareBuildSuffix(ByVal a As String, ByVal b As String) As Long
    Dim da As Long, db As Long
    da = LastNumberAfterDash(a)
    db = LastNumberAfterDash(b)

    If da > db Then
        CompareBuildSuffix = 1
    ElseIf da < db Then
        CompareBuildSuffix = -1
    Else
        CompareBuildSuffix = 0
    End If
End Function

Private Function LastNumberAfterDash(ByVal value As String) As Long
    Dim parts() As String
    Dim i As Long

    parts = Split(value, "-")
    For i = UBound(parts) To 1 Step -1
        If IsNumeric(parts(i)) Then
            LastNumberAfterDash = CLng(parts(i))
            Exit Function
        End If
    Next i
End Function

Private Sub ScheduleReplacement(ByVal newFile As String, ByVal oldFile As String)
    Dim scriptPath As String
    Dim logPath As String
    Dim scriptText As String
    Dim fso As Object
    Dim ts As Object
    Dim shell As Object

    scriptPath = Environ$("TEMP") & "\PCD-Excel-updater.vbs"
    logPath = Environ$("TEMP") & "\PCD-Excel-updater.log"

    On Error GoTo CreateError

    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    If fso.FileExists(logPath) Then fso.DeleteFile logPath, True
    On Error GoTo CreateError

    scriptText = _
        "Option Explicit" & vbCrLf & _
        "Dim fso, shell, newFile, oldFile, logPath, i, replaced" & vbCrLf & _
        "Set fso = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
        "Set shell = CreateObject(""WScript.Shell"")" & vbCrLf & _
        "newFile = WScript.Arguments(0)" & vbCrLf & _
        "oldFile = WScript.Arguments(1)" & vbCrLf & _
        "logPath = WScript.Arguments(2)" & vbCrLf & _
        "LogLine ""START""" & vbCrLf & _
        "LogLine ""NEW="" & newFile" & vbCrLf & _
        "LogLine ""OLD="" & oldFile" & vbCrLf & _
        "For i = 1 To 120" & vbCrLf & _
        "  On Error Resume Next" & vbCrLf & _
        "  Err.Clear" & vbCrLf & _
        "  If fso.FileExists(oldFile) Then fso.DeleteFile oldFile, True" & vbCrLf & _
        "  If Err.Number = 0 Then" & vbCrLf & _
        "    LogLine ""DELETE OK attempt="" & i" & vbCrLf & _
        "    Err.Clear" & vbCrLf & _
        "    fso.MoveFile newFile, oldFile" & vbCrLf & _
        "    If Err.Number = 0 Then" & vbCrLf & _
        "      replaced = True" & vbCrLf & _
        "      LogLine ""MOVE OK attempt="" & i" & vbCrLf & _
        "      Exit For" & vbCrLf & _
        "    Else" & vbCrLf & _
        "      LogLine ""MOVE ERROR "" & Err.Number & "" "" & Err.Description" & vbCrLf & _
        "    End If" & vbCrLf & _
        "  Else" & vbCrLf & _
        "    LogLine ""DELETE ERROR "" & Err.Number & "" "" & Err.Description" & vbCrLf & _
        "  End If" & vbCrLf & _
        "  Err.Clear" & vbCrLf & _
        "  On Error GoTo 0" & vbCrLf & _
        "  WScript.Sleep 1000" & vbCrLf & _
        "Next" & vbCrLf & _
        "If replaced Then" & vbCrLf & _
        "  LogLine ""REPLACED OK""" & vbCrLf & _
        "  LogLine ""LAUNCHING "" & oldFile" & vbCrLf & _
        "  shell.Run Chr(34) & oldFile & Chr(34), 1, False" & vbCrLf & _
        "Else" & vbCrLf & _
        "  LogLine ""REPLACEMENT FAILED after 120 attempts""" & vbCrLf & _
        "End If" & vbCrLf & _
        "LogLine ""END""" & vbCrLf & _
        "Sub LogLine(ByVal message)" & vbCrLf & _
        "  Dim logFile" & vbCrLf & _
        "  On Error Resume Next" & vbCrLf & _
        "  Set logFile = fso.OpenTextFile(logPath, 8, True)" & vbCrLf & _
        "  logFile.WriteLine Now & "" | "" & message" & vbCrLf & _
        "  logFile.Close" & vbCrLf & _
        "End Sub"

    Set ts = fso.CreateTextFile(scriptPath, True, False)
    ts.Write scriptText
    ts.Close

    If Not fso.FileExists(scriptPath) Then
        Err.Raise vbObjectError + 1020, , "Updater nije uspeo da napravi VBS fajl."
    End If

    Set shell = CreateObject("WScript.Shell")
    shell.Run "wscript.exe " & QuoteArg(scriptPath) & " " & _
              QuoteArg(newFile) & " " & QuoteArg(oldFile) & " " & _
              QuoteArg(logPath), 0, False
    Exit Sub

CreateError:
    Err.Raise vbObjectError + 1021, , "Updater nije uspeo da pripremi eksterni updater." & vbCrLf & Err.Description
End Sub

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = """" & Replace(value, """", """""") & """"
End Function
