Attribute VB_Name = "PCD_Updater"
Option Explicit

Private Const VERSION_URL As String = "https://api.github.com/repos/milance78/PCD-excel-version-/contents/VERSION.json?ref=main"
Private Const ARTIFACT_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-latest.xlsm"
Private Const UPDATE_TIMEOUT_SECONDS As Long = 30

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr

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

    Dim currentVersion As String
    Dim remoteVersion As String
    Dim remoteSha256 As String
    Dim manifestText As String
    Dim tempPath As String
    Dim answer As VbMsgBoxResult
    Dim localSha256 As String
    Dim cacheBust As String

    currentVersion = Trim$(CStr(ThisWorkbook.Worksheets("Intervention en cours").Range("H2").Value))
    If Len(currentVersion) = 0 Then currentVersion = "0.0.0"

    Application.StatusBar = "PCD Excel: proveravam novu verziju..."
    cacheBust = CStr(Timer)

    manifestText = Base64Decode(JsonValue(HttpGetText(VERSION_URL & "&t=" & cacheBust), "content"))
    remoteVersion = JsonValue(manifestText, "version")
    remoteSha256 = LCase$(JsonValue(manifestText, "sha256"))

    If Len(remoteVersion) = 0 Then Err.Raise vbObjectError + 1001, , "GitHub nije vratio broj verzije."
    If Len(remoteSha256) <> 64 Then Err.Raise vbObjectError + 1002, , "GitHub nije vratio ispravan SHA-256."

    If CompareVersions(remoteVersion, currentVersion) <= 0 Then
        Application.StatusBar = False
        MsgBox "Koristis najnoviju dostupnu verziju: " & currentVersion, _
               vbInformation, "PCD Excel"
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

    tempPath = DownloadUpdate(remoteVersion, cacheBust)
    If Len(tempPath) = 0 Then Err.Raise vbObjectError + 1003, , "Preuzimanje nove verzije nije uspelo."

    Application.StatusBar = "PCD Excel: proveravam integritet nove verzije..."
    localSha256 = LCase$(FileSha256(tempPath))

    If localSha256 <> remoteSha256 Then
        On Error Resume Next
        Kill tempPath
        On Error GoTo UpdateError
        Err.Raise vbObjectError + 1004, , _
            "SHA-256 kontrola nije prosla." & vbCrLf & _
            "Ocekivani: " & remoteSha256 & vbCrLf & _
            "Dobijeni: " & localSha256
    End If

    If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        Err.Raise vbObjectError + 1005, , "Automatsko azuriranje je podrzano za .xlsm fajl."
    End If

    ScheduleReplacement tempPath, ThisWorkbook.FullName, CurrentExcelProcessId
    Application.StatusBar = False

    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Saved = True
    Err.Clear
    Application.Quit
    Err.Clear
    On Error GoTo 0
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
    http.setRequestHeader "Pragma", "no-cache"
    http.setRequestHeader "If-Modified-Since", "Sat, 01 Jan 2000 00:00:00 GMT"
    http.Send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1010, , "GitHub HTTP greska: " & http.Status & " " & http.StatusText
    End If

    HttpGetText = CStr(http.responseText)
End Function

Private Function DownloadUpdate(ByVal remoteVersion As String, ByVal cacheBust As String) As String
    Dim http As Object
    Dim stream As Object
    Dim tempPath As String

    tempPath = Environ$("TEMP") & "\PCD-Excel-update-" & Replace(remoteVersion, ".", "_") & ".xlsm"

    On Error Resume Next
    Kill tempPath
    On Error GoTo 0

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "GET", ARTIFACT_URL & "?v=" & Replace(remoteVersion, " ", "%20") & "&pcd=" & cacheBust, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.setRequestHeader "Pragma", "no-cache"
    http.setRequestHeader "If-Modified-Since", "Sat, 01 Jan 2000 00:00:00 GMT"
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

    outputPath = Environ$("TEMP") & "\PCD-sha256-" & Format$(Timer * 1000, "0") & ".txt"

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

Private Function Base64Decode(ByVal encoded As String) As String
    Dim xml As Object
    Dim node As Object
    Dim bytes() As Byte
    Dim stream As Object

    encoded = Replace(encoded, vbCr, "")
    encoded = Replace(encoded, vbLf, "")
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = encoded
    bytes = node.nodeTypedValue

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.Position = 0
    stream.Type = 2
    stream.Charset = "utf-8"
    Base64Decode = stream.ReadText
    stream.Close
End Function

Private Function JsonValue(ByVal json As String, ByVal key As String) As String
    Dim re As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = Chr(34) & key & Chr(34) & "\s*:\s*" & Chr(34) & "([^" & Chr(34) & "]*)" & Chr(34)
    Set matches = re.Execute(json)

    If matches.Count > 0 Then JsonValue = matches(0).SubMatches(0)
End Function

Private Function CompareVersions(ByVal a As String, ByVal b As String) As Long
    Dim buildA As Long, buildB As Long

    buildA = BuildNumber(a)
    buildB = BuildNumber(b)

    If buildA > buildB Then
        CompareVersions = 1
    ElseIf buildA < buildB Then
        CompareVersions = -1
    Else
        CompareVersions = 0
    End If
End Function

Private Function BuildNumber(ByVal value As String) As Long
    Dim parts() As String
    Dim i As Long

    value = Trim$(value)
    If Len(value) = 0 Then Exit Function

    If IsNumeric(value) Then
        BuildNumber = CLng(value)
        Exit Function
    End If

    parts = Split(value, "-")
    For i = UBound(parts) To 0 Step -1
        If IsNumeric(parts(i)) Then
            BuildNumber = CLng(parts(i))
            Exit Function
        End If
    Next i
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

Private Function CurrentExcelProcessId() As Long
    Dim processId As Long
    GetWindowThreadProcessId Application.Hwnd, processId
    CurrentExcelProcessId = processId
End Function

Private Sub ScheduleReplacement(ByVal newFile As String, ByVal oldFile As String, ByVal processId As Long)
    Dim scriptPath As String
    Dim logPath As String
    Dim scriptText As String
    Dim fso As Object
    Dim ts As Object
    Dim shellResult As LongPtr
    Dim wscriptPath As String
    Dim launchParams As String

    scriptPath = Environ$("TEMP") & "\PCD-Excel-updater.vbs"
    logPath = Environ$("TEMP") & "\PCD-Excel-updater.log"

    On Error GoTo CreateError
    Set fso = CreateObject("Scripting.FileSystemObject")

    On Error Resume Next
    If fso.FileExists(logPath) Then fso.DeleteFile logPath, True
    If fso.FileExists(scriptPath) Then fso.DeleteFile scriptPath, True
    On Error GoTo CreateError

    scriptText = "Option Explicit" & vbCrLf
    AppendVbsLine scriptText, "Dim fso, shell, xl, wb, logPath"
    AppendVbsLine scriptText, "Set fso = CreateObject(""Scripting.FileSystemObject"")"
    AppendVbsLine scriptText, "Set shell = CreateObject(""WScript.Shell"")"
    AppendVbsLine scriptText, "logPath = WScript.Arguments(3)"
    AppendVbsLine scriptText, "LogLine ""START"""
    AppendVbsLine scriptText, "LogLine ""NEW="" & WScript.Arguments(0)"
    AppendVbsLine scriptText, "LogLine ""OLD="" & WScript.Arguments(1)"
    AppendVbsLine scriptText, "WScript.Sleep 5000"
    AppendVbsLine scriptText, "LogLine ""STARTING EXCEL COM"""
    AppendVbsLine scriptText, "On Error Resume Next"
    AppendVbsLine scriptText, "Set xl = CreateObject(""Excel.Application"")"
    AppendVbsLine scriptText, "If Err.Number <> 0 Then LogLine ""EXCEL CREATE ERROR "" & Err.Number & "" "" & Err.Description: WScript.Quit 10"
    AppendVbsLine scriptText, "Err.Clear"
    AppendVbsLine scriptText, "xl.Visible = False"
    AppendVbsLine scriptText, "xl.DisplayAlerts = False"
    AppendVbsLine scriptText, "xl.AskToUpdateLinks = False"
    AppendVbsLine scriptText, "LogLine ""OPENING NEW XLSM"""
    AppendVbsLine scriptText, "Set wb = xl.Workbooks.Open(WScript.Arguments(0), 0, False)"
    AppendVbsLine scriptText, "If Err.Number <> 0 Then LogLine ""OPEN ERROR "" & Err.Number & "" "" & Err.Description: xl.Quit: WScript.Quit 11"
    AppendVbsLine scriptText, "Err.Clear"
    AppendVbsLine scriptText, "LogLine ""SAVING TO SHAREPOINT URL"""
    AppendVbsLine scriptText, "wb.SaveAs WScript.Arguments(1), 52, , , False, False, 1, 2, False"
    AppendVbsLine scriptText, "If Err.Number <> 0 Then LogLine ""SAVEAS ERROR "" & Err.Number & "" "" & Err.Description: wb.Close False: xl.Quit: WScript.Quit 12"
    AppendVbsLine scriptText, "Err.Clear"
    AppendVbsLine scriptText, "LogLine ""SAVEAS OK"""
    AppendVbsLine scriptText, "wb.Close False"
    AppendVbsLine scriptText, "Set wb = Nothing"
    AppendVbsLine scriptText, "LogLine ""OPENING SHAREPOINT COPY"""
    AppendVbsLine scriptText, "Set wb = xl.Workbooks.Open(WScript.Arguments(1), 0, False)"
    AppendVbsLine scriptText, "If Err.Number <> 0 Then LogLine ""REOPEN ERROR "" & Err.Number & "" "" & Err.Description: xl.Quit: WScript.Quit 13"
    AppendVbsLine scriptText, "Err.Clear"
    AppendVbsLine scriptText, "xl.Visible = True"
    AppendVbsLine scriptText, "LogLine ""REOPEN OK"""
    AppendVbsLine scriptText, "LogLine ""END"""
    AppendVbsLine scriptText, "WScript.Quit 0"
    AppendVbsLine scriptText, "Sub LogLine(ByVal value)"
    AppendVbsLine scriptText, "On Error Resume Next"
    AppendVbsLine scriptText, "Dim t"
    AppendVbsLine scriptText, "Set t = fso.OpenTextFile(logPath, 8, True)"
    AppendVbsLine scriptText, "t.WriteLine Now & "" | "" & value"
    AppendVbsLine scriptText, "t.Close"
    AppendVbsLine scriptText, "End Sub"

    Set ts = fso.CreateTextFile(scriptPath, True, False)
    ts.Write scriptText
    ts.Close

    wscriptPath = Environ$("WINDIR") & "\System32\wscript.exe"
    launchParams = QuoteArg(scriptPath) & " " & QuoteArg(newFile) & " " & _
                   QuoteArg(oldFile) & " " & QuoteArg(CStr(processId)) & " " & QuoteArg(logPath)

    shellResult = ShellExecute(0, "open", wscriptPath, launchParams, vbNullString, 0)
    If shellResult <= 32 Then
        Err.Raise vbObjectError + 1022, , _
            "Windows nije mogao da pokrene updater. ShellExecute=" & CStr(shellResult)
    End If

    Exit Sub

CreateError:
    Err.Raise vbObjectError + 1021, , "Updater nije uspeo da pripremi SharePoint updater." & vbCrLf & Err.Description
End Sub

Private Sub AppendVbsLine(ByRef scriptText As String, ByVal lineText As String)
    scriptText = scriptText & lineText & vbCrLf
End Sub

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr(34) & Replace(value, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function
