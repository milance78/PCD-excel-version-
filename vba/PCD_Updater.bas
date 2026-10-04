Attribute VB_Name = "PCD_Updater"
Option Explicit

Private Const VERSION_URL As String = "https://api.github.com/repos/milance78/PCD-excel-version-/contents/VERSION.json?ref=main"
Private Const ARTIFACT_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-latest.xlsm"
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

    ReplaceThroughExcelCom tempPath, ThisWorkbook.FullName
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

Private Sub ReplaceThroughExcelCom(ByVal newFile As String, ByVal oldFile As String)
    Dim xl As Object
    Dim wb As Object

    On Error GoTo ReplaceError

    Application.StatusBar = "PCD Excel: upisujem novu verziju na SharePoint..."

    Set xl = CreateObject("Excel.Application")
    If xl Is Nothing Then Err.Raise vbObjectError + 1020, , "Nije moguce pokrenuti pomocni Excel."

    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False

    Set wb = xl.Workbooks.Open(newFile, 0, False)
    If wb Is Nothing Then Err.Raise vbObjectError + 1021, , "Nova XLSM verzija nije mogla da se otvori."

    wb.SaveAs oldFile, 52, , , False, False, 1, 2, False
    wb.Close False
    Set wb = Nothing

    Set wb = xl.Workbooks.Open(oldFile, 0, False)
    If wb Is Nothing Then Err.Raise vbObjectError + 1022, , "SharePoint kopija nije mogla da se ponovo otvori."

    xl.Visible = True
    Set wb = Nothing
    Set xl = Nothing
    Exit Sub

ReplaceError:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set wb = Nothing
    Set xl = Nothing
    On Error GoTo 0
    Err.Raise vbObjectError + 1023, , _
        "Direktna zamena preko Excel COM-a nije uspela." & vbCrLf & Err.Description
End Sub

Private Sub AppendVbsLine(ByRef scriptText As String, ByVal lineText As String)
    scriptText = scriptText & lineText & vbCrLf
End Sub

Private Function QuoteArg(ByVal value As String) As String
    QuoteArg = Chr(34) & Replace(value, Chr(34), Chr(34) & Chr(34)) & Chr(34)
End Function


Private Sub LogUpdaterLaunch(ByVal stage As String, ByVal wscriptPath As String, ByVal launchParams As String)
    Dim fso As Object
    Dim ts As Object
    Dim p As String

    On Error Resume Next
    p = Environ$("TEMP") & "\PCD-Excel-launch.log"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(p, 8, True)
    ts.WriteLine Now & " | " & stage
    ts.WriteLine "WSCRIPT=" & wscriptPath
    ts.WriteLine "PARAMS=" & launchParams
    ts.Close
End Sub