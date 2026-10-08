Attribute VB_Name = "PCD_Updater"
Option Explicit

Private Const SHAREPOINT_LATEST_URL As String = "https://proximuscorp-my.sharepoint.com/personal/milan_pavlovic_proximus_com/Documents/Desktop/PCD-Excel-Version-dev.xlsm"
Private Const UPDATE_TIMEOUT_SECONDS As Long = 30

Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, ByRef lpdwProcessId As Long) As Long

Public Sub CheckForUpdate()
    Dim stage As String
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
    stage = "INIT"

    Dim currentVersion As String
    Dim remoteVersion As String
    Dim manifestText As String
    Dim tempPath As String
    Dim answer As VbMsgBoxResult
    Dim localSha256 As String
    Dim cacheBust As String

    LogCheckPoint "AFTER START"

    stage = "READ LOCAL VERSION"
    currentVersion = Trim$(CStr(ThisWorkbook.Worksheets("Intervention en cours").Range("H2").Value))
    If Len(currentVersion) = 0 Then currentVersion = "0.0.0"

    Application.StatusBar = "PCD Excel: proveravam novu verziju..."
    stage = "BEFORE SHAREPOINT VERSION"
    LogCheckPoint "BEFORE SHAREPOINT VERSION"
    remoteVersion = GetSharePointLatestVersion(SHAREPOINT_LATEST_URL)
    stage = "AFTER SHAREPOINT VERSION"
    LogCheckPoint "AFTER SHAREPOINT VERSION"

    If Len(remoteVersion) = 0 Then Err.Raise vbObjectError + 1001, , "SharePoint nije vratio broj verzije."

    stage = "BEFORE VERSION COMPARE"
    LogCheckPoint "BEFORE VERSION COMPARE"
    If CompareVersions(remoteVersion, currentVersion) <= 0 Then
        LogCheckPoint "VERSION IS CURRENT"
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

    stage = "BEFORE DOWNLOAD"
    LogCheckPoint "BEFORE DOWNLOAD"
    tempPath = DownloadSharePointUpdate(SHAREPOINT_LATEST_URL, remoteVersion)
    stage = "AFTER DOWNLOAD"
    LogCheckPoint "AFTER DOWNLOAD"
    If Len(tempPath) = 0 Then Err.Raise vbObjectError + 1003, , "Preuzimanje nove verzije nije uspelo."

    Application.StatusBar = "PCD Excel: proveravam preuzetu verziju..."
    If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        Err.Raise vbObjectError + 1005, , "Automatsko azuriranje je podrzano za .xlsm fajl."
    End If

    stage = "BEFORE SCHEDULE REPLACEMENT"
    LogCheckPoint "BEFORE SCHEDULE REPLACEMENT"
    ScheduleReplacement tempPath, ThisWorkbook.FullName, CurrentExcelProcessId
    stage = "AFTER SCHEDULE REPLACEMENT"
    LogCheckPoint "AFTER SCHEDULE REPLACEMENT"
    Application.StatusBar = False

    MsgBox "Nova verzija je preuzeta i proverena." & vbCrLf & vbCrLf & _
           "Excel ce sada zatvoriti staru verziju, zameniti je novom i ponovo je otvoriti.", _
           vbInformation, "PCD Excel - azuriranje"

    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    Application.Quit
    Exit Sub

UpdateError:
    LogCheckPoint "ERROR STAGE=" & stage & " NUMBER=" & CStr(Err.Number) & " SOURCE=" & Err.Source & " DESCRIPTION=" & Err.Description
    Application.StatusBar = False
    MsgBox "Azuriranje nije izvrseno." & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "Mozes nastaviti da koristis ovu verziju. Ako je korporativna mreza blokirala GitHub, koristi rucno preuzimanje najnovijeg XLSM fajla.", _
           vbExclamation, "PCD Excel - azuriranje"
End Sub

Private Sub LogCheckPoint(ByVal value As String)
    Dim fso As Object
    Dim ts As Object
    Dim p As String
    On Error Resume Next
    p = Environ$("TEMP") & "\PCD-Excel-checkforupdate.log"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(p, 8, True)
    ts.WriteLine Now & " | " & value
    ts.Close
End Sub

Private Function GetSharePointLatestVersion(ByVal sharePointUrl As String) As String
    Dim xl As Object
    Dim wb As Object
    Dim oldSecurity As Long

    On Error GoTo ErrorHandler
    LogCheckPoint "SP VERSION BEFORE EXCEL"

    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    oldSecurity = xl.AutomationSecurity
    xl.AutomationSecurity = 3

    LogCheckPoint "SP VERSION BEFORE OPEN"
    Set wb = xl.Workbooks.Open(sharePointUrl, 0, True, , , , True, , , , False)
    LogCheckPoint "SP VERSION AFTER OPEN"

    xl.AutomationSecurity = oldSecurity
    GetSharePointLatestVersion = Trim$(CStr(wb.Worksheets("Intervention en cours").Range("H2").Value))

    wb.Close False
    Set wb = Nothing
    xl.Quit
    Set xl = Nothing
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set wb = Nothing
    Set xl = Nothing
    Err.Raise vbObjectError + 1010, , "SharePoint provera verzije nije uspela." & vbCrLf & Err.Description
End Function

Private Function DownloadSharePointUpdate(ByVal sharePointUrl As String, ByVal remoteVersion As String) As String
    Dim xl As Object
    Dim wb As Object
    Dim tempPath As String
    Dim oldSecurity As Long
    Dim openedVersion As String

    tempPath = Environ$("TEMP") & "\PCD-Excel-update-" & Replace(remoteVersion, ".", "_") & ".xlsm"

    On Error Resume Next
    Kill tempPath
    On Error GoTo ErrorHandler

    LogCheckPoint "SP DOWNLOAD BEFORE EXCEL"
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    oldSecurity = xl.AutomationSecurity
    xl.AutomationSecurity = 3

    LogCheckPoint "SP DOWNLOAD BEFORE OPEN"
    Set wb = xl.Workbooks.Open(sharePointUrl, 0, True, , , , True, , , , False)
    LogCheckPoint "SP DOWNLOAD AFTER OPEN"

    xl.AutomationSecurity = oldSecurity
    openedVersion = Trim$(CStr(wb.Worksheets("Intervention en cours").Range("H2").Value))
    If CompareVersions(openedVersion, remoteVersion) <> 0 Then
        Err.Raise vbObjectError + 1011, , "SharePoint fajl se promenio tokom preuzimanja."
    End If

    LogCheckPoint "SP DOWNLOAD BEFORE SAVECOPYAS"
    wb.SaveCopyAs tempPath
    LogCheckPoint "SP DOWNLOAD AFTER SAVECOPYAS"

    wb.Close False
    Set wb = Nothing
    xl.Quit
    Set xl = Nothing

    DownloadSharePointUpdate = tempPath
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Not xl Is Nothing Then xl.Quit
    Set wb = Nothing
    Set xl = Nothing
    Err.Raise vbObjectError + 1012, , "Preuzimanje sa SharePoint-a nije uspelo." & vbCrLf & Err.Description
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
    Dim shell As Object

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

    Set shell = CreateObject("WScript.Shell")
    shell.Run "wscript.exe " & QuoteArg(scriptPath) & " " & QuoteArg(newFile) & " " & _
              QuoteArg(oldFile) & " " & QuoteArg(CStr(processId)) & " " & QuoteArg(logPath), 0, False
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
