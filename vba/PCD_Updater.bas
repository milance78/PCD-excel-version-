Attribute VB_Name = "PCD_Updater"
Option Explicit

' PCD Excel updater
' The workbook stays offline at runtime except when the user explicitly checks for updates.
' GitHub is distribution only.

Private Const VERSION_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/VERSION.json"
Private Const ARTIFACT_BASE_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/"

Public Sub CheckForUpdate(Optional ByVal silent As Boolean = False)
    On Error GoTo Fail

    Dim manifest As String, remoteVersion As String, localVersion As String
    Dim artifactName As String, artifactSha As String
    Dim answer As VbMsgBoxResult

    manifest = HttpGet(VERSION_URL)
    remoteVersion = JsonString(manifest, "version")
    artifactName = JsonString(manifest, "artifact")
    artifactSha = LCase$(JsonString(manifest, "sha256"))
    localVersion = GetLocalVersion()

    If remoteVersion = "" Or artifactName = "" Then Err.Raise vbObjectError + 700, , "VERSION.json is incomplete."

    If Not IsNewerVersion(remoteVersion, localVersion) Then
        If Not silent Then MsgBox "Vous utilisez déjà la version " & localVersion & ".", vbInformation, "PCD Excel"
        Exit Sub
    End If

    answer = MsgBox("Une nouvelle version de PCD Excel est disponible." & vbCrLf & vbCrLf & _
                    "Version actuelle : " & localVersion & vbCrLf & _
                    "Nouvelle version : " & remoteVersion & vbCrLf & vbCrLf & _
                    "Voulez-vous la télécharger ?", vbQuestion + vbYesNo, "Mise à jour PCD")
    If answer <> vbYes Then Exit Sub

    InstallUpdate remoteVersion, artifactName, artifactSha
    Exit Sub

Fail:
    If Not silent Then MsgBox "La vérification de mise à jour a échoué :" & vbCrLf & Err.Description, vbExclamation, "PCD Excel"
End Sub

Private Sub InstallUpdate(ByVal newVersion As String, ByVal artifactName As String, ByVal expectedSha As String)
    Dim folder As String, newFile As String, currentFile As String, scriptFile As String
    folder = Environ$("TEMP") & "\PCD-Excel-Update"
    EnsureFolder folder

    newFile = folder & "\" & artifactName
    currentFile = ThisWorkbook.FullName

    DownloadFile ARTIFACT_BASE_URL & artifactName, newFile

    If Len(expectedSha) > 0 Then
        If LCase$(Sha256CertUtil(newFile)) <> LCase$(expectedSha) Then
            Kill newFile
            Err.Raise vbObjectError + 701, , "Le contrôle SHA-256 a échoué. La mise à jour n'est pas installée."
        End If
    End If

    scriptFile = folder & "\install-update.vbs"
    WriteUpdateScript scriptFile, currentFile, newFile

    Shell "wscript.exe " & Chr$(34) & scriptFile & Chr$(34), vbHide
    MsgBox "La version " & newVersion & " a été téléchargée et vérifiée." & vbCrLf & _
           "Excel va se fermer puis ouvrir la nouvelle version.", vbInformation, "Mise à jour PCD"
    ThisWorkbook.Close SaveChanges:=True
End Sub

Private Function HttpGet(ByVal url As String) As String
    Dim x As Object
    Set x = CreateObject("WinHttp.WinHttpRequest.5.1")
    x.Open "GET", url, False
    x.SetRequestHeader "User-Agent", "PCD-Excel-Updater"
    x.Send
    If x.Status <> 200 Then Err.Raise vbObjectError + 702, , "HTTP " & x.Status & " pour " & url
    HttpGet = x.ResponseText
End Function

Private Sub DownloadFile(ByVal url As String, ByVal target As String)
    Dim x As Object, stm As Object
    Set x = CreateObject("WinHttp.WinHttpRequest.5.1")
    x.Open "GET", url, False
    x.SetRequestHeader "User-Agent", "PCD-Excel-Updater"
    x.Send
    If x.Status <> 200 Then Err.Raise vbObjectError + 703, , "Téléchargement impossible. HTTP " & x.Status

    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write x.ResponseBody
    stm.SaveToFile target, 2
    stm.Close
End Sub

Private Function GetLocalVersion() As String
    GetLocalVersion = "0.1.0"
    On Error Resume Next
    GetLocalVersion = ThisWorkbook.Worksheets("Config").Range("B2").Value
    On Error GoTo 0
End Function

Private Function JsonString(ByVal json As String, ByVal key As String) As String
    Dim r As Object, m As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Global = False
    r.IgnoreCase = True
    r.Pattern = Chr$(34) & key & Chr$(34) & "\s*:\s*" & Chr$(34) & "([^" & Chr$(34) & "]*)" & Chr$(34)
    Set m = r.Execute(json)
    If m.Count > 0 Then JsonString = m(0).SubMatches(0)
End Function

Private Function IsNewerVersion(ByVal remote As String, ByVal local As String) As Boolean
    Dim a() As String, b() As String, i As Long, av As Long, bv As Long
    a = Split(Replace(remote, "-", "."), ".")
    b = Split(Replace(local, "-", "."), ".")
    For i = 0 To 2
        av = 0: bv = 0
        If i <= UBound(a) Then av = Val(a(i))
        If i <= UBound(b) Then bv = Val(b(i))
        If av > bv Then IsNewerVersion = True: Exit Function
        If av < bv Then Exit Function
    Next
End Function

Private Function Sha256CertUtil(ByVal filePath As String) As String
    Dim sh As Object, ex As Object, s As String
    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec("cmd /c certutil -hashfile " & Chr$(34) & filePath & Chr$(34) & " SHA256")
    s = ex.StdOut.ReadAll
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    s = Replace(s, "SHA256 hash of " & filePath & ":", "")
    s = Replace(s, "CertUtil: -hashfile command completed successfully.", "")
    Sha256CertUtil = Replace(s, " ", "")
End Function

Private Sub WriteUpdateScript(ByVal scriptFile As String, ByVal oldFile As String, ByVal newFile As String)
    Dim f As Integer, q As String, s As String
    q = Chr$(34)
    s = "Option Explicit" & vbCrLf & _
        "Dim sh, fso, oldFile, newFile, backupFile" & vbCrLf & _
        "Set sh = CreateObject(" & q & "WScript.Shell" & q & ")" & vbCrLf & _
        "Set fso = CreateObject(" & q & "Scripting.FileSystemObject" & q & ")" & vbCrLf & _
        "oldFile = " & q & oldFile & q & vbCrLf & _
        "newFile = " & q & newFile & q & vbCrLf & _
        "backupFile = oldFile & " & q & ".previous" & q & vbCrLf & _
        "Do While IsRunningExcel()" & vbCrLf & "  WScript.Sleep 500" & vbCrLf & "Loop" & vbCrLf & _
        "If fso.FileExists(backupFile) Then fso.DeleteFile backupFile, True" & vbCrLf & _
        "fso.MoveFile oldFile, backupFile" & vbCrLf & _
        "fso.CopyFile newFile, oldFile, True" & vbCrLf & _
        "sh.Run " & q & "excel.exe " & q & " & " & q & oldFile & q & ", 1, False" & vbCrLf & _
        "fso.DeleteFile WScript.ScriptFullName, True" & vbCrLf & _
        "Function IsRunningExcel()" & vbCrLf & _
        "  Dim p, svc, col" & vbCrLf & _
        "  Set svc = GetObject(" & q & "winmgmts:\\.\root\cimv2" & q & ")" & vbCrLf & _
        "  Set col = svc.ExecQuery(" & q & "Select * from Win32_Process where Name='EXCEL.EXE'" & q & ")" & vbCrLf & _
        "  IsRunningExcel = (col.Count > 0)" & vbCrLf & _
        "End Function" & vbCrLf
    f = FreeFile
    Open scriptFile For Output As #f
    Print #f, s
    Close #f
End Sub

Private Sub EnsureFolder(ByVal folder As String)
    If Dir(folder, vbDirectory) = "" Then MkDir folder
End Sub
