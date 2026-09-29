package main

import (
    "fmt"
    "os"
    "strings"

    "github.com/kay-ws/ovba-writer/vbaproject"
)

func read(path string) []byte {
    b, err := os.ReadFile(path)
    if err != nil {
        panic(err)
    }
    return b
}

func write(path string, b []byte) {
    if err := os.WriteFile(path, b, 0644); err != nil {
        panic(err)
    }
}

func stripNameAndOptionExplicit(s string) string {
    lines := strings.Split(strings.ReplaceAll(s, "\r\n", "\n"), "\n")
    out := make([]string, 0, len(lines))

    for _, line := range lines {
        t := strings.TrimSpace(line)

        if strings.HasPrefix(t, "Attribute VB_Name ") ||
            strings.EqualFold(t, "Option Explicit") {
            continue
        }

        out = append(out, line)
    }

    return strings.TrimRight(strings.Join(out, "\r\n"), "\r\n")
}

func main() {
    if len(os.Args) != 5 {
        panic("usage: inject base.bin magic.bas updater.bas output.bin")
    }

    basePath, magicPath, updaterPath, outPath :=
        os.Args[1], os.Args[2], os.Args[3], os.Args[4]

    p, err := vbaproject.Read(read(basePath))
    if err != nil {
        panic(err)
    }

    magic := stripNameAndOptionExplicit(string(read(magicPath)))
    _ = updaterPath

    // DEV-28 diagnostic:
    // Add the complete ScheduleReplacement + QuoteArg code,
    // but do not call them.
    module1Source := "Attribute VB_Name = \"Module1\"\r\n" +
        "Option Explicit\r\n" +
        "Private Const VERSION_URL As String = \"https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/VERSION.json\"\r\n" +
        "Private Const ARTIFACT_URL As String = \"https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-dev.xlsm\"\r\n" +
        "Private Const UPDATE_TIMEOUT_SECONDS As Long = 30\r\n" +
        magic + "\r\n\r\n" +

        "Public Sub CheckForUpdate()\r\n" +
        "    MsgBox \"DEV-28.1: Module1 loads with split ScheduleReplacement scriptText.\", vbInformation, \"PCD DIJAGNOSTIKA\"\r\n" +
        "End Sub\r\n\r\n" +

        "Private Function HttpGetText(ByVal url As String) As String\r\n" +
        "    Dim http As Object\r\n" +
        "    Set http = CreateObject(\"MSXML2.XMLHTTP.6.0\")\r\n" +
        "    http.Open \"GET\", url, False\r\n" +
        "    http.setRequestHeader \"Cache-Control\", \"no-cache\"\r\n" +
        "    http.Send\r\n" +
        "    If http.Status < 200 Or http.Status >= 300 Then\r\n" +
        "        Err.Raise vbObjectError + 1010, , \"GitHub HTTP greska: \" & http.Status & \" \" & http.StatusText\r\n" +
        "    End If\r\n" +
        "    HttpGetText = CStr(http.responseText)\r\n" +
        "End Function\r\n\r\n" +

        "Private Function JsonValue(ByVal json As String, ByVal key As String) As String\r\n" +
        "    Dim re As Object\r\n" +
        "    Dim matches As Object\r\n" +
        "    Set re = CreateObject(\"VBScript.RegExp\")\r\n" +
        "    re.Global = False\r\n" +
        "    re.IgnoreCase = True\r\n" +
        "    re.Pattern = \"\"\"\" & key & \"\"\"\" & \"\\s*:\\s*\"\"\"([^\"\"]*)\"\"\"\"\r\n" +
        "    Set matches = re.Execute(json)\r\n" +
        "    If matches.Count > 0 Then JsonValue = matches(0).SubMatches(0)\r\n" +
        "End Function\r\n\r\n" +

        "Private Function CompareVersions(ByVal a As String, ByVal b As String) As Long\r\n" +
        "    Dim pa() As String, pb() As String\r\n" +
        "    Dim i As Long, na As Long, nb As Long\r\n" +
        "    Dim aCore As String, bCore As String\r\n" +
        "    aCore = Split(a, \"-\")(0)\r\n" +
        "    bCore = Split(b, \"-\")(0)\r\n" +
        "    pa = Split(aCore, \".\")\r\n" +
        "    pb = Split(bCore, \".\")\r\n" +
        "    For i = 0 To 2\r\n" +
        "        na = 0: nb = 0\r\n" +
        "        If i <= UBound(pa) And IsNumeric(pa(i)) Then na = CLng(pa(i))\r\n" +
        "        If i <= UBound(pb) And IsNumeric(pb(i)) Then nb = CLng(pb(i))\r\n" +
        "        If na > nb Then CompareVersions = 1: Exit Function\r\n" +
        "        If na < nb Then CompareVersions = -1: Exit Function\r\n" +
        "    Next i\r\n" +
        "    CompareVersions = CompareBuildSuffix(a, b)\r\n" +
        "End Function\r\n\r\n" +

        "Private Function CompareBuildSuffix(ByVal a As String, ByVal b As String) As Long\r\n" +
        "    Dim da As Long, db As Long\r\n" +
        "    da = LastNumberAfterDash(a)\r\n" +
        "    db = LastNumberAfterDash(b)\r\n" +
        "    If da > db Then\r\n" +
        "        CompareBuildSuffix = 1\r\n" +
        "    ElseIf da < db Then\r\n" +
        "        CompareBuildSuffix = -1\r\n" +
        "    Else\r\n" +
        "        CompareBuildSuffix = 0\r\n" +
        "    End If\r\n" +
        "End Function\r\n\r\n" +

        "Private Function LastNumberAfterDash(ByVal value As String) As Long\r\n" +
        "    Dim parts() As String\r\n" +
        "    Dim i As Long\r\n" +
        "    parts = Split(value, \"-\")\r\n" +
        "    For i = UBound(parts) To 1 Step -1\r\n" +
        "        If IsNumeric(parts(i)) Then\r\n" +
        "            LastNumberAfterDash = CLng(parts(i))\r\n" +
        "            Exit Function\r\n" +
        "        End If\r\n" +
        "    Next i\r\n" +
        "End Function\r\n\r\n" +

        "Private Function DownloadUpdate(ByVal remoteVersion As String) As String\r\n" +
        "    Dim http As Object\r\n" +
        "    Dim stream As Object\r\n" +
        "    Dim tempPath As String\r\n" +
        "    tempPath = Environ$(\"TEMP\") & \"PCD-Excel-update-\" & Replace(remoteVersion, \".\", \"_\") & \".xlsm\"\r\n" +
        "    On Error Resume Next\r\n" +
        "    Kill tempPath\r\n" +
        "    On Error GoTo 0\r\n" +
        "    Set http = CreateObject(\"MSXML2.XMLHTTP.6.0\")\r\n" +
        "    http.Open \"GET\", ARTIFACT_URL & \"?v=\" & Replace(remoteVersion, \" \", \"%20\") & \"&t=\" & CStr(Timer), False\r\n" +
        "    http.setRequestHeader \"Cache-Control\", \"no-cache\"\r\n" +
        "    http.Send\r\n" +
        "    If http.Status < 200 Or http.Status >= 300 Then\r\n" +
        "        Err.Raise vbObjectError + 1011, , \"Preuzimanje XLSM fajla nije uspelo: HTTP \" & http.Status\r\n" +
        "    End If\r\n" +
        "    Set stream = CreateObject(\"ADODB.Stream\")\r\n" +
        "    stream.Type = 1\r\n" +
        "    stream.Open\r\n" +
        "    stream.Write http.responseBody\r\n" +
        "    stream.SaveToFile tempPath, 2\r\n" +
        "    stream.Close\r\n" +
        "    DownloadUpdate = tempPath\r\n" +
        "End Function\r\n\r\n" +

        "Private Function FileSha256(ByVal filePath As String) As String\r\n" +
        "    Dim outputPath As String\r\n" +
        "    Dim shell As Object\r\n" +
        "    Dim fso As Object\r\n" +
        "    Dim ts As Object\r\n" +
        "    Dim text As String\r\n" +
        "    Dim matches As Object\r\n" +
        "    Dim re As Object\r\n" +
        "    outputPath = Environ$(\"TEMP\") & \"PCD-sha256-\" & Format$(Timer * 1000, \"0\") & \".txt\"\r\n" +
        "    Set shell = CreateObject(\"WScript.Shell\")\r\n" +
        "    shell.Run \"cmd.exe /c certutil -hashfile \" & QuoteArg(filePath) & \" SHA256 > \" & QuoteArg(outputPath), 0, True\r\n" +
        "    Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n" +
        "    If Not fso.FileExists(outputPath) Then Err.Raise vbObjectError + 1012, , \"Windows nije mogao da izracuna SHA-256.\"\r\n" +
        "    Set ts = fso.OpenTextFile(outputPath, 1, False)\r\n" +
        "    text = ts.ReadAll\r\n" +
        "    ts.Close\r\n" +
        "    On Error Resume Next\r\n" +
        "    fso.DeleteFile outputPath, True\r\n" +
        "    On Error GoTo 0\r\n" +
        "    Set re = CreateObject(\"VBScript.RegExp\")\r\n" +
        "    re.Global = False\r\n" +
        "    re.IgnoreCase = True\r\n" +
        "    re.Pattern = \"([0-9A-Fa-f]{64})\"\r\n" +
        "    Set matches = re.Execute(text)\r\n" +
        "    If matches.Count = 0 Then Err.Raise vbObjectError + 1013, , \"Windows nije vratio SHA-256 vrednost.\"\r\n" +
        "    FileSha256 = LCase$(matches(0).Value)\r\n" +
        "End Function\r\n\r\n" +

        "Private Sub ScheduleReplacement(ByVal newFile As String, ByVal oldFile As String)\r\n" +
        "    Dim scriptPath As String\r\n" +
        "    Dim logPath As String\r\n" +
        "    Dim scriptText As String\r\n" +
        "    Dim fso As Object\r\n" +
        "    Dim ts As Object\r\n" +
        "    Dim shell As Object\r\n" +
        "\r\n" +
        "    scriptPath = Environ$(\"TEMP\") & \"\\PCD-Excel-updater.vbs\"\r\n" +
        "    logPath = Environ$(\"TEMP\") & \"\\PCD-Excel-updater.log\"\r\n" +
        "\r\n" +
        "    On Error GoTo CreateError\r\n" +
        "\r\n" +
        "    Set fso = CreateObject(\"Scripting.FileSystemObject\")\r\n" +
        "\r\n" +
        "    On Error Resume Next\r\n" +
        "    If fso.FileExists(logPath) Then fso.DeleteFile logPath, True\r\n" +
        "    On Error GoTo CreateError\r\n" +
        "\r\n" +
        "    scriptText = \"Option Explicit\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"Dim fso, shell, newFile, oldFile, logPath, i, replaced\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"Set fso = CreateObject(\"\"Scripting.FileSystemObject\"\")\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"Set shell = CreateObject(\"\"WScript.Shell\"\")\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"newFile = WScript.Arguments(0)\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"oldFile = WScript.Arguments(1)\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"logPath = WScript.Arguments(2)\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"LogLine \"\"START\"\"\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"LogLine \"\"NEW=\"\" & newFile\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"LogLine \"\"OLD=\"\" & oldFile\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"For i = 1 To 120\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  On Error Resume Next\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  Err.Clear\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  If fso.FileExists(oldFile) Then fso.DeleteFile oldFile, True\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  If Err.Number = 0 Then\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"    LogLine \"\"DELETE OK attempt=\"\" & i\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"    Err.Clear\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"    fso.MoveFile newFile, oldFile\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"    If Err.Number = 0 Then\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"      replaced = True\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"      LogLine \"\"MOVE OK attempt=\"\" & i\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"      Exit For\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"    Else\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"      LogLine \"\"MOVE ERROR \"\" & Err.Number & \"\" \"\" & Err.Description\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"    End If\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  Else\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"    LogLine \"\"DELETE ERROR \"\" & Err.Number & \"\" \"\" & Err.Description\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  End If\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  Err.Clear\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  On Error GoTo 0\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  WScript.Sleep 1000\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"Next\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"If replaced Then\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  LogLine \"\"REPLACED OK\"\"\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  LogLine \"\"LAUNCHING \"\" & oldFile\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  shell.Run Chr(34) & oldFile & Chr(34), 1, False\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"Else\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  LogLine \"\"REPLACEMENT FAILED after 120 attempts\"\"\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"End If\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"LogLine \"\"END\"\"\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"Sub LogLine(ByVal message)\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  Dim logFile\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  On Error Resume Next\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  Set logFile = fso.OpenTextFile(logPath, 8, True)\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  logFile.WriteLine Now & \"\" | \"\" & message\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"  logFile.Close\" & vbCrLf\r\n" +
        "    scriptText = scriptText & \"End Sub\"\r\n" +
        "    scriptText = scriptText & \r\n" +
        "    Set ts = fso.CreateTextFile(scriptPath, True, False)\r\n" +
        "    ts.Write scriptText\r\n" +
        "    ts.Close\r\n" +
        "\r\n" +
        "    If Not fso.FileExists(scriptPath) Then\r\n" +
        "        Err.Raise vbObjectError + 1020, , \"Updater nije uspeo da napravi VBS fajl.\"\r\n" +
        "    End If\r\n" +
        "\r\n" +
        "    Set shell = CreateObject(\"WScript.Shell\")\r\n" +
        "    shell.Run \"wscript.exe \" & QuoteArg(scriptPath) & \" \" & _\r\n" +
        "              QuoteArg(newFile) & \" \" & QuoteArg(oldFile) & \" \" & _\r\n" +
        "              QuoteArg(logPath), 0, False\r\n" +
        "    Exit Sub\r\n" +
        "\r\n" +
        "CreateError:\r\n" +
        "    Err.Raise vbObjectError + 1021, , \"Updater nije uspeo da pripremi eksterni updater.\" & vbCrLf & Err.Description\r\n" +
        "End Sub\r\n\r\n" +

        "Private Function QuoteArg(ByVal value As String) As String\r\n" +
        "    QuoteArg = \"\"\"\" & Replace(value, \"\"\"\", \"\"\"\"\") & \"\"\"\"\r\n" +
        "End Function\r\n"

    foundModule1, foundSheet2 := false, false

    for i := range p.Modules {
        switch p.Modules[i].Name {

        case "Module1":
            normalized, err :=
                vbaproject.NormalizeModuleSource(
                    vbaproject.ModuleStd,
                    module1Source,
                    nil,
                )

            if err != nil {
                panic(err)
            }

            p.Modules[i].Source = normalized
            foundModule1 = true

        case "Sheet2":
            eventSource :=
                "Private Sub Worksheet_Change(ByVal Target As Range)\r\n" +
                "    If Intersect(Target, Me.Range(\"B5\")) Is Nothing Then Exit Sub\r\n" +
                "    If Len(Trim$(CStr(Target.Value))) < 10 Then Exit Sub\r\n" +
                "    On Error GoTo CleanFail\r\n" +
                "    Application.EnableEvents = False\r\n" +
                "    ParseMagicImportText CStr(Target.Value), False\r\n" +
                "CleanFail:\r\n" +
                "    Application.EnableEvents = True\r\n" +
                "End Sub\r\n"

            existing := p.Modules[i]

            normalized, err :=
                vbaproject.NormalizeModuleSource(
                    vbaproject.ModuleDocument,
                    eventSource,
                    &existing,
                )

            if err != nil {
                panic(err)
            }

            p.Modules[i].Source = normalized
            foundSheet2 = true
        }
    }

    if !foundModule1 || !foundSheet2 {
        panic(fmt.Sprintf(
            "expected base modules not found: Module1=%v Sheet2=%v",
            foundModule1,
            foundSheet2,
        ))
    }

    out, err := vbaproject.Write(p)
    if err != nil {
        panic(err)
    }

    // Read the generated project back and validate the expected public macros.
    check, err := vbaproject.Read(out)
    if err != nil {
        panic(fmt.Sprintf(
            "post-write VBA validation failed: %v",
            err,
        ))
    }

    validated := false

    for _, m := range check.Modules {
        if m.Name == "Module1" {

            if !strings.Contains(
                m.Source,
                "Public Sub ImportMagicFromSheet()",
            ) {
                panic(
                    "post-write VBA validation: ImportMagicFromSheet missing",
                )
            }

            if !strings.Contains(
                m.Source,
                "Public Sub CheckForUpdate()",
            ) {
                panic(
                    "post-write VBA validation: CheckForUpdate stub missing",
                )
            }

            if !strings.Contains(
                m.Source,
                "Private Const VERSION_URL As String",
            ) ||
                !strings.Contains(
                    m.Source,
                    "Private Const ARTIFACT_URL As String",
                ) ||
                !strings.Contains(
                    m.Source,
                    "Private Function HttpGetText(ByVal url As String)",
                ) ||
                !strings.Contains(
                    m.Source,
                    "Private Function JsonValue(ByVal json As String, ByVal key As String)",
                ) ||
                !strings.Contains(
                    m.Source,
                    "Private Function CompareVersions(ByVal a As String, ByVal b As String)",
                ) ||
                !strings.Contains(
                    m.Source,
                    "Private Function DownloadUpdate(ByVal remoteVersion As String)",
                ) ||
                !strings.Contains(
                    m.Source,
                    "Private Function FileSha256(ByVal filePath As String)",
                ) ||
                !strings.Contains(
                    m.Source,
                    "Private Sub ScheduleReplacement(ByVal newFile As String, ByVal oldFile As String)",
                ) ||
                !strings.Contains(
                    m.Source,
                    "Private Function QuoteArg(ByVal value As String)",
                ) {
                panic(
                    "post-write VBA validation: DEV-28 updater declarations missing",
                )
            }

            if strings.Contains(
                m.Source,
                "Dim localSha256 As String",
            ) {
                panic(
                    "post-write VBA validation: full updater code unexpectedly present in DEV-28",
                )
            }

            firstProc := len(m.Source)

            for _, token := range []string{
                "Private Function ",
                "Public Function ",
                "Private Sub ",
                "Public Sub ",
            } {
                if i := strings.Index(m.Source, token); i >= 0 && i < firstProc {
                    firstProc = i
                }
            }

            if i := strings.Index(m.Source, "Private Const "); i >= firstProc {
                panic(
                    "post-write VBA validation: Private Const appears after first procedure",
                )
            }

            validated = true
        }
    }

    if !validated {
        panic("post-write VBA validation: Module1 not found")
    }

    write(outPath, out)

    fmt.Printf(
        "Injected and validated PCD VBA: %s (%d bytes)\n",
        outPath,
        len(out),
    )
}
