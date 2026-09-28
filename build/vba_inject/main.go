package main

import (
    "fmt"
    "os"
    "strings"

    "github.com/kay-ws/ovba-writer/vbaproject"
)

func read(path string) []byte {
    b, err := os.ReadFile(path)
    if err != nil { panic(err) }
    return b
}

func write(path string, b []byte) {
    if err := os.WriteFile(path, b, 0644); err != nil { panic(err) }
}

func stripNameAndOptionExplicit(s string) string {
    lines := strings.Split(strings.ReplaceAll(s, "\r\n", "\n"), "\n")
    out := make([]string, 0, len(lines))
    for _, line := range lines {
        t := strings.TrimSpace(line)
        if strings.HasPrefix(t, "Attribute VB_Name ") || strings.EqualFold(t, "Option Explicit") {
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
    basePath, magicPath, updaterPath, outPath := os.Args[1], os.Args[2], os.Args[3], os.Args[4]
    p, err := vbaproject.Read(read(basePath))
    if err != nil { panic(err) }

    magic := stripNameAndOptionExplicit(string(read(magicPath)))
    _ = updaterPath

    // DEV-23 diagnostic: add updater constants plus HttpGetText, but do not call it.
    // This isolates whether the updater declarations themselves make Module1 unloadable.
    module1Source := "Attribute VB_Name = \"Module1\"\r\n" +
        "Option Explicit\r\n" +
        "Private Const VERSION_URL As String = \"https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/VERSION.json\"\r\n" +
        "Private Const ARTIFACT_URL As String = \"https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-dev.xlsm\"\r\n" +
        "Private Const UPDATE_TIMEOUT_SECONDS As Long = 30\r\n" +
        magic + "\r\n\r\n" +
        "Public Sub CheckForUpdate()\r\n" +
        "    MsgBox \"DEV-23: Module1 loads with HttpGetText.\", vbInformation, \"PCD DIJAGNOSTIKA\"\r\n" +
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
        "End Function\r\n"

    foundModule1, foundSheet2 := false, false
    for i := range p.Modules {
        switch p.Modules[i].Name {
        case "Module1":
            normalized, err := vbaproject.NormalizeModuleSource(vbaproject.ModuleStd, module1Source, nil)
            if err != nil { panic(err) }
            p.Modules[i].Source = normalized
            foundModule1 = true
        case "Sheet2":
            eventSource := "Private Sub Worksheet_Change(ByVal Target As Range)\r\n" +
                "    If Intersect(Target, Me.Range(\"B5\")) Is Nothing Then Exit Sub\r\n" +
                "    If Len(Trim$(CStr(Target.Value))) < 10 Then Exit Sub\r\n" +
                "    On Error GoTo CleanFail\r\n" +
                "    Application.EnableEvents = False\r\n" +
                "    ParseMagicImportText CStr(Target.Value), False\r\n" +
                "CleanFail:\r\n" +
                "    Application.EnableEvents = True\r\n" +
                "End Sub\r\n"
            existing := p.Modules[i]
            normalized, err := vbaproject.NormalizeModuleSource(vbaproject.ModuleDocument, eventSource, &existing)
            if err != nil { panic(err) }
            p.Modules[i].Source = normalized
            foundSheet2 = true
        }
    }
    if !foundModule1 || !foundSheet2 {
        panic(fmt.Sprintf("expected base modules not found: Module1=%v Sheet2=%v", foundModule1, foundSheet2))
    }
    out, err := vbaproject.Write(p)
    if err != nil { panic(err) }

    // Read the generated project back and validate the expected public macros.
    check, err := vbaproject.Read(out)
    if err != nil { panic(fmt.Sprintf("post-write VBA validation failed: %v", err)) }
    validated := false
    for _, m := range check.Modules {
        if m.Name == "Module1" {
            if !strings.Contains(m.Source, "Public Sub ImportMagicFromSheet()") {
                panic("post-write VBA validation: ImportMagicFromSheet missing")
            }
            if !strings.Contains(m.Source, "Public Sub CheckForUpdate()") {
                panic("post-write VBA validation: CheckForUpdate stub missing")
            }
            if !strings.Contains(m.Source, "Private Const VERSION_URL As String") || !strings.Contains(m.Source, "Private Const ARTIFACT_URL As String") || !strings.Contains(m.Source, "Private Function HttpGetText(ByVal url As String)") {
                panic("post-write VBA validation: DEV-23 updater declarations missing")
            }
            if strings.Contains(m.Source, "Dim localSha256 As String") {
                panic("post-write VBA validation: full updater code unexpectedly present in DEV-23")
            }
            firstProc := len(m.Source)
            for _, token := range []string{"Private Function ", "Public Function ", "Private Sub ", "Public Sub "} {
                if i := strings.Index(m.Source, token); i >= 0 && i < firstProc { firstProc = i }
            }
            if i := strings.Index(m.Source, "Private Const "); i >= firstProc {
                panic("post-write VBA validation: Private Const appears after first procedure")
            }
            validated = true
        }
    }
    if !validated { panic("post-write VBA validation: Module1 not found") }

    write(outPath, out)
    fmt.Printf("Injected and validated PCD VBA: %s (%d bytes)\n", outPath, len(out))
}
