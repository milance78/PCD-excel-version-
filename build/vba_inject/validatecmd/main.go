package main

import (
    "fmt"
    "os"
    "strings"

    "github.com/kay-ws/ovba-writer/vbaproject"
)

func fail(msg string) {
    panic("VBA PRE-FLIGHT FAILED: " + msg)
}

// checkVBAStrings verifies the most common generator failure: an unterminated
// VBA string caused by a physical CR/LF inside a string literal. In VBA, a
// string literal cannot cross a physical source line. A doubled quote ("")
// inside a string is an escaped quote and does not close the string.
func checkVBAStrings(moduleName, source string) {
    lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
    for lineNo, line := range lines {
        inString := false
        for i := 0; i < len(line); i++ {
            if line[i] != '"' {
                continue
            }
            if inString && i+1 < len(line) && line[i+1] == '"' {
                i++
                continue
            }
            inString = !inString
        }
        if inString {
            fail(fmt.Sprintf("%s line %d: unterminated string literal", moduleName, lineNo+1))
        }
    }
}

// checkProcedureBlocks catches missing/misplaced End Sub / End Function and
// common block terminator mistakes before Excel ever sees the generated file.
func checkProcedureBlocks(moduleName, source string) {
    lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
    procedure := ""
    for lineNo, raw := range lines {
        line := strings.TrimSpace(raw)
        lower := strings.ToLower(line)

        if strings.HasPrefix(lower, "public sub ") ||
            strings.HasPrefix(lower, "private sub ") ||
            strings.HasPrefix(lower, "sub ") {
            if procedure != "" {
                fail(fmt.Sprintf("%s line %d: nested/unclosed procedure %q", moduleName, lineNo+1, procedure))
            }
            procedure = "Sub"
            continue
        }

        if strings.HasPrefix(lower, "public function ") ||
            strings.HasPrefix(lower, "private function ") ||
            strings.HasPrefix(lower, "function ") {
            if procedure != "" {
                fail(fmt.Sprintf("%s line %d: nested/unclosed procedure %q", moduleName, lineNo+1, procedure))
            }
            procedure = "Function"
            continue
        }

        if strings.EqualFold(line, "End Sub") {
            if procedure != "Sub" {
                fail(fmt.Sprintf("%s line %d: unexpected End Sub", moduleName, lineNo+1))
            }
            procedure = ""
            continue
        }

        if strings.EqualFold(line, "End Function") {
            if procedure != "Function" {
                fail(fmt.Sprintf("%s line %d: unexpected End Function", moduleName, lineNo+1))
            }
            procedure = ""
        }
    }

    if procedure != "" {
        fail(fmt.Sprintf("%s: procedure %q is not closed", moduleName, procedure))
    }
}

func validateModule1(source string, required string) {
    if !strings.Contains(source, required) {
        fail("Module1 does not contain required text: " + required)
    }

    // Module-level declarations must appear before the first procedure.
    firstProc := len(source)
    for _, token := range []string{
        "Private Function ",
        "Public Function ",
        "Private Sub ",
        "Public Sub ",
    } {
        if i := strings.Index(source, token); i >= 0 && i < firstProc {
            firstProc = i
        }
    }
    if i := strings.Index(source, "Private Const "); i >= firstProc {
        fail("Module1 has a Private Const declaration after the first procedure")
    }

    checkVBAStrings("Module1", source)
    checkProcedureBlocks("Module1", source)
}

func main() {
    if len(os.Args) != 3 {
        panic("usage: validate <vbaProject.bin> <required-text>")
    }

    data, err := os.ReadFile(os.Args[1])
    if err != nil {
        panic(err)
    }

    p, err := vbaproject.Read(data)
    if err != nil {
        panic(fmt.Sprintf("cannot read VBA project: %v", err))
    }

    foundModule1 := false
    foundSheet2 := false

    for _, m := range p.Modules {
        switch m.Name {
        case "Module1":
            foundModule1 = true
            validateModule1(m.Source, os.Args[2])
        case "Sheet2":
            foundSheet2 = true
            checkVBAStrings("Sheet2", m.Source)
            checkProcedureBlocks("Sheet2", m.Source)
        }
    }

    if !foundModule1 {
        fail("Module1 not found")
    }
    if !foundSheet2 {
        fail("Sheet2 not found")
    }

    fmt.Println("VALID: VBA pre-flight checks passed for Module1 and Sheet2")
}
