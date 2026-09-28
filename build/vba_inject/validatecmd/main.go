package main

import (
    "fmt"
    "os"
    "strings"

    "github.com/kay-ws/ovba-writer/vbaproject"
)

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

    required := os.Args[2]
    for _, m := range p.Modules {
        if m.Name == "Module1" {
            if strings.Contains(m.Source, required) {
                fmt.Printf("VALID: Module1 contains required text: %s\n", required)
                return
            }
            panic("FINAL XLSM VBA validation failed: Module1 does not contain required text")
        }
    }

    panic("FINAL XLSM VBA validation failed: Module1 not found")
}
