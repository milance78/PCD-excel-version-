package main

import (
	"fmt"
	"strings"
)

type blockFrame struct {
	kind string
	line int
}

func maskForStructure(line string) string {
	var b strings.Builder
	inString := false

	for i := 0; i < len(line); i++ {
		c := line[i]

		if inString {
			b.WriteByte(' ')
			if c == '"' {
				if i+1 < len(line) && line[i+1] == '"' {
					b.WriteByte(' ')
					i++
				} else {
					inString = false
				}
			}
			continue
		}

		if c == '"' {
			inString = true
			b.WriteByte(' ')
			continue
		}

		if c == '\'' {
			for j := i; j < len(line); j++ {
				b.WriteByte(' ')
			}
			break
		}

		b.WriteByte(c)
	}

	return b.String()
}

func checkVBAContinuationRules(moduleName, source string) {
	lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")

	for i, raw := range lines {
		masked := strings.TrimSpace(maskForStructure(raw))
		if masked == "" {
			continue
		}

		if strings.HasSuffix(masked, "_") {
			if i+1 >= len(lines) {
				fail(fmt.Sprintf("%s line %d: continuation '_' has no following line", moduleName, i+1))
			}

			next := strings.TrimSpace(maskForStructure(lines[i+1]))
			if next == "" {
				fail(fmt.Sprintf("%s line %d: continuation '_' is followed by a blank line", moduleName, i+1))
			}
		}
	}
}

func popVBABlock(moduleName string, stack *[]blockFrame, want string, line int) {
	if len(*stack) == 0 {
		fail(fmt.Sprintf("%s line %d: unexpected end of %s block", moduleName, line, want))
	}

	got := (*stack)[len(*stack)-1]
	if got.kind != want {
		fail(fmt.Sprintf(
			"%s line %d: closes %s opened at line %d; expected %s",
			moduleName, line, got.kind, got.line, want,
		))
	}

	*stack = (*stack)[:len(*stack)-1]
}

func checkVBABlockStructure(moduleName, source string) {
	lines := strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n")
	stack := make([]blockFrame, 0)

	for i, raw := range lines {
		lineNo := i + 1
		line := strings.TrimSpace(maskForStructure(raw))
		if line == "" {
			continue
		}

		l := strings.ToLower(line)

		// Procedure nesting is checked separately by checkProcedureBlocks.
		if strings.HasPrefix(l, "public sub ") ||
			strings.HasPrefix(l, "private sub ") ||
			strings.HasPrefix(l, "sub ") ||
			strings.HasPrefix(l, "public function ") ||
			strings.HasPrefix(l, "private function ") ||
			strings.HasPrefix(l, "function ") ||
			strings.HasPrefix(l, "public property ") ||
			strings.HasPrefix(l, "private property ") ||
			strings.HasPrefix(l, "property ") ||
			l == "end sub" ||
			l == "end function" ||
			l == "end property" {
			continue
		}

		// A single-line If has code after Then and therefore has no End If.
		if strings.HasPrefix(l, "if ") {
			if p := strings.Index(l, " then"); p >= 0 {
				rest := strings.TrimSpace(l[p+5:])
				if rest == "" {
					stack = append(stack, blockFrame{"If", lineNo})
				}
			}
		}

		if strings.HasPrefix(l, "for ") {
			stack = append(stack, blockFrame{"For", lineNo})
		}

		if l == "do" || strings.HasPrefix(l, "do ") {
			stack = append(stack, blockFrame{"Do", lineNo})
		}

		if strings.HasPrefix(l, "select case ") {
			stack = append(stack, blockFrame{"Select", lineNo})
		}

		if strings.HasPrefix(l, "with ") {
			stack = append(stack, blockFrame{"With", lineNo})
		}

		if strings.HasPrefix(l, "while ") {
			stack = append(stack, blockFrame{"While", lineNo})
		}

		if strings.HasPrefix(l, "end if") {
			popVBABlock(moduleName, &stack, "If", lineNo)
		}

		if strings.HasPrefix(l, "next") {
			popVBABlock(moduleName, &stack, "For", lineNo)
		}

		if strings.HasPrefix(l, "loop") {
			popVBABlock(moduleName, &stack, "Do", lineNo)
		}

		if strings.HasPrefix(l, "end select") {
			popVBABlock(moduleName, &stack, "Select", lineNo)
		}

		if strings.HasPrefix(l, "end with") {
			popVBABlock(moduleName, &stack, "With", lineNo)
		}

		if strings.HasPrefix(l, "wend") {
			popVBABlock(moduleName, &stack, "While", lineNo)
		}
	}

	if len(stack) != 0 {
		top := stack[len(stack)-1]
		fail(fmt.Sprintf(
			"%s: unclosed %s block opened at line %d",
			moduleName, top.kind, top.line,
		))
	}
}

func validateEmbeddedVBScript(moduleName, source string) {
	// ScheduleReplacement creates VBScript one source line at a time:
	//     scriptText = scriptText & "..." & vbCrLf
	// Reconstruct the static fragments and validate their block structure.
	var fragments []string

	for _, raw := range strings.Split(strings.ReplaceAll(source, "\r\n", "\n"), "\n") {
		t := strings.TrimSpace(raw)
		if !strings.HasPrefix(strings.ToLower(t), "scripttext =") {
			continue
		}

		firstQuote := strings.IndexByte(t, '"')
		if firstQuote < 0 {
			continue
		}

		var b strings.Builder
		inString := false

		for i := firstQuote; i < len(t); i++ {
			c := t[i]

			if !inString {
				if c == '"' {
					inString = true
				}
				continue
			}

			if c == '"' {
				if i+1 < len(t) && t[i+1] == '"' {
					b.WriteByte('"')
					i++
					continue
				}

				fragments = append(fragments, b.String())
				b.Reset()
				inString = false
				continue
			}

			b.WriteByte(c)
		}
	}

	if len(fragments) == 0 {
		fail(moduleName + ": no scriptText fragments found for embedded VBScript")
	}

	vbs := strings.Join(fragments, "\n")
	lines := strings.Split(vbs, "\n")

	type vbsFrame struct {
		kind string
		line int
	}

	stack := make([]vbsFrame, 0)

	for i, raw := range lines {
		lineNo := i + 1
		line := strings.TrimSpace(raw)

		if line == "" {
			fail(fmt.Sprintf("%s generated VBScript line %d is blank", moduleName, lineNo))
		}

		l := strings.ToLower(line)

		if strings.HasPrefix(l, "if ") &&
			strings.HasSuffix(l, " then") {
			stack = append(stack, vbsFrame{"If", lineNo})
		}

		if strings.HasPrefix(l, "for ") {
			stack = append(stack, vbsFrame{"For", lineNo})
		}

		if l == "do" || strings.HasPrefix(l, "do ") {
			stack = append(stack, vbsFrame{"Do", lineNo})
		}

		if strings.HasPrefix(l, "end if") {
			if len(stack) == 0 || stack[len(stack)-1].kind != "If" {
				fail(fmt.Sprintf("%s generated VBScript line %d: unexpected End If", moduleName, lineNo))
			}
			stack = stack[:len(stack)-1]
		}

		if strings.HasPrefix(l, "next") {
			if len(stack) == 0 || stack[len(stack)-1].kind != "For" {
				fail(fmt.Sprintf("%s generated VBScript line %d: unexpected Next", moduleName, lineNo))
			}
			stack = stack[:len(stack)-1]
		}

		if strings.HasPrefix(l, "loop") {
			if len(stack) == 0 || stack[len(stack)-1].kind != "Do" {
				fail(fmt.Sprintf("%s generated VBScript line %d: unexpected Loop", moduleName, lineNo))
			}
			stack = stack[:len(stack)-1]
		}
	}

	if len(stack) != 0 {
		top := stack[len(stack)-1]
		fail(fmt.Sprintf(
			"%s generated VBScript: unclosed %s block opened at generated line %d",
			moduleName, top.kind, top.line,
		))
	}
}
