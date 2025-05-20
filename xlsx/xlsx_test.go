package xlsx

import (
	"encoding/xml"
	"os"
	"strings"
	"testing"
)

func TestXlsxToHTML(t *testing.T) {
	DebugHTML = true
	f, err := os.Open("test.xlsx")
	if err != nil {
		t.Fatalf("failed to open test.xlsx: %v", err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatalf("failed to stat test.xlsx: %v", err)
	}

	html, err := XLSXToHTML(f, info.Size())
	if err != nil {
		t.Fatalf("failed to convert xlsx to html: %v", err)
	}
	// write html to file
	err = os.WriteFile("test.xlsx.html", []byte(html), 0644)
	if err != nil {
		t.Fatalf("failed to write test.html: %v", err)
	}
}
func TestXlsxToHTML2(t *testing.T) {
	DebugHTML = true
	f, err := os.Open("test2.xlsx")
	if err != nil {
		t.Fatalf("failed to open test.xlsx: %v", err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatalf("failed to stat test.xlsx: %v", err)
	}

	html, err := XLSXToHTML(f, info.Size())
	if err != nil {
		t.Fatalf("failed to convert xlsx to html: %v", err)
	}
	// write html to file
	err = os.WriteFile("test2.xlsx.html", []byte(html), 0644)
	if err != nil {
		t.Fatalf("failed to write test.html: %v", err)
	}
}

func TestXlsxToHTML3(t *testing.T) {
	DebugHTML = true
	f, err := os.Open("test3.xlsx")
	if err != nil {
		t.Fatalf("failed to open test.xlsx: %v", err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatalf("failed to stat test.xlsx: %v", err)
	}

	html, err := XLSXToHTML(f, info.Size())
	if err != nil {
		t.Fatalf("failed to convert xlsx to html: %v", err)
	}
	// write html to file
	err = os.WriteFile("test3.xlsx.html", []byte(html), 0644)
	if err != nil {
		t.Fatalf("failed to write test.html: %v", err)
	}
}

func SprintXML(a any) string {
	var b strings.Builder
	enc := xml.NewEncoder(&b)
	enc.Indent("", "  ")
	enc.Encode(a)
	return b.String()
}
