package convert

import (
	"os"
	"testing"
)

func TestXlsxToHTML(t *testing.T) {
	f, err := os.Open("test.xlsx")
	if err != nil {
		t.Fatalf("failed to open test.xlsx: %v", err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatalf("failed to stat test.xlsx: %v", err)
	}
	html, err := XlsxToHTML(f, info.Size())
	if err != nil {
		t.Fatalf("XlsxToHTML failed: %v", err)
	}
	if len(html) == 0 {
		t.Error("XlsxToHTML returned empty HTML")
	}
	// write html to file
	err = os.WriteFile("test.xlsx.html", []byte(html), 0644)
	if err != nil {
		t.Fatalf("failed to write test.html: %v", err)
	}
}
