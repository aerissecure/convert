package convert

import (
	"os"
	"testing"
)

func TestDocxToHTML(t *testing.T) {
	f, err := os.Open("test.docx")
	if err != nil {
		t.Fatalf("failed to open test.docx: %v", err)
	}
	defer f.Close()
	info, err := f.Stat()
	if err != nil {
		t.Fatalf("failed to stat test.docx: %v", err)
	}
	html, err := DocxToHTML(f, info.Size())
	if err != nil {
		t.Fatalf("DocxToHTML failed: %v", err)
	}
	if len(html) == 0 {
		t.Error("DocxToHTML returned empty HTML")
	}
	// write html to file
	err = os.WriteFile("test.docx.html", []byte(html), 0644)
	if err != nil {
		t.Fatalf("failed to write test.html: %v", err)
	}
}
