package xlsx

import (
	"encoding/xml"
	"fmt"
	"os"
	"strings"
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
	ir, err := ParseWorkbookModel(f, info.Size())
	if err != nil {
		t.Fatalf("failed to parse workbook model: %v", err)
	}

	for _, sheet := range ir.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if cell != nil && cell.Ref == "F98" {
					// Doesn't have any things directly attached,
					// 	<CT_Cell r="F98" s="7" t="s">
					// 	<ma:v>404</ma:v>
					// </CT_Cell>
					fmt.Println(SprintXML(cell.Cell.X()))
					fmt.Println(cell.String())
				}
			}
		}
	}

	html := RenderWorkbookHTML(ir)

	// write html to file
	err = os.WriteFile("test.xlsx.html", []byte(html), 0644)
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
