package convert

import (
	"fmt"
	"io"
	"strings"

	"github.com/unidoc/unioffice/document"
)

// DocxToHTML converts a DOCX file (from r, with given size) to HTML.
func DocxToHTML(r io.ReaderAt, size int64) (string, error) {
	doc, err := document.Read(r, size)
	if err != nil {
		return "", err
	}

	var sb strings.Builder
	sb.WriteString("<html><body>\n")
	for _, para := range doc.Paragraphs() {
		style := para.Style()
		if strings.HasPrefix(style, "Heading") {
			sb.WriteString(fmt.Sprintf("<h1>")) // For simplicity, treat all as h1
		} else {
			sb.WriteString("<p>")
		}
		for _, run := range para.Runs() {
			text := run.Text()
			if text == "" {
				continue
			}
			if run.Properties().IsBold() {
				sb.WriteString("<b>")
			}
			if run.Properties().IsItalic() {
				sb.WriteString("<i>")
			}
			sb.WriteString(text)
			if run.Properties().IsItalic() {
				sb.WriteString("</i>")
			}
			if run.Properties().IsBold() {
				sb.WriteString("</b>")
			}
		}
		if strings.HasPrefix(style, "Heading") {
			sb.WriteString("</h1>\n")
		} else {
			sb.WriteString("</p>\n")
		}
	}
	sb.WriteString("</body></html>\n")
	return sb.String(), nil
}
