package docx

import (
	"fmt"
	"html"
	"io"
	"regexp"
	"strings"
)

// DebugHTML controls whether extra data attributes with raw style info are included in the rendered HTML output.
var DebugHTML bool

// DocxToHTML is a convenience wrapper that converts a DOCX reader to HTML
// using the intermediate representation defined in this package.
func DocxToHTML(r io.ReaderAt, size int64) (string, error) {
	ir, err := ParseDocumentModel(r, size)
	if err != nil {
		return "", err
	}
	return RenderDocumentHTML(ir), nil
}

// -----------------------------------------------------------------------------
// Helpers for sanitising CSS values – copied from xlsx/html.go for consistency.
// -----------------------------------------------------------------------------
var (
	fontFamilySafeRe = regexp.MustCompile(`[^a-zA-Z0-9 ,_-]+`)
	hexColorRe       = regexp.MustCompile(`^[0-9a-fA-F]{3}([0-9a-fA-F]{3})?$`)
)

// sanitizeFontFamily strips any characters that are not considered safe for a
// CSS font-family declaration.  This prevents breaking out of the CSS context
// and injecting arbitrary directives.
func sanitizeFontFamily(s string) string {
	return fontFamilySafeRe.ReplaceAllString(s, "")
}

// sanitizeColor ensures the value is a valid 3- or 6-digit hexadecimal string.
// Any invalid input results in an empty string, preventing potential CSS or
// markup injection.
func sanitizeColor(s string) string {
	if hexColorRe.MatchString(s) {
		return s
	}
	return ""
}

// -----------------------------------------------------------------------------
// Run-level helpers
// -----------------------------------------------------------------------------

func runStyleToCSS(s RunStyle) string {
	var b strings.Builder
	if s.FontFamily != "" {
		b.WriteString(fmt.Sprintf("font-family:'%s';", sanitizeFontFamily(s.FontFamily)))
	}
	if s.FontSizePt > 0 {
		b.WriteString(fmt.Sprintf("font-size:%.1fpt;", s.FontSizePt))
	}
	if s.FontColor != "" {
		if safe := sanitizeColor(s.FontColor); safe != "" {
			b.WriteString(fmt.Sprintf("color:#%s;", safe))
		}
	}
	if s.Bold {
		b.WriteString("font-weight:bold;")
	}
	if s.Italic {
		b.WriteString("font-style:italic;")
	}
	if s.Underline && s.Strike {
		b.WriteString("text-decoration:underline line-through;")
	} else if s.Underline {
		b.WriteString("text-decoration:underline;")
	} else if s.Strike {
		b.WriteString("text-decoration:line-through;")
	}
	switch s.VerticalAlign {
	case "superscript":
		b.WriteString("vertical-align:super;")
	case "subscript":
		b.WriteString("vertical-align:sub;")
	}
	return b.String()
}

// -----------------------------------------------------------------------------
// Paragraph-level helpers
// -----------------------------------------------------------------------------

func paragraphStyleToCSS(s ParagraphStyle) string {
	var b strings.Builder
	// Alignment
	switch s.Alignment {
	case "center":
		b.WriteString("text-align:center;")
	case "right":
		b.WriteString("text-align:right;")
	case "justify":
		b.WriteString("text-align:justify;")
	default:
		// left is default – nothing to emit
	}
	// Spacing (top/bottom margin in pt => convert to px ~ 1pt = 1.333px)
	if s.SpaceBeforePt > 0 {
		b.WriteString(fmt.Sprintf("margin-top:%.0fpt;", s.SpaceBeforePt))
	}
	if s.SpaceAfterPt > 0 {
		b.WriteString(fmt.Sprintf("margin-bottom:%.0fpt;", s.SpaceAfterPt))
	}
	// Indent (convert px)
	if s.IndentLeftPx > 0 {
		b.WriteString(fmt.Sprintf("padding-left:%.0fpx;", s.IndentLeftPx))
	}
	if s.IndentRightPx > 0 {
		b.WriteString(fmt.Sprintf("padding-right:%.0fpx;", s.IndentRightPx))
	}
	return b.String()
}

// -----------------------------------------------------------------------------
// Table cell helpers
// -----------------------------------------------------------------------------

func cellStyleToCSS(s TableCellStyle) string {
	var b strings.Builder
	if s.BackgroundColor != "" {
		if safe := sanitizeColor(s.BackgroundColor); safe != "" {
			b.WriteString(fmt.Sprintf("background-color:#%s;", safe))
		}
	}
	if s.VerticalAlign != "" {
		switch s.VerticalAlign {
		case "top":
			b.WriteString("vertical-align:top;")
		case "middle":
			b.WriteString("vertical-align:middle;")
		default:
			b.WriteString("vertical-align:bottom;")
		}
	}
	return b.String()
}

// -----------------------------------------------------------------------------
// Paragraph & Run rendering
// -----------------------------------------------------------------------------

func renderRunsHTML(runs []RenderRun) string {
	var b strings.Builder
	for _, run := range runs {
		text := html.EscapeString(run.Text)
		text = strings.ReplaceAll(text, "\n", "<br>")
		css := runStyleToCSS(run.Style)
		debugAttr := ""
		if DebugHTML {
			debugAttr = fmt.Sprintf(" data-run-style=\"%s\"", html.EscapeString(run.Style.String()))
		}
		if css != "" {
			b.WriteString(fmt.Sprintf("<span style=\"%s\"%s>%s</span>", css, debugAttr, text))
		} else {
			b.WriteString(fmt.Sprintf("<span%s>%s</span>", debugAttr, text))
		}
	}
	return b.String()
}

func renderParagraphHTML(p RenderParagraph) string {
	var tag string
	if p.Style.HeadingLevel > 0 && p.Style.HeadingLevel <= 6 {
		tag = fmt.Sprintf("h%d", p.Style.HeadingLevel)
	} else {
		tag = "p"
	}
	css := paragraphStyleToCSS(p.Style)
	debugAttr := ""
	if DebugHTML {
		debugAttr = fmt.Sprintf(" data-para-style=\"%s\"", html.EscapeString(p.Style.String()))
	}
	if css != "" {
		return fmt.Sprintf("<%s style=\"%s\"%s>%s</%s>\n", tag, css, debugAttr, renderRunsHTML(p.Runs), tag)
	}
	return fmt.Sprintf("<%s%s>%s</%s>\n", tag, debugAttr, renderRunsHTML(p.Runs), tag)
}

// -----------------------------------------------------------------------------
// Table rendering
// -----------------------------------------------------------------------------

func renderTableHTML(t RenderTable) string {
	var b strings.Builder
	b.WriteString("<table style=\"border-collapse:collapse;\">\n")
	for _, row := range t.Rows {
		b.WriteString("  <tr>")
		for _, cell := range row.Cells {
			// Guard against nil cells (shouldn't happen normally)
			var cellHTML string
			if len(cell.Paragraphs) == 0 {
				cellHTML = "&nbsp;"
			} else {
				var paraB strings.Builder
				for _, p := range cell.Paragraphs {
					paraB.WriteString(renderParagraphHTML(p))
				}
				cellHTML = paraB.String()
			}

			css := cellStyleToCSS(cell.Style)
			spanAttr := ""
			if cell.ColSpan > 1 {
				spanAttr += fmt.Sprintf(" colspan=\"%d\"", cell.ColSpan)
			}
			if cell.RowSpan > 1 {
				spanAttr += fmt.Sprintf(" rowspan=\"%d\"", cell.RowSpan)
			}
			if cell.WidthPx > 0 {
				css += fmt.Sprintf("width:%.0fpx;", cell.WidthPx)
			}
			debugAttr := ""
			if DebugHTML {
				debugAttr = fmt.Sprintf(" data-cell-style=\"%s\"", html.EscapeString(cell.Style.String()))
			}
			if css != "" {
				b.WriteString(fmt.Sprintf("    <td%s style=\"%s border:1px solid #333; padding:4px;\"%s>%s</td>", spanAttr, css, debugAttr, cellHTML))
			} else {
				b.WriteString(fmt.Sprintf("    <td%s style=\"border:1px solid #333; padding:4px;\"%s>%s</td>", spanAttr, debugAttr, cellHTML))
			}
		}
		b.WriteString("  </tr>\n")
	}
	b.WriteString("</table>\n")
	return b.String()
}

// -----------------------------------------------------------------------------
// Top-level rendering entry point
// -----------------------------------------------------------------------------

// RenderDocumentHTML converts the DocumentModel into an HTML string.
func RenderDocumentHTML(m DocumentModel) string {
	var b strings.Builder
	b.WriteString("<html><body>\n")

	if len(m.Blocks) > 0 {
		for _, blk := range m.Blocks {
			if blk.Paragraph != nil {
				b.WriteString(renderParagraphHTML(*blk.Paragraph))
			} else if blk.Table != nil {
				b.WriteString(renderTableHTML(*blk.Table))
			}
		}
	} else {
		// Fallback to legacy behaviour if Blocks not populated
		for _, p := range m.Paragraphs {
			b.WriteString(renderParagraphHTML(p))
		}
		for _, tbl := range m.Tables {
			b.WriteString(renderTableHTML(tbl))
		}
	}

	b.WriteString("</body></html>\n")
	return b.String()
}

func DOCXToHTML(r io.ReaderAt, size int64) (string, error) {
	ir, err := ParseDocumentModel(r, size)
	if err != nil {
		return "", err
	}
	return RenderDocumentHTML(ir), nil
}
