package xlsx

import (
	"fmt"
	"html"
	"strings"
)

// RenderWorkbookHTML converts the IR into an HTML string.
func RenderWorkbookHTML(m WorkbookModel) string {
	var builder strings.Builder

	// 1. Collect unique cell styles and count property values
	type propCount map[string]int
	fontFamilyCount := make(propCount)
	fontSizeCount := make(map[float64]int)
	borderColorCount := make(propCount)
	hAlignCount := make(propCount)
	vAlignCount := make(propCount)
	fontColorCount := make(propCount)
	bgColorCount := make(propCount)
	wrapTextCount := make(map[bool]int)
	indentPxCount := make(map[float64]int)

	styleMap := make(map[CellStyle]string) // CellStyle -> class name
	styleList := make([]CellStyle, 0)      // To preserve order
	classIdx := 1
	styledCells := 0

	for _, sheet := range m.Sheets {
		for _, row := range sheet.Rows {
			for _, cell := range row.Cells {
				if cell == nil {
					continue
				}
				styledCells++
				st := cell.Style
				if st.FontFamily != "" {
					fontFamilyCount[st.FontFamily]++
				}
				if st.FontSizePt > 0 {
					fontSizeCount[st.FontSizePt]++
				}
				if st.BorderColor != "" {
					borderColorCount[st.BorderColor]++
				}
				if st.HorizontalAlign != "" {
					hAlignCount[st.HorizontalAlign]++
				}
				if st.VerticalAlign != "" {
					vAlignCount[st.VerticalAlign]++
				}
				if st.FontColor != "" {
					fontColorCount[st.FontColor]++
				}
				if st.BackgroundColor != "" {
					bgColorCount[st.BackgroundColor]++
				}
				wrapTextCount[st.WrapText]++
				if st.IndentPx > 0 {
					indentPxCount[st.IndentPx]++
				}
				if _, exists := styleMap[st]; !exists {
					className := fmt.Sprintf("cellstyle%d", classIdx)
					styleMap[st] = className
					styleList = append(styleList, st)
					classIdx++
				}
			}
		}
	}

	// Helper to find most common value with count
	mostCommonStr := func(m propCount) (string, int) {
		max := 0
		val := ""
		for k, v := range m {
			if v > max {
				max = v
				val = k
			}
		}
		return val, max
	}
	mostCommonFloat := func(m map[float64]int) (float64, int) {
		max := 0
		var val float64
		for k, v := range m {
			if v > max {
				max = v
				val = k
			}
		}
		return val, max
	}
	mostCommonBool := func(m map[bool]int) (bool, int) {
		max := 0
		val := false
		for k, v := range m {
			if v > max {
				max = v
				val = k
			}
		}
		return val, max
	}

	// 2. Compute defaults
	defaultFontFamily, ffCount := mostCommonStr(fontFamilyCount)
	if ffCount <= styledCells/2 {
		defaultFontFamily = ""
	}
	defaultFontSize, fsCount := mostCommonFloat(fontSizeCount)
	if fsCount <= styledCells/2 {
		defaultFontSize = 0
	}
	defaultBorderColor, bcCount := mostCommonStr(borderColorCount)
	if bcCount <= styledCells/2 {
		defaultBorderColor = ""
	}
	defaultHAlign, haCount := mostCommonStr(hAlignCount)
	if haCount <= styledCells/2 {
		defaultHAlign = ""
	}
	defaultVAlign, vaCount := mostCommonStr(vAlignCount)
	if vaCount <= styledCells/2 {
		defaultVAlign = ""
	}
	defaultFontColor, fcCount := mostCommonStr(fontColorCount)
	if fcCount <= styledCells/2 {
		defaultFontColor = ""
	}
	defaultBgColor, bgCount := mostCommonStr(bgColorCount)
	if bgCount <= styledCells/2 {
		defaultBgColor = ""
	}
	// For wrap text and indent, we typically don't want defaults
	defaultWrapText, _ := mostCommonBool(wrapTextCount)
	defaultIndentPx := 0.0 // no default indent

	// 3. Basic CSS
	builder.WriteString(`<style>
`)
	builder.WriteString(`.table { border-collapse: collapse; table-layout: fixed; margin-bottom: 2em; }
`)
	builder.WriteString(`.table td { padding: 4px 8px;`)
	if defaultFontFamily != "" {
		builder.WriteString(fmt.Sprintf(" font-family:'%s';", defaultFontFamily))
	}
	if defaultFontSize > 0 {
		builder.WriteString(fmt.Sprintf(" font-size:%.1fpt;", defaultFontSize))
	}
	if defaultFontColor != "" {
		builder.WriteString(fmt.Sprintf(" color:#%s;", defaultFontColor))
	}
	if defaultBgColor != "" {
		builder.WriteString(fmt.Sprintf(" background-color:#%s;", defaultBgColor))
	}
	if defaultBorderColor != "" {
		builder.WriteString(fmt.Sprintf(" border:1px solid #%s;", defaultBorderColor))
	} else {
		builder.WriteString(" border:1px solid #333;")
	}
	// Handle default wrap behaviour
	if !defaultWrapText {
		// No wrapping: prevent text spillover
		builder.WriteString(" white-space:nowrap; overflow:hidden;")
	}
	if defaultHAlign != "" {
		switch defaultHAlign {
		case "center", "centerContinuous", "distributed":
			builder.WriteString(" text-align:center;")
		case "right":
			builder.WriteString(" text-align:right;")
		case "justify":
			builder.WriteString(" text-align:justify;")
		default:
			builder.WriteString(" text-align:left;")
		}
	}
	if defaultVAlign != "" {
		if defaultVAlign == "top" {
			builder.WriteString(" vertical-align:top;")
		} else if defaultVAlign == "middle" {
			builder.WriteString(" vertical-align:middle;")
		} else {
			builder.WriteString(" vertical-align:bottom;")
		}
	}
	// WrapText and IndentPx are less common as defaults, so skip for now
	builder.WriteString(` }
`)
	builder.WriteString(`.sheet { margin-bottom: 2em; }
`)

	// 4. Render cell style classes (only properties that differ from default)
	for i, style := range styleList {
		className := fmt.Sprintf("cellstyle%d", i+1)
		css := styleToCSSDiff(style, defaultFontFamily, defaultFontSize, defaultBorderColor, defaultHAlign, defaultVAlign, defaultFontColor, defaultBgColor, defaultWrapText, defaultIndentPx)
		if css != "" {
			builder.WriteString(fmt.Sprintf(".%s { %s }\n", className, css))
		}
	}
	builder.WriteString(`</style>
`)

	for _, sheet := range m.Sheets {
		totalPx := 0.0
		for _, w := range sheet.ColWidths {
			totalPx += w
		}
		builder.WriteString(fmt.Sprintf(`<div class="sheet" data-name="%s">
`, html.EscapeString(sheet.Name)))
		builder.WriteString(`<div style="width:100%;overflow-x:auto;">
`)
		builder.WriteString(fmt.Sprintf(`<table class="table" style="width:%.0fpx;">
`, totalPx))
		builder.WriteString("  <colgroup>\n")
		for i, w := range sheet.ColWidths {
			style := fmt.Sprintf(" style=\"width:%.0fpx;\"", w)
			if sheet.ColHidden[i] {
				style = " style=\"display:none;\""
			}
			builder.WriteString(fmt.Sprintf("    <col%s>\n", style))
		}
		builder.WriteString("  </colgroup>\n")

		for _, row := range sheet.Rows {
			rowStyle := fmt.Sprintf("height:%.0fpx;", row.HeightPx)
			if row.Hidden {
				rowStyle += "display:none;"
			}
			builder.WriteString(fmt.Sprintf("  <tr style=\"%s\">\n", rowStyle))
			for colIdx := 0; colIdx < len(row.Cells); colIdx++ {
				cell := row.Cells[colIdx]
				// Blank cell
				if cell == nil {
					builder.WriteString("    <td></td>\n")
					continue
				}

				// Prepare attributes
				className := styleMap[cell.Style]
				spanAttr := ""
				if cell.ColSpan > 1 {
					spanAttr += fmt.Sprintf(" colspan=\"%d\"", cell.ColSpan)
				}
				if cell.RowSpan > 1 {
					spanAttr += fmt.Sprintf(" rowspan=\"%d\"", cell.RowSpan)
				}

				escaped := html.EscapeString(cell.Value)
				// Excel stores explicit line breaks as \n; preserve them in HTML
				escaped = strings.ReplaceAll(escaped, "\n", "<br>")
				builder.WriteString(fmt.Sprintf("    <td data-cell=\"%s\"%s class=\"%s\">%s</td>\n",
					cell.Ref, spanAttr, className, escaped))

				// Skip over columns that are covered by this cell's colspan so we don't emit extra cells
				if cell.ColSpan > 1 {
					colIdx += cell.ColSpan - 1
				}
			}
			builder.WriteString("  </tr>\n")
		}
		builder.WriteString("</table>\n</div>\n</div>\n")
	}
	return builder.String()
}

// styleToCSSDiff returns only the CSS properties from s that differ from the provided defaults.
func styleToCSSDiff(s CellStyle, defFontFamily string, defFontSize float64, defBorderColor, defHAlign, defVAlign, defFontColor, defBgColor string, defWrapText bool, defIndentPx float64) string {
	var b strings.Builder
	if s.FontFamily != "" && s.FontFamily != defFontFamily {
		b.WriteString(fmt.Sprintf("font-family:'%s';", s.FontFamily))
	}
	if s.FontSizePt > 0 && s.FontSizePt != defFontSize {
		b.WriteString(fmt.Sprintf("font-size:%.1fpt;", s.FontSizePt))
	}
	if s.FontColor != "" && s.FontColor != defFontColor {
		b.WriteString(fmt.Sprintf("color:#%s;", s.FontColor))
	}
	if s.BackgroundColor != "" && s.BackgroundColor != defBgColor {
		b.WriteString(fmt.Sprintf("background-color:#%s;", s.BackgroundColor))
	}
	if s.BorderColor != "" && s.BorderColor != defBorderColor {
		b.WriteString(fmt.Sprintf("border:1px solid #%s;", s.BorderColor))
	}
	if s.HorizontalAlign != "" && s.HorizontalAlign != defHAlign {
		switch s.HorizontalAlign {
		case "center", "centerContinuous", "distributed":
			b.WriteString("text-align:center;")
		case "right":
			b.WriteString("text-align:right;")
		case "justify":
			b.WriteString("text-align:justify;")
		default:
			b.WriteString("text-align:left;")
		}
	}
	if s.VerticalAlign != "" && s.VerticalAlign != defVAlign {
		if s.VerticalAlign == "top" {
			b.WriteString("vertical-align:top;")
		} else if s.VerticalAlign == "middle" {
			b.WriteString("vertical-align:middle;")
		} else {
			b.WriteString("vertical-align:bottom;")
		}
	}
	// Only output wrap/indent if different from default
	if s.WrapText != defWrapText {
		if s.WrapText {
			b.WriteString("white-space:normal;")
		} else {
			b.WriteString("white-space:nowrap;overflow:hidden;")
		}
	}
	if s.IndentPx > 0 && s.IndentPx != defIndentPx {
		if strings.Contains(b.String(), "text-align:right") {
			b.WriteString(fmt.Sprintf("padding-right:%.0fpx;", s.IndentPx))
		} else {
			b.WriteString(fmt.Sprintf("padding-left:%.0fpx;", s.IndentPx))
		}
	}
	return b.String()
}
