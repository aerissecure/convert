package convert

import (
	"fmt"
	"html"
	"io"
	"strings"

	"github.com/unidoc/unioffice/schema/soo/dml"
	"github.com/unidoc/unioffice/schema/soo/sml"
	"github.com/unidoc/unioffice/spreadsheet"
	"github.com/unidoc/unioffice/spreadsheet/reference"
)

// TODO: Set a default font family and size, only add to style if differs.

// We need to display with a table instead of divs

// Note: Google Drive preview renders to canvas and also renders to <table>, but
// it hides table and maps between it and canvas for search. Gives them more
// accurate rendering.

// Helper to extract the underlying font XML struct from a style ID
func GetFontProps(ss spreadsheet.StyleSheet, styleID uint32) *sml.CT_Font {
	if int(styleID) < 0 || int(styleID) >= len(ss.X().CellXfs.Xf) {
		return nil
	}
	xf := ss.X().CellXfs.Xf[styleID]
	if xf.FontIdAttr == nil {
		return nil
	}
	fontIdx := int(*xf.FontIdAttr)
	if fontIdx < 0 || fontIdx >= len(ss.X().Fonts.Font) {
		return nil
	}
	return ss.X().Fonts.Font[fontIdx]
}

// Helper to extract the underlying fill XML struct from a style ID
func GetFillProps(ss spreadsheet.StyleSheet, styleID uint32) *sml.CT_Fill {
	if int(styleID) < 0 || int(styleID) >= len(ss.X().CellXfs.Xf) {
		return nil
	}
	xf := ss.X().CellXfs.Xf[styleID]
	if xf.FillIdAttr == nil {
		return nil
	}
	fillIdx := int(*xf.FillIdAttr)
	if fillIdx < 0 || fillIdx >= len(ss.X().Fills.Fill) {
		return nil
	}
	return ss.X().Fills.Fill[fillIdx]
}

// Helper to extract the underlying border XML struct from a style ID
func GetBorderProps(ss spreadsheet.StyleSheet, styleID uint32) *sml.CT_Border {
	if int(styleID) < 0 || int(styleID) >= len(ss.X().CellXfs.Xf) {
		return nil
	}
	xf := ss.X().CellXfs.Xf[styleID]
	if xf.BorderIdAttr == nil {
		return nil
	}
	borderIdx := int(*xf.BorderIdAttr)
	if borderIdx < 0 || borderIdx >= len(ss.X().Borders.Border) {
		return nil
	}
	return ss.X().Borders.Border[borderIdx]
}

// ThemeColorToRGB resolves a theme color index (0-based) to an RGB hex string (e.g., "FFFFFF").
// It does not apply tint. Returns false if the index is invalid or the color cannot be resolved.
func ThemeColorToRGB(wb *spreadsheet.Workbook, themeIdx int) (string, bool) {
	themes := wb.Themes() // Your own method returning []*dml.Theme
	if len(themes) == 0 || themes[0] == nil {
		return "", false
	}
	clrScheme := themes[0].ThemeElements.ClrScheme

	// Map themeIdx to the corresponding color field
	var clr *dml.CT_Color
	switch themeIdx {
	case 0:
		clr = clrScheme.Dk1
	case 1:
		clr = clrScheme.Lt1
	case 2:
		clr = clrScheme.Dk2
	case 3:
		clr = clrScheme.Lt2
	case 4:
		clr = clrScheme.Accent1
	case 5:
		clr = clrScheme.Accent2
	case 6:
		clr = clrScheme.Accent3
	case 7:
		clr = clrScheme.Accent4
	case 8:
		clr = clrScheme.Accent5
	case 9:
		clr = clrScheme.Accent6
	case 10:
		clr = clrScheme.Hlink
	case 11:
		clr = clrScheme.FolHlink
	default:
		return "", false
	}

	if clr == nil {
		return "", false
	}

	if clr.SrgbClr != nil && clr.SrgbClr.ValAttr != "" {
		return clr.SrgbClr.ValAttr, true
	} else if clr.SysClr != nil && clr.SysClr.LastClrAttr != nil {
		return *clr.SysClr.LastClrAttr, true
	}
	return "", false
}

func XlsxToHTML(r io.ReaderAt, size int64) (string, error) {

	reference.ColumnToIndex("A")

	ss, err := spreadsheet.Read(r, size)
	if err != nil {
		return "", err
	}

	var builder strings.Builder

	// Add CSS for table-like rendering
	builder.WriteString(`<style>
`)
	builder.WriteString(`.table { border-collapse: collapse; table-layout: fixed; margin-bottom: 2em; }
`)
	builder.WriteString(`.table td { border: 1px solid #333; padding: 4px 8px; vertical-align: bottom; }
`)
	builder.WriteString(`.sheet { margin-bottom: 2em; }
`)
	builder.WriteString(`</style>
`)

	sheets := ss.Sheets()
	if len(sheets) == 0 {
		return "<table class=\"table\"></table>", nil
	}

	for _, sheet := range sheets {

		// TODO: use the custom column width if set, perhaps this shoudl be figured out for each cell?
		// if sheet.Column(1).X().CustomWidthAttr; use sheet.Column(1).X().WidthAttr

		// Preprocess merges: build maps keyed by row/col of master cell.
		mergeMaster := make(map[[2]int]struct{ rowSpan, colSpan int })
		skipCells := make(map[[2]int]bool)
		if sheet.X().MergeCells != nil {
			for _, mc := range sheet.X().MergeCells.MergeCell {
				from, to, err := reference.ParseRangeReference(mc.RefAttr)
				if err != nil {
					continue // if parsing fails, ignore merge
				}
				// Convert to zero-based indices to match our iteration counters
				fromRow := int(from.RowIdx - 1)
				fromCol := int(from.ColumnIdx)
				toRow := int(to.RowIdx - 1)
				toCol := int(to.ColumnIdx)

				rowSpan := toRow - fromRow + 1
				colSpan := toCol - fromCol + 1

				key := [2]int{fromRow, fromCol}
				mergeMaster[key] = struct{ rowSpan, colSpan int }{rowSpan, colSpan}

				// Mark the rest of the cells as skipped
				for r := fromRow; r <= toRow; r++ {
					for c := fromCol; c <= toCol; c++ {
						if r == fromRow && c == fromCol {
							continue
						}
						skipCells[[2]int{r, c}] = true
					}
				}
			}
		}
		sheetName := html.EscapeString(sheet.Name())
		builder.WriteString(fmt.Sprintf(`<div class="sheet" data-name="%s">
`, sheetName))
		rows := sheet.Rows()

		// Determine maximum number of columns present so we can emit a <colgroup>
		maxCols := 0
		for _, r := range rows {
			if len(r.Cells()) > maxCols {
				maxCols = len(r.Cells())
			}
		}

		builder.WriteString(`<div style="width:100%;overflow-x:auto;">
`)

		// Compute total pixel width of the columns to set as min-width on table
		var totalPx float64
		for c := 0; c < maxCols; c++ {
			colObj := sheet.Column(uint32(c + 1))
			if colObj.X().CustomWidthAttr != nil && *colObj.X().CustomWidthAttr {
				totalPx += *colObj.X().WidthAttr * 8.3
			} else {
				// Approximate default column width (using Excel's default ~8.43 characters)
				totalPx += 8.43 * 8.3
			}
		}

		builder.WriteString(fmt.Sprintf(`<table class="table" style="min-width:%.0fpx;">
`, totalPx))
		// Emit column definitions so the browser uses column-level widths instead of
		// repeating the width style on every individual cell.
		builder.WriteString("  <colgroup>\n")
		for c := 0; c < maxCols; c++ {
			colObj := sheet.Column(uint32(c + 1)) // 1-based in unioffice expects uint32

			colStyle := ""
			if colObj.X().CustomWidthAttr != nil && *colObj.X().CustomWidthAttr {
				colStyle = fmt.Sprintf(" style=\"width:%.2fpx;\"", *colObj.X().WidthAttr*8.3)
			}
			// Hidden columns – use display:none on the <col>
			if colObj.X().HiddenAttr != nil && *colObj.X().HiddenAttr {
				colStyle = " style=\"display:none;\""
			}
			builder.WriteString(fmt.Sprintf("    <col%s>\n", colStyle))
		}
		builder.WriteString("  </colgroup>\n")

		for _, row := range rows {
			rowNum0 := int(row.RowNumber()) - 1
			// Build row style for custom height and hidden
			rowStyle := ""
			if row.X().CustomHeightAttr != nil && *row.X().CustomHeightAttr {
				// Set custom height on row using row.X().HtAttr (height in points)
				rowStyle += fmt.Sprintf("height:%.2fpt;", *row.X().HtAttr)
			}

			if row.IsHidden() {
				// Set hidden on row using row.X().HiddenAttr
				rowStyle += "display:none;"
			}

			builder.WriteString(fmt.Sprintf("  <tr style=\"%s\">\n", rowStyle))
			cells := row.Cells()
			cellMap := make(map[int]spreadsheet.Cell)
			for _, cell := range cells {
				if colName, err := cell.Column(); err == nil {
					idx := int(reference.ColumnToIndex(colName))
					cellMap[idx] = cell
				}
			}

			for colIdx1 := 0; colIdx1 < maxCols; colIdx1++ {
				// Skip cells covered by a merge range (non-master cells)
				if skipCells[[2]int{rowNum0, colIdx1}] {
					continue
				}

				var (
					colStyle string
					attr     string
					cellVal  string
				)

				// If this (row,col) is a merge master, collect rowspan/colspan
				if info, ok := mergeMaster[[2]int{rowNum0, colIdx1}]; ok {
					if info.colSpan > 1 {
						attr += fmt.Sprintf(" colspan=\"%d\"", info.colSpan)
					}
					if info.rowSpan > 1 {
						attr += fmt.Sprintf(" rowspan=\"%d\"", info.rowSpan)
					}
				}

				// If we have an actual cell object, extract styles and value
				if cell, ok := cellMap[colIdx1]; ok {
					// --- Add style extraction from cell style ---
					if cell.X().SAttr != nil {
						styleID := *cell.X().SAttr
						font := GetFontProps(ss.StyleSheet, styleID)
						fill := GetFillProps(ss.StyleSheet, styleID)
						border := GetBorderProps(ss.StyleSheet, styleID)
						xf := ss.StyleSheet.X().CellXfs.Xf[styleID]

						if font != nil && len(font.Name) > 0 {
							colStyle += fmt.Sprintf("font-family:'%s';", font.Name[0].ValAttr)
						}
						if font != nil && len(font.Sz) > 0 {
							colStyle += fmt.Sprintf("font-size:%.1fpt;", font.Sz[0].ValAttr)
						}
						if font != nil && len(font.Color) > 0 && font.Color[0].RgbAttr != nil && *font.Color[0].RgbAttr != "" {
							colStyle += fmt.Sprintf("color:#%s;", *font.Color[0].RgbAttr)
						}
						if fill != nil && fill.PatternFill != nil && fill.PatternFill.FgColor != nil {
							fg := fill.PatternFill.FgColor
							if fg.RgbAttr != nil && *fg.RgbAttr != "" {
								colStyle += fmt.Sprintf("background-color:#%s;", *fg.RgbAttr)
							} else if fg.ThemeAttr != nil {
								if hex, ok := ThemeColorToRGB(ss, int(*fg.ThemeAttr)); ok {
									colStyle += fmt.Sprintf("background-color:#%s;", hex)
								}
							}
						}
						if border != nil && border.Left != nil && border.Left.Color != nil && border.Left.Color.RgbAttr != nil && *border.Left.Color.RgbAttr != "" {
							colStyle += fmt.Sprintf("border-left: 1px solid #%s;", *border.Left.Color.RgbAttr)
						}

						// Alignment
						if xf.Alignment != nil {
							switch xf.Alignment.HorizontalAttr.String() {
							case "left", "general":
								colStyle += "text-align:left;"
							case "center", "centerContinuous", "distributed":
								colStyle += "text-align:center;"
							case "right":
								colStyle += "text-align:right;"
							case "justify":
								colStyle += "text-align:justify;"
							}
							switch xf.Alignment.VerticalAttr.String() {
							case "top":
								colStyle += "vertical-align:top;"
							case "center":
								colStyle += "vertical-align:middle;"
							case "bottom":
								// default already bottom
							default:
							}
							if xf.Alignment.WrapTextAttr != nil {
								if *xf.Alignment.WrapTextAttr {
									colStyle += "white-space:normal;"
								} else {
									colStyle += "white-space:nowrap;"
								}
							}
							if xf.Alignment.IndentAttr != nil && *xf.Alignment.IndentAttr > 0 {
								indentPx := float64(*xf.Alignment.IndentAttr) * 8.0
								// Apply padding-left by default unless right-aligned
								if strings.Contains(colStyle, "text-align:right") {
									colStyle += fmt.Sprintf("padding-right:%.0fpx;", indentPx)
								} else {
									colStyle += fmt.Sprintf("padding-left:%.0fpx;", indentPx)
								}
							}
						}
					}
					// --- End style extraction ---
					cellVal = html.EscapeString(cell.GetFormattedValue())
				} else {
					cellVal = ""
				}

				builder.WriteString(fmt.Sprintf("    <td data-row=\"%d\" data-col=\"%d\"%s style=\"%s\">%s</td>\n",
					rowNum0, colIdx1, attr, colStyle, cellVal))
			}
			builder.WriteString("  </tr>\n")
		}
		builder.WriteString("</table>\n") // close table
		builder.WriteString("</div>\n")   // close sheet
		builder.WriteString("</div>\n")   // close scroll wrapper
	}
	return builder.String(), nil
}
