package xlsx

import (
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/unidoc/unioffice/schema/soo/sml"
	"github.com/unidoc/unioffice/spreadsheet"
	"github.com/unidoc/unioffice/spreadsheet/reference"
)

// ParseWorkbookModel reads an XLSX from r/size and returns the intermediate representation.
func ParseWorkbookModel(r io.ReaderAt, size int64) (WorkbookModel, error) {
	wb, err := spreadsheet.Read(r, size)
	if err != nil {
		return WorkbookModel{}, err
	}

	var model WorkbookModel

	for _, sheet := range wb.Sheets() {
		// ---- find max column ----
		maxCols := 0
		for _, row := range sheet.Rows() {
			if len(row.Cells()) > maxCols {
				maxCols = len(row.Cells())
			}
		}

		// Column metadata
		colWidths := make([]float64, maxCols)
		colHidden := make([]bool, maxCols)
		for c := 0; c < maxCols; c++ {
			colObj := sheet.Column(uint32(c + 1))
			if colObj.X().CustomWidthAttr != nil && *colObj.X().CustomWidthAttr {
				colWidths[c] = *colObj.X().WidthAttr * 8.3
			} else {
				colWidths[c] = 8.43 * 8.3 // default approximation
			}
			if colObj.X().HiddenAttr != nil {
				colHidden[c] = *colObj.X().HiddenAttr
			}
		}

		rs := RenderSheet{
			Name:      sheet.Name(),
			ColWidths: colWidths,
			ColHidden: colHidden,
		}

		// --- process merges ---
		mergeMaster := make(map[[2]int]struct{ rowSpan, colSpan int })
		skipCells := make(map[[2]int]bool)
		if sheet.X().MergeCells != nil {
			for _, mc := range sheet.X().MergeCells.MergeCell {
				from, to, err := reference.ParseRangeReference(mc.RefAttr)
				if err != nil {
					continue
				}
				fromRow := int(from.RowIdx - 1)
				fromCol := int(from.ColumnIdx)
				toRow := int(to.RowIdx - 1)
				toCol := int(to.ColumnIdx)
				mergeMaster[[2]int{fromRow, fromCol}] = struct{ rowSpan, colSpan int }{toRow - fromRow + 1, toCol - fromCol + 1}

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

		// --- build rows ---
		for _, row := range sheet.Rows() {
			rowIdx := int(row.RowNumber()) - 1
			if rowIdx >= len(rs.Rows) {
				// grow slice to accommodate sparse rows
				newRows := make([]RenderRow, rowIdx-len(rs.Rows)+1)
				rs.Rows = append(rs.Rows, newRows...)
			}

			rr := &rs.Rows[rowIdx]
			rr.Cells = make([]*RenderCell, maxCols)
			rr.Hidden = row.IsHidden()
			if row.X().CustomHeightAttr != nil && *row.X().CustomHeightAttr {
				rr.HeightPx = *row.X().HtAttr * 1.333 // pt -> px
			} else {
				rr.HeightPx = 15.0 * 1.333 // Excel default 15pt
			}

			for _, cell := range row.Cells() {
				colName, err := cell.Column()
				if err != nil {
					continue
				}
				colIdx := int(reference.ColumnToIndex(colName))
				if skipCells[[2]int{rowIdx, colIdx}] {
					continue
				}
				// style
				var st CellStyle
				if cell.X().SAttr != nil {
					styleID := *cell.X().SAttr
					font := GetFontProps(wb.StyleSheet, styleID)
					fill := GetFillProps(wb.StyleSheet, styleID)
					border := GetBorderProps(wb.StyleSheet, styleID)
					xf := wb.StyleSheet.X().CellXfs.Xf[styleID]
					if font != nil && len(font.Name) > 0 {
						st.FontFamily = font.Name[0].ValAttr
					}
					if font != nil && len(font.Sz) > 0 {
						st.FontSizePt = font.Sz[0].ValAttr
					}
					if font != nil && len(font.Color) > 0 && font.Color[0].RgbAttr != nil {
						st.FontColor = normalizeColor(*font.Color[0].RgbAttr)
					}
					if fill != nil && fill.PatternFill != nil && fill.PatternFill.FgColor != nil {
						fg := fill.PatternFill.FgColor
						if fg.RgbAttr != nil {
							st.BackgroundColor = normalizeColor(*fg.RgbAttr)
						} else if fg.ThemeAttr != nil {
							if hex, ok := ThemeColorToRGB(wb, int(*fg.ThemeAttr)); ok {
								st.BackgroundColor = hex
							}
						}
					}
					if border != nil && border.Left != nil && border.Left.Color != nil && border.Left.Color.RgbAttr != nil {
						st.BorderColor = normalizeColor(*border.Left.Color.RgbAttr)
					}
					if xf.Alignment != nil {
						st.HorizontalAlign = xf.Alignment.HorizontalAttr.String()
						switch xf.Alignment.VerticalAttr.String() {
						case "top":
							st.VerticalAlign = "top"
						case "center":
							st.VerticalAlign = "middle"
						default:
							st.VerticalAlign = "bottom"
						}
						if xf.Alignment.WrapTextAttr != nil {
							st.WrapText = *xf.Alignment.WrapTextAttr
						}
						if xf.Alignment.IndentAttr != nil {
							st.IndentPx = float64(*xf.Alignment.IndentAttr) * 8.0
						}
					}
				}

				rc := &RenderCell{
					Cell:  cell,
					Ref:   fmt.Sprintf("%s%d", colName, rowIdx+1),
					Value: cell.GetFormattedValue(),
					// Runs will be populated below if rich text present
					ColSpan: 1,
					RowSpan: 1,
					Style:   st,
				}

				// Check for rich-text runs
				rt := cellRichTextString(cell, wb)
				if rt != nil && len(rt.R) > 0 {
					fmt.Println(rc.Ref)
					// Prefer runs if present, else fallback on plain text T
					if len(rt.R) > 0 {
						for _, r := range rt.R {
							text := r.T
							run := RenderRun{Text: text}
							if rp := r.RPr; rp != nil {
								if rp.RFont != nil {
									run.FontFamily = rp.RFont.ValAttr
								}
								if rp.Sz != nil {
									run.FontSizePt = rp.Sz.ValAttr
								}
								if rp.Color != nil {
									if rp.Color.RgbAttr != nil {
										run.FontColor = normalizeColor(*rp.Color.RgbAttr)
									} else if rp.Color.ThemeAttr != nil {
										themeIdx := int(*rp.Color.ThemeAttr)
										// Skip Light1 (theme 1) which typically represents default automatic font color (black) in Excel.
										if themeIdx != 1 {
											if hex, ok := ThemeColorToRGB(wb, themeIdx); ok {
												run.FontColor = hex
											}
										}
									}
								}
								run.Bold = rp.B != nil
								run.Italic = rp.I != nil
								run.Strike = rp.Strike != nil
								run.Underline = rp.U != nil
								if rp.VertAlign != nil {
									run.VerticalAlign = rp.VertAlign.ValAttr.String()
								}
							}
							rc.Runs = append(rc.Runs, run)
						}
					} else if rt.T != nil {
						// Single run of plain text; keep consistency
						rc.Runs = []RenderRun{{Text: *rt.T}}
					}
				}
				// check if this cell is a merge master
				if info, ok := mergeMaster[[2]int{rowIdx, colIdx}]; ok {
					rc.RowSpan = info.rowSpan
					rc.ColSpan = info.colSpan
				}

				rr.Cells[colIdx] = rc
			}
		}

		model.Sheets = append(model.Sheets, rs)
	}

	return model, nil
}

func cellRichTextString(cell spreadsheet.Cell, w *spreadsheet.Workbook) *sml.CT_Rst {
	x := cell.X()
	if x.Is != nil {
		return x.Is
	}
	if x.TAttr == sml.ST_CellTypeS {
		if x.V == nil {
			return nil
		}
		id, err := strconv.Atoi(*x.V)
		if err != nil {
			return nil
		}

		ssx := w.SharedStrings.X()
		if id < 0 || id >= len(ssx.Si) {
			return nil
		}

		return ssx.Si[id]
	}
	return nil
}

// normalizeColor converts an 8-digit ARGB hex (as used in XLSX) to a 6-digit RGB string.
// If the string is already 6 digits (or any other length), it is returned unchanged.
func normalizeColor(hex string) string {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) == 8 {
		return hex[2:]
	}
	return hex
}
