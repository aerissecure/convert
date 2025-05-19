package xlsx

import (
	"fmt"

	"github.com/unidoc/unioffice/spreadsheet"
)

// Intermediate representation for XLSX.

// Pixel values are floats to allow fractional widths/heights if desired.

// CellStyle captures the limited set of Excel styles we currently support.
type CellStyle struct {
	FontFamily      string  // e.g. "Calibri"
	FontSizePt      float64 // original size in points
	FontColor       string  // "RRGGBB"
	BackgroundColor string  // "RRGGBB"
	BorderColor     string  // we use left-border color as representative
	HorizontalAlign string  // left|center|right|justify
	VerticalAlign   string  // top|middle|bottom
	WrapText        bool
	IndentPx        float64 // computed indent in pixels
}

func (s CellStyle) String() string {
	return fmt.Sprintf("FontFamily: %s, FontSizePt: %f, FontColor: %s, BackgroundColor: %s, BorderColor: %s, HorizontalAlign: %s, VerticalAlign: %s, WrapText: %t, IndentPx: %f", s.FontFamily, s.FontSizePt, s.FontColor, s.BackgroundColor, s.BorderColor, s.HorizontalAlign, s.VerticalAlign, s.WrapText, s.IndentPx)
}

// RenderCell is the IR for a single cell (or merged master).
type RenderCell struct {
	Cell    spreadsheet.Cell
	Ref     string    // e.g. "A1"
	Value   string    // already formatted value
	ColSpan int       // 1 if not merged
	RowSpan int       // 1 if not merged
	Style   CellStyle // resolved style
}

func (c RenderCell) String() string {
	return fmt.Sprintf("Ref: %s, Value: %s, ColSpan: %d, RowSpan: %d, Style: %s", c.Ref, c.Value, c.ColSpan, c.RowSpan, c.Style.String())
}

// RenderRow represents one logical row in a sheet.
type RenderRow struct {
	HeightPx float64 // resolved height in px
	Hidden   bool
	Cells    []*RenderCell // length == ColCount of parent sheet; may contain nil for blank cells
}

func (r RenderRow) String() string {
	return fmt.Sprintf("HeightPx: %f, Hidden: %t, Cells: %d", r.HeightPx, r.Hidden, len(r.Cells))
}

// RenderSheet is the intermediate representation of a worksheet.
type RenderSheet struct {
	Name      string
	ColWidths []float64   // per column pixel widths, len == ColCount
	ColHidden []bool      // true if column hidden
	Rows      []RenderRow // in order
}

func (s RenderSheet) String() string {
	return fmt.Sprintf("Name: %s, ColWidths: %v, ColHidden: %v, Rows: %d", s.Name, s.ColWidths, s.ColHidden, len(s.Rows))
}

// WorkbookModel is the top-level IR containing all sheets.
type WorkbookModel struct {
	Sheets []RenderSheet
}
