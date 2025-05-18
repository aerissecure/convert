package convert

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

// RenderCell is the IR for a single cell (or merged master).
type RenderCell struct {
	Value   string    // already formatted value
	ColSpan int       // 1 if not merged
	RowSpan int       // 1 if not merged
	Style   CellStyle // resolved style
}

// RenderRow represents one logical row in a sheet.
type RenderRow struct {
	HeightPx float64 // resolved height in px
	Hidden   bool
	Cells    []*RenderCell // length == ColCount of parent sheet; may contain nil for blank cells
}

// RenderSheet is the intermediate representation of a worksheet.
type RenderSheet struct {
	Name      string
	ColWidths []float64   // per column pixel widths, len == ColCount
	ColHidden []bool      // true if column hidden
	Rows      []RenderRow // in order
}

// WorkbookModel is the top-level IR containing all sheets.
type WorkbookModel struct {
	Sheets []RenderSheet
}
