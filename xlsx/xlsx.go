package xlsx

import (
	"github.com/unidoc/unioffice/schema/soo/dml"
	"github.com/unidoc/unioffice/schema/soo/sml"
	"github.com/unidoc/unioffice/spreadsheet"
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
