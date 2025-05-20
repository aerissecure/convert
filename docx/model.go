package convert

import (
	"fmt"
	"time"

	"github.com/unidoc/unioffice/document"
)

// Intermediate representation (IR) for DOCX documents.
//
// The purpose of these types is to provide a Go-native structure that captures
// just the information our converter cares about – no more and no less.  They
// intentionally mirror the level of detail found in the XLSX IR so development
// against the two formats feels familiar.
//
// All colours are expressed as 6-character RGB hex strings without the leading
// "#" (e.g. "FF0000" for red).

// -----------------------------------------------------------------------------
// Document-level information
// -----------------------------------------------------------------------------

// DocProperties captures the common document properties that users typically
// care about (title, author, …).  The field list can be expanded later if we
// need deeper metadata.
type DocProperties struct {
	Title       string
	Subject     string
	Author      string
	Keywords    string
	Description string
	Created     time.Time
	Modified    time.Time
}

func (p DocProperties) String() string {
	return fmt.Sprintf("Title: %q, Subject: %q, Author: %q, Keywords: %q, Description: %q, Created: %s, Modified: %s",
		p.Title, p.Subject, p.Author, p.Keywords, p.Description, p.Created.Format(time.RFC3339), p.Modified.Format(time.RFC3339))
}

// -----------------------------------------------------------------------------
// Run-level information
// -----------------------------------------------------------------------------

// RunStyle captures the character formatting for a run of text.
type RunStyle struct {
	FontFamily    string  // e.g. "Calibri"
	FontSizePt    float64 // size in points
	FontColor     string  // "RRGGBB"
	Bold          bool
	Italic        bool
	Underline     bool
	Strike        bool
	VerticalAlign string // "superscript" | "subscript" | "baseline"
}

func (s RunStyle) String() string {
	return fmt.Sprintf("FontFamily: %s, FontSizePt: %f, FontColor: %s, Bold: %t, Italic: %t, Underline: %t, Strike: %t, VerticalAlign: %s",
		s.FontFamily, s.FontSizePt, s.FontColor, s.Bold, s.Italic, s.Underline, s.Strike, s.VerticalAlign)
}

// RenderRun represents a single run (\<w:r>) within a paragraph.
type RenderRun struct {
	Run   document.Run // underlying run – useful for callers that need direct access
	Text  string       // already expanded/decoded text for the run
	Style RunStyle     // resolved run style
}

func (r RenderRun) String() string {
	return fmt.Sprintf("Text: %q, Style: [%s]", r.Text, r.Style.String())
}

// -----------------------------------------------------------------------------
// Paragraph-level information
// -----------------------------------------------------------------------------

// ParagraphStyle captures paragraph-level formatting.
type ParagraphStyle struct {
	Alignment     string  // "left" | "center" | "right" | "justify"
	LineSpacingPt float64 // leading – 0 means default/single
	SpaceBeforePt float64 // spacing before paragraph in points
	SpaceAfterPt  float64 // spacing after paragraph in points
	IndentLeftPx  float64 // left indent in pixels
	IndentRightPx float64 // right indent in pixels
	HeadingLevel  int     // 0 means normal paragraph, 1-6 for headings
	ListType      string  // "ordered" | "unordered" | "none"
	ListLevel     int     // nesting level (0-based)
}

func (s ParagraphStyle) String() string {
	return fmt.Sprintf("Alignment: %s, LineSpacingPt: %f, SpaceBeforePt: %f, SpaceAfterPt: %f, IndentLeftPx: %f, IndentRightPx: %f, HeadingLevel: %d, ListType: %s, ListLevel: %d",
		s.Alignment, s.LineSpacingPt, s.SpaceBeforePt, s.SpaceAfterPt, s.IndentLeftPx, s.IndentRightPx, s.HeadingLevel, s.ListType, s.ListLevel)
}

// RenderParagraph is the IR for a paragraph.
type RenderParagraph struct {
	Paragraph document.Paragraph // underlying paragraph – may be handy for later processing
	Runs      []RenderRun        // constituent runs
	Style     ParagraphStyle     // resolved paragraph style
}

func (p RenderParagraph) String() string {
	return fmt.Sprintf("Runs: %d, Style: [%s]", len(p.Runs), p.Style.String())
}

// -----------------------------------------------------------------------------
// Table-level information
// -----------------------------------------------------------------------------

// TableCellStyle represents the limited set of cell properties we are currently
// interested in (borders/shading could be added later).
type TableCellStyle struct {
	BackgroundColor string // fill colour – "RRGGBB"
	VerticalAlign   string // "top" | "middle" | "bottom"
}

func (s TableCellStyle) String() string {
	return fmt.Sprintf("BackgroundColor: %s, VerticalAlign: %s", s.BackgroundColor, s.VerticalAlign)
}

// RenderTableCell is the IR for a single table cell.  It can contain multiple
// paragraphs.
type RenderTableCell struct {
	Paragraphs []RenderParagraph // content
	ColSpan    int               // 1 if not horizontally merged
	RowSpan    int               // 1 if not vertically merged
	WidthPx    float64           // resolved width in px (0 means auto)
	Style      TableCellStyle    // resolved style
}

func (c RenderTableCell) String() string {
	return fmt.Sprintf("Paragraphs: %d, ColSpan: %d, RowSpan: %d, WidthPx: %f, Style: [%s]", len(c.Paragraphs), c.ColSpan, c.RowSpan, c.WidthPx, c.Style.String())
}

// RenderTableRow represents a row within a table.
type RenderTableRow struct {
	Cells    []RenderTableCell // cells, length equals column count of parent table
	HeightPx float64           // resolved height in px (0 means auto)
}

func (r RenderTableRow) String() string {
	return fmt.Sprintf("Cells: %d, HeightPx: %f", len(r.Cells), r.HeightPx)
}

// RenderTable is the IR for a table – rows in order.
type RenderTable struct {
	Rows []RenderTableRow // in order
}

func (t RenderTable) String() string {
	return fmt.Sprintf("Rows: %d", len(t.Rows))
}

// -----------------------------------------------------------------------------
// Block ordering
// -----------------------------------------------------------------------------

// DocumentBlock represents a top-level block element in the DOCX body – either
// a paragraph or a table.  Exactly one of Paragraph/Table will be non-nil.
type DocumentBlock struct {
	Paragraph *RenderParagraph
	Table     *RenderTable
}

// -----------------------------------------------------------------------------
// Top-level document model
// -----------------------------------------------------------------------------

type DocumentModel struct {
	Properties DocProperties

	// The document body is represented as a sequence of paragraphs and tables
	// in the order they appear.  For compatibility we keep dedicated slices
	// too, but the primary ordering source is Blocks.
	Blocks     []DocumentBlock
	Paragraphs []RenderParagraph
	Tables     []RenderTable
}

func (d DocumentModel) String() string {
	return fmt.Sprintf("Blocks: %d, Paragraphs: %d, Tables: %d, Properties: [%s]", len(d.Blocks), len(d.Paragraphs), len(d.Tables), d.Properties.String())
}
