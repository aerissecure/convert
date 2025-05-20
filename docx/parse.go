package docx

import (
	"io"

	"github.com/unidoc/unioffice/document"
	"github.com/unidoc/unioffice/schema/soo/wml"
)

// ParseDocumentModel reads a DOCX document from the provided reader and size
// and builds a DocumentModel intermediate representation.  The current
// implementation focuses on text content and basic structure (paragraphs and
// tables).  Most styling information is left at zero-values for now – the
// HTML renderer will gracefully fall back to defaults when style attributes
// are empty.
func ParseDocumentModel(r io.ReaderAt, size int64) (DocumentModel, error) {
	doc, err := document.Read(r, size)
	if err != nil {
		return DocumentModel{}, err
	}

	var mdl DocumentModel

	// ---- Build lookup maps from underlying XML ptr -> high-level wrapper ----
	pMap := make(map[*wml.CT_P]document.Paragraph)
	for _, p := range doc.Paragraphs() {
		pMap[p.X()] = p
	}

	tMap := make(map[*wml.CT_Tbl]document.Table)
	for _, tbl := range doc.Tables() {
		tMap[tbl.X()] = tbl
	}

	// ---- Walk body elements in order ----
	body := doc.X().Body
	if body == nil {
		// Empty document
		return mdl, nil
	}

	for _, bl := range body.EG_BlockLevelElts {
		for _, c := range bl.EG_ContentBlockContent {
			// Paragraphs
			for _, cp := range c.P {
				if par, ok := pMap[cp]; ok {
					rp := convertParagraph(par)
					mdl.Paragraphs = append(mdl.Paragraphs, rp)
					rpCopy := rp
					mdl.Blocks = append(mdl.Blocks, DocumentBlock{Paragraph: &rpCopy})
				}
			}
			// Tables
			for _, ct := range c.Tbl {
				if tbl, ok := tMap[ct]; ok {
					rt := convertTable(tbl)
					mdl.Tables = append(mdl.Tables, rt)
					rtCopy := rt
					mdl.Blocks = append(mdl.Blocks, DocumentBlock{Table: &rtCopy})
				}
			}
		}
	}

	return mdl, nil
}

// convertRun builds a RenderRun from a unioffice Run. Styling information is
// currently resolved on a best-effort basis.  Where a style attribute cannot
// be determined it is simply left at the zero value.
func convertRun(r document.Run) RenderRun {
	return RenderRun{
		Run:   r,
		Text:  r.Text(),
		Style: RunStyle{}, // default/empty style
	}
}

// convertParagraph converts a unioffice Paragraph into the RenderParagraph IR.
func convertParagraph(p document.Paragraph) RenderParagraph {
	rp := RenderParagraph{Paragraph: p}

	for _, run := range p.Runs() {
		rp.Runs = append(rp.Runs, convertRun(run))
	}

	// Paragraph style left as zero-values for now.
	rp.Style = ParagraphStyle{}

	return rp
}

// convertTable converts a unioffice Table into the RenderTable IR.
func convertTable(t document.Table) RenderTable {
	rt := RenderTable{}

	for _, row := range t.Rows() {
		rr := RenderTableRow{}

		for _, cell := range row.Cells() {
			rc := RenderTableCell{
				ColSpan: 1,
				RowSpan: 1,
			}

			for _, p := range cell.Paragraphs() {
				rc.Paragraphs = append(rc.Paragraphs, convertParagraph(p))
			}

			rr.Cells = append(rr.Cells, rc)
		}

		rt.Rows = append(rt.Rows, rr)
	}

	return rt
}
