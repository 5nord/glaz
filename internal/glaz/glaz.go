package glaz

import (
	"bytes"
	"fmt"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
)

const (
	sheetOffset = 2 // Offset to January worksheet
	rowOffset   = 9 // Offset to first row containing work times

	beginWork  Column = "F"
	endWork    Column = "G"
	beginPause Column = "H"
	endPause   Column = "I"
)

type Column string

type Sheet struct {
	wb  *excelize.File
	err error
}

type Day struct {
	t     time.Time
	Work  Span
	Pause Span
}

func (d Day) String() string {
	b := &bytes.Buffer{}
	fmt.Fprintf(b, "work: %s - %s", d.Work.Begin.Format("15:04"), d.Work.End.Format("15:04"))

	workBegin := d.Work.Begin
	if workBegin.IsZero() {
		workBegin = time.Now()
	}

	workEnd := d.Work.End
	if workEnd.IsZero() {
		workEnd = time.Now()
	}

	worked := workEnd.Sub(workBegin)
	if worked < 0 {
		worked = 0
	}

	paused := d.Pause.End.Sub(d.Pause.Begin)
	if paused > 0 {
		fmt.Fprintf(b, ", paused: %.2gh", paused.Hours())
		worked -= paused
	}
	if worked > 0 {
		fmt.Fprintf(b, ", worked: %.2gh", worked.Hours())
	}

	return b.String()
}

type Span struct {
	Begin, End time.Time
}

func (s *Sheet) Today() Day {
	return s.Day(time.Now())
}

func (s *Sheet) Day(t time.Time) Day {
	return Day{
		t:     t,
		Work:  Span{s.getCellTime(t, beginWork), s.getCellTime(t, endWork)},
		Pause: Span{s.getCellTime(t, beginPause), s.getCellTime(t, endPause)},
	}
}

func (s *Sheet) Update(day Day) error {
	s.setCellTime(day.t, beginWork, day.Work.Begin)
	s.setCellTime(day.t, endWork, day.Work.End)
	s.setCellTime(day.t, beginPause, day.Pause.Begin)
	s.setCellTime(day.t, endPause, day.Pause.End)

	if s.err == nil {
		s.wb.UpdateLinkedValue()
	}

	if s.err == nil {
		s.wb.Save()
	}

	return s.err
}

func (s *Sheet) Worksheet(t time.Time) string {
	return s.wb.GetSheetName(int(t.Month()) + sheetOffset - 1)
}

func Cell(t time.Time, f Column) string {
	return fmt.Sprintf("%s%d", f, Row(t))
}

func Row(t time.Time) int {
	return int(t.Day()) + rowOffset - 1
}

func (s *Sheet) Err() error {
	return s.err
}

func (s *Sheet) getCellTime(t time.Time, col Column) (value time.Time) {
	if s.err != nil {
		return
	}

	ws := s.Worksheet(t)
	cell := Cell(t, col)

	// When reading a cell we get an incomplete date (e.g. "5:00" instead of
	// "5:00 PM"), due to custom formats. Therefore we have to overwrite the
	// style with something more useful, like
	// https://github.com/360EntSecGroup-Skylar/excelize/blob/1cbb05d/styles.go#L48
	orig, _ := s.wb.GetCellStyle(ws, cell)
	style, _ := s.wb.NewStyle(`{"number_format":18}`)
	s.wb.SetCellStyle(ws, cell, cell, style)

	v, err := s.wb.GetCellValue(ws, cell)
	if err != nil {
		s.err = err
		return
	}

	// Restore the original style of read cell.
	s.wb.SetCellStyle(ws, cell, cell, orig)

	if v == "" {
		return
	}

	value, s.err = time.Parse("3:04 pm", v)
	return time.Date(t.Year(), t.Month(), t.Day(), value.Hour(), value.Minute(), value.Second(), 0, time.Local)
}

func (s *Sheet) setCellTime(t time.Time, col Column, value time.Time) {
	if value.IsZero() || s.err != nil {
		return
	}

	ws := s.Worksheet(t)
	cell := Cell(t, col)

	// First we must change Location to UTC without changing the hours,
	// because excelize requires UTC times.
	// GLAZ sheet requires the date to be 30/12/1899 to operate correctly.
	value = time.Date(1899, 12, 31, value.Hour(), value.Minute(), value.Second(), 0, time.UTC)
	s.err = s.wb.SetCellValue(ws, cell, value)
}

func OpenFile(file string) (*Sheet, error) {
	wb, err := excelize.OpenFile(file)
	return &Sheet{
		wb: wb,
	}, err
}
