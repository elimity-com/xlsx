package xlsx

import (
	"strconv"
	"time"
)

// Use this if you want to have an empty cell in your row
var EmptyStreamCell = StreamCell{}

// StreamCell holds the data, style and type of cell for streaming.
type StreamCell struct {
	cellData  string
	cellStyle StreamStyle
	cellType  CellType
	hyperlink Hyperlink
}

// NewStreamCell creates a new cell containing the given data with the given style and type.
func NewStreamCell(cellData string, cellStyle StreamStyle, cellType CellType) StreamCell {
	return StreamCell{
		cellData:  cellData,
		cellStyle: cellStyle,
		cellType:  cellType,
	}
}

// NewBooleanStreamCell creates a new cell containing the given boolean value.
func NewBooleanStreamCell(boolean bool) StreamCell{
	return NewStyledBooleanStreamCell(boolean, StreamStyleDefaultString)
}

// NewStyledBooleanStreamCell creates a new cell containing the given boolean value, and that has the given style.
func NewStyledBooleanStreamCell(boolean bool, style StreamStyle) StreamCell{
	var cellValue string
	if boolean {
		cellValue = "1"
	} else {
		cellValue = "0"
	}
	return NewStreamCell(cellValue, style, CellTypeBool)
}

// NewHyperlinkStreamCell creates a new cell containing the given hyperlink, displayData and tooltip.
// DisplayData and tooltip are not necessary. If an empty string is passed then they won't be set.
// You need to pas a valid URL starting with "http://" or "https://" for excel to recognize it
// as an external link.
// ! Important Note: Streaming Hyperlinks is impossible without breaking the rule of never building up memory.
// ! The structure of xlsx files requires writing data in 3 different places in 2 files and this requires using memory.
// ! Keep this in mind if you want to stream Hyperlinks.
func NewHyperlinkStreamCell(hyperlink string, displayData string, tooltip string) StreamCell {
	return NewStyledHyperlinkStreamCell(hyperlink, displayData, tooltip, StreamStyleHyperlink)
}

// NewStyledHyperlinkStreamCell creates a new cell containing the given hyperlink, displayData and tooltip.
// The cell will also have the specified style.
// DisplayData and tooltip are not necessary. If an empty string is passed then they won't be set.
// You need to pas a valid URL starting with "http://" or "https://" for excel to recognize it
// as an external link.
// ! Important Note: Streaming Hyperlinks is impossible without breaking the rule of never building up memory.
// ! The structure of xlsx files requires writing data in 3 different places in 2 files and this requires using memory.
// ! Keep this in mind if you want to stream Hyperlinks.
func NewStyledHyperlinkStreamCell(hyperlink string, displayData string, tooltip string, style StreamStyle) StreamCell {
	link := Hyperlink{Link: hyperlink}
	cellData := hyperlink
	if displayData != "" {
		link.DisplayString = displayData
		cellData = displayData
	}
	if tooltip != "" {
		link.Tooltip = tooltip
	}
	return StreamCell{
		cellData:  cellData,
		cellStyle: style,
		cellType:  CellTypeString,
		hyperlink: link,
	}
}

// NewStringStreamCell creates a new cell that holds string data, is of type string and uses general formatting.
func NewStringStreamCell(cellData string) StreamCell {
	return NewStreamCell(cellData, StreamStyleDefaultString, CellTypeString)
}

// NewStyledStringStreamCell creates a new cell that holds a string and is styled according to the given style.
func NewStyledStringStreamCell(cellData string, cellStyle StreamStyle) StreamCell {
	return NewStreamCell(cellData, cellStyle, CellTypeString)
}

// NewIntegerStreamCell creates a new cell that holds an integer value (represented as string),
// is formatted as a standard integer and is of type numeric.
func NewIntegerStreamCell(cellData int) StreamCell {
	return NewStreamCell(strconv.Itoa(cellData), StreamStyleDefaultInteger, CellTypeNumeric)
}

// NewStyledIntegerStreamCell creates a new cell that holds an integer value (represented as string)
// and is styled according to the given style.
func NewStyledIntegerStreamCell(cellData int, cellStyle StreamStyle) StreamCell {
	return NewStreamCell(strconv.Itoa(cellData), cellStyle, CellTypeNumeric)
}

// NewDateStreamCell creates a new cell that holds a date value and is formatted as dd-mm-yyyy
// and is of type numeric.
func NewDateStreamCell(t time.Time) StreamCell {
	excelTime := TimeToExcelTime(t, false)
	return NewStreamCell(strconv.Itoa(int(excelTime)), StreamStyleDefaultDate, CellTypeNumeric)
}
