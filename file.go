package xlsx

import "io"

var (
	DefaultSheetName = "Sheet1"
	DefaultFontName  = "Microsoft YaHei"
	DefaultFontSize  = 11
)

// File define for operation xlsx file
type File interface {
	// SaveFile save xlsx file
	SaveFile(name string) error

	// Save save to steam
	Save(w io.Writer) error

	// OpenSheet open a exist Sheet by name
	//
	// Example:
	//
	//     sheet := file.OpenSheet("Sheet1")
	//
	// return nil if sheet not exist
	OpenSheet(name string) Sheet

	// NewSheet create a new Sheet with sheet name
	// Example:
	//
	//     sheet := file.NewSheet("Sheet2")
	NewSheet(name string) Sheet

	// Sheets return all sheet for operator
	Sheets() []Sheet
}

// NewFile create a xlsx File
func NewFile(opts ...Option) File {
	return newFile(opts...)
}

// OpenFile open a xlsx file for operator
func OpenFile(name string) (File, error) {
	return openFile(name)
}

// OpenFileReader open a stream for operator
func OpenFileReader(r io.ReaderAt, size int64) (File, error) {
	return openFileReader(r, size)
}
