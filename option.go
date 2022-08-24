package xlsx

import "github.com/gofika/xlsx/internal"

type Option interface {
	Apply(*internal.Settings)
}

func WithDefaultFontName(name string) Option {
	return withDefaultFontName{name}
}

type withDefaultFontName struct{ fontName string }

func (w withDefaultFontName) Apply(s *internal.Settings) {
	s.DefaultFontName = w.fontName
}

func WithDefaultFontSize(size int) Option {
	return withDefaultFontSize{size}
}

type withDefaultFontSize struct{ fontSize int }

func (w withDefaultFontSize) Apply(s *internal.Settings) {
	s.DefaultFontSize = w.fontSize
}

func WithDefaultSheetName(sheetName string) Option {
	return withDefaultSheetName{sheetName}
}

type withDefaultSheetName struct{ sheetName string }

func (w withDefaultSheetName) Apply(s *internal.Settings) {
	s.DefaultSheetName = w.sheetName
}
