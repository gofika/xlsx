package xlsx

import (
	"strconv"

	"github.com/gofika/xlsx/packaging"
)

// sharedStrings shared strings operator
type sharedStrings struct {
	file *fileImpl
	ss   map[string]int
}

func newSharedStrings(file *fileImpl) *sharedStrings {
	return &sharedStrings{
		file: file,
		ss:   make(map[string]int),
	}
}

func (s *sharedStrings) getSharedStrings() *packaging.XSharedStrings {
	return s.getFile().SharedStrings
}

func (s *sharedStrings) getFile() *packaging.XFile {
	return s.file.xFile
}

// Append append string or get string if str exist
// return string id
func (s *sharedStrings) Append(str string) (stringID string) {
	sst := s.getSharedStrings()
	needUpdateRelationships := sst.Count == 0
	sst.Count++ // ref count
	// for i, si := range sst.Si {
	// 	if str == si.T { // has exist one
	// 		return strconv.Itoa(i + 1)
	// 	}
	// }
	if idx, ok := s.ss[str]; ok {
		return strconv.Itoa(idx)
	}

	// need create new one
	sst.Si = append(sst.Si, &packaging.XSi{
		T: str,
		PhoneticPr: &packaging.XPhoneticPr{
			FontID: "0", // TODO: need ref from styles
			Type:   "noConversion",
		},
	})
	sst.UniqueCount++
	id := sst.UniqueCount - 1
	s.ss[str] = id

	if needUpdateRelationships { // need update relationships
		file := s.getFile()
		// update workbook.xml.rels
		file.WorkbookRelationships = packaging.NewWorkbookXRelationships(file)

		// update [Content_Types].xml
		file.ContentTypes = packaging.NewXContentTypes(file.WorkbookRelationships)
	}
	return strconv.Itoa(id)
}

func (s *sharedStrings) Get(stringID string) string {
	sst := s.getSharedStrings()
	si, err := strconv.Atoi(stringID)
	if err != nil || sst.Si == nil || si >= len(sst.Si) {
		return ""
	}
	return sst.Si[si].T
}

func (s *sharedStrings) calcMap() {
	sst := s.getSharedStrings()
	s.ss = make(map[string]int)
	for i, si := range sst.Si {
		s.ss[si.T] = i
	}
}
