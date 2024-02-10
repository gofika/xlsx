package packaging

import (
	"encoding/xml"
	"strconv"
)

type BoolAttr int

func (b BoolAttr) Value() bool {
	return b != 0
}

func NewBool(b bool) BoolAttr {
	if b {
		return BoolAttr(1)
	}
	return BoolAttr(0)
}

func (b BoolAttr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	var err error
	if b == 0 {
		err = e.EncodeElement("0", start)
	} else {
		err = e.EncodeElement("1", start)
	}
	return err
}

func (b *BoolAttr) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var s string
	err := d.DecodeElement(&s, &start)
	if err != nil {
		return err
	}
	if s == "1" {
		*b = 1
	} else {
		*b = 0
	}
	return nil
}

type OmitIntAttr int

func (i OmitIntAttr) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if i == 0 {
		return xml.Attr{}, nil
	}
	return xml.Attr{Name: name, Value: strconv.Itoa(int(i))}, nil
}

func (i OmitIntAttr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if i == 0 {
		return nil
	}
	return e.EncodeElement(int(i), start)
}
