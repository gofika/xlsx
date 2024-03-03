package packaging

import (
	"encoding/xml"
	"strconv"

	"github.com/shopspring/decimal"
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

type OmitDecimalAttr decimal.Decimal

func (i OmitDecimalAttr) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if i.Value().IsZero() {
		return xml.Attr{}, nil
	}
	return xml.Attr{Name: name, Value: i.Value().String()}, nil
}

func (a OmitDecimalAttr) Value() decimal.Decimal {
	return decimal.Decimal(a)
}

func NewOmitDecimal(v decimal.Decimal) OmitDecimalAttr {
	if !v.IsZero() {
		return OmitDecimalAttr(v)
	}
	return OmitDecimalAttr(decimal.Zero)
}

func (a OmitDecimalAttr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	var err error
	if a.Value().IsZero() {
		err = e.EncodeElement("", start)
	} else {
		err = e.EncodeElement(a.Value().String(), start)
	}
	return err
}

type OmitUIntAttr uint

func (i OmitUIntAttr) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if i == 0 {
		return xml.Attr{}, nil
	}
	return xml.Attr{Name: name, Value: strconv.Itoa(int(i))}, nil
}

func (i OmitUIntAttr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if i == 0 {
		return nil
	}
	return e.EncodeElement(int(i), start)
}

type OmitUByteAttr uint8

func (i OmitUByteAttr) MarshalXMLAttr(name xml.Name) (xml.Attr, error) {
	if i == 0 {
		return xml.Attr{}, nil
	}
	return xml.Attr{Name: name, Value: strconv.Itoa(int(i))}, nil
}

func (i OmitUByteAttr) MarshalXML(e *xml.Encoder, start xml.StartElement) error {
	if i == 0 {
		return nil
	}
	return e.EncodeElement(int(i), start)
}
