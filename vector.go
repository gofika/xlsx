package xlsx

import "github.com/gofika/xlsx/packaging"

// VariantTypes
const (
	VariantTypeVariant = "variant"
	VariantTypeVTLPSTR = "lpstr"
)

// vector for operation xml vt:vector node
type vector struct {
	vector *packaging.XVector
}

func newVector(xv *packaging.XVector) *vector {
	return &vector{vector: xv}
}

// GetIntPairs get variant type pairs
//
// Example:
//
//	InputXML:
//	<vt:vector size="2" baseType="variant">
//	    <vt:variant>
//	        <vt:lpstr>Worksheets</vt:lpstr>
//	    </vt:variant>
//	    <vt:variant>
//	        <vt:i4>1</vt:i4>
//	    </vt:variant>
//	</vt:vector>
//
//	vector.GetIntPairs() // => map[Worksheets 1]
func (v *vector) GetIntPairs() (pairs map[string]int) {
	pairs = make(map[string]int)
	if v.vector.BaseType != VariantTypeVariant {
		return
	}
	name := ""
	for i, variant := range v.vector.Variant {
		if i%2 == 0 {
			name = variant.Lpstr
		} else {
			pairs[name] = int(variant.I4)
		}
	}
	return
}

// SetIntPairs set pairs for variant type
//
// Example:
//
//	variantPairs := map[string]int {
//	    "Worksheets": 1,
//	}
//	vector.SetIntPairs(variantPairs)
//
//	OutputXML:
//	<vt:vector size="2" baseType="variant">
//	    <vt:variant>
//	        <vt:lpstr>Worksheets</vt:lpstr>
//	    </vt:variant>
//	    <vt:variant>
//	        <vt:i4>1</vt:i4>
//	    </vt:variant>
//	</vt:vector>
func (v *vector) SetIntPairs(pairs map[string]int) {
	v.vector.BaseType = VariantTypeVariant
	v.vector.Variant = []*packaging.XVariant{}
	for name, value := range pairs {
		v.vector.Variant = append(v.vector.Variant, &packaging.XVariant{Lpstr: name})
		v.vector.Variant = append(v.vector.Variant, &packaging.XVariant{I4: int32(value)})
	}
	v.vector.Size = len(v.vector.Variant)
}

// GetStringArray get lpstr array
//
// Example:
//
//	InputXML:
//	<vt:vector size="2" baseType="lpstr">
//	    <vt:lpstr>Sheet1</vt:lpstr>
//	    <vt:lpstr>Sheet2</vt:lpstr>
//	</vt:vector>
//
//	vector.GetStringArray() // => [Sheet1 Sheet2]
func (v *vector) GetStringArray() []string {
	if v.vector.BaseType != VariantTypeVTLPSTR {
		return []string{}
	}
	return v.vector.Lpstr
}

// SetStringArray set lpstr array
//
// Example:
//
//	vector.SetStringArray([]string{"Sheet1", "Sheet2"})
//
//	OutputXML:
//	<vt:vector size="2" baseType="lpstr">
//	    <vt:lpstr>Sheet1</vt:lpstr>
//	    <vt:lpstr>Sheet2</vt:lpstr>
//	</vt:vector>
func (v *vector) SetStringArray(strArray []string) {
	v.vector.BaseType = VariantTypeVTLPSTR
	v.vector.Lpstr = strArray
	v.vector.Size = len(strArray)
}

// AppendString append lpstr to array
//
// Example:
//
//	InputXML:
//	<vt:vector size="2" baseType="lpstr">
//	    <vt:lpstr>Sheet1</vt:lpstr>
//	</vt:vector>
//
//	vector.AppendString("Sheet2")
//
//	OutputXML:
//	<vt:vector size="2" baseType="lpstr">
//	    <vt:lpstr>Sheet1</vt:lpstr>
//	    <vt:lpstr>Sheet2</vt:lpstr>
//	</vt:vector>
func (v *vector) AppendString(str string) {
	if v.vector.BaseType != VariantTypeVTLPSTR {
		return
	}
	v.vector.Lpstr = append(v.vector.Lpstr, str)
	v.vector.Size = len(v.vector.Lpstr)
}
