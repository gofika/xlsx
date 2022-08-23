package packaging

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"os"
	"path"
	"reflect"
	"strconv"
	"testing"

	"github.com/stretchr/testify/assert"
)

func XMLMarshalAppendHead(v any) (ret string, err error) {
	output, err := xml.Marshal(v)
	if err != nil {
		return
	}
	ret = fmt.Sprintf("%s%s", XMLProlog, output)
	return
}

func XMLMarshalAppendHeadIndent(v any) (ret string, err error) {
	output, err := xml.MarshalIndent(v, "", "    ")
	if err != nil {
		return
	}
	ret = fmt.Sprintf("%s%s", XMLProlog, output)
	return
}

var templatePath = path.Join("../test_docs/template")

var (
	needWriteTestFile, _ = strconv.ParseBool(os.Getenv("NEED_WRITE_TEST_FILE"))
)

func hasSameStructFields(va, vb reflect.Value, level string, traceLog *bytes.Buffer) bool {
	if va.Kind() != vb.Kind() { // not same kind
		traceLog.WriteString(fmt.Sprintf("%sKind: %s != %s\n", level, va.Kind(), vb.Kind()))
		return false
	}
	switch va.Kind() {
	case reflect.Invalid:
		traceLog.WriteString(fmt.Sprintf("%s[reflect.Invalid] Kind: %s | %s\n", level, va.Kind(), vb.Kind()))
		return true
	case reflect.Bool, reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64, reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64, reflect.Uintptr, reflect.Float32, reflect.Float64, reflect.Complex64, reflect.Complex128:
		return true
	case reflect.Chan, reflect.Func, reflect.String, reflect.UnsafePointer, reflect.Interface:
		return true
	case reflect.Map:
		traceLog.WriteString(fmt.Sprintf("%s[reflect.Map] return false", level))
		return false
	case reflect.Ptr:
		va = reflect.New(va.Type().Elem()).Elem()
		vb = reflect.New(vb.Type().Elem()).Elem()
		return hasSameStructFields(va, vb, level+"\t", traceLog)
	case reflect.Slice:
		ta := reflect.New(reflect.MakeSlice(reflect.New(va.Type()).Elem().Type(), 1, 1).Index(0).Type()).Elem()
		tb := reflect.New(reflect.MakeSlice(reflect.New(vb.Type()).Elem().Type(), 1, 1).Index(0).Type()).Elem()
		return hasSameStructFields(ta, tb, level+"\t", traceLog)
	case reflect.Struct:
		traceLog.WriteString(fmt.Sprintf("%s[Struct] %s | %s\n", level, va.Type().Name(), vb.Type().Name()))
		if va.NumField() != vb.NumField() {
			traceLog.WriteString(fmt.Sprintf("%s[Struct] NumField %d != %d\n", level, va.NumField(), vb.NumField()))
			return false
		}
		for i := 0; i < va.NumField(); i++ {
			fa := va.Field(i)
			fb := vb.Field(i)
			ta := va.Type().Field(i)
			tb := vb.Type().Field(i)
			if fa.Kind() != fb.Kind() {
				traceLog.WriteString(fmt.Sprintf("%s[Field] %s(%s) != %s(%s)\n", level, ta.Name, ta.Type, tb.Name, tb.Type))
				return false
			}
			if fa.Kind() == reflect.Ptr {
				fa = reflect.New(fa.Type().Elem()).Elem()
				fb = reflect.New(fb.Type().Elem()).Elem()
			}

			traceLog.WriteString(fmt.Sprintf("\t%s[Field] %s(%s) | %s(%s)\n", level, ta.Name, ta.Type, tb.Name, tb.Type))
			if !hasSameStructFields(fa, fb, level+"\t", traceLog) {
				return false
			}
		}
	case reflect.Array:
		traceLog.WriteString("[reflect.Array] return false")
		return false
	default:
		traceLog.WriteString(fmt.Sprintf("[default] %sva.Kind:%s vb.Kind:%s\n", level, va.Kind(), vb.Kind()))
	}
	return true
}

// check for ns structs
func TestFixXMLStruct(t *testing.T) {
	buf := bytes.NewBufferString("")
	assert.Equal(t, hasSameStructFields(reflect.ValueOf(&XCoreProperties{}), reflect.ValueOf(&XCorePropertiesU{}), "", buf), true, buf.String())
	buf.Reset()
	assert.Equal(t, hasSameStructFields(reflect.ValueOf(&XExtendedProperties{}), reflect.ValueOf(&XExtendedPropertiesU{}), "", buf), true, buf.String())
	buf.Reset()
	assert.Equal(t, hasSameStructFields(reflect.ValueOf(&XStyleSheet{}), reflect.ValueOf(&XStyleSheetU{}), "", buf), true, buf.String())
	buf.Reset()
	assert.Equal(t, hasSameStructFields(reflect.ValueOf(&XTheme{}), reflect.ValueOf(&XThemeU{}), "", buf), true, buf.String())
	buf.Reset()
	assert.Equal(t, hasSameStructFields(reflect.ValueOf(&XWorkbook{}), reflect.ValueOf(&XWorkbookU{}), "", buf), true, buf.String())
	buf.Reset()
}
