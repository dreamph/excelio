package excelio

import (
	"bytes"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"
	"time"

	"github.com/go-playground/validator/v10"
	"github.com/xuri/excelize/v2"
)

type metaSample struct {
	Code   string    `excel:"Code, Product Code" required:"true"`
	Amount int       `col:"2"`
	When   time.Time `excelcol:"c" fmt:"2006-01-02"`
	hidden string    `excel:"Hidden"`
	Ignore string
}

type readSample struct {
	Code   string    `excel:"Code" required:"true"`
	Name   string    `col:"2"`
	Price  float64   `excelcol:"C"`
	Active bool      `excel:"Active"`
	Since  time.Time `excel:"Since" fmt:"2006-01-02"`
}

type validatedSample struct {
	Code string `excel:"Code" validate:"required"`
	Qty  int    `excel:"Qty" validate:"gt=0"`
}

type writeSample struct {
	Code string    `excel:"Code"`
	Qty  int       `excel:"Qty"`
	When time.Time `excel:"When" fmt:"2006-01-02"`
	Opt  *int      `excel:"Opt"`
}

type writeOrderSample struct {
	Col3 string `excel:"Col3" col:"3"`
	ColA string `excel:"ColA" excelcol:"A"`
	Auto string `excel:"Auto"`
}

func mustWorkbookBytes(t *testing.T, sheet string, rows [][]any, extraCells map[string]any) []byte {
	t.Helper()

	f := excelize.NewFile()
	defaultSheet := f.GetSheetName(f.GetActiveSheetIndex())
	if sheet == "" {
		sheet = defaultSheet
	} else if sheet != defaultSheet {
		idx, err := f.NewSheet(sheet)
		if err != nil {
			t.Fatalf("NewSheet() error: %v", err)
		}
		_ = f.DeleteSheet(defaultSheet)
		f.SetActiveSheet(idx)
	}

	for i, row := range rows {
		axis := fmt.Sprintf("A%d", i+1)
		if err := f.SetSheetRow(sheet, axis, &row); err != nil {
			t.Fatalf("SetSheetRow(%s) error: %v", axis, err)
		}
	}
	for cell, value := range extraCells {
		if err := f.SetCellValue(sheet, cell, value); err != nil {
			t.Fatalf("SetCellValue(%s) error: %v", cell, err)
		}
	}

	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	_ = f.Close()
	return buf.Bytes()
}

func mustOpenWorkbook(t *testing.T, b []byte) *excelize.File {
	t.Helper()
	f, err := excelize.OpenReader(bytes.NewReader(b))
	if err != nil {
		t.Fatalf("OpenReader() error: %v", err)
	}
	return f
}

func mustCell(t *testing.T, f *excelize.File, sheet, cell string) string {
	t.Helper()
	v, err := f.GetCellValue(sheet, cell)
	if err != nil {
		t.Fatalf("GetCellValue(%s) error: %v", cell, err)
	}
	return v
}

func TestApplyDefaults(t *testing.T) {
	o := &Options{}
	applyDefaults(o)
	if o.SheetIndex != 0 || o.HeaderRow != 1 || o.FirstDataRow != 2 {
		t.Fatalf("unexpected defaults: %+v", *o)
	}

	o = &Options{HeaderRow: 5}
	applyDefaults(o)
	if o.FirstDataRow != 6 {
		t.Fatalf("expected FirstDataRow=6, got %d", o.FirstDataRow)
	}

	o = &Options{SheetIndex: -10, HeaderRow: 1, FirstDataRow: 7}
	applyDefaults(o)
	if o.SheetIndex != 0 || o.FirstDataRow != 7 {
		t.Fatalf("unexpected normalized options: %+v", *o)
	}
}

func TestOptionHelpers(t *testing.T) {
	o := &Options{}
	Sheet("Products")(o)
	SheetAt(2)(o)
	Header(3)(o)
	StartRow(9)(o)
	ErrCol(11)(o)
	v := validator.New()
	UseValidator(v)(o)

	if o.SheetName != "Products" || o.SheetIndex != 2 || o.HeaderRow != 3 || o.FirstDataRow != 9 || o.ErrorColumnIndex != 11 || o.GoValidator != v {
		t.Fatalf("unexpected options after helpers: %+v", *o)
	}

	type row struct{ Name string }
	var seen []string
	OnStreamRow(func(rowIdx, logicalIdx int, r *row, rowErrs []RowError) error {
		if r == nil {
			seen = append(seen, "nil")
			return nil
		}
		seen = append(seen, r.Name)
		return nil
	})(o)

	p := row{Name: "ptr"}
	if err := o.streamHandler(1, 1, &p, nil); err != nil {
		t.Fatalf("streamHandler ptr error: %v", err)
	}
	if err := o.streamHandler(2, 2, row{Name: "val"}, nil); err != nil {
		t.Fatalf("streamHandler val error: %v", err)
	}
	if err := o.streamHandler(3, 3, nil, nil); err != nil {
		t.Fatalf("streamHandler nil error: %v", err)
	}

	want := []string{"ptr", "val", "nil"}
	if !reflect.DeepEqual(seen, want) {
		t.Fatalf("unexpected streamed rows: got %v want %v", seen, want)
	}

	// nil handler should keep the existing handler unchanged.
	old := o.streamHandler
	OnStreamRow[row](nil)(o)
	if reflect.ValueOf(old).Pointer() != reflect.ValueOf(o.streamHandler).Pointer() {
		t.Fatalf("expected nil OnStreamRow to preserve existing handler")
	}
}

func TestSplitAndTrim(t *testing.T) {
	if got := splitAndTrim(""); got != nil {
		t.Fatalf("expected nil for empty input, got %v", got)
	}
	got := splitAndTrim(" A, ,B ,  C ")
	want := []string{"A", "B", "C"}
	if !reflect.DeepEqual(got, want) {
		t.Fatalf("unexpected split result: got %v want %v", got, want)
	}
}

func TestGetTypeMetaAndFindFieldByName(t *testing.T) {
	meta, err := getTypeMeta(reflect.TypeOf(metaSample{}))
	if err != nil {
		t.Fatalf("getTypeMeta(metaSample) error: %v", err)
	}
	if len(meta.Fields) != 3 {
		t.Fatalf("expected 3 mapped fields, got %d", len(meta.Fields))
	}

	fmCode := meta.FindFieldByName("Code")
	if fmCode == nil || !fmCode.Required {
		t.Fatalf("expected Code field to be present and required")
	}
	if meta.HeaderToField["code"] != fmCode || meta.HeaderToField["product code"] != fmCode {
		t.Fatalf("expected header aliases to resolve to Code field")
	}

	fmAmount := meta.FindFieldByName("Amount")
	if fmAmount == nil || fmAmount.ColIndexTag != 1 {
		t.Fatalf("expected Amount col index tag to be 1, got %+v", fmAmount)
	}

	fmWhen := meta.FindFieldByName("When")
	if fmWhen == nil || fmWhen.ColLetterTag != "C" || fmWhen.TimeFormat != "2006-01-02" {
		t.Fatalf("unexpected When metadata: %+v", fmWhen)
	}

	meta2, err := getTypeMeta(reflect.TypeOf(&metaSample{}))
	if err != nil {
		t.Fatalf("getTypeMeta(ptr) error: %v", err)
	}
	if meta2.FindFieldByName("Code") == nil {
		t.Fatalf("expected pointer type to resolve same metadata")
	}

	if _, err := getTypeMeta(reflect.TypeOf(123)); err == nil {
		t.Fatalf("expected non-struct metadata error")
	}
}

func TestColumnHelpers(t *testing.T) {
	if got := colLetter(-1); got != "" {
		t.Fatalf("expected empty letter for -1, got %q", got)
	}
	if got := colLetter(0); got != "A" {
		t.Fatalf("expected A, got %q", got)
	}
	if got := colLetter(25); got != "Z" {
		t.Fatalf("expected Z, got %q", got)
	}
	if got := colLetter(26); got != "AA" {
		t.Fatalf("expected AA, got %q", got)
	}
	if got := colLetter(701); got != "ZZ" {
		t.Fatalf("expected ZZ, got %q", got)
	}

	if got := colIndexFromLetter(" "); got != -1 {
		t.Fatalf("expected -1 for empty letter, got %d", got)
	}
	if got := colIndexFromLetter("a"); got != 0 {
		t.Fatalf("expected 0 for a, got %d", got)
	}
	if got := colIndexFromLetter("AZ"); got != 51 {
		t.Fatalf("expected 51 for AZ, got %d", got)
	}
	if got := colIndexFromLetter("A1"); got != -1 {
		t.Fatalf("expected -1 for invalid letter A1, got %d", got)
	}
}

func TestParseHeader(t *testing.T) {
	data := mustWorkbookBytes(t, "Products", [][]any{
		{"skip"},
		{" Code ", "Qty", ""},
	}, nil)
	f := mustOpenWorkbook(t, data)
	defer f.Close()

	h, err := parseHeader(f, "Products", 2)
	if err != nil {
		t.Fatalf("parseHeader() error: %v", err)
	}
	if h[0] != "Code" || h[1] != "Qty" || h[2] != "" {
		t.Fatalf("unexpected header map: %#v", h)
	}

	if _, err := parseHeader(f, "Products", 5); err == nil {
		t.Fatalf("expected header row not found error")
	}
}

func TestParseBool(t *testing.T) {
	trueValues := []string{"1", "true", "T", " yes ", "Y", "on"}
	for _, raw := range trueValues {
		v, err := parseBool(raw)
		if err != nil || !v {
			t.Fatalf("expected %q to parse true, got v=%v err=%v", raw, v, err)
		}
	}

	falseValues := []string{"0", "false", "F", " no ", "N", "off"}
	for _, raw := range falseValues {
		v, err := parseBool(raw)
		if err != nil || v {
			t.Fatalf("expected %q to parse false, got v=%v err=%v", raw, v, err)
		}
	}

	if _, err := parseBool("maybe"); err == nil {
		t.Fatalf("expected invalid bool error")
	}
}

func TestExcelSerialToTime(t *testing.T) {
	if _, err := excelSerialToTime(0); err == nil {
		t.Fatalf("expected invalid serial error")
	}

	tm, err := excelSerialToTime(1)
	if err != nil {
		t.Fatalf("excelSerialToTime(1) error: %v", err)
	}
	want := time.Date(1899, 12, 31, 0, 0, 0, 0, time.UTC)
	if !tm.Equal(want) {
		t.Fatalf("unexpected serial date: got %v want %v", tm, want)
	}

	tm, err = excelSerialToTime(2.5)
	if err != nil {
		t.Fatalf("excelSerialToTime(2.5) error: %v", err)
	}
	want = time.Date(1900, 1, 1, 12, 0, 0, 0, time.UTC)
	if !tm.Equal(want) {
		t.Fatalf("unexpected serial datetime: got %v want %v", tm, want)
	}
}

func TestParseTime(t *testing.T) {
	fm := &fieldMeta{TimeFormat: "02/01/2006"}
	tm, err := parseTime("07/03/2026", fm)
	if err != nil {
		t.Fatalf("custom parse error: %v", err)
	}
	if !tm.Equal(time.Date(2026, 3, 7, 0, 0, 0, 0, time.UTC)) {
		t.Fatalf("unexpected custom time: %v", tm)
	}

	tm, err = parseTime("2026-03-07T12:34:56Z", nil)
	if err != nil {
		t.Fatalf("RFC3339 parse error: %v", err)
	}
	if !tm.Equal(time.Date(2026, 3, 7, 12, 34, 56, 0, time.UTC)) {
		t.Fatalf("unexpected RFC3339 time: %v", tm)
	}

	tm, err = parseTime("2026-03-07", nil)
	if err != nil {
		t.Fatalf("layout parse error: %v", err)
	}
	if tm.Year() != 2026 || tm.Month() != 3 || tm.Day() != 7 {
		t.Fatalf("unexpected layout time: %v", tm)
	}

	tm, err = parseTime("2", nil)
	if err != nil {
		t.Fatalf("serial parse error: %v", err)
	}
	if !tm.Equal(time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)) {
		t.Fatalf("unexpected serial time: %v", tm)
	}

	if _, err := parseTime("", nil); err == nil {
		t.Fatalf("expected empty time error")
	}
	if _, err := parseTime("not-time", nil); err == nil {
		t.Fatalf("expected invalid time error")
	}
}

func TestSetFieldValueAndConvertAndSet(t *testing.T) {
	type conv struct {
		S  string
		I  int
		U  uint
		F  float64
		B  bool
		T  time.Time
		PI *int
		X  struct{ A int }
	}

	obj := &conv{}
	v := reflect.ValueOf(obj).Elem()

	if err := setFieldValue(v.FieldByName("S"), nil, "abc"); err != nil {
		t.Fatalf("set string error: %v", err)
	}
	if err := setFieldValue(v.FieldByName("I"), nil, "42"); err != nil {
		t.Fatalf("set int error: %v", err)
	}
	if err := setFieldValue(v.FieldByName("U"), nil, "7"); err != nil {
		t.Fatalf("set uint error: %v", err)
	}
	if err := setFieldValue(v.FieldByName("F"), nil, "3.14"); err != nil {
		t.Fatalf("set float error: %v", err)
	}
	if err := setFieldValue(v.FieldByName("B"), nil, "yes"); err != nil {
		t.Fatalf("set bool error: %v", err)
	}
	if err := setFieldValue(v.FieldByName("T"), &fieldMeta{TimeFormat: "2006-01-02"}, "2026-03-07"); err != nil {
		t.Fatalf("set time error: %v", err)
	}

	if got := obj.S; got != "abc" {
		t.Fatalf("unexpected string value: %q", got)
	}
	if obj.I != 42 || obj.U != 7 || obj.F != 3.14 || !obj.B {
		t.Fatalf("unexpected numeric/bool values: %+v", obj)
	}

	if err := setFieldValue(v.FieldByName("PI"), nil, ""); err != nil {
		t.Fatalf("set nil pointer error: %v", err)
	}
	if obj.PI != nil {
		t.Fatalf("expected PI to stay nil")
	}

	if err := setFieldValue(v.FieldByName("PI"), nil, "9"); err != nil {
		t.Fatalf("set pointer error: %v", err)
	}
	if obj.PI == nil || *obj.PI != 9 {
		t.Fatalf("unexpected pointer value: %+v", obj.PI)
	}

	if err := setFieldValue(v.FieldByName("B"), nil, "not-bool"); err == nil {
		t.Fatalf("expected bool conversion error")
	}
	if err := convertAndSet(v.FieldByName("F"), nil, "bad-number"); err == nil {
		t.Fatalf("expected float conversion error")
	}
	if err := setFieldValue(v.FieldByName("X"), nil, "anything"); err == nil {
		t.Fatalf("expected unsupported kind error")
	}
}

func TestBuildFieldColIndex(t *testing.T) {
	type mapping struct {
		ByIdx    string `excel:"HeaderIdx" col:"2"`
		ByLetter string `excel:"HeaderLetter" excelcol:"D"`
		ByHeader string `excel:"Price,Cost"`
	}

	meta, err := getTypeMeta(reflect.TypeOf(mapping{}))
	if err != nil {
		t.Fatalf("getTypeMeta(mapping) error: %v", err)
	}

	idxMap := buildFieldColIndex(meta, map[string]int{"price": 4})
	if idxMap[meta.FindFieldByName("ByIdx")] != 1 {
		t.Fatalf("expected ByIdx at col 1")
	}
	if idxMap[meta.FindFieldByName("ByLetter")] != 3 {
		t.Fatalf("expected ByLetter at col 3")
	}
	if idxMap[meta.FindFieldByName("ByHeader")] != 4 {
		t.Fatalf("expected ByHeader at col 4")
	}
}

func TestBuildRowError(t *testing.T) {
	fm := &fieldMeta{FieldName: "Code", ColumnNames: []string{"Product Code"}}
	err := errors.New("bad value")
	re := buildRowError(5, 2, fm, 1, map[int]string{1: "Display Name"}, []string{"A", "B"}, err)

	if re.ExcelRowIndex != 5 || re.LogicalIndex != 2 {
		t.Fatalf("unexpected row indices: %+v", re)
	}
	if re.ColIndex != 2 || re.ColLetter != "B" {
		t.Fatalf("unexpected column coordinates: %+v", re)
	}
	if re.Field != "Code" || re.Column != "Display Name" || re.Value != "B" || !errors.Is(re.Err, err) {
		t.Fatalf("unexpected row error payload: %+v", re)
	}

	re = buildRowError(7, 3, fm, 9, nil, []string{"x"}, err)
	if re.Column != "Product Code" || re.ColLetter != "" || re.Value != "" {
		t.Fatalf("expected fallback column/value handling, got %+v", re)
	}
}

func TestIsRowEmpty(t *testing.T) {
	if !isRowEmpty([]string{"", " ", "\t"}) {
		t.Fatalf("expected row to be empty")
	}
	if isRowEmpty([]string{"", "x"}) {
		t.Fatalf("expected row with value to be non-empty")
	}
}

func TestResolveSheet(t *testing.T) {
	f := excelize.NewFile()
	defer f.Close()
	if _, err := f.NewSheet("Data"); err != nil {
		t.Fatalf("NewSheet(Data) error: %v", err)
	}

	o := &Options{}
	applyDefaults(o)
	sheet, err := resolveSheet(f, o)
	if err != nil {
		t.Fatalf("resolve default sheet error: %v", err)
	}
	if sheet == "" {
		t.Fatalf("expected non-empty default sheet")
	}

	o.SheetName = "Data"
	resolvedAgain, err := resolveSheet(f, o)
	if err != nil {
		t.Fatalf("resolve cached sheet error: %v", err)
	}
	if resolvedAgain != sheet {
		t.Fatalf("expected cached sheet %q, got %q", sheet, resolvedAgain)
	}

	o2 := &Options{SheetName: "Data"}
	sheet, err = resolveSheet(f, o2)
	if err != nil || sheet != "Data" {
		t.Fatalf("expected explicit Data sheet, got sheet=%q err=%v", sheet, err)
	}

	o3 := &Options{SheetIndex: 99}
	if _, err := resolveSheet(f, o3); err == nil {
		t.Fatalf("expected out-of-range sheet index error")
	}
}

func TestReadFromExcelFile(t *testing.T) {
	data := mustWorkbookBytes(t, "Products", [][]any{
		{"Code", "Name", "Price", "Active", "Since"},
		{"P1", "Item 1", "12.5", "yes", "2026-03-01"},
		{"", "Item 2", "9.9", "true", "2026-03-02"},
		{"P3", "Item 3", "bad", "false", "2026-03-03"},
		{"", "", "", "", ""},
		{"P4", "Item 4", "7.25", "off", "2"},
	}, nil)

	f := mustOpenWorkbook(t, data)
	defer f.Close()

	o := &Options{
		SheetName:    "Products",
		HeaderRow:    1,
		FirstDataRow: 2,
		RowIndexMapper: func(excelRow, dataIdx int) int {
			return excelRow + 100
		},
	}
	applyDefaults(o)

	rows, errs, err := readFromExcelFile[readSample](f, o)
	if err != nil {
		t.Fatalf("readFromExcelFile() error: %v", err)
	}
	if len(rows) != 2 {
		t.Fatalf("expected 2 valid rows, got %d", len(rows))
	}
	if rows[0].Code != "P1" || rows[1].Code != "P4" {
		t.Fatalf("unexpected mapped rows: %+v", rows)
	}
	if len(errs) != 2 {
		t.Fatalf("expected 2 row errors, got %d (%v)", len(errs), errs)
	}

	sort.Slice(errs, func(i, j int) bool { return errs[i].ExcelRowIndex < errs[j].ExcelRowIndex })
	if errs[0].ExcelRowIndex != 3 || errs[0].LogicalIndex != 103 || errs[0].Field != "Code" || !strings.Contains(errs[0].Err.Error(), "required value is empty") {
		t.Fatalf("unexpected first error: %+v", errs[0])
	}
	if errs[1].ExcelRowIndex != 4 || errs[1].LogicalIndex != 104 || errs[1].Field != "Price" || !strings.Contains(errs[1].Err.Error(), "invalid syntax") {
		t.Fatalf("unexpected second error: %+v", errs[1])
	}

	if _, _, err := readFromExcelFile[int](f, o); err == nil {
		t.Fatalf("expected readFromExcelFile[int] to fail for non-struct type")
	}
}

func TestReadAndStreamAPIs(t *testing.T) {
	data := mustWorkbookBytes(t, "Sheet1", [][]any{
		{"Code", "Qty"},
		{"A", "1"},
		{"", "2"},
		{"B", "0"},
	}, nil)
	v := validator.New()

	rows, errs, err := Read[validatedSample](bytes.NewReader(data), Header(1), StartRow(2), UseValidator(v))
	if err != nil {
		t.Fatalf("Read() error: %v", err)
	}
	if len(rows) != 1 || rows[0].Code != "A" {
		t.Fatalf("unexpected Read() rows: %+v", rows)
	}
	if len(errs) != 2 {
		t.Fatalf("expected 2 validation errors from Read(), got %d", len(errs))
	}

	if _, err := Stream[validatedSample](bytes.NewReader(data), Header(1), StartRow(2)); err == nil {
		t.Fatalf("expected Stream() error when handler is missing")
	}

	var callbackRows []int
	var callbackStates []string
	streamErrs, err := Stream[validatedSample](
		bytes.NewReader(data),
		Header(1),
		StartRow(2),
		UseValidator(v),
		OnStreamRow(func(rowIdx, logicalIdx int, obj *validatedSample, rowErrs []RowError) error {
			callbackRows = append(callbackRows, rowIdx)
			if len(rowErrs) > 0 {
				callbackStates = append(callbackStates, "err")
				if obj != nil {
					t.Fatalf("expected nil obj for invalid row %d", rowIdx)
				}
				return nil
			}
			callbackStates = append(callbackStates, obj.Code)
			return nil
		}),
	)
	if err != nil {
		t.Fatalf("Stream() error: %v", err)
	}
	if len(streamErrs) != 2 {
		t.Fatalf("expected 2 stream errors, got %d", len(streamErrs))
	}
	if !reflect.DeepEqual(callbackRows, []int{2, 3, 4}) {
		t.Fatalf("unexpected callback rows: %v", callbackRows)
	}
	if !reflect.DeepEqual(callbackStates, []string{"A", "err", "err"}) {
		t.Fatalf("unexpected callback states: %v", callbackStates)
	}

	sentinel := errors.New("stop-stream")
	streamErrs, err = Stream[validatedSample](
		bytes.NewReader(data),
		Header(1),
		StartRow(2),
		UseValidator(v),
		OnStreamRow(func(rowIdx, logicalIdx int, obj *validatedSample, rowErrs []RowError) error {
			if rowIdx == 3 {
				return sentinel
			}
			return nil
		}),
	)
	if !errors.Is(err, sentinel) {
		t.Fatalf("expected sentinel error, got %v", err)
	}
	if len(streamErrs) != 1 || streamErrs[0].ExcelRowIndex != 3 {
		t.Fatalf("expected accumulated errors before stop, got %+v", streamErrs)
	}

	if _, _, err := Read[validatedSample](bytes.NewReader([]byte("not-an-xlsx"))); err == nil {
		t.Fatalf("expected Read() to fail for invalid workbook bytes")
	}
}

func TestFileAPIsAndErrorWriteBack(t *testing.T) {
	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "input.xlsx")

	data := mustWorkbookBytes(t, "Products", [][]any{
		{"Code", "Qty", "Errors"},
		{"A", "1", ""},
		{"B", "bad", ""},
	}, nil)
	if err := os.WriteFile(path, data, 0o644); err != nil {
		t.Fatalf("WriteFile input error: %v", err)
	}

	rows, errs, err := ReadFile[validatedSample](path, Sheet("Products"), Header(1), StartRow(2))
	if err != nil {
		t.Fatalf("ReadFile() error: %v", err)
	}
	if len(rows) != 1 || len(errs) != 1 || errs[0].ExcelRowIndex != 3 {
		t.Fatalf("unexpected ReadFile output rows=%v errs=%v", rows, errs)
	}

	streamErrs, err := StreamFile[validatedSample](
		path,
		Sheet("Products"),
		Header(1),
		StartRow(2),
		ErrCol(3),
		OnStreamRow(func(rowIdx, logicalIdx int, obj *validatedSample, rowErrs []RowError) error { return nil }),
	)
	if err != nil {
		t.Fatalf("StreamFile() error: %v", err)
	}
	if len(streamErrs) != 1 || streamErrs[0].ExcelRowIndex != 3 {
		t.Fatalf("unexpected StreamFile errors: %+v", streamErrs)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile(path) error: %v", err)
	}
	defer f.Close()
	msg := mustCell(t, f, "Products", "C3")
	if !strings.Contains(msg, "invalid syntax") {
		t.Fatalf("expected StreamFile to write error message into C3, got %q", msg)
	}

	if _, err := StreamFile[validatedSample](path, Sheet("Products"), Header(1), StartRow(2)); err == nil {
		t.Fatalf("expected StreamFile() error when handler is missing")
	}
}

func TestWriteErrorsToAndWriteErrors(t *testing.T) {
	input := mustWorkbookBytes(t, "Products", [][]any{
		{"Code", "Name"},
		{"A", "Item"},
		{"B", "Item"},
	}, map[string]any{"J2": "old"})

	errs := []RowError{
		{ExcelRowIndex: 2, Err: errors.New("first")},
		{ExcelRowIndex: 2, Err: errors.New("second")},
		{ExcelRowIndex: 3, Err: errors.New("third")},
		{ExcelRowIndex: 0, Err: errors.New("ignored")},
	}

	var out bytes.Buffer
	if err := WriteErrorsTo(&out, bytes.NewReader(input), errs, Sheet("Products"), ErrCol(10)); err != nil {
		t.Fatalf("WriteErrorsTo() error: %v", err)
	}

	f := mustOpenWorkbook(t, out.Bytes())
	defer f.Close()
	if got := mustCell(t, f, "Products", "J2"); got != "old\nfirst\nsecond" {
		t.Fatalf("unexpected merged errors in J2: %q", got)
	}
	if got := mustCell(t, f, "Products", "J3"); got != "third" {
		t.Fatalf("unexpected error in J3: %q", got)
	}

	var copied bytes.Buffer
	if err := WriteErrorsTo(&copied, bytes.NewReader(input), nil); err != nil {
		t.Fatalf("WriteErrorsTo() copy-only error: %v", err)
	}
	if !bytes.Equal(copied.Bytes(), input) {
		t.Fatalf("expected copy-only WriteErrorsTo to preserve bytes")
	}

	if err := WriteErrorsTo(&bytes.Buffer{}, bytes.NewReader(input), errs); err == nil {
		t.Fatalf("expected WriteErrorsTo to require ErrCol when errs are present")
	}

	tmpDir := t.TempDir()
	path := filepath.Join(tmpDir, "errors.xlsx")
	if err := os.WriteFile(path, input, 0o644); err != nil {
		t.Fatalf("WriteFile(errors.xlsx) error: %v", err)
	}

	if err := WriteErrors(path, nil); err != nil {
		t.Fatalf("WriteErrors should no-op on empty errors: %v", err)
	}
	if err := WriteErrors(path, errs); err == nil {
		t.Fatalf("expected WriteErrors to require ErrCol when errs are present")
	}
	if err := WriteErrors(path, errs, Sheet("Products"), ErrCol(10)); err != nil {
		t.Fatalf("WriteErrors(path) error: %v", err)
	}

	f2, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("OpenFile(errors.xlsx) error: %v", err)
	}
	defer f2.Close()
	if got := mustCell(t, f2, "Products", "J3"); got != "third" {
		t.Fatalf("expected written error in J3, got %q", got)
	}
}

func TestBuildFieldOrderForWriteAndFieldHeaderName(t *testing.T) {
	meta, err := getTypeMeta(reflect.TypeOf(writeOrderSample{}))
	if err != nil {
		t.Fatalf("getTypeMeta(writeOrderSample) error: %v", err)
	}

	fields, idx, maxCol := buildFieldOrderForWrite(meta)
	if maxCol != 2 {
		t.Fatalf("expected maxCol=2, got %d", maxCol)
	}
	if idx[meta.FindFieldByName("ColA")] != 0 || idx[meta.FindFieldByName("Auto")] != 1 || idx[meta.FindFieldByName("Col3")] != 2 {
		t.Fatalf("unexpected index map: %+v", idx)
	}

	gotNames := make([]string, 0, len(fields))
	for _, fm := range fields {
		gotNames = append(gotNames, fm.FieldName)
	}
	wantNames := []string{"ColA", "Auto", "Col3"}
	if !reflect.DeepEqual(gotNames, wantNames) {
		t.Fatalf("unexpected field order: got %v want %v", gotNames, wantNames)
	}

	if fieldHeaderName(nil) != "" {
		t.Fatalf("expected empty header name for nil field")
	}
	if fieldHeaderName(meta.FindFieldByName("Col3")) != "Col3" {
		t.Fatalf("expected header name from excel tag")
	}
}

func TestValueToCell(t *testing.T) {
	if got := valueToCell(reflect.Value{}, nil); got != "" {
		t.Fatalf("expected empty string for invalid reflect value, got %#v", got)
	}

	s := "abc"
	if got := valueToCell(reflect.ValueOf(&s), nil); got != "abc" {
		t.Fatalf("expected pointer string value, got %#v", got)
	}

	var ps *string
	if got := valueToCell(reflect.ValueOf(ps), nil); got != "" {
		t.Fatalf("expected nil pointer to map to empty string, got %#v", got)
	}

	now := time.Date(2026, 3, 7, 14, 15, 16, 0, time.UTC)
	if got := valueToCell(reflect.ValueOf(now), nil); got != "2026-03-07 14:15:16" {
		t.Fatalf("unexpected default time format: %#v", got)
	}
	if got := valueToCell(reflect.ValueOf(now), &fieldMeta{TimeFormat: "2006-01-02"}); got != "2026-03-07" {
		t.Fatalf("unexpected custom time format: %#v", got)
	}

	zero := time.Time{}
	if got := valueToCell(reflect.ValueOf(zero), nil); got != "" {
		t.Fatalf("expected zero time to map to empty string, got %#v", got)
	}

	type other struct{ A int }
	if got := valueToCell(reflect.ValueOf(other{A: 9}), nil); got != "{9}" {
		t.Fatalf("unexpected fallback string value: %#v", got)
	}
}

func TestStreamWriterAndWriteAPIs(t *testing.T) {
	if _, err := NewStreamWriter[writeSample](nil); err == nil {
		t.Fatalf("expected NewStreamWriter(nil) error")
	}
	if _, err := NewStreamWriter[int](&bytes.Buffer{}); err == nil {
		t.Fatalf("expected NewStreamWriter[int] to fail for non-struct type")
	}

	optVal := 99
	rows := []writeSample{
		{Code: "A", Qty: 1, When: time.Date(2026, 3, 7, 0, 0, 0, 0, time.UTC), Opt: &optVal},
		{Code: "B", Qty: 2, When: time.Date(2026, 3, 8, 0, 0, 0, 0, time.UTC)},
	}

	var buf bytes.Buffer
	sw, err := NewStreamWriter[writeSample](&buf, Sheet("Export"), Header(1), StartRow(2))
	if err != nil {
		t.Fatalf("NewStreamWriter() error: %v", err)
	}
	if err := sw.WriteRow(nil); err != nil {
		t.Fatalf("WriteRow(nil) should be no-op: %v", err)
	}
	if err := sw.WriteRows(rows); err != nil {
		t.Fatalf("WriteRows() error: %v", err)
	}
	if err := sw.Close(); err != nil {
		t.Fatalf("Close() error: %v", err)
	}
	if err := sw.Close(); err != nil {
		t.Fatalf("Close() second call error: %v", err)
	}

	f := mustOpenWorkbook(t, buf.Bytes())
	defer f.Close()
	if got := mustCell(t, f, "Export", "A1"); got != "Code" {
		t.Fatalf("unexpected header A1: %q", got)
	}
	if got := mustCell(t, f, "Export", "B1"); got != "Qty" {
		t.Fatalf("unexpected header B1: %q", got)
	}
	if got := mustCell(t, f, "Export", "C1"); got != "When" {
		t.Fatalf("unexpected header C1: %q", got)
	}
	if got := mustCell(t, f, "Export", "D1"); got != "Opt" {
		t.Fatalf("unexpected header D1: %q", got)
	}
	if got := mustCell(t, f, "Export", "A2"); got != "A" {
		t.Fatalf("unexpected row value A2: %q", got)
	}
	if got := mustCell(t, f, "Export", "D2"); got != "99" {
		t.Fatalf("unexpected row value D2: %q", got)
	}
	if got := mustCell(t, f, "Export", "D3"); got != "" {
		t.Fatalf("expected empty optional pointer in D3, got %q", got)
	}

	var out bytes.Buffer
	if err := Write(&out, rows, Sheet("Out"), Header(1), StartRow(2)); err != nil {
		t.Fatalf("Write() error: %v", err)
	}
	fo := mustOpenWorkbook(t, out.Bytes())
	defer fo.Close()
	if got := mustCell(t, fo, "Out", "A3"); got != "B" {
		t.Fatalf("unexpected Write() row value A3: %q", got)
	}

	tmpPath := filepath.Join(t.TempDir(), "written.xlsx")
	if err := WriteFile(tmpPath, rows, Sheet("FileOut"), Header(1), StartRow(2)); err != nil {
		t.Fatalf("WriteFile() error: %v", err)
	}
	ff, err := excelize.OpenFile(tmpPath)
	if err != nil {
		t.Fatalf("OpenFile(written.xlsx) error: %v", err)
	}
	defer ff.Close()
	if got := mustCell(t, ff, "FileOut", "A2"); got != "A" {
		t.Fatalf("unexpected WriteFile() value: %q", got)
	}

	var nilSW *StreamWriter[writeSample]
	if err := nilSW.Close(); err != nil {
		t.Fatalf("nil StreamWriter Close should be safe, got %v", err)
	}
	if err := nilSW.WriteRow(&rows[0]); err == nil {
		t.Fatalf("expected nil StreamWriter WriteRow to fail")
	}
}
