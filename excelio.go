package excelio

import (
	"fmt"
	"io"
	"reflect"
	"strconv"
	"strings"
	"sync"
	"time"

	"github.com/go-playground/validator/v10"
	"github.com/xuri/excelize/v2"
)

/*
Package excelio

High-level features:

  - Map Excel rows to Go structs using tags:
      - `excel:"Code"`    → match header text
      - `col:"2"`         → match column index (1-based)
      - `excelcol:"C"`    → match column letter
  - Type conversion for:
      - string, int*, uint*, float*, bool, time.Time (with custom format and Excel serial support)
      - Pointer types for the above
  - Validation via go-playground/validator
  - Streaming read APIs (low memory):
      - StreamFile / Stream + WithStreamRead handler
  - Error tracking:
      - RowError provides row/column/field/value/error details
      - WriteErrors: write error messages back into an existing Excel file (by path)
      - WriteErrorsTo: write a new Excel file with error messages to an io.Writer

For lowest memory usage:
  - Prefer StreamFile / Stream with WithStreamRead(...)
For simplicity:
  - Use ReadFile / Read to get a []T and a []RowError.
*/

/* =========================================================
 *  Public Types
 * ========================================================= */

// RowError represents a detailed error for a specific row/column/field.
type RowError struct {
	ExcelRowIndex int    // Physical row index in Excel (1-based)
	LogicalIndex  int    // Logical data index (1,2,3,...) after skipping header/empty rows
	ColIndex      int    // Column index (1-based)
	ColLetter     string // Column letter, e.g. "A", "B", "C"
	Field         string // Struct field name
	Column        string // Column header or configured display name
	Value         string // Raw cell value
	Err           error  // Underlying error
}

// RowHandler is the per-row callback used by streaming APIs.
// If obj == nil, row is invalid (errors in rowErrs).
// If rowErrs is non-empty, obj may still be non-nil if you choose to treat soft errors.
type RowHandler[T any] func(rowIdx, logicalIdx int, obj *T, rowErrs []RowError) error

// GenericRowHandler is an internal, type-erased handler stored in Options.
type GenericRowHandler func(rowIdx, logicalIdx int, obj any, rowErrs []RowError) error

// Option is the configuration option type for Read/Stream APIs.
type Option func(*Options)

/* =========================================================
 *  Options
 * ========================================================= */

// Options control how Excel is read and mapped.
type Options struct {
	// Sheet selection:
	SheetName  string // If empty, SheetIndex is used
	SheetIndex int    // 0-based index; used if SheetName is empty

	// Row layout:
	HeaderRow    int // Header row index (1-based). 0 = no header
	FirstDataRow int // First data row index (1-based)

	// Row index mapper:
	//   If not nil, logical index = RowIndexMapper(ExcelRowIndex, dataIdx)
	//   Otherwise, logical index = dataIdx (1-based count of non-empty data rows).
	RowIndexMapper func(excelRow int, dataIdx int) int

	// Validation:
	GoValidator *validator.Validate

	// Error column:
	//   If > 0, WriteErrors / WriteErrorsTo / StreamFile can write error messages
	//   into this 1-based column index.
	ErrorColumnIndex int

	// Internal cache:
	sheetResolved string

	// Internal streaming handler:
	streamHandler GenericRowHandler
}

// applyDefaults fills in default values for unspecified options.
func applyDefaults(o *Options) {
	if o.SheetIndex < 0 {
		o.SheetIndex = 0
	}
	// Default layout: header at row 1, data at row 2.
	if o.HeaderRow == 0 && o.FirstDataRow == 0 {
		o.HeaderRow = 1
		o.FirstDataRow = 2
	}
	// If HeaderRow is set but FirstDataRow is not, assume next row.
	if o.HeaderRow > 0 && o.FirstDataRow == 0 {
		o.FirstDataRow = o.HeaderRow + 1
	}
}

/* =========================================================
 *  Option Helpers (public API)
 * ========================================================= */

// Sheet selects a sheet by name.
func Sheet(name string) Option {
	return func(o *Options) { o.SheetName = name }
}

// SheetAt selects a sheet by index (0-based).
func SheetAt(idx int) Option {
	return func(o *Options) { o.SheetIndex = idx }
}

// Header sets the header row index (1-based).
// If FirstDataRow is not set, it's automatically set to header+1.
func Header(row int) Option {
	return func(o *Options) {
		o.HeaderRow = row
		if row > 0 && o.FirstDataRow == 0 {
			o.FirstDataRow = row + 1
		}
	}
}

// StartRow sets the first data row index (1-based).
func StartRow(row int) Option {
	return func(o *Options) { o.FirstDataRow = row }
}

// ErrCol sets the 1-based error column index.
func ErrCol(idx int) Option {
	return func(o *Options) { o.ErrorColumnIndex = idx }
}

// UseValidator sets the go-playground/validator instance used for struct validation.
func UseValidator(v *validator.Validate) Option {
	return func(o *Options) { o.GoValidator = v }
}

// OnStreamRow registers a per-row handler for Stream / StreamFile.
// This is required for streaming APIs; if omitted, Stream/StreamFile will return an error.
func OnStreamRow[T any](h RowHandler[T]) Option {
	return func(o *Options) {
		if h == nil {
			return
		}
		o.streamHandler = func(rowIdx, logicalIdx int, obj any, rowErrs []RowError) error {
			var p *T
			if obj != nil {
				if cast, ok := obj.(*T); ok {
					p = cast
				} else if val, ok := obj.(T); ok {
					p = &val
				}
			}
			return h(rowIdx, logicalIdx, p, rowErrs)
		}
	}
}

/* =========================================================
 *  Type Metadata & Tags
 * ========================================================= */

// fieldMeta stores mapping info for a single struct field.
type fieldMeta struct {
	Index        []int
	FieldName    string
	ColumnNames  []string // From `excel:"Code,Name,..."`
	ColIndexTag  int      // From `col:"2"` (0-based). -1 = none
	ColLetterTag string   // From `excelcol:"C"` (normalized uppercase)

	Required   bool   // From tag `required:"true"` or `required:"1"`
	TimeFormat string // From tag `fmt:"2006-01-02"`

	// lastColIndex is the last column index used for this field in the current row.
	// Used for mapping validator errors back to the column.
	lastColIndex int
}

// typeMeta stores metadata for a struct type.
type typeMeta struct {
	Fields        []*fieldMeta
	HeaderToField map[string]*fieldMeta // header text (lowercased) -> field
	FieldByName   map[string]*fieldMeta // struct field name -> field
}

var metaCache sync.Map // map[reflect.Type]*typeMeta

// splitAndTrim splits a comma-separated string and trims each part.
func splitAndTrim(s string) []string {
	if s == "" {
		return nil
	}
	parts := strings.Split(s, ",")
	out := make([]string, 0, len(parts))
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p != "" {
			out = append(out, p)
		}
	}
	return out
}

// getTypeMeta builds and caches metadata for a struct type T.
func getTypeMeta(t reflect.Type) (*typeMeta, error) {
	if v, ok := metaCache.Load(t); ok {
		return v.(*typeMeta), nil
	}

	if t.Kind() == reflect.Ptr {
		t = t.Elem()
	}
	if t.Kind() != reflect.Struct {
		return nil, fmt.Errorf("excelio: type %s is not struct", t.String())
	}

	m := &typeMeta{
		HeaderToField: make(map[string]*fieldMeta),
		FieldByName:   make(map[string]*fieldMeta),
	}

	numField := t.NumField()
	for i := 0; i < numField; i++ {
		f := t.Field(i)
		// Skip unexported fields.
		if f.PkgPath != "" {
			continue
		}

		excelTag := f.Tag.Get("excel")
		colTag := f.Tag.Get("col")
		excelColTag := f.Tag.Get("excelcol")

		// Only consider fields that have at least one mapping tag.
		if excelTag == "" && colTag == "" && excelColTag == "" {
			continue
		}

		fm := &fieldMeta{
			Index:       f.Index,
			FieldName:   f.Name,
			ColumnNames: splitAndTrim(excelTag),
			Required:    f.Tag.Get("required") == "1" || strings.ToLower(f.Tag.Get("required")) == "true",
			TimeFormat:  f.Tag.Get("fmt"),
			ColIndexTag: -1,
		}

		// Header-based mapping.
		for _, name := range fm.ColumnNames {
			key := strings.ToLower(name)
			m.HeaderToField[key] = fm
		}

		// Index-based mapping: col:"2"
		if colTag != "" {
			if n, err := strconv.Atoi(strings.TrimSpace(colTag)); err == nil && n > 0 {
				fm.ColIndexTag = n - 1 // internal 0-based
			}
		}

		// Letter-based mapping: excelcol:"C"
		if excelColTag != "" {
			fm.ColLetterTag = strings.ToUpper(strings.TrimSpace(excelColTag))
		}

		m.Fields = append(m.Fields, fm)
		m.FieldByName[f.Name] = fm
	}

	metaCache.Store(t, m)
	return m, nil
}

// FindFieldByName returns the fieldMeta for a given struct field name.
func (m *typeMeta) FindFieldByName(name string) *fieldMeta {
	if m.FieldByName == nil {
		return nil
	}
	return m.FieldByName[name]
}

/* =========================================================
 *  Column Helpers
 * ========================================================= */

// colLetter converts a 0-based column index to an Excel column letter ("A", "B", ...).
func colLetter(idx int) string {
	n := idx + 1
	if n <= 0 {
		return ""
	}
	var result []byte
	for n > 0 {
		n--
		r := 'A' + (n % 26)
		result = append([]byte{byte(r)}, result...)
		n /= 26
	}
	return string(result)
}

// colIndexFromLetter converts a column letter ("A", "B", ...) to a 0-based index.
func colIndexFromLetter(s string) int {
	s = strings.TrimSpace(strings.ToUpper(s))
	if s == "" {
		return -1
	}
	n := 0
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c < 'A' || c > 'Z' {
			return -1
		}
		n = n*26 + int(c-'A'+1)
	}
	return n - 1
}

/* =========================================================
 *  Header Helpers
 * ========================================================= */

// parseHeader reads the header row and returns a map[columnIndex]headerText.
func parseHeader(f *excelize.File, sheet string, headerRow int) (map[int]string, error) {
	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	curRow := 0
	for rows.Next() {
		curRow++
		cols, err := rows.Columns()
		if err != nil {
			return nil, err
		}
		if curRow == headerRow {
			m := make(map[int]string, len(cols))
			for i, c := range cols {
				m[i] = strings.TrimSpace(c)
			}
			return m, nil
		}
	}

	return nil, fmt.Errorf("excelio: header row %d not found", headerRow)
}

/* =========================================================
 *  Type Conversion
 * ========================================================= */

// parseBool converts various common boolean strings into bool.
func parseBool(raw string) (bool, error) {
	s := strings.TrimSpace(strings.ToLower(raw))
	switch s {
	case "1", "true", "t", "yes", "y", "on":
		return true, nil
	case "0", "false", "f", "no", "n", "off":
		return false, nil
	}
	return false, fmt.Errorf("invalid bool: %q", raw)
}

// excelSerialToTime converts an Excel serial date (1900 date system) to time.Time (UTC).
// Note: This uses the common "1899-12-30" base to match Excel's 1900 system behavior.
func excelSerialToTime(serial float64) (time.Time, error) {
	if serial <= 0 {
		return time.Time{}, fmt.Errorf("invalid excel serial: %f", serial)
	}
	const secondsInDay = 24 * 60 * 60

	days := int64(serial)
	frac := serial - float64(days)

	base := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	t := base.AddDate(0, 0, int(days))

	sec := int64(frac*secondsInDay + 0.5)
	t = t.Add(time.Duration(sec) * time.Second)

	return t, nil
}

// parseTime attempts to parse a time value from the cell text.
// It tries in this order:
//  1. Custom format from fieldMeta.TimeFormat
//  2. RFC3339
//  3. Several common date/time layouts
//  4. Excel serial number
func parseTime(raw string, fm *fieldMeta) (time.Time, error) {
	s := strings.TrimSpace(raw)
	if s == "" {
		return time.Time{}, fmt.Errorf("empty time")
	}

	// 1. Custom format.
	if fm != nil && fm.TimeFormat != "" {
		if t, err := time.Parse(fm.TimeFormat, s); err == nil {
			return t, nil
		}
	}

	// 2. RFC3339.
	if t, err := time.Parse(time.RFC3339, s); err == nil {
		return t, nil
	}

	// 3. Common formats.
	layouts := []string{
		"2006-01-02",
		"02/01/2006",
		"02-01-2006",
		"2006/01/02",
		"02/01/2006 15:04",
		"2006-01-02 15:04",
		"02-01-2006 15:04",
	}
	for _, layout := range layouts {
		if t, err := time.Parse(layout, s); err == nil {
			return t, nil
		}
	}

	// 4. Excel serial number.
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		if t, err2 := excelSerialToTime(f); err2 == nil {
			return t, nil
		}
	}

	return time.Time{}, fmt.Errorf("cannot parse time: %q", raw)
}

// setFieldValue sets a field value from a raw string, handling pointer and non-pointer types.
func setFieldValue(field reflect.Value, fm *fieldMeta, raw string) error {
	// Handle pointer types: if value is empty, keep nil; otherwise allocate and set.
	if field.Kind() == reflect.Ptr {
		trim := strings.TrimSpace(raw)
		if trim == "" {
			return nil
		}
		elem := reflect.New(field.Type().Elem()).Elem()
		if err := convertAndSet(elem, fm, raw); err != nil {
			return err
		}
		field.Set(elem.Addr())
		return nil
	}
	return convertAndSet(field, fm, raw)
}

// convertAndSet performs conversion for the underlying concrete kind.
func convertAndSet(field reflect.Value, fm *fieldMeta, raw string) error {
	trim := strings.TrimSpace(raw)

	switch field.Kind() {
	case reflect.String:
		field.SetString(raw)
		return nil

	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
		i, err := strconv.ParseInt(trim, 10, 64)
		if err != nil {
			return err
		}
		field.SetInt(i)
		return nil

	case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		u, err := strconv.ParseUint(trim, 10, 64)
		if err != nil {
			return err
		}
		field.SetUint(u)
		return nil

	case reflect.Float32, reflect.Float64:
		f, err := strconv.ParseFloat(trim, 64)
		if err != nil {
			return err
		}
		field.SetFloat(f)
		return nil

	case reflect.Bool:
		b, err := parseBool(trim)
		if err != nil {
			return err
		}
		field.SetBool(b)
		return nil

	case reflect.Struct:
		if field.Type() == reflect.TypeOf(time.Time{}) {
			tm, err := parseTime(raw, fm)
			if err != nil {
				return err
			}
			field.Set(reflect.ValueOf(tm))
			return nil
		}
	}

	// If you want to be more permissive, you could skip unsupported kinds instead.
	return fmt.Errorf("unsupported kind %s for value %q", field.Kind(), raw)
}

/* =========================================================
 *  Row Mapping
 * ========================================================= */

// buildFieldColIndex resolves the final column index for each fieldMeta,
// combining header-based, index-based (`col`) and letter-based (`excelcol`) mapping.
func buildFieldColIndex(meta *typeMeta, headerIndex map[string]int) map[*fieldMeta]int {
	fieldColIndex := make(map[*fieldMeta]int, len(meta.Fields))
	for _, fm := range meta.Fields {
		// 1. Explicit index: col:"2".
		if fm.ColIndexTag >= 0 {
			fieldColIndex[fm] = fm.ColIndexTag
			continue
		}

		// 2. Letters: excelcol:"C".
		if fm.ColLetterTag != "" {
			if idx := colIndexFromLetter(fm.ColLetterTag); idx >= 0 {
				fieldColIndex[fm] = idx
				continue
			}
		}

		// 3. Header-based: excel:"Code,Name"
		if len(fm.ColumnNames) > 0 && len(headerIndex) > 0 {
			for _, name := range fm.ColumnNames {
				if idx, ok := headerIndex[strings.ToLower(strings.TrimSpace(name))]; ok {
					fieldColIndex[fm] = idx
					break
				}
			}
		}
	}
	return fieldColIndex
}

// buildRowError creates a RowError populated with row/column information.
func buildRowError(rowIdx, logicalIdx int, fm *fieldMeta, colIdx int, headerMap map[int]string, cols []string, err error) RowError {
	var raw, colName, colLet string
	if colIdx >= 0 && colIdx < len(cols) {
		raw = cols[colIdx]
		colLet = colLetter(colIdx)
		if headerMap != nil {
			if h, ok := headerMap[colIdx]; ok {
				colName = h
			}
		}
	}
	if colName == "" && fm != nil && len(fm.ColumnNames) > 0 {
		colName = fm.ColumnNames[0]
	}

	fieldName := ""
	if fm != nil {
		fieldName = fm.FieldName
	}

	return RowError{
		ExcelRowIndex: rowIdx,
		LogicalIndex:  logicalIdx,
		ColIndex:      colIdx + 1,
		ColLetter:     colLet,
		Field:         fieldName,
		Column:        colName,
		Value:         raw,
		Err:           err,
	}
}

// mapRow maps a single row (slice of cell values) into a struct T,
// returning the object, the row's errors, and whether it is valid.
func mapRow[T any](
	t reflect.Type,
	meta *typeMeta,
	fieldColIndex map[*fieldMeta]int,
	headerMap map[int]string,
	o *Options,
	rowIdx, logicalIdx int,
	cols []string,
) (T, []RowError, bool) {
	var zero T
	v := reflect.New(t).Elem()
	rowHasError := false
	var rowErrs []RowError

	// Map raw values into struct fields.
	for _, fm := range meta.Fields {
		colIdx, ok := fieldColIndex[fm]
		if !ok {
			continue
		}
		fm.lastColIndex = colIdx

		// Column out of range.
		if colIdx < 0 || colIdx >= len(cols) {
			if fm.Required {
				rowHasError = true
				rowErrs = append(rowErrs, buildRowError(
					rowIdx, logicalIdx, fm, colIdx, headerMap, cols,
					fmt.Errorf("required column out of range"),
				))
			}
			continue
		}

		raw := cols[colIdx]
		trim := strings.TrimSpace(raw)

		// Empty value.
		if trim == "" {
			if fm.Required {
				rowHasError = true
				rowErrs = append(rowErrs, buildRowError(
					rowIdx, logicalIdx, fm, colIdx, headerMap, cols,
					fmt.Errorf("required value is empty"),
				))
			}
			continue
		}

		field := v.FieldByIndex(fm.Index)
		if !field.CanSet() {
			continue
		}

		if err := setFieldValue(field, fm, raw); err != nil {
			rowHasError = true
			rowErrs = append(rowErrs, buildRowError(
				rowIdx, logicalIdx, fm, colIdx, headerMap, cols, err,
			))
		}
	}

	obj := v.Interface().(T)

	// Struct-level validation using go-playground/validator (if configured).
	if o.GoValidator != nil {
		if e := o.GoValidator.Struct(obj); e != nil {
			if verrs, ok := e.(validator.ValidationErrors); ok {
				for _, fe := range verrs {
					rowHasError = true
					fm := meta.FindFieldByName(fe.StructField())
					colIdx := -1
					if fm != nil {
						colIdx = fm.lastColIndex
					}

					displayName := fe.Field()
					if fm != nil && len(fm.ColumnNames) > 0 {
						displayName = fm.ColumnNames[0]
					}

					rowErrs = append(rowErrs, buildRowError(
						rowIdx, logicalIdx, fm, colIdx, headerMap, cols,
						fmt.Errorf("column '%s' failed on '%s': %s",
							displayName, fe.Tag(), fe.Error()),
					))
				}
			} else {
				rowHasError = true
				rowErrs = append(rowErrs, RowError{
					ExcelRowIndex: rowIdx,
					LogicalIndex:  logicalIdx,
					Err:           fmt.Errorf("struct validation error: %w", e),
				})
			}
		}
	}

	if rowHasError {
		return zero, rowErrs, false
	}
	return obj, nil, true
}

/* =========================================================
 *  Sheet Resolve
 * ========================================================= */

// resolveSheet resolves the sheet name to use, based on SheetName or SheetIndex.
func resolveSheet(f *excelize.File, o *Options) (string, error) {
	if o.sheetResolved != "" {
		return o.sheetResolved, nil
	}
	sheet := o.SheetName
	if sheet == "" {
		sheets := f.GetSheetList()
		if len(sheets) == 0 {
			return "", fmt.Errorf("excelio: no sheets")
		}
		if o.SheetIndex < 0 || o.SheetIndex >= len(sheets) {
			return "", fmt.Errorf("excelio: sheet index %d out of range", o.SheetIndex)
		}
		sheet = sheets[o.SheetIndex]
	}
	o.sheetResolved = sheet
	return sheet, nil
}

/* =========================================================
 *  Core: read/stream from *excelize.File
 * ========================================================= */

// isRowEmpty checks whether all cells in a row are empty (after trimming).
func isRowEmpty(cols []string) bool {
	for _, c := range cols {
		if strings.TrimSpace(c) != "" {
			return false
		}
	}
	return true
}

// readFromExcelFile implements the core "read everything into slice" logic.
func readFromExcelFile[T any](f *excelize.File, o *Options) ([]T, []RowError, error) {
	sheet, err := resolveSheet(f, o)
	if err != nil {
		return nil, nil, err
	}

	var headerMap map[int]string
	headerIndex := make(map[string]int)

	// Build header index if header row is configured.
	if o.HeaderRow > 0 {
		headerMap, err = parseHeader(f, sheet, o.HeaderRow)
		if err != nil {
			return nil, nil, err
		}
		for idx, name := range headerMap {
			n := strings.ToLower(strings.TrimSpace(name))
			if n != "" {
				headerIndex[n] = idx
			}
		}
	}

	t := reflect.TypeOf((*T)(nil)).Elem()
	meta, err := getTypeMeta(t)
	if err != nil {
		return nil, nil, err
	}

	fieldColIndex := buildFieldColIndex(meta, headerIndex)

	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, nil, err
	}
	defer rows.Close()

	var result []T
	var errs []RowError
	rowIdx := 0
	dataIdx := 0

	for rows.Next() {
		rowIdx++
		cols, err := rows.Columns()
		if err != nil {
			errs = append(errs, RowError{
				ExcelRowIndex: rowIdx,
				Err:           fmt.Errorf("read row: %w", err),
			})
			continue
		}

		if rowIdx < o.FirstDataRow {
			continue
		}
		if isRowEmpty(cols) {
			continue
		}

		dataIdx++
		logicalIdx := dataIdx
		if o.RowIndexMapper != nil {
			logicalIdx = o.RowIndexMapper(rowIdx, dataIdx)
		}

		obj, rowErrs, ok := mapRow[T](t, meta, fieldColIndex, headerMap, o, rowIdx, logicalIdx, cols)
		if len(rowErrs) > 0 {
			errs = append(errs, rowErrs...)
		}
		if ok {
			result = append(result, obj)
		}
	}

	return result, errs, nil
}

// streamFromExcelFile implements the core streaming logic using Options.streamHandler.
func streamFromExcelFile[T any](f *excelize.File, o *Options) ([]RowError, error) {
	if o.streamHandler == nil {
		return nil, fmt.Errorf("excelio: WithStreamRead() is required for Stream/StreamFile")
	}

	sheet, err := resolveSheet(f, o)
	if err != nil {
		return nil, err
	}

	var headerMap map[int]string
	headerIndex := make(map[string]int)

	// Build header index if header row is configured.
	if o.HeaderRow > 0 {
		headerMap, err = parseHeader(f, sheet, o.HeaderRow)
		if err != nil {
			return nil, err
		}
		for idx, name := range headerMap {
			n := strings.ToLower(strings.TrimSpace(name))
			if n != "" {
				headerIndex[n] = idx
			}
		}
	}

	t := reflect.TypeOf((*T)(nil)).Elem()
	meta, err := getTypeMeta(t)
	if err != nil {
		return nil, err
	}

	fieldColIndex := buildFieldColIndex(meta, headerIndex)

	rows, err := f.Rows(sheet)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	rowIdx := 0
	dataIdx := 0

	var allErrs []RowError
	var fatalErr error

	for rows.Next() {
		rowIdx++
		cols, err := rows.Columns()
		if err != nil {
			// Row read error: still pass to handler for logging/use.
			re := RowError{
				ExcelRowIndex: rowIdx,
				Err:           fmt.Errorf("read row: %w", err),
			}
			allErrs = append(allErrs, re)
			if hErr := o.streamHandler(rowIdx, -1, nil, []RowError{re}); hErr != nil {
				fatalErr = hErr
				break
			}
			continue
		}

		if rowIdx < o.FirstDataRow {
			continue
		}
		if isRowEmpty(cols) {
			continue
		}

		dataIdx++
		logicalIdx := dataIdx
		if o.RowIndexMapper != nil {
			logicalIdx = o.RowIndexMapper(rowIdx, dataIdx)
		}

		obj, rowErrs, ok := mapRow[T](t, meta, fieldColIndex, headerMap, o, rowIdx, logicalIdx, cols)
		if len(rowErrs) > 0 {
			allErrs = append(allErrs, rowErrs...)
		}

		var objAny any
		if ok {
			objCopy := obj // ensure address is stable
			objAny = &objCopy
		}

		if hErr := o.streamHandler(rowIdx, logicalIdx, objAny, rowErrs); hErr != nil {
			fatalErr = hErr
			break
		}
	}

	if fatalErr != nil {
		return allErrs, fatalErr
	}
	return allErrs, nil
}

/* =========================================================
 *  Public API: Read / Stream
 * ========================================================= */

// ReadFile reads an Excel file from a file path and returns:
//   - a slice of successfully mapped objects
//   - a slice of RowError for all rows with issues
func ReadFile[T any](path string, opts ...Option) ([]T, []RowError, error) {
	var o Options
	for _, opt := range opts {
		opt(&o)
	}
	applyDefaults(&o)

	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, nil, err
	}
	defer f.Close()

	return readFromExcelFile[T](f, &o)
}

// Read reads an Excel file from an io.Reader (e.g. HTTP upload, memory buffer)
// and returns:
//   - a slice of successfully mapped objects
//   - a slice of RowError for all rows with issues
func Read[T any](r io.Reader, opts ...Option) ([]T, []RowError, error) {
	var o Options
	for _, opt := range opts {
		opt(&o)
	}
	applyDefaults(&o)

	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, nil, err
	}
	defer f.Close()

	return readFromExcelFile[T](f, &o)
}

// StreamFile streams an Excel file from a file path, calling the handler
// supplied via WithStreamRead(...) for each non-empty data row.
// It returns a slice of RowError for all rows with issues.
// If ErrCol(...) is set and there are errors, it will also write error messages
// back into the original file in the specified column.
func StreamFile[T any](path string, opts ...Option) ([]RowError, error) {
	var o Options
	for _, opt := range opts {
		opt(&o)
	}
	applyDefaults(&o)

	if o.streamHandler == nil {
		return nil, fmt.Errorf("excelio: WithStreamRead() is required for StreamFile")
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	allErrs, err := streamFromExcelFile[T](f, &o)
	if err != nil {
		return allErrs, err
	}

	// Auto-write errors back to the file if configured.
	if o.ErrorColumnIndex > 0 && len(allErrs) > 0 {
		if err := WriteErrors(path, allErrs, opts...); err != nil {
			return allErrs, err
		}
	}

	return allErrs, nil
}

// Stream streams an Excel file from an io.Reader, calling the handler supplied
// via WithStreamRead(...) for each non-empty data row.
// It returns a slice of RowError for all rows with issues.
// This variant does not modify the original source (no path), but you can
// later call WriteErrorsTo(...) if you want to produce a new file with errors.
func Stream[T any](r io.Reader, opts ...Option) ([]RowError, error) {
	var o Options
	for _, opt := range opts {
		opt(&o)
	}
	applyDefaults(&o)

	if o.streamHandler == nil {
		return nil, fmt.Errorf("excelio: WithStreamRead() is required for Stream")
	}

	f, err := excelize.OpenReader(r)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	return streamFromExcelFile[T](f, &o)
}

/* =========================================================
 *  WriteErrors: path & io.Writer versions
 * ========================================================= */

// WriteErrors writes error messages into an existing Excel file identified by path.
// It uses ErrCol(...) to determine which column to write to.
func WriteErrors(path string, errs []RowError, opts ...Option) error {
	if len(errs) == 0 {
		return nil
	}

	var o Options
	for _, opt := range opts {
		opt(&o)
	}
	if o.ErrorColumnIndex <= 0 {
		return fmt.Errorf("excelio: ErrCol() / ErrorColumnIndex must be > 0 for WriteErrors")
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		return err
	}
	defer f.Close()

	return writeErrorsToExcelFile(f, errs, &o, nil)
}

// WriteErrorsTo writes error messages into a copy of the Excel file read from r,
// and writes the resulting file to w. This is useful for HTTP responses or
// streaming to cloud storage without touching the original file.
//
// If errs is empty, it simply copies the input stream r to w.
func WriteErrorsTo(w io.Writer, r io.Reader, errs []RowError, opts ...Option) error {
	if len(errs) == 0 {
		_, err := io.Copy(w, r)
		return err
	}

	var o Options
	for _, opt := range opts {
		opt(&o)
	}
	if o.ErrorColumnIndex <= 0 {
		return fmt.Errorf("excelio: ErrCol() / ErrorColumnIndex must be > 0 for WriteErrorsTo")
	}

	f, err := excelize.OpenReader(r)
	if err != nil {
		return err
	}
	defer f.Close()

	return writeErrorsToExcelFile(f, errs, &o, w)
}

// writeErrorsToExcelFile writes the provided RowError list into the given
// *excelize.File f, using the sheet from Options and error column from Options.
// If w == nil, it saves the file to disk (f.Save()).
// If w != nil, it writes the Excel content to w (f.Write(w)).
func writeErrorsToExcelFile(f *excelize.File, errs []RowError, o *Options, w io.Writer) error {
	sheet, err := resolveSheet(f, o)
	if err != nil {
		return err
	}

	errColIdx := o.ErrorColumnIndex - 1
	errColLetter := colLetter(errColIdx)

	for _, re := range errs {
		if re.ExcelRowIndex <= 0 {
			continue
		}
		cell := fmt.Sprintf("%s%d", errColLetter, re.ExcelRowIndex)
		oldVal, _ := f.GetCellValue(sheet, cell)
		msg := re.Err.Error()
		if oldVal != "" {
			msg = oldVal + "\n" + msg
		}
		if setErr := f.SetCellValue(sheet, cell, msg); setErr != nil {
			return setErr
		}
	}

	if w != nil {
		return f.Write(w)
	}
	return f.Save()
}
