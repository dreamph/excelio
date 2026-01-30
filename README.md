# excelio

**The missing link between Excel and Go structs.**

Stop wrestling with cell coordinates and manual type conversions. `excelio` lets you map Excel rows directly to Go structs using simple tags—and handles millions of rows without breaking a sweat.

```go
type Product struct {
    Code   string    `excel:"Code"`
    Name   string    `col:"2"`
    Price  float64   `excelcol:"C"`
    Active bool      `excel:"Active"`
    Since  time.Time `excel:"Since" fmt:"2006-01-02"`
}

products, errs, _ := excelio.ReadFile[Product]("products.xlsx")
```

That's it. No cell references. No type casting. Just data.

---

## Installation

```bash
go get github.com/dreamph/excelio
```

---

## Why excelio?

| Problem | excelio Solution |
|---------|------------------|
| "I have a 2M row Excel file" | Stream mode—constant memory, process row by row |
| "Column positions keep changing" | Map by header name: `excel:"Product Code"` |
| "I need to validate data" | Built-in `go-playground/validator` support |
| "Users need to see what's wrong" | Write errors back into the Excel file |
| "I'm getting type conversion errors" | Automatic conversion for all common types |

---

## Three Ways to Map Columns

```go
type Product struct {
    Code  string `excel:"Code"`     // by header text
    Name  string `col:"2"`          // by column number (1-based)
    Price string `excelcol:"C"`     // by Excel letter
}
```

Mix and match as needed. Header-based mapping is most resilient to column reordering.

---

## Reading Excel Files

### Simple Read (Load All)

```go
products, rowErrs, err := excelio.ReadFile[Product](
    "products.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
)
```

### From io.Reader (HTTP Upload)

```go
func handleUpload(w http.ResponseWriter, r *http.Request) {
    file, _, _ := r.FormFile("excel")
    products, rowErrs, err := excelio.Read[Product](file)
    // process products...
}
```

### Stream Read (Millions of Rows)

Process one row at a time—memory stays flat regardless of file size.

```go
rowErrs, err := excelio.StreamFile[Product](
    "products.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
    excelio.OnStreamRow(func(rowIdx, logicalIdx int, p *Product, rowErrs []RowError) error {
        if len(rowErrs) > 0 {
            log.Printf("Row %d failed: %v", rowIdx, rowErrs)
            return nil // continue processing
        }

        // Insert to DB, send to queue, etc.
        db.Insert(p)
        return nil
    }),
)
```

---

## Writing Excel Files

### Simple Write

```go
err := excelio.WriteFile("output.xlsx", products,
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
)
```

### HTTP Response

```go
func handleExport(w http.ResponseWriter, r *http.Request) {
    w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    w.Header().Set("Content-Disposition", `attachment; filename="export.xlsx"`)
    excelio.Write(w, products)
}
```

### Stream Write (Big Data)

Generate massive files without memory pressure.

```go
sw, _ := excelio.NewStreamWriterFile[Product]("big.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
)
defer sw.Close()

for _, p := range millionsOfProducts {
    sw.WriteRow(&p)
}
```

---

## Validation

Integrates with `go-playground/validator` out of the box.

```go
type Product struct {
    Code  string  `excel:"Code"  validate:"required"`
    Price float64 `excel:"Price" validate:"required,gt=0"`
    Email string  `excel:"Email" validate:"email"`
}

validate := validator.New()
products, rowErrs, err := excelio.ReadFile[Product](
    "products.xlsx",
    excelio.UseValidator(validate),
)

// rowErrs contains field-aware validation errors
for _, e := range rowErrs {
    fmt.Printf("Row %d, Column %s (%s): %v\n",
        e.ExcelRowIndex, e.ColLetter, e.Field, e.Err)
}
```

---

## Error Write-Back

Write error messages directly into the Excel file for users to review and fix.

```go
// Read with validation
products, rowErrs, _ := excelio.ReadFile[Product]("input.xlsx")

// Write errors back to column J
excelio.WriteErrors("input.xlsx", rowErrs, excelio.ErrCol(10))
```

Or create a new file with errors:

```go
excelio.WriteErrorsTo(w, inputReader, rowErrs, excelio.ErrCol(10))
```

---

## Type Conversion

Automatic conversion for:

| Go Type | Supported Formats |
|---------|-------------------|
| `string` | As-is |
| `int`, `int8`...`int64` | Numeric strings |
| `uint`, `uint8`...`uint64` | Numeric strings |
| `float32`, `float64` | Numeric strings |
| `bool` | `true/false`, `yes/no`, `1/0`, `on/off`, `t/f`, `y/n` |
| `time.Time` | RFC3339, common formats, Excel serial dates |
| `*T` (pointers) | Empty = nil, otherwise converted |

Custom time formats via the `fmt` tag:

```go
type Record struct {
    Created time.Time `excel:"Created" fmt:"02/01/2006"`
}
```

---

## RowError Structure

Every error includes full context for debugging or user feedback:

```go
type RowError struct {
    ExcelRowIndex int    // Physical row (1-based)
    LogicalIndex  int    // Data row index (excludes header)
    ColIndex      int    // Column number (1-based)
    ColLetter     string // "A", "B", "C"...
    Field         string // Struct field name
    Column        string // Header text
    Value         string // Raw cell value
    Err           error  // The actual error
}
```

---

## Options Reference

| Option | Description |
|--------|-------------|
| `Sheet("Name")` | Select sheet by name |
| `SheetAt(0)` | Select sheet by index (0-based) |
| `Header(1)` | Header row number (1-based) |
| `StartRow(2)` | First data row (1-based) |
| `ErrCol(10)` | Column for error write-back (1-based) |
| `UseValidator(v)` | Enable go-playground/validator |
| `OnStreamRow(fn)` | Streaming row handler |

---

## Performance

- **Metadata caching** — struct tags parsed once per type
- **Streaming I/O** — process any file size with constant memory
- **Zero reflection per row** — field mapping resolved at initialization
- **Reusable row buffers** — minimal allocations during writes

---

## Credits

Built on top of:
- [excelize](https://github.com/xuri/excelize) — Excel file manipulation
- [go-playground/validator](https://github.com/go-playground/validator) — Struct validation

---

## License

MIT
