# **excelio**
🔥 Simple • Fast • Memory-Efficient Excel → Go Struct Mapping & Writing

---

## ✨ Features
- **Tag Mapping**
  ```go
  excel:"Code"      // by header
  col:"2"           // by column index
  excelcol:"C"      // by Excel letter
  ```
- **Streaming Read / Write** — process millions of rows with low memory
- **Strong Type Conversion**
  `string, int, uint, float, bool, time.Time (+ excel serial)`
- **Validation** — works with `go-playground/validator`
- **Error Tracking**
  - detailed row/column errors
  - write-back errors into Excel
- **Works with `io.Reader / io.Writer`**

---

# 🚀 Quick Start

### Struct
```go
type Product struct {
    Code   string    `excel:"Code"`
    Name   string    `col:"2"`
    Price  float64   `excelcol:"C"`
    Active bool      `excel:"Active" validate:"required"`
    Since  time.Time `excel:"Since" fmt:"2006-01-02"`
}
```

---

# Read (Normal)
Read everything into memory.

```go
rows, errs, err := excelio.ReadFile[Product](
    "products.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
)
```

Or from upload:

```go
rows, errs, err := excelio.Read[Product](file)
```

---

# Stream Read (Low Memory)
Process row-by-row.

```go
_, err := excelio.StreamFile[Product](
    "products.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),

    excelio.OnStreamRow(func(rowIdx, logicalIdx int, p *Product, rowErrs []excelio.RowError) error {
        if len(rowErrs) > 0 {
            fmt.Println("Row Error", rowErrs)
            return nil
        }

        fmt.Println("Product", p)
        return nil
    }),
)
```

---

# Write (Normal)
```go
err := excelio.WriteFile(
    "out.xlsx",
    products,
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
)
```

Write to HTTP response:
```go
excelio.Write(w, products)
```

---

# Stream Write (Big Data)
Write millions of rows efficiently.

```go
sw, _ := excelio.NewStreamWriterFile[Product](
    "big.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
)
defer sw.Close()

for _, p := range bigProducts {
    sw.WriteRow(&p)
}
```

HTTP streaming
```go
sw, _ := excelio.NewStreamWriter(w)
defer sw.Close()
sw.WriteRows(products)
```

---

# Validation (go-playground)
```go
validate := validator.New()

rows, errs, err := excelio.ReadFile[Product](
    "products.xlsx",
    excelio.UseValidator(validate),
)
```

Automatically produces field-aware RowError.

---

# Error Write-Back to Excel

```go
excelio.WriteErrors("products.xlsx", rowErrs, excelio.ErrCol(10))
```

Or create a new file output:

```go
excelio.WriteErrorsTo(w, r, rowErrs, excelio.ErrCol(10))
```

---

# 🧠 RowError Structure
```go
type RowError struct {
    ExcelRowIndex int
    LogicalIndex  int
    ColIndex      int
    ColLetter     string
    Field         string
    Column        string
    Value         string
    Err           error
}
```

---

# ⚙️ Options Summary
| Option | Description |
|--------|------------|
| `Sheet("Name")` | Select sheet |
| `SheetAt(0)` | Select sheet by index |
| `Header(1)` | Header row |
| `StartRow(2)` | First data row |
| `ErrCol(10)` | Error Column |
| `UseValidator(v)` | Enable validation |
| `OnStreamRow(fn)` | Stream handler |

---

# 🧵 Designed for Production
- Zero reflection per row after metadata cache
- Metadata caching with sync.Map
- Streaming I/O = constant RAM
- Works great with DB pipelines
- Friendly API + powerful control

---

# Credits
Powered by:
- `excelize`
- `go-playground/validator`
