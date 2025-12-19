# `excelio`  
### ⚡ Fast • Stream • Map Excel → Go Struct • Validate • Write Back Errors

> Production-ready Excel ingestion library for Go  
> Map Excel rows → Go structs automatically  
> Support streaming (low memory), validation, and writing error messages back into Excel

---

## ✨ Features

- 🔥 **Auto-map Excel → struct**  
  via tags:
  - `excel:"Code"` → header name  
  - `col:"2"` → column index  
  - `excelcol:"C"` → Excel column letter

- ⚡ **Streaming mode (Low RAM)**  
  process millions of rows without loading entire sheet

- 🛡️ Validation support  
  via `go-playground/validator`

- 🧠 Smart type conversion
  - `string`
  - `int / uint`
  - `float`
  - `bool`
  - `time.Time` (custom format + Excel serial dates)

- 📝 Write error messages back to Excel  
  - Modify existing file
  - Or create a new one (HTTP friendly)

- 🧩 Io.Reader / Io.Writer support  
  perfect for **HTTP Upload APIs**

---

## 🚀 Install

```sh
go get github.com/yourorg/excelio
```

---

## 🧬 Define Your Struct

```go
type Product struct {
	Code   string    `excel:"Code"   validate:"required"`
	Name   string    `col:"2"        validate:"required"`
	Price  float64   `excelcol:"C"   validate:"required,gt=0"`
	Active bool      `excel:"Active" validate:"required"`
	Since  time.Time `excel:"Since"  fmt:"2006-01-02"`
}
```

**Supported tags**

| Tag | Meaning |
|------|--------|
| `excel:"Header"` | Map by column header |
| `col:"2"` | Map by index (1-based) |
| `excelcol:"C"` | Map by Excel column letter |
| `fmt:"..."` | Custom time format |
| `validate:"..."` | go-validator rules |
| `required:"true"` | Required at mapping stage |

---

# 🟢 1️⃣ Read Entire Sheet

```go
v := validator.New()

products, errs, err := excelio.ReadFile[Product](
    "products.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
    excelio.UseValidator(v),
)
```

✔️ `products` → valid rows  
⚠️ `errs` → list of row errors

---

# 🟠 2️⃣ Streaming Mode (Ultra Fast)

Process rows **without loading entire sheet**
Perfect for big files.

```go
rowErrs, err := excelio.StreamFile[Product](
    "products.xlsx",
    excelio.Sheet("Products"),
    excelio.Header(1),
    excelio.StartRow(2),
    excelio.ErrCol(10), // put errors into column J
    excelio.OnStreamRow(func(rowIdx, logicalIdx int, p *Product, errs []excelio.RowError) error {
        if len(errs) > 0 {
            fmt.Println("❌ Row:", rowIdx, errs)
            return nil
        }

        fmt.Println("✅", *p)
        // Insert to DB here
        return nil
    }),
)
```

✔️ `rowErrs` → summary  
✔️ Automatically writes errors back to Excel file

---

# 🔵 3️⃣ HTTP Upload → Stream

Use with `io.Reader`  
Zero temp file needed

```go
rowErrs, err := excelio.Stream[Product](
    file, // multipart.File
    excelio.SheetAt(0),
    excelio.Header(1),
    excelio.StartRow(2),
    excelio.OnStreamRow(func(rowIdx, logicalIdx int, p *Product, errs []excelio.RowError) error {
        if len(errs) == 0 && p != nil {
            // process
        }
        return nil
    }),
)
```

---

# 🟥 4️⃣ Return Excel With Error Column

Perfect for APIs that validate user Excel uploads.

Upload → Validate → Return highlighted Excel

```go
excelio.WriteErrorsTo(
    w,                     // io.Writer (HTTP Response)
    bytes.NewReader(buf),  // original file
    rowErrs,
    excelio.SheetAt(0),
    excelio.ErrCol(10),
)
```

Client downloads Excel with error messages auto-filled 😎

---

## ⚙️ Options Cheat Sheet

| Option | Purpose |
|--------|--------|
| `Sheet("Name")` | Select sheet by name |
| `SheetAt(0)` | Select sheet by index |
| `Header(1)` | Header row number |
| `StartRow(2)` | First data row |
| `ErrCol(10)` | Error output column |
| `UseValidator(v)` | Enable validator |
| `OnStreamRow(fn)` | Streaming handler |

---

## 🧪 RowError Structure

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

So you can return rich responses
(JSON, logs, UI highlights etc.)

---

## ⚡ Performance Notes

- **Streaming API** uses:
  - no full sheet materialization
  - efficient reflection metadata cache
- Designed for:
  - Large corporate Excel imports
  - Financial / enterprise usage
- Memory stays small even on big sheets

---

## ❤️ Designed For Humans

- Clean minimal API  
- Zero magic hidden behavior  
- Works great in production

---

## 📌 Roadmap

- Parallel streaming mode  
- Custom converters per field  
- Nested struct support  
- Built-in Excel template generator  

---

## 🧑‍💻 Contribute

PRs welcome 🎉  
Open issues  
Discuss architecture  
Let’s build the best Excel ingestion library in Go

---

## ⭐ Final Words

If your system imports Excel,
**excelio makes it safe, fast, and developer-friendly.**

Enjoy building 🚀
