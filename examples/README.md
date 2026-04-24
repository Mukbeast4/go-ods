# Examples

Runnable samples demonstrating the public API of `go-ods`. Each sub-directory is a self-contained `main` package; run any of them with:

```bash
go run ./examples/<name>
```

Each example writes its output ODS next to the current working directory (except `roundtrip`, which cleans up after itself).

| Example | Focus |
|---------|-------|
| `quickstart` | Minimal Hello World: cells, typed setters, a `SUM` formula, save |
| `formulas` | Several formula functions and `RecalcAll` |
| `styling` | `NewStyle` / `SetCellStyle` with fonts, fills, borders, rotation, `Range` |
| `autofilter` | `SetAutoFilter`, `SetFilterCriteria`, `SetSort` on a tabular dataset |
| `streaming` | `AppendRows` in batches + `NewRowIterator` to scan 10k rows |
| `roundtrip` | Open an existing file, mutate a counter, save it back |
| `images` | Embed a PNG anchored to a cell with `AddImageFromBytes` |

To open any produced file:

```bash
open quickstart.ods          # macOS
xdg-open quickstart.ods      # Linux
```
