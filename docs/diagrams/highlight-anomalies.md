# Highlight anomalies logic

The decision steps in `highlight_anomalies.py`, from argument checks to the saved workbook.

Sources: [scripts/highlight_anomalies.py](../../scripts/highlight_anomalies.py), [scripts/_common.py](../../scripts/_common.py)

```mermaid
flowchart TD
    A["main(): parse args"] --> B{"output same as input?"}
    B -- yes --> E2["error, return 2"]
    B -- no --> C["pd.read_excel(input)"]
    C --> D{"column in df.columns?"}
    D -- no --> E2
    D -- yes --> F["pd.to_numeric(column, errors='coerce')"]
    F --> G{"any numeric values?"}
    G -- no --> E2
    G -- yes --> H["find_outliers(nums, k): outside Q1 - k*IQR and Q3 + k*IQR"]
    H --> I["load_workbook(input), first sheet"]
    I --> J["fill each outlier cell yellow (row + 2)"]
    J --> K["wb.save(output)"]
    K --> L["print count and Excel rows, return 0"]
```
