# Components

Which scripts, helpers, libraries and files make up the repo.

Sources: [scripts/_common.py](../../scripts/_common.py), [scripts/merge_workbooks.py](../../scripts/merge_workbooks.py), [scripts/dedupe_rows.py](../../scripts/dedupe_rows.py), [scripts/validate_columns.py](../../scripts/validate_columns.py), [scripts/pivot_summary.py](../../scripts/pivot_summary.py), [scripts/highlight_anomalies.py](../../scripts/highlight_anomalies.py), [scripts/make_samples.py](../../scripts/make_samples.py), [tests/test_scripts.py](../../tests/test_scripts.py)

```mermaid
flowchart LR
    subgraph files["Excel files"]
        IN["sample_input/*.xlsx"]
        OUT["sample_output/*.xlsx"]
    end

    subgraph scripts["scripts/"]
        MS["make_samples.py"]
        MW["merge_workbooks.py"]
        DR["dedupe_rows.py"]
        VC["validate_columns.py"]
        PS["pivot_summary.py"]
        HA["highlight_anomalies.py"]
        CM["_common.py: existing_file"]
    end

    subgraph libs["Libraries"]
        PD["pandas"]
        OX["openpyxl"]
    end

    T["tests/test_scripts.py (pytest)"]

    MS -->|writes| IN
    IN --> MW
    MW --> OUT
    OUT --> DR
    OUT --> VC
    DR --> OUT
    OUT --> PS
    OUT --> HA
    PS --> OUT
    HA --> OUT

    MW -.->|uses| CM
    DR -.->|uses| CM
    VC -.->|uses| CM
    PS -.->|uses| CM
    HA -.->|uses| CM

    MW -.-> PD
    DR -.-> PD
    VC -.-> PD
    PS -.-> PD
    HA -.-> PD
    HA -.-> OX
    MS -.-> PD

    T -.->|imports and runs| scripts
```
