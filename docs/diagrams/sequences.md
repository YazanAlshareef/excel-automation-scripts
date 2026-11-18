# Sequences

Two key flows: merging workbooks, and validating columns.

## Merge workbooks

Command line run of `merge_workbooks.py`: input checks, the output-equals-input guard, then merge and write.

Sources: [scripts/merge_workbooks.py](../../scripts/merge_workbooks.py), [scripts/_common.py](../../scripts/_common.py)

```mermaid
sequenceDiagram
    actor User
    participant Main as main()
    participant Common as existing_file
    participant Merge as merge()
    participant PD as pandas

    User->>Main: inputs and -o output
    Main->>Common: check each input (argparse type)
    Common-->>Main: path, or error (exit 2)
    Main->>Main: output same as an input?
    alt same path
        Main-->>User: error on stderr, return 2
    else different path
        Main->>Merge: merge(paths)
        loop each input file
            Merge->>PD: read_excel(path)
            PD-->>Merge: DataFrame
            Merge->>Merge: add source_file column
        end
        Merge->>PD: concat(frames, ignore_index=True)
        PD-->>Merge: merged DataFrame
        Merge-->>Main: merged DataFrame
        Main->>PD: to_excel(output, index=False)
        Main-->>User: row count message, return 0
    end
```

## Validate columns

Command line run of `validate_columns.py`: how the exit code is chosen.

Sources: [scripts/validate_columns.py](../../scripts/validate_columns.py), [scripts/_common.py](../../scripts/_common.py)

```mermaid
sequenceDiagram
    actor User
    participant Main as main()
    participant Common as existing_file
    participant PD as pandas
    participant Val as validate()

    User->>Main: input and --required columns
    Main->>Common: check input (argparse type)
    Common-->>Main: path, or error (exit 2)
    Main->>PD: read_excel(input)
    PD-->>Main: DataFrame
    Main->>Val: validate(df, required)
    Val->>Val: report missing columns
    Val->>Val: report blank or whitespace-only cells with Excel row numbers
    Val-->>Main: list of problems
    alt problems found
        Main-->>User: INVALID and the list, return 1
    else no problems
        Main-->>User: OK message, return 0
    end
```
