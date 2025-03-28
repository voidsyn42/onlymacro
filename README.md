# onlymacro

🧠 Small tool to convert simple VBA macros into JSON-compatible structures for OnlyOffice macros.

> ⚠️ For educational and compatibility testing only.

## Features

- Extracts logic from minimal `Sub` / `Function` blocks in VBA
- Wraps logic in OnlyOffice-style JSON macros (uses `onDocumentOpen`)
- NOT a full VBA parser!

## Usage

```bash
python3 onlymacro.py examples/macro_vba.txt
```