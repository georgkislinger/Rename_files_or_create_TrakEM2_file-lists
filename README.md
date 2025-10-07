# TrakEM2 Helper GUI — 2025-10 Flexible Coordinates

A PowerShell GUI utility designed to simplify and standardize file handling for **TrakEM2** projects.  
Supports both **filename renaming** and **TSV (import list) generation** for complex multi-dimensional datasets.

---

## Features

- **Flexible Coordinate Extraction**  
  Handles any coordinate combination — `x`, `y`, `z`, `xy`, `xz`, `yz`, `xyz` — automatically filling missing coordinates with `0`.

- **Custom Template Support**  
  Fully customizable filename templates such as  
  `Tile_#z#-#x#___#y#-#r#` or `#x#_#z#`.  
  Includes automatic guessing of common vendor formats (Zeiss, Thermo).

- **Automatic Tile Dimension Detection**  
  Detects image width and height from the first image in the dataset.

- **Two Main Modes**  
  1. **Rename Mode:** Batch rename files with optional dry-run.  
  2. **Import Mode:** Generate TrakEM2-compatible TSV or copy matched files.

- **Automatic Guessing**  
  - Guess coordinate order from numeric sequences (`z`, `xz`, `xyz`, etc.).  
  - Detect custom patterns from filenames automatically.

- **Path Formatting Options**  
  Output paths can be **relative**, **full**, or **filename-only**, with optional inclusion of base folder.

- **Dark-Themed GUI**  
  Modern WPF interface with responsive layout and interactive controls.

---

## Requirements

- **Windows PowerShell 5.1 or later**
- **.NET Framework 4.7.2 or later**
- **Image formats supported by System.Drawing** (e.g. `.tif`, `.png`, `.jpg`)

---

## Installation

1. Save the script as  
   ```text
   TrakEM2_Helper_GUI.ps1
   ```

2. Unblock the file if downloaded from the internet:
   ```powershell
   Unblock-File .\TrakEM2_Helper_GUI.ps1
   ```

3. Run PowerShell **as Administrator** (recommended).

4. Launch the GUI:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\TrakEM2_Helper_GUI.ps1
   ```
5. Alternatively download and use the executable or compile it yourself with ps2exe
---

## Usage Overview

### 1. **Select Data Folder**
   - Click **Browse…** to choose your image directory.
   - Optionally specify a default extension (e.g., `tif`).

### 2. **Choose Mode**
   - **Rename files:** Prefix filenames based on folder hierarchy.  
   - **Create TrakEM2 Import TSV / Copy:** Build import tables or copy matched files.

---

### 3. **Rename Mode**
   - Choose the folder **level** (`parent`, `grandparent`, or `greatgrand`) used as prefix.
   - Set the extension and enable **Dry Run** to preview renames.
   - Click **Run Rename** to execute.

---

### 4. **Import Mode**
   - Enter tile parameters (`TileSizeX`, `TileSizeY`, `OverlapX`, `OverlapY`).
   - Choose bit depth (8-bit or 16-bit) and Z offset.
   - Select a template type:
     - **Zeiss** (e.g., `Tile_r###-c###_S_##`)
     - **Thermo** (e.g., `Tile_###-###-###_...`)
     - **Sequential Numbers** — use numeric order like `z`, `xz`, `xyz`
     - **Custom Pattern** — define your own format (`#x#`, `#y#`, `#z#`)
   - Use **Auto**, **Guess Order**, or **Guess Pattern** to detect parameters automatically.
   - Click **Create TSV** to write the TrakEM2 import list.
   - Click **Copy Files** to copy matched images into a new subfolder.

---

## Output Files

| Type | Description |
|------|--------------|
| **RawImgList.txt** | Tab-separated import table for TrakEM2 |
| **copied_rawdata/** | Optional copy of matched files |
| **Renamed files** | Updated filenames (if Rename mode used) |

---

## Notes

- Missing coordinates (`x`, `y`, or `z`) are automatically set to `0`.
- Supports any filename pattern with embedded numeric coordinates.
- The **Log / Preview** panel shows detailed operation feedback.
- Overwrites are prevented — duplicate TSVs and folders are auto-numbered.

---

## Example Custom Templates

| Example Pattern | Description |
|------------------|-------------|
| `#z#` | Simple Z-indexed stack |
| `Tile_#z#-#x#___#y#-#r#` | Flexible multi-coordinate pattern |
| `#x#_#z#` | Two-dimensional (XZ) pattern |
| `Tile_r#y#-c#x#_S_#z#_#r#` | Zeiss format |
| `Tile_#y#-#x#-#r#_0-000.s#z#_e00` | Thermo format |

---

## Version

**Release:** 2025-10  
**Author:** Georg Kislinger  
**License:** MIT  

---

## Troubleshooting

- If the GUI does not open, ensure PowerShell execution policy allows script execution:
  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  ```
- If image size detection fails, check that your images are supported by `.NET System.Drawing`.
- Log output can be copied directly from the preview window for debugging.

---

## Changelog

### v2025-10
- Added support for **arbitrary coordinate combinations**.
- Improved **custom pattern detection** and **order guessing**.
- Enhanced **auto-dimension detection** and **preview display**.
- Optimized path handling for mixed-depth folder structures.

---
