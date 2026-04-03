# ScrubAllProperties-Windows

`Scrub-FileProperties.ps1` creates a copy of each file and attempts to remove all writable Windows file properties (metadata), similar to **File Explorer → Properties → Details → Remove Properties and Personal Information → Create a copy with all possible properties removed**.

## Usage

```powershell
# Process files in current directory (default)
.\Scrub-FileProperties.ps1

# Process a specific directory
.\Scrub-FileProperties.ps1 -Path 'C:\Photos'

# Include subdirectories
.\Scrub-FileProperties.ps1 -Path 'C:\Photos' -Recurse

# Restrict file types
.\Scrub-FileProperties.ps1 -Filter '*.jpg'
```

## Notes

- The script only supports Windows.
- Copies are created next to originals with `_scrubbed` suffix.
- ExifTool is a required dependency when matched files include image/video formats; media metadata removal will fail without it.
- ExifTool verification intentionally ignores file/stat and bookkeeping fields (`SourceFile`, `File:*`, `ExifTool:*`, `Composite:*`) plus a small explicit set of structural container tags (for example dimensions/encoding/container brand markers) that are not user-authored metadata.
- If ExifTool verification still reports residual metadata tags after the first scrub pass, the script attempts a format-aware fallback before declaring failure:
  - Video containers (`.mp4`, `.m4v`, `.mov`, `.avi`, `.mkv`, `.wmv`): remux with `ffmpeg` (`-map_metadata -1 -map_chapters -1 -c copy`) when available.
  - Supported still images (`.jpg`, `.jpeg`, `.jpe`, `.jfif`, `.png`, `.tif`, `.tiff`): rewrite through `System.Drawing` save/load.
  - Other media types are left as-is and reported with warnings.
- The script records first-pass and second-pass residual tag sets separately in warnings so fallback effectiveness is visible per file.
- Tradeoffs: remux/rewrite fallback can change container layout or image encoding details, may not preserve all non-metadata stream-level features exactly, and can be skipped/fail when required tooling (for example `ffmpeg`) is unavailable.
- Some formats/properties may be read-only or unsupported by the property handler. In those cases the script keeps the copied file and logs a warning.
