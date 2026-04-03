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
- Some formats/properties may be read-only or unsupported by the property handler. In those cases the script keeps the copied file and logs a warning.
