[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path = (Get-Location).Path,
    [switch]$Recurse,
    [string]$Filter = '*',   # Fix 1: was '\*' (stray backslash — worked by accident on Windows)
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$mediaExtensions = @(
    '.jpg', '.jpeg', '.jpe', '.jfif', '.heic', '.heif', '.png', '.tif', '.tiff', '.webp',
    '.mp4', '.m4v', '.mov', '.avi', '.mkv', '.wmv'
)

$videoExtensions = @(
    '.mp4', '.m4v', '.mov', '.avi', '.mkv', '.wmv'
)

$imageRewriteExtensions = @(
    '.jpg', '.jpeg', '.jpe', '.jfif', '.png', '.tif', '.tiff'
)

# $IsWindows exists in PowerShell 6+, but not in Windows PowerShell 5.1.
$runningOnWindows = if (Get-Variable -Name IsWindows -ErrorAction SilentlyContinue) {
    [bool]$IsWindows
} else {
    $env:OS -eq 'Windows_NT'
}

if (-not $runningOnWindows) {
    throw 'This script only supports Microsoft Windows.'
}

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    throw "Path is not a valid directory: $Path"
}

if (-not ('PropertyScrubber.PropertyHelpers' -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

namespace PropertyScrubber {
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct PROPERTYKEY {
        public Guid fmtid;
        public uint pid;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct PROPVARIANT {
        [FieldOffset(0)] public ushort vt;
        [FieldOffset(8)] public IntPtr pointerValue;
        [FieldOffset(8)] public int int32Value;
        [FieldOffset(8)] public uint uint32Value;
        [FieldOffset(8)] public long int64Value;
        [FieldOffset(8)] public ulong uint64Value;
        [FieldOffset(8)] public double doubleValue;
        [FieldOffset(8)] public short boolValue;
    }

    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPropertyStore {
        uint GetCount(out uint cProps);
        uint GetAt(uint iProp, out PROPERTYKEY pkey);
        uint GetValue(ref PROPERTYKEY key, out PROPVARIANT pv);
        uint SetValue(ref PROPERTYKEY key, ref PROPVARIANT pv);
        uint Commit();
    }

    public static class NativeMethods {
        public static readonly Guid IID_IPropertyStore = new Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99");

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        public static extern uint SHGetPropertyStoreFromParsingName(
            string pszPath,
            IntPtr pbc,
            GETPROPERTYSTOREFLAGS flags,
            ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IPropertyStore propertyStore);

        [DllImport("ole32.dll", PreserveSig = true)]
        public static extern uint PropVariantClear(ref PROPVARIANT pvar);
    }

    [Flags]
    public enum GETPROPERTYSTOREFLAGS : uint {
        GPS_DEFAULT                 = 0x00000000,
        GPS_HANDLERPROPERTIESONLY   = 0x00000001,
        GPS_READWRITE               = 0x00000002,
        GPS_TEMPORARY               = 0x00000004,
        GPS_FASTPROPERTIESONLY      = 0x00000008,
        GPS_OPENSLOWITEM            = 0x00000010,
        GPS_DELAYCREATION           = 0x00000020,
        GPS_BESTEFFORT              = 0x00000040,
        GPS_NO_OPLOCK               = 0x00000080,
        GPS_PREFERQUERYPROPERTIES   = 0x00000100,
        GPS_EXTRINSICPROPERTIES     = 0x00000200,
        GPS_EXTRINSICPROPERTIESONLY = 0x00000400,
        GPS_VOLATILEPROPERTIES      = 0x00000800,
        GPS_VOLATILEPROPERTIESONLY  = 0x00001000,
        GPS_MASK_VALID              = 0x00001fff
    }

    public static class PropertyHelpers {
        private const ushort VT_EMPTY = 0;
        private const ushort VT_NULL  = 1;

        public static uint ClearWritableProperties(string path) {
            Guid iid = NativeMethods.IID_IPropertyStore;
            IPropertyStore store;
            uint hr = NativeMethods.SHGetPropertyStoreFromParsingName(
                path, IntPtr.Zero,
                GETPROPERTYSTOREFLAGS.GPS_READWRITE,
                ref iid, out store);
            if (hr != 0) return hr;

            uint count;
            hr = store.GetCount(out count);
            if (hr != 0) { Marshal.ReleaseComObject(store); return hr; }

            for (uint i = 0; i < count; i++) {
                PROPERTYKEY key;
                hr = store.GetAt(i, out key);
                if (hr != 0) continue;

                PROPVARIANT existing;
                hr = store.GetValue(ref key, out existing);
                if (hr != 0) continue;

                NativeMethods.PropVariantClear(ref existing);

                PROPVARIANT empty = new PROPVARIANT();
                empty.vt = VT_EMPTY;
                // Ignore read-only or unsupported properties.
                store.SetValue(ref key, ref empty);
                NativeMethods.PropVariantClear(ref empty);
            }

            hr = store.Commit();
            Marshal.ReleaseComObject(store);
            return hr;
        }

        // Fix 6: replaces HasPropertyValue + hardcoded PKEY spot-checks with a full
        // re-enumeration of the writable property store.  Returns the count of
        // properties that still carry a non-empty value after scrubbing.
        public static uint CountNonEmptyWritableProperties(string path, out uint nonEmptyCount) {
            nonEmptyCount = 0;
            Guid iid = NativeMethods.IID_IPropertyStore;
            IPropertyStore store;
            uint hr = NativeMethods.SHGetPropertyStoreFromParsingName(
                path, IntPtr.Zero,
                GETPROPERTYSTOREFLAGS.GPS_READWRITE,
                ref iid, out store);
            if (hr != 0) return hr;

            uint count;
            hr = store.GetCount(out count);
            if (hr != 0) { Marshal.ReleaseComObject(store); return hr; }

            for (uint i = 0; i < count; i++) {
                PROPERTYKEY key;
                hr = store.GetAt(i, out key);
                if (hr != 0) continue;

                PROPVARIANT value;
                hr = store.GetValue(ref key, out value);
                if (hr != 0) continue;

                try {
                    if (value.vt != VT_EMPTY && value.vt != VT_NULL) nonEmptyCount++;
                } finally {
                    NativeMethods.PropVariantClear(ref value);
                }
            }

            Marshal.ReleaseComObject(store);
            return 0;
        }
    }
}
"@
}

$script:ExifToolCommand = $null

function Get-ScrubbedFileName {
    param(
        [Parameter(Mandatory)]
        [string]$OriginalPath
    )
    $directory = Split-Path -Path $OriginalPath -Parent
    $baseName  = [System.IO.Path]::GetFileNameWithoutExtension($OriginalPath)
    $extension = [System.IO.Path]::GetExtension($OriginalPath)
    $candidate = Join-Path $directory "$baseName`_scrubbed$extension"
    $counter = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $directory "$baseName`_scrubbed_$counter$extension"
        $counter++
    }
    return $candidate
}

function Test-IsMediaFile {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )
    $extension = [System.IO.Path]::GetExtension($FilePath)
    return $mediaExtensions -contains $extension.ToLowerInvariant()
}

function Test-IsVideoFile {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )
    $extension = [System.IO.Path]::GetExtension($FilePath)
    return $videoExtensions -contains $extension.ToLowerInvariant()
}

function Get-ExifToolCommand {
    if ($script:ExifToolCommand) {
        return $script:ExifToolCommand
    }
    foreach ($candidate in @('exiftool', 'exiftool.exe')) {
        $command = Get-Command -Name $candidate -ErrorAction SilentlyContinue
        if ($command) {
            $script:ExifToolCommand = $command.Source
            return $script:ExifToolCommand
        }
    }
    return $null
}

function Get-FfmpegCommand {
    foreach ($candidate in @('ffmpeg', 'ffmpeg.exe')) {
        $command = Get-Command -Name $candidate -ErrorAction SilentlyContinue
        if ($command) {
            return $command.Source
        }
    }
    return $null
}

function Test-IsUnsupportedPropertyKeyHRESULT {
    param(
        [Parameter(Mandatory)]
        [uint32]$HResult
    )
    # Common HRESULTs when a property key is unavailable through the current
    # file's property handler.
    return @(
        0x80070490,  # ERROR_NOT_FOUND
        0x80004002,  # E_NOINTERFACE
        0x80070032   # ERROR_NOT_SUPPORTED
    ) -contains $HResult
}

# Fix 6: rewritten to use CountNonEmptyWritableProperties, which re-enumerates
# the entire writable property store rather than spot-checking a hardcoded list
# of ~19 PKEYs.  This catches any residual property, not just known ones.
function Invoke-PropertyKeyVerification {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    $nonEmptyCount = [uint32]0
    $hr = [PropertyScrubber.PropertyHelpers]::CountNonEmptyWritableProperties(
        $FilePath, [ref]$nonEmptyCount)

    if ($hr -ne 0) {
        if (Test-IsUnsupportedPropertyKeyHRESULT -HResult $hr) {
            return [pscustomobject]@{
                Attempted      = $true
                Succeeded      = $false
                Status         = 'Unsupported'
                ResidualFields = @()
                Message        = 'Verification unsupported: writable property store not accessible for this file type.'
            }
        }
        return [pscustomobject]@{
            Attempted      = $true
            Succeeded      = $false
            Status         = 'Failed'
            ResidualFields = @()
            Message        = ('Property store verification failed (HRESULT 0x{0:X8}).' -f $hr)
        }
    }

    if ($nonEmptyCount -eq 0) {
        return [pscustomobject]@{
            Attempted      = $true
            Succeeded      = $true
            Status         = 'VerifiedClean'
            ResidualFields = @()
            Message        = 'No non-empty writable properties remain in property store.'
        }
    }

    $noun = if ($nonEmptyCount -eq 1) { 'property' } else { 'properties' }
    return [pscustomobject]@{
        Attempted      = $true
        Succeeded      = $false
        Status         = 'ResidualFieldsFound'
        ResidualFields = @("$nonEmptyCount writable $noun remain non-empty")
        Message        = "$nonEmptyCount writable $noun remain non-empty after scrub."
    }
}

function Invoke-MediaMetadataScrub {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    $tool = Get-ExifToolCommand
    if (-not $tool) {
        return [pscustomobject]@{
            Attempted = $false
            Succeeded = $false
            Status    = 'Failed'
            Message   = 'ExifTool is required for media metadata scrubbing but was not found.'
        }
    }

    $arguments = @(
        '-overwrite_original',
        '-all=',
        '-P',
        '-m',
        '-q',
        '-q',
        $FilePath
    )

    $stdout    = $null
    $exitCode  = 0

    try {
        $stdout   = & $tool @arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    catch {
        return [pscustomobject]@{
            Attempted = $true
            Succeeded = $false
            Status    = 'Failed'
            Message   = "ExifTool invocation failed: $($_.Exception.Message)"
        }
    }

    if ($exitCode -eq 0) {
        return [pscustomobject]@{
            Attempted = $true
            Succeeded = $true
            Status    = 'Succeeded'
            Message   = 'Embedded media metadata removed with ExifTool.'
        }
    }

    $combinedOutput = if ($stdout) { ($stdout | Out-String).Trim() } else { 'No output.' }
    return [pscustomobject]@{
        Attempted = $true
        Succeeded = $false
        Status    = 'Failed'
        Message   = "ExifTool failed with exit code $exitCode. $combinedOutput"
    }
}

function Test-ExcludedExifToolVerificationTag {
    param(
        [Parameter(Mandatory)]
        [string]$TagName
    )

    # Exclude file-system/bookkeeping fields.
    $excludedPrefixes = @(
        'File:',
        'ExifTool:',
        'Composite:'
    )

    # Structural/container tags that are not user-authored metadata.
    # These are checked against the BARE tag name (after stripping any
    # group prefix) so that group-qualified names emitted by ExifTool -G1
    # — e.g. "JFIF:JFIFVersion", "PNG:Compression", "PNG:Filter" — match
    # correctly.  Fix 3: the original code checked the full string, so
    # "JFIF:JFIFVersion" never matched "JFIFVersion" and structural tags
    # were incorrectly reported as residual metadata.
    $excludedTags = @(
        'JFIFVersion',
        'MajorBrand',
        'MinorVersion',
        'CompatibleBrands',
        'HandlerType',
        'PrimaryItemReference',
        'ImageWidth',
        'ImageHeight',
        'BitDepth',
        'ColorType',
        'Compression',
        'Filter',
        'Interlace',
        'Encoding'
    )

    if ($TagName -eq 'SourceFile') {
        return $true
    }

    foreach ($prefix in $excludedPrefixes) {
        if ($TagName.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    # Strip group prefix (e.g. "JFIF:JFIFVersion" -> "JFIFVersion") before
    # checking against $excludedTags.
    $bareName = if ($TagName -match '^[^:]+:(.+)$') { $Matches[1] } else { $TagName }

    foreach ($excludedTag in $excludedTags) {
        if ($bareName.Equals($excludedTag, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Invoke-ExifToolMetadataVerification {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    $tool = Get-ExifToolCommand
    if (-not $tool) {
        return [pscustomobject]@{
            Attempted      = $false
            Succeeded      = $false
            Status         = 'Failed'
            ResidualFields = @()
            Message        = 'ExifTool is required for media metadata verification but was not found.'
        }
    }

    $arguments = @(
        '-j',
        '-a',
        '-G1',
        '-s',
        '-sort',
        $FilePath
    )

    try {
        $raw      = & $tool @arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    catch {
        return [pscustomobject]@{
            Attempted      = $true
            Succeeded      = $false
            Status         = 'Failed'
            ResidualFields = @()
            Message        = "ExifTool verification failed: $($_.Exception.Message)"
        }
    }

    if ($exitCode -ne 0) {
        $output = if ($raw) { ($raw | Out-String).Trim() } else { 'No output.' }
        return [pscustomobject]@{
            Attempted      = $true
            Succeeded      = $false
            Status         = 'Failed'
            ResidualFields = @()
            Message        = "ExifTool verification returned exit code $exitCode. $output"
        }
    }

    try {
        $records = $raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        return [pscustomobject]@{
            Attempted      = $true
            Succeeded      = $false
            Status         = 'Failed'
            ResidualFields = @()
            Message        = "ExifTool verification output was not valid JSON: $($_.Exception.Message)"
        }
    }

    $record = @($records)[0]
    if (-not $record) {
        return [pscustomobject]@{
            Attempted      = $true
            Succeeded      = $true
            Status         = 'VerifiedClean'
            ResidualFields = @()
            Message        = 'No ExifTool metadata records were returned.'
        }
    }

    $residual = New-Object System.Collections.Generic.List[string]
    foreach ($property in $record.PSObject.Properties) {
        if (-not (Test-ExcludedExifToolVerificationTag -TagName $property.Name)) {
            $residual.Add($property.Name)
        }
    }

    $residualFields = @($residual | Sort-Object -Unique)
    if ($residualFields.Count -eq 0) {
        return [pscustomobject]@{
            Attempted      = $true
            Succeeded      = $true
            Status         = 'VerifiedClean'
            ResidualFields = @()
            Message        = 'No residual metadata tags detected by ExifTool (excluding file-system/stat fields).'
        }
    }

    return [pscustomobject]@{
        Attempted      = $true
        Succeeded      = $false
        Status         = 'ResidualFieldsFound'
        ResidualFields = $residualFields
        Message        = ('Residual metadata tags detected by ExifTool: {0}' -f ($residualFields -join ', '))
    }
}

function Invoke-MediaMetadataRewriteFallback {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    $extension = [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
    $tempPath  = Join-Path (Split-Path -Path $FilePath -Parent) (
        '{0}.rewrite{1}' -f [System.Guid]::NewGuid().ToString('N'), $extension)

    try {
        if (Test-IsVideoFile -FilePath $FilePath) {
            $ffmpeg = Get-FfmpegCommand
            if (-not $ffmpeg) {
                return [pscustomobject]@{
                    Attempted = $false
                    Succeeded = $false
                    Status    = 'Skipped'
                    Message   = 'Remux fallback skipped: ffmpeg was not found in PATH.'
                }
            }

            $arguments = @(
                '-y',
                '-i', $FilePath,
                '-map', '0',
                '-map_metadata', '-1',
                '-map_chapters', '-1',
                '-dn',
                '-c', 'copy',
                $tempPath
            )

            $stdout   = & $ffmpeg @arguments 2>&1
            $exitCode = $LASTEXITCODE

            if ($exitCode -ne 0 -or -not (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
                $details = if ($stdout) { ($stdout | Out-String).Trim() } else { 'No output.' }
                return [pscustomobject]@{
                    Attempted = $true
                    Succeeded = $false
                    Status    = 'Failed'
                    Message   = "Video remux fallback failed (exit code $exitCode). $details"
                }
            }
        }
        elseif ($imageRewriteExtensions -contains $extension) {
            Add-Type -AssemblyName System.Drawing

            # Fix 5: use an explicit extension-keyed format object rather than
            # Image.RawFormat, which can return an unrecognised GUID for JPEG
            # variants (.jpe, .jfif) and some TIFF configurations, causing
            # Save() to silently fall back to BMP or throw.
            $formatMap = @{
                '.jpg'  = [System.Drawing.Imaging.ImageFormat]::Jpeg
                '.jpeg' = [System.Drawing.Imaging.ImageFormat]::Jpeg
                '.jpe'  = [System.Drawing.Imaging.ImageFormat]::Jpeg
                '.jfif' = [System.Drawing.Imaging.ImageFormat]::Jpeg
                '.png'  = [System.Drawing.Imaging.ImageFormat]::Png
                '.tif'  = [System.Drawing.Imaging.ImageFormat]::Tiff
                '.tiff' = [System.Drawing.Imaging.ImageFormat]::Tiff
            }
            $imageFormat = $formatMap[$extension]

            $image = [System.Drawing.Image]::FromFile($FilePath)
            try {
                $image.Save($tempPath, $imageFormat)
            }
            finally {
                $image.Dispose()
            }
        }
        else {
            # Fix 7: HEIC, HEIF, WEBP and other formats land here.  Clarify
            # that this is a known limitation, not a generic configuration gap.
            return [pscustomobject]@{
                Attempted = $false
                Succeeded = $false
                Status    = 'Skipped'
                Message   = ("Rewrite fallback skipped for '$extension': no in-box Windows rewrite path is " +
                             "available for this format.  ExifTool is the only supported scrub method; if " +
                             "residual tags remain they cannot be removed by this script.")
            }
        }

        if (-not (Test-Path -LiteralPath $tempPath -PathType Leaf)) {
            return [pscustomobject]@{
                Attempted = $true
                Succeeded = $false
                Status    = 'Failed'
                Message   = 'Rewrite/remux fallback did not produce an output file.'
            }
        }

        Move-Item -LiteralPath $tempPath -Destination $FilePath -Force

        return [pscustomobject]@{
            Attempted = $true
            Succeeded = $true
            Status    = 'Succeeded'
            Message   = 'Format-aware rewrite/remux fallback completed.'
        }
    }
    catch {
        return [pscustomobject]@{
            Attempted = $true
            Succeeded = $false
            Status    = 'Failed'
            Message   = "Rewrite/remux fallback failed: $($_.Exception.Message)"
        }
    }
    finally {
        if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-PropertyStoreScrub {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath
    )

    try {
        $hr = [PropertyScrubber.PropertyHelpers]::ClearWritableProperties($FilePath)
        if ($hr -eq 0) {
            return [pscustomobject]@{
                Attempted = $true
                Succeeded = $true
                Status    = 'Succeeded'
                Message   = 'Windows writable property store values were cleared.'
            }
        }

        $hex = ('0x{0:X8}' -f $hr)
        return [pscustomobject]@{
            Attempted = $true
            Succeeded = $false
            Status    = 'Failed'
            Message   = "Property store commit returned HRESULT $hex."
        }
    }
    catch {
        $message = $_.Exception.Message
        $isUnsupportedHandler =
            $message -match 'bitmap codec does not support the bitmap property' -or
            $message -match '0x88982F41'

        if ($isUnsupportedHandler) {
            return [pscustomobject]@{
                Attempted = $true
                Succeeded = $false
                Status    = 'Unsupported'
                Message   = 'Property store handler does not support writable bitmap properties for this format.'
            }
        }

        return [pscustomobject]@{
            Attempted = $true
            Succeeded = $false
            Status    = 'Failed'
            Message   = "Property store scrub failed: $message"
        }
    }
}

$enumerationOptions = @{
    LiteralPath = $Path
    File        = $true
    Filter      = $Filter
}

if ($Recurse) {
    $enumerationOptions.Recurse = $true
}

$files = Get-ChildItem @enumerationOptions

if (-not $files) {
    Write-Host 'No files matched the provided criteria.'
    return
}

$matchedMediaFiles = @($files | Where-Object { Test-IsMediaFile -FilePath $_.FullName })
if ($matchedMediaFiles.Count -gt 0 -and -not (Get-ExifToolCommand)) {
    throw ('ExifTool is required to scrub metadata from matched media files. Install ExifTool and retry. First media file: {0}' -f $matchedMediaFiles[0].FullName)
}

$results = New-Object System.Collections.Generic.List[object]

function Get-ResultOutcome {
    param(
        [Parameter(Mandatory)]
        [pscustomobject]$Result
    )

    if (-not $Result.Copied) {
        return 'Failed'
    }

    $scrubFailure =
        ($Result.IsMedia -and -not $Result.MediaMetadataRemoved) -or
        ($Result.PropertyStoreStatus -eq 'Failed') -or
        ($Result.VerificationStatus -eq 'ResidualFieldsFound') -or
        ($Result.VerificationStatus -eq 'Failed')

    if ($scrubFailure) {
        return 'Failed'
    }

    if ($Result.Warnings.Count -gt 0) {
        return 'Partial'
    }

    return 'FullSuccess'
}

foreach ($file in $files) {
    $fileResult = [ordered]@{
        SourceFile            = $file.FullName
        OutputFile            = $null
        Copied                = $false
        IsMedia               = Test-IsMediaFile -FilePath $file.FullName
        MediaMetadataRemoved  = $false
        PropertyStoreCleared  = $false
        PropertyStoreStatus   = 'NotAttempted'
        VerificationStatus    = 'NotChecked'
        RewriteFallbackStatus = 'NotAttempted'
        Outcome               = 'NotProcessed'
        ResidualFields        = @()
        Warnings              = New-Object System.Collections.Generic.List[string]
    }

    try {
        # Fix 2: always use collision-safe naming regardless of -Force.
        # The original code bypassed Get-ScrubbedFileName when -Force was set,
        # which could silently overwrite an input file whose name already ended
        # in _scrubbed.  -Force is now forwarded only to Copy-Item, where its
        # standard meaning is "copy over read-only destination attributes".
        $destination = Get-ScrubbedFileName -OriginalPath $file.FullName

        Copy-Item -LiteralPath $file.FullName -Destination $destination -Force:$Force
        $fileResult.OutputFile = $destination
        $fileResult.Copied     = $true

        if ($fileResult.IsMedia) {
            $mediaScrub = Invoke-MediaMetadataScrub -FilePath $destination
            if ($mediaScrub.Succeeded) {
                $fileResult.MediaMetadataRemoved = $true
            }
            else {
                $fileResult.Warnings.Add($mediaScrub.Message)
            }
        }

        $propertyScrub = Invoke-PropertyStoreScrub -FilePath $destination
        $fileResult.PropertyStoreStatus = $propertyScrub.Status
        if ($propertyScrub.Succeeded) {
            $fileResult.PropertyStoreCleared = $true
        }
        elseif ($propertyScrub.Status -eq 'Unsupported') {
            if (-not $fileResult.MediaMetadataRemoved) {
                $fileResult.Warnings.Add($propertyScrub.Message)
            }
        }
        else {
            $fileResult.Warnings.Add($propertyScrub.Message)
        }

        if ($fileResult.IsMedia) {
            $mediaVerification = Invoke-ExifToolMetadataVerification -FilePath $destination
            $fileResult.VerificationStatus = $mediaVerification.Status

            if ($mediaVerification.Status -eq 'ResidualFieldsFound') {
                $firstPassResidualFields = @($mediaVerification.ResidualFields)
                $fileResult.Warnings.Add(('First-pass residual metadata tags: {0}' -f ($firstPassResidualFields -join ', ')))

                $rewriteFallback = Invoke-MediaMetadataRewriteFallback -FilePath $destination
                $fileResult.RewriteFallbackStatus = $rewriteFallback.Status
                $fileResult.Warnings.Add($rewriteFallback.Message)

                if ($rewriteFallback.Succeeded) {
                    $secondPassVerification = Invoke-ExifToolMetadataVerification -FilePath $destination
                    $fileResult.VerificationStatus = $secondPassVerification.Status

                    if ($secondPassVerification.Status -eq 'ResidualFieldsFound') {
                        $secondPassResidualFields = @($secondPassVerification.ResidualFields)
                        $fileResult.ResidualFields = $secondPassResidualFields
                        $fileResult.Warnings.Add(('Second-pass residual metadata tags after fallback: {0}' -f ($secondPassResidualFields -join ', ')))
                    }
                    elseif (-not $secondPassVerification.Succeeded) {
                        $fileResult.ResidualFields = @($secondPassVerification.ResidualFields)
                        $fileResult.Warnings.Add($secondPassVerification.Message)
                    }
                }
                else {
                    $fileResult.ResidualFields = $firstPassResidualFields
                }
            }
            elseif (-not $mediaVerification.Succeeded) {
                $fileResult.ResidualFields = @($mediaVerification.ResidualFields)
                $fileResult.Warnings.Add($mediaVerification.Message)
            }
        }
        else {
            $nonMediaVerification = Invoke-PropertyKeyVerification -FilePath $destination
            $fileResult.VerificationStatus = $nonMediaVerification.Status
            if (-not $nonMediaVerification.Succeeded) {
                $fileResult.ResidualFields = @($nonMediaVerification.ResidualFields)
                $fileResult.Warnings.Add($nonMediaVerification.Message)
            }
        }
    }
    catch {
        $fileResult.Warnings.Add($_.Exception.Message)
        Write-Warning "Failed to process '$($file.FullName)': $($_.Exception.Message)"
    }

    # Fix 4: compute Outcome via Get-ResultOutcome before printing the per-file
    # console message so both always agree.  The original code used a separate
    # inline $isFullSuccess expression whose conditions were not identical to
    # Get-ResultOutcome, causing the printed message and the Outcome column to
    # contradict each other for some edge cases (e.g. VerificationStatus = 'Unsupported').
    $typedResult = [pscustomobject]$fileResult
    $typedResult | Add-Member -NotePropertyName Outcome -NotePropertyValue (Get-ResultOutcome -Result $typedResult) -Force
    $results.Add($typedResult)

    if ($fileResult.Copied) {
        if ($typedResult.Outcome -eq 'FullSuccess') {
            Write-Host "Scrubbed copy created: $($fileResult.OutputFile)"
        }
        else {
            Write-Warning "Scrub completed with warnings for '$($file.FullName)'. See summary output for details."
        }
    }
}

Write-Host ''
Write-Host 'Scrub summary:'

$fullSuccess = 0
$partial     = 0
$failed      = 0

foreach ($result in $results) {
    switch ($result.Outcome) {
        'FullSuccess' { $fullSuccess++ }
        'Partial'     { $partial++ }
        'Failed'      { $failed++ }
        default       { $failed++ }
    }
}

Write-Host ("  Full success : {0}" -f $fullSuccess)
Write-Host ("  Partial      : {0}" -f $partial)
Write-Host ("  Failed       : {0}" -f $failed)
Write-Host ''

$results |
    Select-Object SourceFile, OutputFile, Copied, IsMedia, MediaMetadataRemoved,
        PropertyStoreCleared, PropertyStoreStatus, VerificationStatus, RewriteFallbackStatus, Outcome,
        @{ Name = 'ResidualFields'; Expression = { $_.ResidualFields -join '; ' } },
        @{ Name = 'Warnings';       Expression = { $_.Warnings -join ' | ' } } |
    Format-Table -AutoSize
