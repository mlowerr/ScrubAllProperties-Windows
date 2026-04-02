[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string]$Path = (Get-Location).Path,

    [switch]$Recurse,

    [string]$Filter = '*',

    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# `$IsWindows exists in PowerShell 6+, but not in Windows PowerShell 5.1.
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
        GPS_DEFAULT = 0x00000000,
        GPS_HANDLERPROPERTIESONLY = 0x00000001,
        GPS_READWRITE = 0x00000002,
        GPS_TEMPORARY = 0x00000004,
        GPS_FASTPROPERTIESONLY = 0x00000008,
        GPS_OPENSLOWITEM = 0x00000010,
        GPS_DELAYCREATION = 0x00000020,
        GPS_BESTEFFORT = 0x00000040,
        GPS_NO_OPLOCK = 0x00000080,
        GPS_PREFERQUERYPROPERTIES = 0x00000100,
        GPS_EXTRINSICPROPERTIES = 0x00000200,
        GPS_EXTRINSICPROPERTIESONLY = 0x00000400,
        GPS_VOLATILEPROPERTIES = 0x00000800,
        GPS_VOLATILEPROPERTIESONLY = 0x00001000,
        GPS_MASK_VALID = 0x00001fff
    }

    public static class PropertyHelpers {
        private const ushort VT_EMPTY = 0;

        public static uint ClearWritableProperties(string path) {
            Guid iid = NativeMethods.IID_IPropertyStore;
            IPropertyStore store;
            uint hr = NativeMethods.SHGetPropertyStoreFromParsingName(
                path,
                IntPtr.Zero,
                GETPROPERTYSTOREFLAGS.GPS_READWRITE,
                ref iid,
                out store);

            if (hr != 0) {
                return hr;
            }

            uint count;
            hr = store.GetCount(out count);
            if (hr != 0) {
                Marshal.ReleaseComObject(store);
                return hr;
            }

            for (uint i = 0; i < count; i++) {
                PROPERTYKEY key;
                hr = store.GetAt(i, out key);
                if (hr != 0) {
                    continue;
                }

                PROPVARIANT existing;
                hr = store.GetValue(ref key, out existing);
                if (hr != 0) {
                    continue;
                }

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
    }
}
"@
}

function Get-ScrubbedFileName {
    param(
        [Parameter(Mandatory)]
        [string]$OriginalPath
    )

    $directory = Split-Path -Path $OriginalPath -Parent
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalPath)
    $extension = [System.IO.Path]::GetExtension($OriginalPath)

    $candidate = Join-Path $directory "$baseName`_scrubbed$extension"
    $counter = 2

    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $directory "$baseName`_scrubbed_$counter$extension"
        $counter++
    }

    return $candidate
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

foreach ($file in $files) {
    try {
        $destination = if ($Force) {
            Join-Path $file.DirectoryName ("{0}_scrubbed{1}" -f $file.BaseName, $file.Extension)
        } else {
            Get-ScrubbedFileName -OriginalPath $file.FullName
        }

        Copy-Item -LiteralPath $file.FullName -Destination $destination -Force:$Force

        $hr = [PropertyScrubber.PropertyHelpers]::ClearWritableProperties($destination)
        if ($hr -ne 0) {
            $hex = ('0x{0:X8}' -f $hr)
            Write-Warning "Copied '$($file.FullName)' to '$destination', but metadata scrub returned HRESULT $hex."
            continue
        }

        Write-Host "Scrubbed copy created: $destination"
    }
    catch {
        Write-Warning "Failed to process '$($file.FullName)': $($_.Exception.Message)"
    }
}
