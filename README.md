# VBE-Decoder

A PowerShell function to decode VBScript Encoded (.vbe) files.

VBScript.Encode (`screnc.exe`) creates "encoded" scripts that are obfuscated but NOT encrypted. This tool reverses the encoding using the known character substitution table.

## Quick Start

```powershell
# Load the function
. .\VBE-Decoder.ps1

# Decode a .vbe file
ConvertFrom-VBE -Path "script.vbe"

# Decode and save to file
ConvertFrom-VBE -Path "encoded.vbe" | Set-Content "decoded.vbs"
```

## Usage

### Decode from file

```powershell
ConvertFrom-VBE -Path "C:\Scripts\logon.vbe"
```

### Decode from string

```powershell
ConvertFrom-VBE -EncodedScript '#@~^DgAAAA==\ko$K6,JCV^GJqAQAAA==^#~@'
```

### Decode from pipeline

```powershell
Get-Content "script.vbe" -Raw | ConvertFrom-VBE
```

### Batch decode all .vbe files in a directory

```powershell
Get-ChildItem "C:\Scripts" -Filter "*.vbe" | ForEach-Object {
    $decoded = ConvertFrom-VBE -Path $_.FullName
    $decoded | Set-Content ($_.FullName -replace '\.vbe$', '.vbs')
}
```

## How VBE Encoding Works

VBE uses a character substitution cipher with a rotating combination table:

1. The encoded data is wrapped in markers: `#@~^XXXXXX==<data>YYYYYY==^#~@`
2. Special escape sequences handle characters like `<`, `>`, `@`, and newlines
3. Each printable character is mapped through a 128-entry substitution table
4. A 64-entry combination table determines which of 3 possible decodings to use, cycling through positions

The encoding is trivially reversible — it provides obfuscation, not security.

## Requirements

- PowerShell 5.1+ (Windows PowerShell)
- No external dependencies

## Also Included In

This decoder is also integrated into [adPEAS v2](https://github.com/61106960/adPEAS), where it is used to automatically decode VBE-encoded logon scripts during credential exposure analysis.

## References

- [Didier Stevens - decode-vbe.py](https://blog.didierstevens.com/2016/03/29/decoding-vbe/)

## Author

**Alexander Sturz** — [SEKurity GmbH](https://sekurity.de)
