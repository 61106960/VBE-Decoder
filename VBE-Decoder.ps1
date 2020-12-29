Function Get-DecodedVBE {
<#
.SYNOPSIS
Author: Alexander Sturz (@_61106960_)
Required Dependencies: None
Optional Dependencies: None

.DESCRIPTION
Decodes an encoded 'VBE' string.
It is based on the original Microsoft VB script https://gallery.technet.microsoft.com/Encode-and-Decode-a-VB-a480d74c
and the Python one of Didier Stevens https://github.com/DidierStevens/DidierStevensSuite/blob/master/decode-vbe.py

.PARAMETER EncodedData
Takes the encoded VBE string

.EXAMPLE
Get-DecodedVBE -EncodedData "#@~^C2oAAA==v,sr^+,1ls+=~Hbo.lDkGUxW4k \(/@#@&v~.DkkGxl~ZRX@#@&vPzEO4KD)~PU&3@#@&v,ZGs:==^#~@"

.EXAMPLE
Get-Content "c:\encodedfile.vbe" | Get-DecodedVBE
#>

    [CmdletBinding()]
    Param (
        [Parameter(Position = 0, ValueFromPipeline = $True, ValueFromPipelineByPropertyName = $True)]
        [string]
        $EncodedData
    )

    try {
        # remove start and stop pattern
        $VbeData = $EncodedData -match '#@~\^......==(.+)......==\^#~@'
        if ($VbeData -eq $true) {
            $VbeData = $Matches[1]
            Write-Verbose "[Get-DecodedVBE] Found VBE trailing characters"    
        }

        # replace special characters
        $VbeData = $VbeData.Replace('@&',"`n").Replace('@#',"`r").Replace('@*','>').replace('@!','<').replace('@$','@')

        # initialize dict with static hex values as a list of three characters each, e.g. 65: ['w', 'E', 'B']
        $VbeDecList = @{
            "9" = 0x57,0x6E,0x7B
            "10" = 0x4A,0x4C,0x41
            "11" = 0x0B,0x0B,0x0B
            "12" = 0x0C,0x0C,0x0C
            "13" = 0x4A,0x4C,0x41
            "14" = 0x0E,0x0E,0x0E
            "15" = 0x0F,0x0F,0x0F
            "16" = 0x10,0x10,0x10
            "17" = 0x11,0x11,0x11
            "18" = 0x12,0x12,0x12
            "19" = 0x13,0x13,0x13
            "20" = 0x14,0x14,0x14
            "21" = 0x15,0x15,0x15
            "22" = 0x16,0x16,0x16
            "23" = 0x17,0x17,0x17
            "24" = 0x18,0x18,0x18
            "25" = 0x19,0x19,0x19
            "26" = 0x1A,0x1A,0x1A
            "27" = 0x1B,0x1B,0x1B
            "28" = 0x1C,0x1C,0x1C
            "29" = 0x1D,0x1D,0x1D
            "30" = 0x1E,0x1E,0x1E
            "31" = 0x1F,0x1F,0x1F
            "32" = 0x2E,0x2D,0x32
            "33" = 0x47,0x75,0x30
            "34" = 0x7A,0x52,0x21
            "35" = 0x56,0x60,0x29
            "36" = 0x42,0x71,0x5B
            "37" = 0x6A,0x5E,0x38
            "38" = 0x2F,0x49,0x33
            "39" = 0x26,0x5C,0x3D
            "40" = 0x49,0x62,0x58
            "41" = 0x41,0x7D,0x3A
            "42" = 0x34,0x29,0x35
            "43" = 0x32,0x36,0x65
            "44" = 0x5B,0x20,0x39
            "45" = 0x76,0x7C,0x5C
            "46" = 0x72,0x7A,0x56
            "47" = 0x43,0x7F,0x73
            "48" = 0x38,0x6B,0x66
            "49" = 0x39,0x63,0x4E
            "50" = 0x70,0x33,0x45
            "51" = 0x45,0x2B,0x6B
            "52" = 0x68,0x68,0x62
            "53" = 0x71,0x51,0x59
            "54" = 0x4F,0x66,0x78
            "55" = 0x09,0x76,0x5E
            "56" = 0x62,0x31,0x7D
            "57" = 0x44,0x64,0x4A
            "58" = 0x23,0x54,0x6D
            "59" = 0x75,0x43,0x71
            "60" = 0x4A,0x4C,0x41
            "61" = 0x7E,0x3A,0x60
            "62" = 0x4A,0x4C,0x41
            "63" = 0x5E,0x7E,0x53
            "64" = 0x40,0x4C,0x40
            "65" = 0x77,0x45,0x42
            "66" = 0x4A,0x2C,0x27
            "67" = 0x61,0x2A,0x48
            "68" = 0x5D,0x74,0x72
            "69" = 0x22,0x27,0x75
            "70" = 0x4B,0x37,0x31
            "71" = 0x6F,0x44,0x37
            "72" = 0x4E,0x79,0x4D
            "73" = 0x3B,0x59,0x52
            "74" = 0x4C,0x2F,0x22
            "75" = 0x50,0x6F,0x54
            "76" = 0x67,0x26,0x6A
            "77" = 0x2A,0x72,0x47
            "78" = 0x7D,0x6A,0x64
            "79" = 0x74,0x39,0x2D
            "80" = 0x54,0x7B,0x20
            "81" = 0x2B,0x3F,0x7F
            "82" = 0x2D,0x38,0x2E
            "83" = 0x2C,0x77,0x4C
            "84" = 0x30,0x67,0x5D
            "85" = 0x6E,0x53,0x7E
            "86" = 0x6B,0x47,0x6C
            "87" = 0x66,0x34,0x6F
            "88" = 0x35,0x78,0x79
            "89" = 0x25,0x5D,0x74
            "90" = 0x21,0x30,0x43
            "91" = 0x64,0x23,0x26
            "92" = 0x4D,0x5A,0x76
            "93" = 0x52,0x5B,0x25
            "94" = 0x63,0x6C,0x24
            "95" = 0x3F,0x48,0x2B
            "96" = 0x7B,0x55,0x28
            "97" = 0x78,0x70,0x23
            "98" = 0x29,0x69,0x41
            "99" = 0x28,0x2E,0x34
            "100" = 0x73,0x4C,0x09
            "101" = 0x59,0x21,0x2A
            "102" = 0x33,0x24,0x44
            "103" = 0x7F,0x4E,0x3F
            "104" = 0x6D,0x50,0x77
            "105" = 0x55,0x09,0x3B
            "106" = 0x53,0x56,0x55
            "107" = 0x7C,0x73,0x69
            "108" = 0x3A,0x35,0x61
            "109" = 0x5F,0x61,0x63
            "110" = 0x65,0x4B,0x50
            "111" = 0x46,0x58,0x67
            "112" = 0x58,0x3B,0x51
            "113" = 0x31,0x57,0x49
            "114" = 0x69,0x22,0x4F
            "115" = 0x6C,0x6D,0x46
            "116" = 0x5A,0x4D,0x68
            "117" = 0x48,0x25,0x7C
            "118" = 0x27,0x28,0x36
            "119" = 0x5C,0x46,0x70
            "120" = 0x3D,0x4A,0x6E
            "121" = 0x24,0x32,0x7A
            "122" = 0x79,0x41,0x2F
            "123" = 0x37,0x3D,0x5F
            "124" = 0x60,0x5F,0x4B
            "125" = 0x51,0x4F,0x5A
            "126" = 0x20,0x42,0x2C
            "127" = 0x36,0x65,0x57
            }

        # initialize dict with static int values as a static key to choose the character of the list above, therefore the values are from 0 to 2
        $VbePosList = @{
            "0" = 0
            "1" = 1
            "2" = 2
            "3" = 0
            "4" = 1
            "5" = 2
            "6" = 1
            "7" = 2
            "8" = 2
            "9" = 1
            "10" = 2
            "11" = 1
            "12" = 0
            "13" = 2
            "14" = 1
            "15" = 2
            "16" = 0
            "17" = 2
            "18" = 1
            "19" = 2
            "20" = 0
            "21" = 0
            "22" = 1
            "23" = 2
            "24" = 2
            "25" = 1
            "26" = 0
            "27" = 2
            "28" = 1
            "29" = 2
            "30" = 2
            "31" = 1
            "32" = 0
            "33" = 0
            "34" = 2
            "35" = 1
            "36" = 2
            "37" = 1
            "38" = 2
            "39" = 0
            "40" = 2
            "41" = 0
            "42" = 0
            "43" = 1
            "44" = 2
            "45" = 0
            "46" = 2
            "47" = 1
            "48" = 0
            "49" = 2
            "50" = 1
            "51" = 2
            "52" = 0
            "53" = 0
            "54" = 1
            "55" = 2
            "56" = 2
            "57" = 0
            "58" = 0
            "59" = 1
            "60" = 2
            "61" = 0
            "62" = 2
            "63" = 1
            }

        $CharIndex = -1
        foreach ($character in $VbeData.ToCharArray()) {

            # get hex value of character
            $Byte = [byte]$character

            # increase $index to change modulo result each run
            if ($Byte -lt 128) {
                $CharIndex += 1
            }

            # check if printable character and do the decoding
            if (($Byte -eq 9 -or $Byte -gt 31 -and $Byte -lt 128) -and ($Byte -ne 60 -and $Byte -ne 62 -and $Byte -ne 64)) {
                   $CombinationNumber = $($VbePosList["$($CharIndex % 64)"])
                   [char]$character = $($VbeDecList["$Byte"])[$CombinationNumber]
            }
            $DecodedVBE += $character -join ''
        }
        return $DecodedVBE     
    }
    
    catch { Write-Verbose "[Get-DecodedVBE] No VBE content detected" }
}