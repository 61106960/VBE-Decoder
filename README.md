# VBE-Decoder

This Powershell script decodes an encoded 'VBE' string.
It is based on the original Microsoft VB script https://gallery.technet.microsoft.com/Encode-and-Decode-a-VB-a480d74c
and the Python code of Didier Stevens https://github.com/DidierStevens/DidierStevensSuite/blob/master/decode-vbe.py

# How To Use

```sh
Import-Module .\VBE-Decoder.ps1
Get-DecodedVBE -EncodedData "#@~^C2oAAA==v,sr^+,1ls+=~Hbo.lDkGUxW4k \(/@#@&v~.DkkGxl~ZRX@#@&vPzEO4KD)~PU&3@#@&v,ZGs:==^#~@"

' File Name: MigrationJobs.vbs
' Version: 0.5
' Author: TS3E
```

```sh
. .\VBE-Decoder.ps1
Get-Content .\EncodedVBSFile.vbe | Get-DecodedVBE

'--- set DHCPClassID and delete previous Migration files ---
'--- must run in source domain before actual migration takes place ---

Option Explicit

Dim strCommand, objShell, objStdOut, objFS, ObjWshScriPTExec, stroutpUt, strExecutable, strScRiptname, strScriptPath
Set objfS = CreatEObjeCt("Scripting.FilesystemObject")
strScriptname = WScript.ScRiptFullName
StrScriPtPath = OBjFS.GetPaRentfolderName(sTrscRiptnAme)
strExecutable = strScriptPath & "\psexec.exe"
'!!! replace service account name !!!
'!!! replace DHCPClassID - ClassID must be the one for the source domain !!!
[strip]
```