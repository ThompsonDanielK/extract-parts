# Extract Parts

A PowerShell utility for extracting and flattening RAR archive sets, optionally compressing loose `.nsp` / `.xci` files, and safely moving outputs into a chosen destination folder.

This project was created and iteratively improved with GitHub Copilot as a coding assistant. It is provided as-is for education, automation, and personal use, and it should be used only with content you own or are otherwise legally entitled to process.

## What it does

- scans a source folder for RAR archive sets
- extracts archive contents with WinRAR
- flattens extracted files into a destination folder
- optionally converts loose `.nsp` / `.xci` files to `.nsz` / `.xcz` using NSZ
- can prompt before deleting original archives or loose source files
- records diagnostic logs and a transcript for troubleshooting

## Key safety behavior

- prompts interactively by default
- honors command-line overrides when provided
- avoids destructive cleanup in dry-run or no-delete modes
- verifies output before deleting original source files
- keeps archives unless the user explicitly confirms deletion

## Requirements

- Windows PowerShell 5.1+
- WinRAR CLI (`Rar.exe`)
- NSZ CLI and a compatible legal keys file (`prod.keys` or `keys.txt`)
- Python environment for the NSZ install flow, if needed

## Important legal note

- You must only process files you own or are legally allowed to convert or extract.
- The script does not provide or download keys.
- You are responsible for obtaining any required legal game keys in compliance with local law and platform terms.

## Quick start

1. Open PowerShell in the repository.
2. Run:

```powershell
.\Extract-Parts.ps1
```

3. Pick the source folder and destination folder when prompted.
4. Confirm any cleanup choices that appear.

## Example command-line usage

```powershell
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -OutputFolder "D:\Extracted" -DeleteArchives -DeleteStandalone
.\Extract-Parts.ps1 -NonInteractive -SourceFolder "D:\Roms" -DryRun
.\Extract-Parts.ps1 -NonInteractive -SourceFolder "D:\Roms" -DeleteArchives
.\Extract-Parts.ps1 -SourceFolder "D:\Roms" -MaxParallel 2
```

| Parameter | Meaning |
| --- | --- |
| `-SourceFolder` | Folder containing RAR sets / loose `.nsp`/`.xci` files. Required with `-NonInteractive`. |
| `-OutputFolder` | Destination folder. Defaults to a picker (interactive) or `<source>\Extracted` (non-interactive). |
| `-DeleteArchives` | Delete each RAR set after all its files are verified in the output folder. |
| `-DeleteStandalone` | Delete loose `.nsp`/`.xci` originals after the compressed copy is verified. |
| `-DryRun` | Report what would be processed, then exit. Nothing is extracted, converted, installed or deleted. |
| `-NoDelete` | Process everything but never delete source files. |
| `-NonInteractive` | No dialogs or prompts: yes/no prompts answer "no", duplicate filenames are skipped. Exit code 1 on failures. |
| `-Force` | With `-NonInteractive`, auto-approve dependency installs (never deletions). |
| `-MaxParallel` | Concurrent WinRAR jobs (1-32, default 6). Use a lower value for spinning hard drives. |

## Notes

- If command-line flags are supplied, they take precedence over the interactive prompts.
- If you run in non-interactive mode, the script prefers safe defaults.
- Password-protected archives are reported as failures instead of waiting for a password.
- Temporary files are staged in a hidden `.ExtractParts_staging_*` folder inside the output folder and removed at the end of the run (including after Ctrl+C).
- The script logs details to the local app data logs directory and keeps a transcript for troubleshooting.

## License

This repository does not currently include a project license file. If you intend to redistribute or reuse it publicly, consider adding an explicit license before publishing broadly.
