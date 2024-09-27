# Skriv ut alle Word-filer i en mappe

Dette skriptet skriver ut alle Word-filer i mappen du spesifiserer og i alle undermapper til denne. Word-filene må ha filendelse `.docx`.

## Bruk
1. Lagre `print.ps1` på datamaskinen
2. Høyreklikk på `print.ps1` og velg *Run with PowerShell*
3. Velg mappen med Word-filer.

## Oppsett
Du kan endre navnet på printeren ved å åpne `print.ps1` i en tekstbehandler (for eksempel notisblokk) og endre linje 1 til riktig navn med `$printer = "NAVN PÅ PRINTER"`. Du kan finne navnet på printere i Windows ved å åpne PowerShell og skrive kommandoen `Get-Printer`.
