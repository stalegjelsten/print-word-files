# Skriv ut dokumenter automatisk

Dette programmet skriver ut alle dokumenter og bilder i en mappe og alle undermapper automatisk. Programmet støtter:
- Word-filer (.docx)
- PDF-filer (.pdf)
- HTML-filer (.html, .htm)
- Bildefiler (.jpg, .jpeg, .png, .gif, .bmp)
- Tekstfiler (.txt)

**Viktig:** Originaldokumentene dine endres IKKE. Programmet legger kun til mappenavnet øverst på utskriften, slik at du vet hvilken mappe dokumentet kom fra.

## Hvordan bruke programmet

1. Lagre filen `print.ps1` et sted på datamaskinen din (for eksempel på Skrivebordet)
2. Høyreklikk på `print.ps1` og velg **"Kjør med PowerShell"** eller **"Run with PowerShell"**
3. Et vindu åpnes. Les informasjonen som vises
4. Det åpnes automatisk et vindu hvor du kan velge hvilken mappe du vil skrive ut fra
5. Velg mappen og klikk OK
6. Programmet skriver nå ut alle dokumenter og bilder i mappen og alle undermapper
7. Når det er ferdig, trykk Enter for å lukke vinduet

## Krav

For at programmet skal fungere optimalt trenger du:
- **Microsoft Word** - For å skrive ut Word-dokumenter og HTML-filer
- **Adobe Acrobat Reader DC** - For å skrive ut PDF-filer

Hvis du ikke har disse programmene installert, vil programmet spørre om du vil fortsette uten støtte for disse filtypene.

## Endre innstillinger

Du kan tilpasse programmet ved å åpne `print.ps1` i Notisblokk og endre disse linjene øverst i filen:

- **Linje 8** (`$CONFIG_MARGIN_CM`): Endre sidemarger (i centimeter)
- **Linje 9** (`$CONFIG_IMAGE_WIDTH_CM`): Endre maksimal bildebredde (i centimeter)
- **Linje 10** (`$CONFIG_PRINTER`): Endre hvilken printer som skal brukes

**Standard printer er:** "Microsoft Print to PDF" (lagrer som PDF i stedet for å skrive ut på papir)

## Hva gjør programmet?

- **Word, HTML og bilder:** Programmet legger til mappenavnet øverst på hver side, slik at du vet hvilken mappe filen kom fra
- **PDF-filer:** Skrives ut som de er (uten mappenavn, siden PDF-filer er vanskeligere å endre)
- **Bilder og tekstfiler:** Programmet lager en midlertidig HTML-fil som viser alle bildene og tekstene fra hver mappe samlet, og skriver så ut denne. Når utskriften er ferdig, slettes den midlertidige filen automatisk

## Feilsøking

**Hvis du får feilmelding om "execution policy" når du prøver å kjøre skriptet:**

Dette betyr at datamaskinen din blokkerer PowerShell-skript av sikkerhetsgrunner. Slik fikser du det:

1. Høyreklikk på Start-knappen (nederst til venstre på skjermen)
2. Velg "Windows PowerShell (administrator)" eller "Terminal (administrator)"
3. Når det åpner seg et vindu med hvit eller blå tekst, skriv inn følgende og trykk Enter:
   ```
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
4. Skriv `J` og trykk Enter når du får spørsmål
5. Du kan nå lukke dette vinduet og prøve å kjøre `print.ps1` på nytt

*Dette trenger du bare å gjøre én gang på datamaskinen din*

**Hvis PDF-filer ikke skrives ut:**
- Installer Adobe Acrobat Reader DC (gratis nedlasting fra Adobe)

**Hvis Word- eller HTML-filer ikke skrives ut:**
- Kontroller at Microsoft Word er installert på datamaskinen
