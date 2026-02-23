# ============================================================================
# KONFIGURASJON - Endre disse verdiene for å tilpasse skriptet
# ============================================================================
$CONFIG_MARGIN_CM = 2.0        # Sidemarger i centimeter (standard er 2.0 cm)
$CONFIG_IMAGE_WIDTH_CM = 17.0  # Maksimal bildebredde i centimeter
$CONFIG_PRINTER = "\\TDCSPRN30\Sikker_UtskriftCS"  # Printernavn

# ============================================================================
# SKRIPT FOR UTSKRIFT AV WORD, PDF, HTML FILER OG BILDER
# ============================================================================
#
# HENSIKT:
# Dette skriptet skriver ut alle Word (.docx), PDF (.pdf), HTML (.html/.htm) 
# filer og bilder i en valgt mappe og alle dens undermapper. 
# - For Word og HTML-filer legges mappenavnet til i toppteksten på utskriften
# - For mapper med bilder opprettes det automatisk en HTML-fil som viser bildene
#
# HVORDAN KJØRE SKRIPTET:
# 1. Høyreklikk på filen og velg "Kjør med PowerShell"
#    ELLER
# 2. Åpne PowerShell, naviger til mappen og skriv: .\print.ps1
#
# FØRSTE GANG DU KJØRER SKRIPTET:
# Hvis du får en feilmelding om "execution policy", åpne PowerShell
# (trenger IKKE administrator) og kjør:
#    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Svar "J" (Ja) på spørsmålet som kommer opp.
#
# KRAV FOR PDF-UTSKRIFT:
# Skriptet krever at Adobe Acrobat Reader er installert for automatisk 
# PDF-utskrift. Scriptet finner Adobe Reader automatisk på standardplasseringer.
# Dette påvirker IKKE Adobe Reader sine vanlige innstillinger - du får 
# fortsatt dialogbokser når du bruker Adobe Reader til vanlig.
#
# BILDESTØTTE:
# Skriptet støtter følgende bildeformater: jpg, jpeg, png, gif, bmp
# For hver mappe med bilder opprettes en HTML-fil som viser alle bildene.
# HTML-filen får mappenavnet som overskrift.
# Bildene skaleres automatisk til A4-størrelse for optimal utskrift.
#
# VIKTIG INFORMASJON:
# - Originaldokumentene endres IKKE, kun utskriften får mappenavn i topptekst
# - PDF-filer får IKKE automatisk mappenavn i topptekst (teknisk begrensning)
# - Du må velge en mappe når dialogen åpnes
# - Alle filer i valgt mappe OG undermapper blir skrevet ut
#
# ENDRE PRINTER:
# Bytt ut verdien til $CONFIG_PRINTER øverst i filen (linje 6)
#
# ============================================================================

# Fiks encoding for norske tegn (æ, ø, å)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Sjekk om skriptet kjører på Windows
if (-not $IsWindows -and $null -ne $PSVersionTable.PSVersion.Major -and $PSVersionTable.PSVersion.Major -ge 6) {
  # PowerShell Core 6+ har $IsWindows variabel
  Write-Host "FEIL: Dette skriptet kan kun kjøres på Windows."
  Write-Host "Skriptet krever funksjonalitet i Microsoft Word og Adobe Reader som kun er tilgjengelig i Windows."
  Write-Host ""
  Read-Host "Trykk Enter for å avslutte"
  exit
}

# Last inn nødvendige Windows-assemblies
try {
  Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
  Add-Type -AssemblyName System.Drawing -ErrorAction Stop
  Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
} catch {
  Write-Host "FEIL: Kunne ikke laste nødvendige Windows-komponenter."
  Write-Host "Dette skriptet krever Windows 10 eller nyere."
  Write-Host ""
  Read-Host "Trykk Enter for å avslutte"
  exit
}

Write-Host "Dette programmet skriver ut *alle* Word, PDF, HTML filer og bilder i mappen eller zip-filen du velger."
Write-Host "Printeren som er valgt er: $CONFIG_PRINTER"
Write-Host "Du kan bytte til en annen printer ved å redigere linje 6 i denne fila."

# Aktiverer ANSI/VT escape-koder i Windows-konsollen.
# Kreves for farger via escape-sekvenser og alternativ skjermbuffer.
function Enable-VirtualTerminal {
  $vtCode = @"
  using System;
  using System.Runtime.InteropServices;
  public class VT {
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GetStdHandle(int h);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GetConsoleMode(IntPtr h, out uint m);
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool SetConsoleMode(IntPtr h, uint m);
  }
"@
  try {
    Add-Type -TypeDefinition $vtCode -ErrorAction Stop
  } catch {
    # Typen kan allerede være registrert fra en tidligere kjøring
  }
  $handle = [VT]::GetStdHandle(-11)  # STD_OUTPUT_HANDLE
  $mode = 0
  [VT]::GetConsoleMode($handle, [ref]$mode) | Out-Null
  [VT]::SetConsoleMode($handle, $mode -bor 0x0004) | Out-Null  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
}

# Viser oppsummering i alternate screen buffer og venter på at brukeren trykker Enter.
function Show-Summary {
  param([string[]]$Lines, [bool]$HasErrors = $false)
  $esc = [char]27
  Enable-VirtualTerminal
  $origCursor = [Console]::CursorVisible
  [Console]::CursorVisible = $false
  [Console]::Write("$esc[?1049h")
  try {
    $width = (Get-Host).UI.RawUI.WindowSize.Width
    $yellow = "$esc[33m"; $green = "$esc[32m"; $red = "$esc[31m"; $dim = "$esc[2m"; $reset = "$esc[0m"
    $titleColor = if ($HasErrors) { $red } else { $green }
    $title = if ($HasErrors) { "  UTSKRIFT FULLFØRT MED FEIL" } else { "  UTSKRIFT FULLFØRT" }
    $buf = [System.Text.StringBuilder]::new(2048)
    [void]$buf.Append("$esc[H")
    [void]$buf.AppendLine("$yellow$("=" * $width)$reset")
    [void]$buf.AppendLine("$titleColor$($title.PadRight($width))$reset")
    [void]$buf.AppendLine("$yellow$("=" * $width)$reset")
    [void]$buf.AppendLine("".PadRight($width))
    foreach ($line in $Lines) {
      [void]$buf.AppendLine(("  " + $line).PadRight($width))
    }
    [void]$buf.AppendLine("".PadRight($width))
    [void]$buf.AppendLine("$dim$("  Trykk Enter for å avslutte".PadRight($width))$reset")
    [Console]::Write($buf.ToString())
    do {
      $key = (Get-Host).UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    } while ($key.VirtualKeyCode -ne 0x0D)
  } finally {
    [Console]::Write("$esc[?1049l")
    [Console]::CursorVisible = $origCursor
  }
}

# Bygger hele menyskjermbildet som én streng med ANSI-fargekoder.
# Støtter scrolling: viser kun elementer innenfor viewport ($scrollOffset).
# Hele strengen inkludert cursor-home skrives i ett Console.Write()-kall.
function Build-MenuBuffer {
  param(
    [string]$printer,
    [object[]]$menuItems,
    [int]$selectedIndex,
    [int]$scrollOffset,
    [int]$viewportSize
  )

  $esc = [char]27
  $cyan = "$esc[36m"
  $yellow = "$esc[33m"
  $dim = "$esc[2m"
  $reset = "$esc[0m"
  $width = (Get-Host).UI.RawUI.WindowSize.Width
  $buf = [System.Text.StringBuilder]::new(4096)

  # Cursor home — del av strengen for atomisk write
  [void]$buf.Append("$esc[H")

  $line = ("=" * [Math]::Min(60, $width)).PadRight($width)
  [void]$buf.AppendLine($line)
  [void]$buf.AppendLine(("  UTSKRIFTSINNSTILLINGER").PadRight($width))
  [void]$buf.AppendLine($line)
  [void]$buf.AppendLine(("  Printer: $printer").PadRight($width))

  # Scroll-indikator opp
  if ($scrollOffset -gt 0) {
    [void]$buf.AppendLine("$dim$("  ... $scrollOffset til over ...".PadRight($width))$reset")
  } else {
    [void]$buf.AppendLine("".PadRight($width))
  }

  # Tegn kun synlige elementer
  $endIndex = [Math]::Min($scrollOffset + $viewportSize, $menuItems.Count)
  for ($i = $scrollOffset; $i -lt $endIndex; $i++) {
    $item = $menuItems[$i]
    if ($item.Type -eq "header") {
      [void]$buf.AppendLine("$yellow$("  $($item.Label)".PadRight($width))$reset")
    }
    elseif ($item.Type -eq "separator") {
      [void]$buf.AppendLine("".PadRight($width))
    }
    elseif ($item.Type -eq "info") {
      [void]$buf.AppendLine("$dim$("  $($item.Label)".PadRight($width))$reset")
    }
    elseif ($item.Selectable) {
      $check = if ($item.Checked) { "X" } else { " " }
      $pointer = if ($i -eq $selectedIndex) { ">" } else { " " }
      $text = "  $pointer [$check] $($item.Label)"
      if ($item.Detail) { $text += "  [$($item.Detail)]" }
      if ($i -eq $selectedIndex) {
        [void]$buf.AppendLine("$cyan$($text.PadRight($width))$reset")
      } else {
        [void]$buf.AppendLine($text.PadRight($width))
      }
    }
  }

  # Fyll ut resten av viewport med tomme linjer
  $rendered = $endIndex - $scrollOffset
  for ($j = $rendered; $j -lt $viewportSize; $j++) {
    [void]$buf.AppendLine("".PadRight($width))
  }

  # Scroll-indikator ned
  $remaining = $menuItems.Count - $endIndex
  if ($remaining -gt 0) {
    [void]$buf.AppendLine("$dim$("  ... $remaining til under ...".PadRight($width))$reset")
  } else {
    [void]$buf.AppendLine("".PadRight($width))
  }

  [void]$buf.AppendLine(("  Piltaster: naviger | Mellomrom: endre valg").PadRight($width))
  [void]$buf.AppendLine(("  Enter: start utskrift | Esc: avbryt").PadRight($width))
  [void]$buf.Append($line)

  return $buf.ToString()
}

# Interaktiv innstillingsmeny (TUI).
# Tar over hele terminalen med alternativ skjermbuffer (som vim/htop).
# Skjermbildet bygges som én streng og skrives i ett Console.Write()-kall.
# Støtter scrolling for lange lister.
function Show-PrintSettings {
  param (
    [string]$printer,
    [object[]]$menuItems
  )

  # Finn startvalg: første innstilling, ellers første valgbare
  $selectedIndex = -1
  $firstSelectable = -1
  for ($i = 0; $i -lt $menuItems.Count; $i++) {
    if ($menuItems[$i].Selectable) {
      if ($firstSelectable -eq -1) { $firstSelectable = $i }
      if ($menuItems[$i].Key -and $selectedIndex -eq -1) { $selectedIndex = $i }
    }
  }
  if ($selectedIndex -eq -1) { $selectedIndex = $firstSelectable }

  $esc = [char]27
  Enable-VirtualTerminal

  $originalCursorVisible = [Console]::CursorVisible
  [Console]::CursorVisible = $false

  # Beregn viewport: terminalhøyde minus header (5) og footer (4)
  $windowHeight = (Get-Host).UI.RawUI.WindowSize.Height
  $chromeLines = 9  # 4 header + 1 scroll-opp + 1 scroll-ned + 3 footer
  $viewportSize = [Math]::Max(3, $windowHeight - $chromeLines)

  # Scroll-offset: sørg for at valgt element alltid er synlig
  # Scroll-margin: holder 2 linjer kontekst over/under valgt element
  $scrollMargin = 2
  $scrollOffset = 0
  if ($selectedIndex -gt $viewportSize - $scrollMargin - 1) {
    $scrollOffset = [Math]::Min($selectedIndex - $viewportSize + $scrollMargin + 1, [Math]::Max(0, $menuItems.Count - $viewportSize))
  }

  # Bytt til alternativ skjermbuffer (bevarer original terminalinnhold)
  [Console]::Write("$esc[?1049h")

  try {
    # Tegn menyen første gang
    $buffer = Build-MenuBuffer -printer $printer -menuItems $menuItems -selectedIndex $selectedIndex -scrollOffset $scrollOffset -viewportSize $viewportSize
    [Console]::Write($buffer)

    # Inputløkke med batched tastetrykk
    while ($true) {
      $needsRedraw = $false

      do {
        $press = (Get-Host).UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $vk = $press.VirtualKeyCode

        # Escape - avbryt
        if ($vk -eq 0x1B) { return $false }

        # Enter - bekreft
        if ($vk -eq 0x0D) { return $true }

        # Mellomrom - toggle avkryssing
        if ($vk -eq 0x20) {
          $menuItems[$selectedIndex].Checked = -not $menuItems[$selectedIndex].Checked
          $needsRedraw = $true
        }

        # Pil opp
        if ($vk -eq 0x26) {
          $next = $selectedIndex - 1
          while ($next -ge 0 -and -not $menuItems[$next].Selectable) { $next-- }
          if ($next -ge 0) { $selectedIndex = $next; $needsRedraw = $true }
        }

        # Pil ned
        if ($vk -eq 0x28) {
          $next = $selectedIndex + 1
          while ($next -lt $menuItems.Count -and -not $menuItems[$next].Selectable) { $next++ }
          if ($next -lt $menuItems.Count) { $selectedIndex = $next; $needsRedraw = $true }
        }
      } while ([Console]::KeyAvailable)

      if ($needsRedraw) {
        # Juster scroll-offset med margin slik at kontekst rundt valgt element er synlig
        if ($selectedIndex -lt $scrollOffset + $scrollMargin) {
          $scrollOffset = [Math]::Max(0, $selectedIndex - $scrollMargin)
        }
        if ($selectedIndex -ge $scrollOffset + $viewportSize - $scrollMargin) {
          $scrollOffset = [Math]::Min($selectedIndex - $viewportSize + $scrollMargin + 1, [Math]::Max(0, $menuItems.Count - $viewportSize))
        }

        $buffer = Build-MenuBuffer -printer $printer -menuItems $menuItems -selectedIndex $selectedIndex -scrollOffset $scrollOffset -viewportSize $viewportSize
        [Console]::Write($buffer)
      }
    }
  } finally {
    # Bytt tilbake til original skjermbuffer
    [Console]::Write("$esc[?1049l")
    [Console]::CursorVisible = $originalCursorVisible
  }
}

# Funksjon for å hente bildedimensjoner
function Get-ImageDimensions {
  param (
    [string]$imagePath
  )
  try {
    $img = [System.Drawing.Image]::FromFile($imagePath)
    $width = $img.Width
    $height = $img.Height
    $img.Dispose()
    return @{Width=$width; Height=$height}
  } catch {
    Write-Host "  ADVARSEL: Kunne ikke lese bilde: $(Split-Path $imagePath -Leaf)"
    # Standard dimensjoner: ca 11cm x 6cm ved 96 DPI
    return @{Width=415; Height=227}
  }
}

# Funksjon for å legge til mappenavn i topptekst og sidenummer i bunntekst
function Add-FolderNameToHeader {
  param (
    [object]$doc,
    [string]$folderName
  )

  foreach ($section in $doc.Sections) {
    # Sett marger
    $marginPoints = $CONFIG_MARGIN_CM * 28.35
    $section.PageSetup.TopMargin = $marginPoints
    $section.PageSetup.BottomMargin = $marginPoints
    $section.PageSetup.LeftMargin = $marginPoints
    $section.PageSetup.RightMargin = $marginPoints

    # Legg til mappenavn i topptekst
    $header = $section.Headers.Item(1)
    $existingText = $header.Range.Text.Trim()

    if ($existingText -ne "") {
      $header.Range.InsertBefore($folderName + "`n")
      $header.Range.Paragraphs.Item(1).Range.Font.Size = 10
      $header.Range.Paragraphs.Item(1).Range.Font.Bold = $true
      $header.Range.Paragraphs.Item(1).Range.ParagraphFormat.Alignment = 1
    } else {
      $header.Range.Text = $folderName
      $header.Range.Font.Size = 10
      $header.Range.Font.Bold = $true
      $header.Range.ParagraphFormat.Alignment = 1
    }

    # Legg til sidenummer i bunntekst (Side X av Y)
    $footer = $section.Footers.Item(1)
    $footer.Range.Text = ""

    # Word-konstanter for felttyper
    $wdFieldPage = 33      # Nåværende sidenummer
    $wdFieldNumPages = 26  # Totalt antall sider

    # Sett formatering på bunntekst
    $footer.Range.ParagraphFormat.Alignment = 1  # Midtstill
    $footer.Range.Font.Size = 10

    # Bygg bunntekst: "Side X av Y"
    # Vi bruker Duplicate for å lage kopier av Range-objektet slik at
    # vi kan jobbe med dem uavhengig av hverandre
    $footer.Range.InsertBefore("Side ")

    $tempRange = $footer.Range.Duplicate
    $tempRange.Collapse(0)  # Flytt til slutten
    $tempRange.Fields.Add($tempRange, $wdFieldPage, "", $false) | Out-Null

    $tempRange = $footer.Range.Duplicate
    $tempRange.Collapse(0)
    $tempRange.InsertAfter(" av ")

    $tempRange = $footer.Range.Duplicate
    $tempRange.Collapse(0)
    $tempRange.Fields.Add($tempRange, $wdFieldNumPages, "", $false) | Out-Null

    # Oppdater feltene slik at de vises riktig
    $footer.Range.Fields.Update() | Out-Null
  }
}

# Spør brukeren om de vil velge en zip-fil eller en mappe
$result = [System.Windows.Forms.MessageBox]::Show(
  "Vil du velge en ZIP-fil fra itslearning?`n`n" +
  "Klikk JA for å velge en zip-fil`n" +
  "Klikk NEI for å velge en mappe",
  "Velg filtype",
  [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
  [System.Windows.Forms.MessageBoxIcon]::Question
)

$selectedPath = $null
$isZipFile = $false
$tempExtractPath = $null

if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
  # Brukeren vil velge en zip-fil
  $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $fileDialog.Filter = "Zip-filer (*.zip)|*.zip"
  $fileDialog.Title = "Velg zip-fil fra itslearning"

  if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $zipFilePath = $fileDialog.FileName
    $isZipFile = $true

    # Opprett en midlertidig mappe for å pakke ut zip-filen
    $tempExtractPath = Join-Path $env:TEMP "PrintScript_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    Write-Host "Pakker ut zip-fil til midlertidig mappe..."
    Write-Host "Midlertidig mappe: $tempExtractPath"

    try {
      # Pakk ut zip-filen
      [System.IO.Compression.ZipFile]::ExtractToDirectory($zipFilePath, $tempExtractPath)
      $selectedPath = $tempExtractPath
      Write-Host "Zip-fil pakket ut!"
    } catch {
      Write-Host "FEIL: Kunne ikke pakke ut zip-filen: $_"
      Read-Host "Trykk Enter for å avslutte"
      exit
    }
  }
} elseif ($result -eq [System.Windows.Forms.DialogResult]::No) {
  # Brukeren vil velge en mappe
  $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
  $folderBrowser.Description = "Velg mappen med filer du vil skrive ut"

  if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $selectedPath = $folderBrowser.SelectedPath
  }
} else {
  # Brukeren kansellerte
  Read-Host "Avbrutt. Trykk Enter for å avslutte"
  exit
}

# Hvis brukeren klikket OK (enten for fil eller mappe)
if ($selectedPath -ne $null)
{

  # Hent alle Word, PDF og HTML filer i valgt mappe og undermapper
  # VIKTIG: Vi filtrerer bort filer som starter med punktum (.) siden disse ofte er
  # sikkerhetskopier eller skjulte systemfiler (f.eks. .~lock.dokument.docx)
  Write-Host "`nSkanner filer i valgt mappe..."
  $wordFiles = Get-ChildItem -Path $selectedPath -Recurse -Filter "*.docx" | Where-Object { -not $_.Name.StartsWith(".") }
  $pdfFiles = Get-ChildItem -Path $selectedPath -Recurse -Filter "*.pdf" | Where-Object { -not $_.Name.StartsWith(".") }
  $htmlFiles = Get-ChildItem -Path $selectedPath -Recurse -Include "*.html","*.htm" | Where-Object { -not $_.Name.StartsWith(".") }

  # Hent alle bildefiler og tekstfiler (hopp over filer som starter med .)
  $imageFiles = Get-ChildItem -Path $selectedPath -Recurse -Include "*.jpg","*.jpeg","*.png","*.gif","*.bmp" | Where-Object { -not $_.Name.StartsWith(".") }
  $textFiles = Get-ChildItem -Path $selectedPath -Recurse -Filter "*.txt" | Where-Object { -not $_.Name.StartsWith(".") }

  Write-Host "Funnet:"
  Write-Host "  - Word-filer (.docx): $($wordFiles.Count)"
  Write-Host "  - PDF-filer: $($pdfFiles.Count)"
  Write-Host "  - HTML-filer: $($htmlFiles.Count)"
  Write-Host "  - Bildefiler: $($imageFiles.Count)"
  Write-Host "  - Tekstfiler: $($textFiles.Count)"
  Write-Host ""

  # Sjekk om vi trenger Microsoft Word (for .docx eller .html filer)
  $wordAvailable = $true
  $needsWord = ($wordFiles.Count -gt 0 -or $htmlFiles.Count -gt 0 -or $imageFiles.Count -gt 0 -or $textFiles.Count -gt 0)

  if ($needsWord) {
    Write-Host "Sjekker om Microsoft Word er installert..."
    $wordType = [System.Type]::GetTypeFromProgID("Word.Application")
    if ($wordType -ne $null) {
      Write-Host "Microsoft Word funnet."
      Write-Host ""
    } else {
      Write-Host "ADVARSEL: Microsoft Word er ikke installert eller tilgjengelig."
      Write-Host "Word- og HTML-filer vil ikke bli skrevet ut."
      Write-Host ""
      $continue = Read-Host "Vil du fortsette uten Word-støtte? (J/N)"
      if ($continue -ne "J" -and $continue -ne "j") {
        # Rydd opp midlertidig mappe hvis vi pakket ut en zip-fil
        if ($isZipFile -and $tempExtractPath -ne $null -and (Test-Path $tempExtractPath)) {
          Remove-Item -Path $tempExtractPath -Recurse -Force
        }
        exit
      }
      $wordAvailable = $false
      Write-Host ""
    }
  } else {
    Write-Host "Ingen Word- eller HTML-filer funnet, hopper over Word-sjekk."
    Write-Host ""
  }

  # Sjekk om vi trenger Adobe Reader (for PDF-filer)
  $adobeReaderPath = $null
  if ($pdfFiles.Count -gt 0) {
    Write-Host "Sjekker etter Adobe Reader for PDF-utskrift..."

    # Finn Adobe Reader (sjekk vanlige installasjonsplasseringer)
    $adobePaths = @(
      "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
      "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
      "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
      "C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
      "C:\Program Files\Adobe\Reader 11.0\Reader\AcroRd32.exe"
    )

    foreach ($path in $adobePaths) {
      if (Test-Path $path) {
        $adobeReaderPath = $path
        Write-Host "Fant Adobe Reader: $path"
        Write-Host ""
        break
      }
    }

    if ($adobeReaderPath -eq $null) {
      Write-Host "ADVARSEL: Fant ikke Adobe Reader. PDF-filer vil ikke bli skrevet ut automatisk."
      Write-Host "Installer Adobe Acrobat Reader DC for automatisk PDF-utskrift."
      Write-Host ""
      $continue = Read-Host "Vil du fortsette uten PDF-støtte? (J/N)"
      if ($continue -ne "J" -and $continue -ne "j") {
        # Rydd opp midlertidig mappe hvis vi pakket ut en zip-fil
        if ($isZipFile -and $tempExtractPath -ne $null -and (Test-Path $tempExtractPath)) {
          Remove-Item -Path $tempExtractPath -Recurse -Force
        }
        exit
      }
      Write-Host ""
    }
  }

  # Generer kombinerte HTML-filer for mapper med bilder og/eller tekstfiler
  $generatedHtmlFiles = @()
  $foldersWithCombinedHtml = @()

  # Grupper bildefiler etter mappe
  $imagesByFolder = $imageFiles | Group-Object -Property DirectoryName

  # Grupper tekstfiler etter mappe
  $textsByFolder = $textFiles | Group-Object -Property DirectoryName

  # Finn alle mapper som har enten bilder eller tekstfiler
  $allFolders = @()
  $allFolders += $imagesByFolder | ForEach-Object { $_.Name }
  $allFolders += $textsByFolder | ForEach-Object { $_.Name }
  $allFolders = $allFolders | Select-Object -Unique

  foreach ($folderPath in $allFolders) {
    $folderName = Split-Path $folderPath -Leaf

    # Finn bilder i denne mappen
    $images = @($imagesByFolder | Where-Object { $_.Name -eq $folderPath } | ForEach-Object { $_.Group })

    # Finn tekstfiler i denne mappen
    $texts = @($textsByFolder | Where-Object { $_.Name -eq $folderPath } | ForEach-Object { $_.Group })

    # Sjekk om det allerede finnes HTML-filer i denne mappen
    $existingHtml = $htmlFiles | Where-Object { $_.DirectoryName -eq $folderPath }

    # Lag en kombinert HTML-fil
    $htmlFileName = Join-Path $folderPath "${folderName}_kombinert.html"

    # Start HTML-innhold
    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>$folderName</title>
    <style>
        @page {
            size: A4;
            margin: ${CONFIG_MARGIN_CM}cm;
        }

        body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: ${CONFIG_MARGIN_CM}cm;
        }

        h1 {
            text-align: center;
            color: #333;
            font-size: 24pt;
            margin: 0 0 1cm 0;
            page-break-after: avoid;
        }

        .image-container {
            width: 100%;
        }

        .image-wrapper {
            page-break-inside: avoid;
            page-break-after: always;
            text-align: center;
            margin-bottom: 1cm;
        }

        .image-wrapper:last-child {
            page-break-after: auto;
        }

        .image-wrapper img {
            display: block;
            margin: 0 auto;
        }

        .image-caption {
            margin-top: 0.5cm;
            font-size: 10pt;
            color: #666;
            page-break-before: avoid;
        }

        @media print {
            body {
                margin: 0;
                padding: 0;
            }

            .image-wrapper {
                page-break-inside: avoid;
                page-break-after: always;
            }

            .image-wrapper:last-child {
                page-break-after: auto;
            }
        }
    </style>
</head>
<body>
    <h1>$folderName</h1>
    <div class="image-container">
"@

    # Legg til tekstfiler øverst hvis de finnes
    if ($texts.Count -gt 0) {
      Write-Host "  Legger til $($texts.Count) tekstfil(er) i: $folderName"

      foreach ($txtFile in $texts) {
        $txtContent = Get-Content $txtFile.FullName -Raw -Encoding UTF8
        # HTML-encode tekstinnholdet for sikker visning
        $txtContentEncoded = $txtContent -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&#39;'
        # Legg til tekstinnholdet i en formatert div
        $htmlContent += "`n<!-- Innhold fra: $($txtFile.Name) -->`n"
        $htmlContent += "<div style='white-space: pre-wrap; font-family: monospace; margin-bottom: 1cm; padding: 0.5cm; border: 1px solid #ccc; background-color: #f9f9f9;'>"
        $htmlContent += $txtContentEncoded
        $htmlContent += "</div>`n"
      }
    }

    # Hvis det finnes eksisterende HTML-filer, ekstraher body-innholdet
    if ($existingHtml) {
      Write-Host "  Kombinerer $($existingHtml.Count) HTML-fil(er) i: $folderName"

      foreach ($htmlFile in $existingHtml) {
        $existingContent = Get-Content $htmlFile.FullName -Raw -Encoding UTF8

        # Ekstraher innhold mellom <body> og </body>
        if ($existingContent -match '(?s)<body[^>]*>(.*?)</body>') {
          $bodyContent = $matches[1]
          # Legg til body-innholdet
          $htmlContent += "`n<!-- Innhold fra: $($htmlFile.Name) -->`n"
          $htmlContent += $bodyContent
          $htmlContent += "`n"
        }
      }

      # Marker at denne mappen har en kombinert fil
      $foldersWithCombinedHtml += $folderPath
    }

    # Finn bilder som IKKE er referert i eksisterende HTML-filer
    $newImages = $images
    if ($existingHtml) {
      $referencedImages = @()
      foreach ($htmlFile in $existingHtml) {
        $existingContent = Get-Content $htmlFile.FullName -Raw -Encoding UTF8
        $matches = [regex]::Matches($existingContent, 'src="([^"]+)"')
        foreach ($match in $matches) {
          $imgSrc = $match.Groups[1].Value
          $referencedImages += Split-Path $imgSrc -Leaf
        }
      }
      $newImages = $images | Where-Object { $referencedImages -notcontains $_.Name }

      if ($newImages.Count -gt 0) {
        Write-Host "  Legger til $($newImages.Count) nye bilder på bunnen"
      }
    }

    # Legg til nye bilder hvis det finnes noen
    if ($newImages.Count -gt 0) {
      # Hent dimensjoner for alle bilder
      $imageDimensions = @{}

      foreach ($image in $newImages) {
        $dims = Get-ImageDimensions -imagePath $image.FullName
        $imageDimensions[$image.Name] = $dims
      }

      # Legg til hvert nye bilde med beregnet bredde og høyde i piksler
      foreach ($image in $newImages) {
        $relativePath = $image.Name
        $dims = $imageDimensions[$image.Name]
        $originalWidth = $dims.Width
        $originalHeight = $dims.Height

        # Maksimal bredde i piksler (ca 17cm ved 96 DPI)
        $maxWidthPx = 650

        # Beregn skalerte dimensjoner basert på aspect ratio
        if ($originalWidth -gt $maxWidthPx) {
          # Bildet er bredere enn max - skaler ned
          $scaleFactor = $maxWidthPx / $originalWidth
          $scaledWidth = $maxWidthPx
          $scaledHeight = [math]::Round($originalHeight * $scaleFactor)
        } else {
          # Bildet er smalere enn max - behold original størrelse
          $scaledWidth = $originalWidth
          $scaledHeight = $originalHeight
        }

        $htmlContent += @"

        <div class="image-wrapper">
            <img src="$relativePath" alt="$($image.BaseName)" width="$scaledWidth" height="$scaledHeight" style="display: block; margin: 0 auto;">
            <div class="image-caption">$($image.Name)</div>
        </div>
"@
      }
    }

    # Avslutt HTML
    $htmlContent += @"

    </div>
</body>
</html>
"@

    # Skriv HTML-filen
    $htmlContent | Out-File -FilePath $htmlFileName -Encoding UTF8

    $itemCount = $texts.Count + $images.Count + $(if ($existingHtml) { $existingHtml.Count } else { 0 })
    Write-Host "  Opprettet kombinert HTML-fil: $itemCount elementer ($($texts.Count) txt, $($existingHtml.Count) html, $($images.Count) bilder)"

    # Legg til i listen over genererte filer
    $generatedHtmlFiles += Get-Item $htmlFileName
  }
  
  # Fjern originale HTML-filer fra mapper som har kombinerte filer
  $htmlFilesToPrint = @($htmlFiles | Where-Object { $foldersWithCombinedHtml -notcontains $_.DirectoryName })

  # Legg til de genererte (kombinerte) HTML-filene
  $htmlFilesToPrint = @($htmlFilesToPrint) + @($generatedHtmlFiles)

  # Kombiner alle filer til én liste
  $allFiles = @()
  $allFiles += $wordFiles
  $allFiles += $pdfFiles
  $allFiles += $htmlFilesToPrint

  $fileCounter = 0
  $totalFiles = $allFiles.Count
  $failedFiles = @()

  # Bygg menyelementer for TUI (filer + innstillinger)
  $menuItems = @()

  if ($wordFiles.Count -gt 0) {
    $menuItems += @{ Type = "header"; Label = "Word-filer ($($wordFiles.Count)):"; Selectable = $false }
    foreach ($f in $wordFiles) {
      $menuItems += @{ Type = "item"; Label = $f.Name; Detail = $f.Directory.Name; Checked = $true; Selectable = $true; File = $f }
    }
  }

  if ($pdfFiles.Count -gt 0) {
    $menuItems += @{ Type = "separator"; Selectable = $false }
    $menuItems += @{ Type = "header"; Label = "PDF-filer ($($pdfFiles.Count)):"; Selectable = $false }
    foreach ($f in $pdfFiles) {
      $menuItems += @{ Type = "item"; Label = $f.Name; Detail = $f.Directory.Name; Checked = $true; Selectable = $true; File = $f }
    }
  }

  if ($htmlFilesToPrint.Count -gt 0) {
    $menuItems += @{ Type = "separator"; Selectable = $false }
    $menuItems += @{ Type = "header"; Label = "HTML-filer ($($htmlFilesToPrint.Count)):"; Selectable = $false }
    foreach ($f in $htmlFilesToPrint) {
      $menuItems += @{ Type = "item"; Label = $f.Name; Detail = $f.Directory.Name; Checked = $true; Selectable = $true; File = $f }
    }
  }

  # Innstillinger
  $menuItems += @{ Type = "separator"; Selectable = $false }
  $menuItems += @{ Type = "header"; Label = "Innstillinger:"; Selectable = $false }
  $menuItems += @{ Type = "item"; Label = "Topptekst og bunntekst (mappenavn + sidenummer)"; Checked = $true; Selectable = $true; Key = "headerFooter" }
  if ($wordFiles.Count -gt 0) {
    $menuItems += @{ Type = "item"; Label = "Skriv ut kommentarer i Word-dokumenter"; Checked = $false; Selectable = $true; Key = "comments" }
  }
  $menuItems += @{ Type = "info"; Label = "Tips: Svart-hvitt? Gå til Innstillinger (trykk Windows + I) › Skrivere og skannere › $CONFIG_PRINTER ›
  Utskriftsinnstillinger"; Selectable = $false }

  # Vis interaktiv innstillingsmeny
  $confirmed = Show-PrintSettings -printer $CONFIG_PRINTER -menuItems $menuItems

  if (-not $confirmed) {
    # Rydd opp midlertidig mappe hvis vi pakket ut en zip-fil
    if ($isZipFile -and $tempExtractPath -ne $null -and (Test-Path $tempExtractPath)) {
      Remove-Item -Path $tempExtractPath -Recurse -Force
    }
    Read-Host "Avbrutt. Trykk Enter for å avslutte"
    exit
  }

  # Les innstillinger fra menyen
  $shouldAddHeaderFooter = ($menuItems | Where-Object { $_.Key -eq "headerFooter" }).Checked
  $printWithComments = $false
  $commentsItem = $menuItems | Where-Object { $_.Key -eq "comments" }
  if ($commentsItem) { $printWithComments = $commentsItem.Checked }
  # Bygg filliste basert på avkryssede filer
  $allFiles = @($menuItems | Where-Object { $_.File -and $_.Checked } | ForEach-Object { $_.File })
  $totalFiles = $allFiles.Count

  Write-Host "Starter utskrift av $totalFiles filer..."

  # Opprett Word-applikasjon for Word og HTML filer
  $wordApp = $null
  if (($wordFiles.Count -gt 0 -or $htmlFilesToPrint.Count -gt 0) -and $wordAvailable) {
    $wordApp = New-Object -ComObject Word.Application
    $wordApp.Visible = $false
    # Sett aktiv printer for Word-applikasjonen
    $wordApp.ActivePrinter = $CONFIG_PRINTER
  }

  # Gå gjennom hver fil og skriv den ut
  foreach ($file in $allFiles)
  {
    $fileCounter++
    $fileExtension = $file.Extension.ToLower()
    
    # Hent mappenavnet (navnet på katalogen filen ligger i)
    $folderName = $file.Directory.Name

    try
    {
      if ($fileExtension -eq ".docx")
      {
        if ($wordApp -eq $null) {
          Write-Host "HOPPET OVER! $fileCounter av $totalFiles. Word: $($file.Name) (Word ikke tilgjengelig)"
          $failedFiles += $file.FullName
          continue
        }

        # Sjekk om filen er åpen i Word/Office.
        # Office erstatter de to første tegnene i filnavnet med ~$ (f.eks. test.docx → ~$st.docx).
        # Vi sjekker begge varianter for sikkerhetsskyld.
        $lockName1 = "~`$" + $file.Name.Substring([Math]::Min(2, $file.Name.Length))
        $lockName2 = "~`$" + $file.Name
        $lockFile1 = Join-Path $file.Directory.FullName $lockName1
        $lockFile2 = Join-Path $file.Directory.FullName $lockName2
        if ((Test-Path $lockFile1) -or (Test-Path $lockFile2)) {
          Write-Host "HOPPET OVER! $fileCounter av $totalFiles. $($file.Name) [Mappe: $folderName]: Filen er åpen i et annet program – lukk den og prøv igjen."
          $failedFiles += $file.FullName
          continue
        }

        # Håndter Word-dokumenter
        $doc = $wordApp.Documents.Open($file.FullName)

        # Legg til mappenavn i topptekst og sett marger (hvis ønsket)
        if ($shouldAddHeaderFooter) {
          Add-FolderNameToHeader -doc $doc -folderName $folderName
        }

        # Skriv ut dokumentet (1 kopi)
        # Item-parameter: 0 = wdPrintDocumentContent (uten kommentarer), 7 = wdPrintMarkup (med kommentarer)
        $wdPrintItem = if ($printWithComments) { 7 } else { 0 }
        $doc.PrintOut([ref]$false, [ref]$false, [ref]0, [ref]"", [ref]"", [ref]"", [ref]$wdPrintItem, [ref]1)

        # Lukk dokumentet uten å lagre endringer
        $doc.Close([ref]$false)

        Write-Host "OK! $fileCounter av $totalFiles. Skrev ut Word-dokumentet: $($file.Name) [Mappe: $folderName]"
      }
      elseif ($fileExtension -eq ".pdf")
      {
        # Håndter PDF-filer med Adobe Reader kommandolinje
        if ($adobeReaderPath -ne $null) {
          # Bruk Adobe Reader kommandolinje for automatisk utskrift
          # /t = print to printer, /h = minimize window
          $processInfo = New-Object System.Diagnostics.ProcessStartInfo
          $processInfo.FileName = $adobeReaderPath
          $processInfo.Arguments = "/t `"$($file.FullName)`" `"$CONFIG_PRINTER`""
          $processInfo.CreateNoWindow = $true
          $processInfo.UseShellExecute = $false

          $process = [System.Diagnostics.Process]::Start($processInfo)
          
          # Vent på at Adobe Reader starter og sender utskriftsjobben
          Start-Sleep -Seconds 3
          
          # Prøv å lukke Adobe Reader-prosessen
          try {
            $process.Kill()
          } catch {
            # Prosessen kan allerede være lukket
          }

          Write-Host "OK! $fileCounter av $totalFiles. Skrev ut PDF: $($file.Name) [Mappe: $folderName]"
          Write-Host "  NB: PDF-filer får ikke automatisk mappenavn i topptekst"
        } else {
          Write-Host "HOPPET OVER! $fileCounter av $totalFiles. PDF: $($file.Name) (Adobe Reader ikke funnet)"
          $failedFiles += $file.FullName
        }
      }
      elseif ($fileExtension -eq ".html" -or $fileExtension -eq ".htm")
      {
        if ($wordApp -eq $null) {
          Write-Host "HOPPET OVER! $fileCounter av $totalFiles. HTML: $($file.Name) (Word ikke tilgjengelig)"
          $failedFiles += $file.FullName
          continue
        }

        # Håndter HTML-filer gjennom Word
        $doc = $wordApp.Documents.Open($file.FullName)

        # Legg til mappenavn i topptekst og sett marger (hvis ønsket)
        if ($shouldAddHeaderFooter) {
          Add-FolderNameToHeader -doc $doc -folderName $folderName
        }

        # Skriv ut dokumentet (1 kopi)
        $doc.PrintOut([ref]$false, [ref]$false, [ref]0, [ref]"", [ref]"", [ref]"", [ref]0, [ref]1)

        # Lukk dokumentet uten å lagre endringer
        $doc.Close([ref]$false)

        Write-Host "OK! $fileCounter av $totalFiles. Skrev ut HTML-fil: $($file.Name) [Mappe: $folderName]"
      }

    } catch
    {
      $errMsg = $_.Exception.Message
      if ($errMsg -match "lock|låst|in use|opptatt|being used|another user|annen bruker") {
        Write-Host "FEIL! $fileCounter av $totalFiles. $($file.Name) [Mappe: $folderName]: Filen er åpen i et annet program – lukk den og prøv igjen."
      } else {
        Write-Host "FEIL! $fileCounter av $totalFiles. Problem med $($file.Name) [Mappe: $folderName]: $_"
      }
      $failedFiles += $file.FullName
    }
  }

  # Bygg oppsummeringslinjer
  $summaryLines = @()
  if ($failedFiles.Count -gt 0) {
    $summaryLines += "Følgende filer ble IKKE skrevet ut:"
    $failedFiles | ForEach-Object { $summaryLines += "  - $_" }
    $summaryLines += ""
  } else {
    $summaryLines += "$totalFiles dokumenter skrevet ut."
    $summaryLines += ""
  }
  $summaryLines += "- Word-filer: $($wordFiles.Count)"
  $summaryLines += "- PDF-filer: $($pdfFiles.Count)"
  $summaryLines += "- HTML-filer skrevet ut: $($htmlFilesToPrint.Count)"
  $summaryLines += "- Bildefiler: $($imageFiles.Count)"
  $summaryLines += "- Tekstfiler: $($textFiles.Count)"

  # Rydd opp genererte HTML-filer automatisk (stille)
  if ($generatedHtmlFiles.Count -gt 0) {
    foreach ($htmlFile in $generatedHtmlFiles) {
      Remove-Item $htmlFile.FullName -Force -ErrorAction SilentlyContinue
    }
  }

  # Rydd opp midlertidig mappe hvis vi pakket ut en zip-fil (stille)
  if ($isZipFile -and $tempExtractPath -ne $null -and (Test-Path $tempExtractPath)) {
    Remove-Item -Path $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
  }

  # Spør om snarvei før vi viser oppsummeringsskjermen
  $scriptPath = $PSCommandPath
  $scriptDir = Split-Path $scriptPath
  $noShortcutFile = Join-Path $scriptDir ".no-shortcut-prompt"
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  $shortcutPath = Join-Path $desktopPath "Print fra itslearning.lnk"

  if (-not (Test-Path $shortcutPath) -and -not (Test-Path $noShortcutFile)) {
    Write-Host ""
    $createShortcut = Read-Host "Vil du opprette en snarvei til dette skriptet på skrivebordet? (J/N)"

    if ($createShortcut -eq "J" -or $createShortcut -eq "j") {
      try {
        $wshShell = New-Object -ComObject WScript.Shell
        $shortcut = $wshShell.CreateShortcut($shortcutPath)
        $shortcut.TargetPath = "powershell.exe"
        $shortcut.Arguments = "-NoExit -ExecutionPolicy Bypass -File `"$scriptPath`""
        $shortcut.WorkingDirectory = $scriptDir
        $shortcut.Description = "Skriv ut elevbesvarelser automatisk"
        $shortcut.IconLocation = "powershell.exe,0"
        $shortcut.Save()
        $summaryLines += ""
        $summaryLines += "Snarvei opprettet på skrivebordet."
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wshShell) | Out-Null
      } catch {
        $summaryLines += ""
        $summaryLines += "Kunne ikke opprette snarvei: $_"
      }
    } else {
      try {
        "" | Out-File -FilePath $noShortcutFile -Encoding UTF8
      } catch { }
    }
  }

  Show-Summary -Lines $summaryLines -HasErrors ($failedFiles.Count -gt 0)

  # Avslutt Word-applikasjonen etter at oppsummeringen er vist
  if ($wordApp -ne $null) {
    $wordApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
  }
} else
{
  Read-Host "Ingen fil eller mappe valgt. Trykk Enter for å avslutte programmet."
}
