# ============================================================================
# KONFIGURASJON - Endre disse verdiene for å tilpasse skriptet
# ============================================================================
$CONFIG_MARGIN_CM = 2.0        # Sidemarger i centimeter (standard er 2.0 cm)
$CONFIG_IMAGE_WIDTH_CM = 17.0  # Maksimal bildebredde i centimeter
$CONFIG_PRINTER = "\\TDCSOM30\Sikker_UtskriftCS"  # Printernavn

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
# 2. Åpne PowerShell, naviger til mappen og skriv: .\print_forbedret_v2.ps1
#
# FØRSTE GANG DU KJØRER SKRIPTET:
# Hvis du får en feilmelding om "execution policy", må du åpne PowerShell 
# som administrator og kjøre følgende kommando:
#    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Svar "Y" (Yes) på spørsmålet som kommer opp.
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
# Bytt printernavn på linje under (der det står $printer = ...)
#
# ============================================================================

# Fiks encoding for norske tegn (æ, ø, å)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "Dette programmet skriver ut *alle* Word, PDF, HTML filer og bilder i mappen eller zip-filen du velger."
Write-Host "Printeren som er valgt er: $CONFIG_PRINTER"
Write-Host "Du kan bytte til en annen printer ved å redigere linje 6 i denne fila."
Write-Host ""
Write-Host "Tips: Du kan velge en zip-fil direkte fra itslearning uten å pakke den ut først!"
Write-Host ""

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem

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
    return @{Width=650; Height=800}
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

# Sjekk om Microsoft Word er installert
Write-Host "Sjekker om Microsoft Word er installert..."
$wordAvailable = $true
try {
  $testWord = New-Object -ComObject Word.Application -ErrorAction Stop
  $testWord.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($testWord) | Out-Null
  Write-Host "Microsoft Word funnet."
} catch {
  Write-Host "ADVARSEL: Microsoft Word er ikke installert eller tilgjengelig."
  Write-Host "Word- og HTML-filer vil ikke bli skrevet ut."
  Write-Host ""
  $continue = Read-Host "Vil du fortsette uten Word-støtte? (J/N)"
  if ($continue -ne "J" -and $continue -ne "j") {
    exit
  }
  $wordAvailable = $false
}

# Finn Adobe Reader (sjekk vanlige installasjonsplasseringer)
$adobePaths = @(
  "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
  "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
  "C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
  "C:\Program Files (x86)\Adobe\Reader 11.0\Reader\AcroRd32.exe",
  "C:\Program Files\Adobe\Reader 11.0\Reader\AcroRd32.exe"
)

$adobeReaderPath = $null
foreach ($path in $adobePaths) {
  if (Test-Path $path) {
    $adobeReaderPath = $path
    Write-Host "Fant Adobe Reader: $path"
    break
  }
}

if ($adobeReaderPath -eq $null) {
  Write-Host "ADVARSEL: Fant ikke Adobe Reader. PDF-filer vil ikke bli skrevet ut automatisk."
  Write-Host "Installer Adobe Acrobat Reader DC for automatisk PDF-utskrift."
  Write-Host ""
  $continue = Read-Host "Vil du fortsette uten PDF-støtte? (J/N)"
  if ($continue -ne "J" -and $continue -ne "j") {
    exit
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
  $wordFiles = Get-ChildItem -Path $selectedPath -Recurse -Filter "*.docx" | Where-Object { -not $_.Name.StartsWith(".") }
  $pdfFiles = Get-ChildItem -Path $selectedPath -Recurse -Filter "*.pdf" | Where-Object { -not $_.Name.StartsWith(".") }
  $htmlFiles = Get-ChildItem -Path $selectedPath -Recurse -Include "*.html","*.htm" | Where-Object { -not $_.Name.StartsWith(".") }

  # Hent alle bildefiler og tekstfiler (hopp over filer som starter med .)
  $imageFiles = Get-ChildItem -Path $selectedPath -Recurse -Include "*.jpg","*.jpeg","*.png","*.gif","*.bmp" | Where-Object { -not $_.Name.StartsWith(".") }
  $textFiles = Get-ChildItem -Path $selectedPath -Recurse -Filter "*.txt" | Where-Object { -not $_.Name.StartsWith(".") }

  # Generer kombinerte HTML-filer for mapper med bilder og/eller tekstfiler
  Write-Host "`nGenererer kombinerte HTML-filer for mapper med bilder og tekstfiler..."
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

  Write-Host "`nFjernet $($htmlFiles.Count - $htmlFilesToPrint.Count + $generatedHtmlFiles.Count) originale HTML-filer (erstattet med kombinerte)"

  # Kombiner alle filer til én liste
  $allFiles = @()
  $allFiles += $wordFiles
  $allFiles += $pdfFiles
  $allFiles += $htmlFilesToPrint

  $fileCounter = 0
  $totalFiles = $allFiles.Count
  $failedFiles = @()

  # Vis oversikt over filer som skal skrives ut
  Write-Host "`n============================================================"
  Write-Host "OVERSIKT OVER FILER SOM SKAL SKRIVES UT ($totalFiles filer)"
  Write-Host "============================================================"

  Write-Host "`nWord-filer ($($wordFiles.Count)):"
  foreach ($file in $wordFiles) {
    Write-Host "  - $($file.Name) [Mappe: $($file.Directory.Name)]"
  }

  Write-Host "`nPDF-filer ($($pdfFiles.Count)):"
  foreach ($file in $pdfFiles) {
    Write-Host "  - $($file.Name) [Mappe: $($file.Directory.Name)]"
  }

  Write-Host "`nHTML-filer ($($htmlFilesToPrint.Count)):"
  foreach ($file in $htmlFilesToPrint) {
    Write-Host "  - $($file.Name) [Mappe: $($file.Directory.Name)]"
  }

  Write-Host "`n============================================================"

  # Spør om topptekst og bunntekst skal legges til
  Write-Host "`nVil du legge til mappenavn i topptekst og sidenummer i bunntekst?"
  Write-Host "(Dette gjelder kun for Word og HTML-filer, ikke PDF-filer)"
  $addHeaderFooter = Read-Host "Trykk Enter for JA, eller skriv 'n' for NEI"

  # Standard er ja (tom input eller noe annet enn 'n')
  $shouldAddHeaderFooter = ($addHeaderFooter -ne "n" -and $addHeaderFooter -ne "N")

  if ($shouldAddHeaderFooter) {
    Write-Host "Legger til topptekst og bunntekst på utskriftene.`n"
  } else {
    Write-Host "Hopper over topptekst og bunntekst.`n"
  }

  Write-Host "`nStarter utskrift av $totalFiles filer..."

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

        # Håndter Word-dokumenter
        $doc = $wordApp.Documents.Open($file.FullName)

        # Legg til mappenavn i topptekst og sett marger (hvis ønsket)
        if ($shouldAddHeaderFooter) {
          Add-FolderNameToHeader -doc $doc -folderName $folderName
        }

        # Skriv ut dokumentet
        $doc.PrintOut()

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

        # Skriv ut dokumentet
        $doc.PrintOut()

        # Lukk dokumentet uten å lagre endringer
        $doc.Close([ref]$false)

        Write-Host "OK! $fileCounter av $totalFiles. Skrev ut HTML-fil: $($file.Name) [Mappe: $folderName]"
      }

    } catch
    {
      Write-Host "FEIL! $fileCounter av $totalFiles. Problem med $($file.Name): $_"
      $failedFiles += $file.FullName
    }
  }

  # Avslutt Word-applikasjonen hvis den ble opprettet
  if ($wordApp -ne $null) {
    $wordApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
  }

  # Frigjør COM-objektressurser
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  if ($failedFiles.Count -gt 0) {
    Write-Host "`nFølgende filer ble ikke skrevet ut:"
    $failedFiles | ForEach-Object { Write-Host $_ }
  } else {
    Write-Host "`n$totalFiles dokumenter skrevet ut."
  }

  Write-Host "`nOppsummering:"
  Write-Host "- Word-filer: $($wordFiles.Count)"
  Write-Host "- PDF-filer: $($pdfFiles.Count)"
  Write-Host "- HTML-filer skrevet ut: $($htmlFilesToPrint.Count) (hvorav $($generatedHtmlFiles.Count) kombinerte)"
  Write-Host "- Bildefiler funnet: $($imageFiles.Count)"
  Write-Host "- Tekstfiler funnet: $($textFiles.Count)"

  # Rydd opp genererte HTML-filer automatisk
  if ($generatedHtmlFiles.Count -gt 0) {
    Write-Host "`nSletter $($generatedHtmlFiles.Count) genererte HTML-fil(er)..."
    foreach ($htmlFile in $generatedHtmlFiles) {
      Remove-Item $htmlFile.FullName -Force
      Write-Host "  Slettet: $($htmlFile.Name)"
    }
  }

  # Rydd opp midlertidig mappe hvis vi pakket ut en zip-fil
  if ($isZipFile -and $tempExtractPath -ne $null -and (Test-Path $tempExtractPath)) {
    Write-Host "`nRydder opp midlertidig mappe..."
    try {
      Remove-Item -Path $tempExtractPath -Recurse -Force
      Write-Host "Midlertidig mappe slettet."
    } catch {
      Write-Host "Kunne ikke slette midlertidig mappe: $tempExtractPath"
      Write-Host "Du kan slette den manuelt hvis du vil."
    }
  }

  Read-Host "`nTrykk Enter for å avslutte programmet"
} else
{
  Read-Host "Ingen fil eller mappe valgt. Trykk Enter for å avslutte programmet."
}
