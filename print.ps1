$printer = "\\TDCSOM30\Sikker_UtskriftCS"
Write-Host "Dette programmet skriver ut *alle* Word filer i mappen du velger."
Write-Host "Printeren som er valgt er: $printer"
Write-Host "Du kan bytte til en annen printer ved å redigere linje 1 i denne fila."

Add-Type -AssemblyName System.Windows.Forms

# Create a new folder browser dialog
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

# Show the dialog and get the result
$dialogResult = $folderBrowser.ShowDialog()

# If the user clicked OK
if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK)
{
  $selectedPath = $folderBrowser.SelectedPath

  # Get all the WORD files in the selected folder and its subfolders
  $files = Get-ChildItem -Path $selectedPath -Recurse -Filter "*.docx"

  # Create a new Word Application COM object
  $wordApp = New-Object -ComObject Word.Application
  $wordApp.Visible = $false

  $fileCounter = 0
  $totalFiles = $files.Count
  $failedFiles = @()

  # Loop through each file and print it
  foreach ($file in $files)
  {

    $fileCounter++
    try
    {
      
      # Open the document
      $doc = $wordApp.Documents.Open($file.FullName)

      # Print the document to the specified printer
      $doc.PrintOut()

      # Close the document without saving
      $doc.Close([ref]$false)

      Write-Host "OK! $fileCounter av $totalFiles. Skrev ut dokumentet $($file.Name)"

    } catch
    {

      Write-Host "FEIL! $fileCounter av $totalFiles. Problem med $($file.Name): $_"
      $failedFiles += $file.FullName

    }

  }

  # Quit the Word Application
  $wordApp.Quit()

  # Release COM object resources
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()

  if ($failedFiles.Count -gt 0) {
    Write-Host "Følgende filer ble ikke skrevet ut:"
    $failedFiles | ForEach-Object { Write-Host $_ }
  } else {
    Write-Host "$totalFiles dokumenter skrevet ut."
  }

  Read-Host "Trykk Enter for å avslutte programmet."
} else
{
  Read-Host "Ingen mappe valgt. Trykk Enter for å avslutte programmet."
}
