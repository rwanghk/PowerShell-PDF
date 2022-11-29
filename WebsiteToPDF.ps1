$url = "https://google.com"
$destination = "C:\Users\user\Desktop\1.pdf"
<##############################>
<# Load Data to Internet Browser #>

$ie = New-Object -ComObject 'InternetExplorer.Application'
while($ie.busy) { Start-Sleep -Milliseconds 50 }
$ie.Navigate($url)
while($ie.busy) { Start-Sleep -Milliseconds 50 }
#Load separate HTML file into the website for custom rendering
#$html = Get-Content -Path './body.html' -Raw; $ie.Document.body.innerHTML = $html; while($ie.busy) { Start-Sleep -Milliseconds 50 }
$ie.Document.parentWindow.execScript('let div = null;')
$ie.Visible = $false

Function Load-Data($ie, $id, $val) {
    $jsCommand = "div = null; div = document.getElementById(""$($id)""); if (div) { div.innerText = ""$($val)""}"
    $document = $ie.Document
    $document.parentWindow.execScript($jsCommand, 'javascript') | Out-Null
}

<##############################>
<# Set up PDF printers as appropriate #>
$printerName = 'PrintPDFUnattended'
$defaultPrinter = (Get-CimInstance -Class Win32_Printer -Filter "Default=true").Name
# https://www.powershellgallery.com/packages/WienExpertsLive/1.14
#Author(s)
#
#Dr. Tobias Weltner
#Copyright
#
#(c) 2019 Tobias Weltner. Use freely at own risk.
Function Install-PDFPrinter {
    <#
        .SYNOPSIS
        Installs a new Printer called "PrintPDFUnattended" which prints to file
 
        .DESCRIPTION
        Uses the built-in "PrintToPDF" printer driver to create a new printer
        that prints unattendedly to a fixed file in the temp folder
 
        .EXAMPLE
        Install-PDFPrinter
        Installs the printer "PrintPDFUnattended". Needs to be run only once.
        To remove the printer again, use this command:
        Remove-Printer PrintPDFUnattended
 
        .NOTES
        Requires the "PrintToPDF" printer driver shipping with Windows 10, Server 2016, or better
 
        .LINK
        URLs to related sites
        The first link is opened by Get-Help -Online Install-PDFPrinter
    #>
    $PrinterDefaultName = 'Microsoft Print to PDF'
    # choose a default path where the PDF is saved:
    $PDFFilePath = "$env:temp\PDFResultFile.pdf"
    # see whether the driver exists
    $ok = @(Get-PrinterDriver -Name $PrinterDefaultName -ea 0).Count -gt 0
    if (!$ok) {
      Write-Warning -Message "Printer driver 'Microsoft Print to PDF' not available."
      Write-Warning -Message 'This driver ships with Windows 10 or Server 2016.'
      Write-Warning -Message "If it is still not available, enable the 'Printing-PrintToPDFServices-Features'"
      Write-Warning -Message 'Example: Enable-WindowsOptionalFeature -Online -FeatureName Printing-PrintToPDFServices-Features'
      return
    }
    # check whether port exists
    $port = Get-PrinterPort -Name $PDFFilePath -ErrorAction SilentlyContinue
    if ($port -eq $null) {
      # create printer port
      Add-PrinterPort -Name $PDFFilePath 
    }
    # add printer
    Add-Printer -DriverName $PrinterDefaultName -Name $printerName -PortName $PDFFilePath 
}

If (Get-CimInstance -Class Win32_Printer -Filter "Name='PrintPDFUnattended'") {
    
} Else {
    Install-PDFPrinter
}

$printer = Get-CimInstance -Class Win32_Printer -Filter "Name='$printerName'"
Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter | Out-Null

<##############################>
<# Load and save PDF #>
#Fix style issue in printing PDF if required
#$ie.Document.body.style.padding = "0 0 0 0"; $ie.Document.body.style.removeAttribute("maxWidth");

Function Print-IE ($ie, $destination) {
  $TempPDF = "$env:temp\PDFResultFile.pdf"
  If (Test-Path -Path $TempPDF) {
      Remove-Item -Path $TempPDF -Force
  }
  Start-Sleep -Milliseconds 50
  while($ie.busy) {
      Start-Sleep -Milliseconds 50
  }
  $ie.ExecWB(6, 2)
  $ok = $false
  do { 
      Start-Sleep -Milliseconds 100
      Write-Host '.' -NoNewline
      $fileExists = Test-Path -Path $TempPDF
      If ($fileExists) {
          Try {
              Move-Item -Path $TempPDF -Destination $destination -Force -ea Stop
              $ok = $true
          } Catch {
              # file is still in use, cannot move
              # try again
          }
      }
  } until ( $ok )
}

#Load data to IE as appropriate, then
#
# Print-IE -ie $ie -destination $destination

#Repeat for other files as appropriate
Write-Host

#Finish
$ie.Quit();

<##############################>
<# Reset printer setting #>

Write-Host "Reset default printer as normal printer"
$printer = Get-CimInstance -Class Win32_Printer -Filter "Name='$($defaultPrinter)'"
Invoke-CimMethod -InputObject $printer -MethodName SetDefaultPrinter | Out-Null
Remove-Printer -Name $printerName

Write-Host "Finised. Press any key to continue..."
Read-Host
