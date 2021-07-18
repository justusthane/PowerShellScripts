Get-Item $PSScriptRoot\*.ps1 | ForEach {
  If ($_.Name -ne "Load-Scripts.ps1") {
    . $_
  }
}
  
