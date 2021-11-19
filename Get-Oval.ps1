<#
.SYNOPSIS
  Get advisory data from RedHat
.DESCRIPTION
  <Brief description of script>
.PARAMETER <Parameter_Name>
  This script querries RedHat OVAL api to obtain RHSA, RHBA and RHEA information. 
.INPUTS
  The script will ask you for a CVE number
.OUTPUTS
  It Creates a directory on your desktop named logs. Inside logs will be <RHSA>.csv
  <RHSA>.csv contains Tabs:
    1notes: Any notes related to the advisory including description, (CVSS) base score and Security Fix(es).
    2mitigation: RedHat's assesment on the vulnerability.
    3packageState: a list of applicable packages, fix state and cpe.
    4affectedRelese: lists the packages being pushed, the cpe associated, release date and advisory ID.
    5relatedCVE: Other CVE being resolved in the advisory.
    6errata: the advisory errata url, type and description
.NOTES
  Version:        1.0
  Author:         Scott B Lichty
  Creation Date:  11/19/2021
  Purpose/Change: Initial script development to quickly gather information on RHEL vulnerabilities.
.EXAMPLE
  N/A
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

param(
    [Parameter(Mandatory)]$cve
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'

#Import Modules & Snap-ins

#----------------------------------------------------------[Declarations]----------------------------------------------------------
#No Declarations needed
#-----------------------------------------------------------[Functions]------------------------------------------------------------


Function Get-RSHA {
  Param ($cve)
  Begin {
    $url = "https://access.redhat.com/hydra/rest/securitydata/cve/$cve.json" 
    $results = Invoke-RestMethod -Uri $url
    Write-Host 'Gahering information from RedHat API...' -ForegroundColor Red
  }
  Process {
    Try {
      $affected = $results.affected_release
      $details = $results.details
      $mitigation = $results.mitigation.value
      $uniqueAdvisories = $resultsRef.advisory | Sort-Object | Get-Unique

      foreach ($uniqueAdvisory in $uniqueAdvisories) {
        $uniqueAdvisory
        $details
        $rhsaUrl = "https://access.redhat.com/hydra/rest/securitydata/cvrf/$($uniqueAdvisory).xml"
        $oval = Invoke-RestMethod -Uri $rhsaUrl

        $notes = $oval.cvrfdoc.DocumentNotes.note.'#text'
        $vul = $oval.cvrfdoc.Vulnerability
        $var = $env:USERPROFILE + "\Desktop"
        $dir = $var + "\logs"
        $bam = $uniqueAdvisory.replace(':', '_')
        $finalCsv = $dir + "\" + $bam + ".xlsx"

        if (Test-Path -Path $dir) {
          #Do Nothing
        }
        else {
          mkdir $dir | Out-Null
        }

        $notes | ForEach-Object {
          $newNotes = $_.Replace(",", "")
          Out-File -InputObject $newNotes -FilePath $dir\1notes.csv -Append
        }

        $mitigation | ForEach-Object {
          $newMit = $_.Replace(",", "")
          Out-File -InputObject $newMit -FilePath $dir\2mitigation.csv -Append
        }

        $vul = $oval.cvrfdoc.Vulnerability

        $vulnerability = $vul | ForEach-Object {
          [PSCustomObject]@{
            CVE         = $_.cve
            Type        = $_.Remediations.Remediation.type  
            Notes       = $_.notes.note.'#text'
            Status      = $_.Involvements.Involvement.Status
            Remediation = $_.Remediations.Remediation.url
          }
          $i++
        }

        $results.package_state | export-csv -Path $dir\3packageState.csv -NoTypeInformation
        $affected | Export-Csv -Path $dir\4affectedRelease.csv -NoTypeInformation
        $vulnerability | Export-CSV -Path  $dir\5relatedCVE.csv -NoTypeInformation
        $oval.cvrfdoc.DocumentReferences.Reference | export-csv -Path $dir\6errata.csv -NoTypeInformation

        Function Merge-CSVFiles {
          Param(
            $CSVPath = "$dir", ## Source CSV Folder
            $XLOutput = $finalCsv ## Output file name
          )
    
          $csvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv
          $Excel = New-Object -ComObject excel.application 
          $Excel.visible = $false
          $Excel.sheetsInNewWorkbook = $csvFiles.Count
          $workbooks = $excel.Workbooks.Add()
          $CSVSheet = 1
    
          Foreach ($CSV in $Csvfiles)
          {
            $worksheets = $workbooks.worksheets
            $CSVFullPath = $CSV.FullName
            $SheetName = ($CSV.name -split "\.")[0]
            $worksheet = $worksheets.Item($CSVSheet)
            $worksheet.Name = $SheetName
            $TxtConnector = ("TEXT;" + $CSVFullPath)
            $CellRef = $worksheet.Range("A1")
            $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)
            $worksheet.QueryTables.item($Connector.name).TextFileCommaDelimiter = $True
            $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
            $worksheet.QueryTables.item($Connector.name).Refresh() | Out-Null
            $worksheet.QueryTables.item($Connector.name).delete() | Out-Null
            $worksheet.UsedRange.EntireColumn.AutoFit() | Out-Null
            $CSVSheet++ | Out-Null
    
          }
    
          $workbooks.SaveAs($XLOutput, 51)
          $workbooks.Saved = $true
          $workbooks.Close()
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
          $excel.Quit()
          [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
          [System.GC]::Collect()
          [System.GC]::WaitForPendingFinalizers()
        }

        Merge-CSVFiles
        Remove-Item $dir\1notes.csv
        Remove-Item $dir\2mitigation.csv
        Remove-Item $dir\3packageState.csv
        Remove-Item $dir\4affectedRelease.csv
        Remove-Item $dir\5relatedCVE.csv
        Remove-Item $dir\6errata.csv
        Invoke-Item $dir
      }
    }
    Catch {
      Write-Host -BackgroundColor Red "Error: $($_.Exception)"
      Break
    }
  }
  End {
    If ($?) {
      Write-Host 'Completed Successfully.' -ForegroundColor Green
      Write-Host ' '
    }
  }
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

Get-RSHA