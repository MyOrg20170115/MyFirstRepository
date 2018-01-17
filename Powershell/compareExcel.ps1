# Diff excel file.
param($aExcelPath = $(Read-Host "Enter one excel path"),
      $bExcelPath = $(Read-Host "Enter other excel path"))

$ErrorActionPreference = "stop"
$DebugPreference = "SilentlyContinue" # Continue SilentlyContinue Stop Inquire

$commandPath = Split-Path -parent $myInvocation.MyCommand.path
$commandName = Split-Path -leaf $myInvocation.MyCommand.path
$commandBaseName = (gci $myInvocation.MyCommand.path).BaseName

Set-Location $commandPath

<# for test
$aExcelPath = Join-Path $commandPath "testdata\aExcel.xlsx"
$bExcelPath = Join-Path $commandPath "testdata\bExcel.xlsx"
#>

# const value
$xlNumbers = 1
$xlCellTypeFormulas = -4123
$xlNone = 0

function main() {

  try {
    # check file
    if (-not (Test-Path $aExcelPath)) {
      Write-Host "$($aExcelPath) is not found !"
      return -1
    }
    if (-not (Test-Path $bExcelPath)) {
      Write-Host "$($bExcelPath) is not found !"
      return -1
    }

    $aExcelPath = $aExcelPath -replace """", ""
    $bExcelPath = $bExcelPath -replace """", ""

    # to abs path
    $aExcelPath = Convert-Path $aExcelPath
    $bExcelPath = Convert-Path $bExcelPath

    $aExcelCheckPath = Join-Path (Split-Path -parent $aExcelPath) ((gci $aExcelPath).BaseName +
                                  "_diffcheck" + (gci $aExcelPath).Extension)
    $bExcelCheckPath = Join-Path (Split-Path -parent $bExcelPath) ((gci $bExcelPath).BaseName +
                                  "_diffcheck" + (gci $bExcelPath).Extension)

    # backup
    cp $aExcelPath $aExcelCheckPath
    cp $bExcelPath $bExcelCheckPath

    $excel = New-Object -ComObject Excel.Application

    $excel.Visible = $false
    $excel.Application.DisplayAlerts = $false
    $excel.Application.ScreenUpdating = $false

    # tmp book
    $tBook = $excel.Workbooks.Add()

    $aBook = $excel.Workbooks.Open($aExcelCheckPath, 0, $false)
    $bBook = $excel.Workbooks.Open($bExcelCheckPath, 0, $false)

    Write-Host "============= excel check start ============="
    diffCheck $aBook $bBook $tBook
    Write-Host "============= excel check end ============="

    $aBook.Save()
    $bBook.Save()

  } catch {

    Write-Host "Error Occured ! $($error[0])"

  } finally {

    if ($excel) {
      $excel.Quit()
    }
  }
}

function rgb($r, $g, $b) {
  return ($b + ($g * 256) + ($r * 65536))
}

function randomColor() {
  $r = $(150..255 | Get-Random)
  $g = $(150..255 | Get-Random)
  $b = $(150..255 | Get-Random)
  return rgb $r $g $b
}

function diffCheck($oneBook, $otherBook, $tBook) {

  trap { Write-Host "[diffCheck]: Error $($_)"; throw $_ }

  $oneBook.Worksheets | % {

    $oneSheet = $_
    $tSheet = $tBook.Worksheets.Item(1)
    $tSheet.Cells.ClearContents()

    if (-not (isExistsWorksheet $otherBook $oneSheet.Name)) {
      Write-Host "$($otherBook.Name) has not $($oneSheet.Name)"
      return
    }
    $otherSheet = $otherBook.Worksheets.Item($oneSheet.Name)

    if ($oneSheet.Visible) {

      Write-Host "Check $($oneSheet.Name) ..."

      # reset color
      $oneSheet.Cells.Interior.ColorIndex = $xlNone
      $otherSheet.Cells.Interior.ColorIndex = $xlNone

      $tmpRange = $tSheet.Range(($oneSheet.Range($oneSheet.UsedRange,
                                 $oneSheet.Range($otherSheet.UsedRange.Address()))).Address())
      # set tmp range to fomula
      $tmpRange.FormulaR1C1 = "=IF('[" + $oneBook.Name + "]" + $oneSheet.Name + "'!RC=" +
                             "'[" + $otherBook.Name + "]" + $otherSheet.Name + "'!RC,"""",1)"

      $cnt = $excel.Application.WorksheetFunction.Count($tmpRange.Cells)

      Write-Host "Different count = [$($cnt)]"

      if ($cnt -gt 0) {

        $tRange = $tmpRange.SpecialCells($xlCellTypeFormulas, $xlNumbers)

        $rangeList = New-Object System.Collections.Generic.List[string]
        $rangeString = ""

        $tRange.Areas | % {

          if ($rangeList.Count -gt 10) {
            if ((($rangeList + $_.Address()) -join ",").Length -lt 255) {
              [void]$rangeList.Add($_.Address())
            } else {
              $rangeString = $rangeList.ToArray() -join ","
              Write-Debug $rangeString.Length
              $color = randomColor
              $oneSheet.Range($rangeString).Interior.Color = $color
              $otherSheet.Range($rangeString).Interior.Color = $color
              $rangeList = New-Object System.Collections.Generic.List[string]
              [void]$rangeList.Add($_.Address())
            }
          } else {
            [void]$rangeList.Add($_.Address())
          }
        }

        if ($rangeList.Count -ne 0) {
          $rangeString = $rangeList.ToArray() -join ","
          Write-Debug $rangeString.Length
          $color = randomColor
          $oneSheet.Range($rangeString).Interior.Color = $color
          $otherSheet.Range($rangeString).Interior.Color = $color
        }
      }

    }
  }
}

function isExistsWorksheet($wBook, $wsName) {

  trap { Write-Host "[isExistsWorksheet]: Error $($_)"; throw $_ }
  $wBook.Worksheets | % {
    if ($_.Name -eq $wsName) {
      return $true
    }
  }
  return $false
}

# call main
Measure-Command { main }