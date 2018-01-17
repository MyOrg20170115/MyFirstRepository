#
# 通常のCSVファイルをExcelにインポートするスクリプト
#
 
Param ($FolderName)
 
# スクリプトの親フォルダのパスを取得する
$path = Split-Path $MyInvocation.MyCommand.Path -Parent
 
# Excel の保存先のパスを作成する
$xlsPath = Join-Path $path $FolderName | Join-Path -ChildPath "Table.xlsx"
#New-Item $xlsPath -type file -Force
 
# Excel を起動する
$xls = New-Object -ComObject Excel.Application
 

 
 
# WorkBook を追加する
$wbMaster = $xls.WorkBooks.Add()
$wbMaster.SaveAs($xlsPath, 51)
$sheetAdd = $FALSE 

 echo $xls.ActiveWorkbook > excel1.txt


$dummypath = Join-Path $path "Book1.xlsx"
$wbdummy = $xls.Workbooks.Open($dummypath)
 
 echo $xls.ActiveWorkbook > excel2.txt
$wbdummy.close() 
  
foreach ($table in Get-Content TableList.txt) {
              if ($sheetAdd) {
                            #[void]$wbMaster.Worksheets.Add()
              }
             
              $sheetAdd = $TRUE
 
 
              #$wbMaster.WorkSheets.Item(1).name = $table
              # シートを選択する
 
              $csvfilepath = Join-Path $path $FolderName | Join-Path -ChildPath ($table + ".csv")
              $xlsxfilepath = Join-Path $path $FolderName | Join-Path -ChildPath ($table + ".xlsx")
              echo $csvfilepath
 
              $wb = $xls.Workbooks.Open($csvfilepath)
#            $wb.SaveAs($xlsxfilepath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookNormal)
              $wb.SaveAs($xlsxfilepath, 51)
 
              $wsMaster = $wbMaster.WorkSheets.Item(1)
              $wb.Worksheets.Item($table).copy($wsMaster)
             

              # WorkBook を閉じる
              $wb.Close()
}
 
$wbMaster.Save()
# WorkBook を閉じる
$wbMaster.Close()
 
# Excel を終了する
$xls.Quit()