#
# �ʏ��CSV�t�@�C����Excel�ɃC���|�[�g����X�N���v�g
#
 
Param ($FolderName)
 
# �X�N���v�g�̐e�t�H���_�̃p�X���擾����
$path = Split-Path $MyInvocation.MyCommand.Path -Parent
 
# Excel �̕ۑ���̃p�X���쐬����
$xlsPath = Join-Path $path $FolderName | Join-Path -ChildPath "Table.xlsx"
#New-Item $xlsPath -type file -Force
 
# Excel ���N������
$xls = New-Object -ComObject Excel.Application
 

 
 
# WorkBook ��ǉ�����
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
              # �V�[�g��I������
 
              $csvfilepath = Join-Path $path $FolderName | Join-Path -ChildPath ($table + ".csv")
              $xlsxfilepath = Join-Path $path $FolderName | Join-Path -ChildPath ($table + ".xlsx")
              echo $csvfilepath
 
              $wb = $xls.Workbooks.Open($csvfilepath)
#            $wb.SaveAs($xlsxfilepath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookNormal)
              $wb.SaveAs($xlsxfilepath, 51)
 
              $wsMaster = $wbMaster.WorkSheets.Item(1)
              $wb.Worksheets.Item($table).copy($wsMaster)
             

              # WorkBook �����
              $wb.Close()
}
 
$wbMaster.Save()
# WorkBook �����
$wbMaster.Close()
 
# Excel ���I������
$xls.Quit()