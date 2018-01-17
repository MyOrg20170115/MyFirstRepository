#
# �ʏ��CSV�t�@�C����Excel�ɃC���|�[�g����X�N���v�g
#

# �X�N���v�g�̐e�t�H���_�̃p�X���擾����
$path = Split-Path $MyInvocation.MyCommand.Path -Parent

# Excel �̕ۑ���̃p�X���쐬����
$xlsPath = Join-Path $path "sample.xlsx"

# Excel ���N������
$xls = New-Object -ComObject Excel.Application

# WorkBook ��ǉ�����
$wb = $xls.WorkBooks.Add()

$sheetAdd = $FALSE

foreach ($table in Get-Content TableList.txt) {

	if ($sheetAdd) {
		[void]$wb.Worksheets.Add()
	}
	
	$sheetAdd = $TRUE

	echo $table
	# CSV �t�@�C���̃p�X���쐬����
	$csvPath = Join-Path $path ( $table + ".csv")

	# CSV �t�@�C���̃G���R�[�f�B���O���w�肷��
	$enc = [System.Text.Encoding]::UTF8

	# CSV �t�@�C�����I�[�v������
	$streamReader = New-Object -TypeName System.IO.StreamReader $csvPath, $enc

	$wb.WorkSheets.Item(1).name = $table

	# �V�[�g��I������
	$ws = $wb.WorkSheets.Item($table)

	# �ϐ�������������
	$i = 1
	$j = 1

	# �P�s���ŏI���R�[�h�܂œǂݍ���
	While (($line = $streamReader.ReadLine()) -ne $null) {

	  # �J���}�ŕ�����𕪊����z��Ɋi�[����
	  $fields = $line.Split(",")

	   # �z������Ԃɏ�������
	   foreach ($field in $fields) {
	 
	    # �Z���̏������u������v�ɂ���
	    $ws.Cells.Item($i, $j).NumberFormat = "@"

	    # �Z���ɒl��ݒ肷��
	    $ws.Cells.Item($i, $j).Value = $field

	    # ����P�i�߂�
	    $j++
	  }

	  # �s���P�i�߂�
	  $i++

	  # �ϐ�������
	  $j = 1

	}

	$lastSheet = $wb.WorkSheets.Item($wb.WorkSheets.Count)
	$wb.Worksheets.Item($table).Move([System.Reflection.Missing]::Value, $lastSheet)
	
	# CSV �t�@�C�������
	$streamReader.Close()

	# �t�@�C���������̏ꍇ�x�����b�Z�[�W��\�����Ȃ��悤�ɂ���
	$xls.DisplayAlerts = $false

	# Excel �t�@�C����ۑ�����
	$wb.SaveAs([ref]$xlsPath.ToString())
	#$wb.SaveAs([ref]$xlsPath.ToString(), -4143) # xls�`���̎��͂�������g�p����

	# �x�����b�Z�[�W�̕\�������ɖ߂�
	$xls.DisplayAlerts = $true

}

# WorkBook �����
$wb.Close()

# Excel ���I������
$xls.Quit()

