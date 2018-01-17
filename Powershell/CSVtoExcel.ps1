#
# 通常のCSVファイルをExcelにインポートするスクリプト
#

# スクリプトの親フォルダのパスを取得する
$path = Split-Path $MyInvocation.MyCommand.Path -Parent

# Excel の保存先のパスを作成する
$xlsPath = Join-Path $path "sample.xlsx"

# Excel を起動する
$xls = New-Object -ComObject Excel.Application

# WorkBook を追加する
$wb = $xls.WorkBooks.Add()

$sheetAdd = $FALSE

foreach ($table in Get-Content TableList.txt) {

	if ($sheetAdd) {
		[void]$wb.Worksheets.Add()
	}
	
	$sheetAdd = $TRUE

	echo $table
	# CSV ファイルのパスを作成する
	$csvPath = Join-Path $path ( $table + ".csv")

	# CSV ファイルのエンコーディングを指定する
	$enc = [System.Text.Encoding]::UTF8

	# CSV ファイルをオープンする
	$streamReader = New-Object -TypeName System.IO.StreamReader $csvPath, $enc

	$wb.WorkSheets.Item(1).name = $table

	# シートを選択する
	$ws = $wb.WorkSheets.Item($table)

	# 変数を初期化する
	$i = 1
	$j = 1

	# １行ずつ最終レコードまで読み込む
	While (($line = $streamReader.ReadLine()) -ne $null) {

	  # カンマで文字列を分割し配列に格納する
	  $fields = $line.Split(",")

	   # 配列を順番に処理する
	   foreach ($field in $fields) {
	 
	    # セルの書式を「文字列」にする
	    $ws.Cells.Item($i, $j).NumberFormat = "@"

	    # セルに値を設定する
	    $ws.Cells.Item($i, $j).Value = $field

	    # 列を１つ進める
	    $j++
	  }

	  # 行を１つ進める
	  $i++

	  # 変数初期化
	  $j = 1

	}

	$lastSheet = $wb.WorkSheets.Item($wb.WorkSheets.Count)
	$wb.Worksheets.Item($table).Move([System.Reflection.Missing]::Value, $lastSheet)
	
	# CSV ファイルを閉じる
	$streamReader.Close()

	# ファイルが既存の場合警告メッセージを表示しないようにする
	$xls.DisplayAlerts = $false

	# Excel ファイルを保存する
	$wb.SaveAs([ref]$xlsPath.ToString())
	#$wb.SaveAs([ref]$xlsPath.ToString(), -4143) # xls形式の時はこちらを使用する

	# 警告メッセージの表示を元に戻す
	$xls.DisplayAlerts = $true

}

# WorkBook を閉じる
$wb.Close()

# Excel を終了する
$xls.Quit()

