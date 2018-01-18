Param(
    [parameter(mandatory=$true)][string]$BeforeStr,
    [parameter(mandatory=$true)][string]$AfterStr,
    [parameter(mandatory=$true)][string]$Comment
)
$Date = Date -Format "yyyy/MM/dd"

$AfterStr = @"
'/ FIX:$Date ${Comment}
${AfterStr}
"@

$FileList = Get-ChildItem */*

foreach($File in $FileList){
    $Content = Get-Content $File
    $LineNumberList = Select-String $BeforeStr $File | Select LineNumber
    foreach($LineNumber in $LineNumberList){
        $Content[[int]$LineNumber.LineNumber-1] = $Content[[int]$LineNumber.LineNumber-1] -Replace  $BeforeStr, $AfterStr
        $Content| Out-File $File -Encoding utf8
    }
}