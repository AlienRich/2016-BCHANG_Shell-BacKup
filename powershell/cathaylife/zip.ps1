#設定時間
#$Date = Get-date -Format yyyy-MM-dd
$Date = "2015-07-07"
$Source = "\\cxlsvr239\d$\FtpDrop" 
#===================================
#刪除作業

#設定刪除天數
#$DelDate = (get-date).adddays(-29)  

#抓取要刪除的資料
#$SelectFile=Get-ChildItem -Path $Source | where {$_.LastWriteTime -lt $DelDate}

#刪除資料
#$SelectFile | Remove-Item
#===================================

#===================================
#ZIP作業


$SourceFolder    = "$source\$Date"
$DestinationFile = "$source\$Date.zip"
$Compression     = "Optimal"  # Optimal, Fastest, NoCompression

Zip-Directory -DestinationFileName $DestinationFile `
    -SourceDirectory $SourceFolder `
    -CompressionLevel $Compression ` #Optional parameter

$SourceFolder | Remove-Item -Recurse  

function Zip-Directory {
    Param(
      [Parameter(Mandatory=$True)][string]$DestinationFileName,
      [Parameter(Mandatory=$True)][string]$SourceDirectory,
      [Parameter(Mandatory=$False)][string]$CompressionLevel = "Optimal",
      [Parameter(Mandatory=$False)][switch]$IncludeParentDir
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $CompressionLevel    = [System.IO.Compression.CompressionLevel]::$CompressionLevel  
    [System.IO.Compression.ZipFile]::CreateFromDirectory($SourceDirectory, $DestinationFileName, $CompressionLevel, $IncludeParentDir)
}