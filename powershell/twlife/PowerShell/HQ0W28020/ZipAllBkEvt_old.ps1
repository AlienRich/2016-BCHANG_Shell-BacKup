function Add-Zip
{
	param([string]$zipfilename)

	if(-not (test-path($zipfilename)))
	{
		set-content $zipfilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
		(dir $zipfilename).IsReadOnly = $false	
	}
	
	$shellApplication = new-object -com shell.application
	$zipPackage = $shellApplication.NameSpace($zipfilename)
	
	foreach($file in $input) 
	{ 
            $zipPackage.CopyHere($file.FullName)
            Start-sleep -milliseconds 500
	}
}

#Function 變數
$FormDate=get-date -Format yyyyMMdd
$SelectFile=Get-ChildItem -Path D:\EventLog\Event\

#Function 會出LOG及壓縮
$SelectFile | Select-Object FullName,Length | Export-Csv D:\Batch\Event_$FormDate.csv -NoTypeInformation -encoding "unicode"
Get-Item -path D:\EventLog\Event | Add-Zip D:\EventLog\$FormDate.zip

#Function Send-Mail&Delect Csv
Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw -from HQFileServer@twlife.com.tw -Subject "mail test" -Attachments D:\Batch\Event_$FormDate.csv

#Function 刪除
#Remove-Item D:\Batch\*.csv
#Remove-Item D:\EventLog\Event\*.*