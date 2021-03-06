#匯入AD模組
Import-Module Activedirectory

#匯入確認帳行清單
$UserAccounts = get-Content C:\Batch\UnLock\UserAccount-list.txt

# 將Server清單匯入
  foreach ($UserAccount in $UserAccounts) {
    #抓取LockOut狀態
    $UserLoutSt=Get-ADUser -Identity $UserAccount -Properties * |%{$_.LockedOut}
    #判斷帳號只要停用立即開啟
    if( $UserLoutSt -like "True" ){
    $UserAccount >>C:\Batch\UnLock\UnLock.txt 
    Unlock-ADAccount $UserAccount }
}

$Message = (Get-Content C:\Batch\UnLock\UnLock.txt) -join "`n"

if ($(Get-Item -Path C:\Batch\UnLock\UnLock.txt).Length -gt 0)
    {Send-MailMessage -SmtpServer TCCGROUP-HT01 -To 5100@tcci.com.tw -From UnLockList@tcci.com.tw -Subject "UnlockList" -body "已被鎖定進行解所的帳號如下:`n$Message" -Encoding ([System.Text.Encoding]::UTF8)
     
     Del C:\Batch\UnLock\UnLock.txt
}

