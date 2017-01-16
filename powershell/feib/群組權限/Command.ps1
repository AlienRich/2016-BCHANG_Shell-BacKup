
#========================================================================
# Generated By: Anders Wahlqvist
# Website: DollarUnderscore (http://dollarunderscore.azurewebsites.net)
#========================================================================

function Get-LocalGroup
{
    <#
    .SYNOPSIS
    Retrieves local groups on a server

    .DESCRIPTION
    This cmdlet will retrieve all local groups found a server

    .EXAMPLE
    Get-LocalGroup -ComputerName 'MyMachine'

    .PARAMETER ComputerName
    The computer or IP-adress to connect to

    #>

    [CmdletBinding()]
    param(
      [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
      [Alias('hostname')]
      [string] $ComputerName = $env:COMPUTERNAME)


    BEGIN { }

    PROCESS {
        $LocalGroupsWMI=Get-WmiObject -ComputerName $ComputerName -Query "Select * from Win32_Group  Where LocalAccount = $true"
        foreach ($LocalGroup in $LocalGroupsWMI) {
            $returnObj = New-Object System.Object
            $returnObj | Add-Member -Type NoteProperty -Name ComputerName -Value $Localgroup.Domain
            $returnObj | Add-Member -Type NoteProperty -Name GroupName -Value $Localgroup.Name
            $returnObj | Add-Member -Type NoteProperty -Name SID -Value $Localgroup.SID

            Write-Output $returnObj
        }
    }

    END { }
}


function Get-LocalGroupMember
{
    <#
    .SYNOPSIS
    Retrieves local groups on a server

    .DESCRIPTION
    This cmdlet will retrieve all local groups found a server

    .EXAMPLE
    Get-LocalGroup -ComputerName 'MyMachine'

    .PARAMETER ComputerName
    The computer or IP-adress to connect to

    #>

    [CmdletBinding()]
    param(
      [Parameter(Mandatory=$false, ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)]
      [Alias('hostname')]
      [string] $ComputerName = $env:COMPUTERNAME,
      [Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$true)]
      [Alias('Group')]
      [string] $GroupName)

    BEGIN { }

    PROCESS {
        $LocalGroupMembersWMI=Get-WmiObject -ComputerName $ComputerName -Query "SELECT * FROM Win32_GroupUser WHERE GroupComponent=`"Win32_Group.Domain='$ComputerName',Name='$GroupName'`""

        foreach ($LocalGroupWMI in $LocalGroupMembersWMI) {
            $data = $LocalGroupWMI.PartComponent -split "\," 
            $domain = ($data[0] -split "=")[1] -replace """",""
            $name = ($data[1] -split "=")[1] -replace """",""
            $type = ($data[0] -split "_|\.")[1]  -replace "Account"
            $arr += ("$domain\$name").Replace("""","") 

            $returnObj = New-Object System.Object
            $returnObj | Add-Member -Type NoteProperty -Name Domain -Value $domain
            $returnObj | Add-Member -Type NoteProperty -Name GroupName -Value $GroupName
            $returnObj | Add-Member -Type NoteProperty -Name Type -Value $type
            $returnObj | Add-Member -Type NoteProperty -Name Name -Value $name

            Write-Output $returnObj
        }
    }

    END { }
}

Function Get-LocalUserAccount{
[CmdletBinding()]
param (
 [parameter(ValueFromPipeline=$true,
   ValueFromPipelineByPropertyName=$true)]
 [string[]]$ComputerName=$env:computername,
 [string]$UserName
)
foreach ($comp in $ComputerName){

    [ADSI]$server="WinNT://$comp"

    if ($UserName){

            foreach ($User in $UserName){
            $server.children |
            where {$_.schemaclassname -eq "user" -and $_.name -eq $user}
            }    
    }else{

            $server.children |
            where {$_.schemaclassname -eq "user"}
        }
    }
}


# 匯入Server 清單
$computers = get-Content "D:\個人資料\09.工具區\powershell\feib\群組權限\computer-list.txt"

# Lets create our variables
$HostInfo = @()


# 將Server清單匯入
  foreach ($computer in $computers) {
  #File Localation
  $File = "D:\個人資料\09.工具區\powershell\feib\群組權限\" + $computer + ".csv"
  Get-Content "D:\個人資料\09.工具區\powershell\feib\群組權限\account.txt" >> $File
  Get-LocalUserAccount -computername $computer | select path >> $File
  Get-Content "D:\個人資料\09.工具區\powershell\feib\群組權限\Group.txt" >> $File
  Get-localgroup -computer $computer | select Computername,GroupName  >> $File
  Get-LocalGroupMember -computername $computer -Group administrators | Select GroupName,Type,Name,Domain >> $File
  Get-LocalGroupMember -computername $computer -Group "Backup Operators" | Select GroupName,Type,Name,Domain >> $File
       
}

# Lets output to the console
#$HostInfo | out-file D:\Date.txt
#Function Send-Mail
#Send-MailMessage -smtpserver hq0w23074 -to bchang@twlife.com.tw -from SIM@twlife.com.tw -Subject "CheckList" -Encoding ([System.Text.Encoding]::UTF8) -Attachment "date.txt"
#Remove-Item c:\Date.txt
