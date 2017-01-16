Import-Module Activedirectory

Get-ADComputer -SearchBase 'CN=computers,dc=pegavision,dc=corpnet'  -Filter {name -like 'pc*'} | Move-ADObject -TargetPath 'OU=PC,OU=Device,dc=pegavision,dc=corpnet'