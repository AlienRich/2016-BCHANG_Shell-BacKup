Import-Module Activedirectory
$inputFile = Import-CSV  C:\Batch\ImportAD\2.csv

foreach($element in $inputFile)
{
    $CN=$element.cn
    $OU="cn=" + "$CN" + ",OU=水泥分廠,OU=KL,OU=TCC,DC=taiwancement,DC=com"
    $upn="$CN" + "@taiwancement.com"
    $Lastname=$element.LastName
    $FirstName=$element.FirstName
    $password=$element.password
    $title=$element.TiTle
    $display=$element.display
    $company=$element.company
    $Description=$element.Description
    $Department=$element.Department
    $Tel=$element.Tel
    
    #dsadd user $OU -disabled no -upn $upn -fn $FirstName -ln $Lastname -display $display -pwd $password -company $company -dept $Department -title $title -desc $Description -tel $tel 
    Rename-ADObject -Identity $OU -NewName $display
}


