#Function �ܼ�
$SelectFile=Get-ChildItem -Path "D:\������Ƨ�\��T�A�Ȥ���-�t�Υ��x��\USB��ƥ洫��\" -Recurse | where {$_.CreationTime -lt (get-date).adddays(-3)}  | Where {!$_.PSIsContainer}

#Function Move�ɮ�
$SelectFile | Remove-Item

