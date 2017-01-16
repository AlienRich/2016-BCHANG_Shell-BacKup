Dim WshShell, bKey
Set WshShell = WScript.CreateObject("WScript.Shell")
Set WSHEnv = WSHShell.Environment
 
'WINSPrimaryServer  or WINSSecondaryServer or WINS server.

strWINS1 = "192.168.100.2"
strWINS2 = "10.7.254.2"

Set cards = GetObject("winmgmts:").InstancesOf ("Win32_NetworkAdapterConfiguration")
 ' myNet = Array("255.255.248.0")
  DNSSrv = Array("192.168.100.200","192.168.100.201","192.168.102.180")
  
  on error resume next 
   
   for each card in cards
    ' find first ip enabled card and set the ip address
    if card. DHCPEnabled  = false then 'to judge the DHCP enable or not.
    if card.IPenabled = true then
     ret = card.SetWinsServer(strWINS1, strWINS2) 
     ret = card.SetDNSServerSearchOrder(DNSSrv)
     exit for
    end If
    End If
   next

    if err.Number <> 0 then
    msgbox "There was an error configuring the network adapter", vbCritical, "Error"
   end if
  on error goto 0
  set cards = nothing
