'/*========================================================================= 
' * Intro VBScript使用ADSI为IIS批量添加屏蔽或允许访问的IP 
' * FileName VBScript-ADSI-IIS-Add-Deny-Grant-IP-Change-MetaBase.xml.vbs 
' *==========================================================================*/ 


'添加要屏蔽的IP或一组计算机，到IIS公共配置，以应用到所有站点 
'如果之前对有些站点单独做过屏蔽IP设置，在些设置不会生效，得在总的网站上设置一下，然后覆盖所有子结点 
Sub AddDenyIP2All(strDenyIp) 
	On Error Resume Next 
	Set SecObj = GetObject("IIS://LocalHost/W3SVC") 
	Set MyIPSec = SecObj.IPSecurity 
	'默认是允许所有ip访问，所以设置为TRUE ，如果默认是拒绝所有ip，则设置为False
	MyIPSec.GrantByDefault = True 
	IPList = MyIPSec.IPDeny 
	i = UBound(IPList) + 1 
	ReDim Preserve IPList(i) 
	IPList(i) = strDenyIp 
	MyIPSec.IPDeny = IPList 
	SecObj.IPSecurity = MyIPSec 
	SecObj.Setinfo 
End Sub 


'添加允许的IP或一组计算机，到IIS公共配置，以应用到所有站点 
'如果之前对有些站点单独做过屏蔽IP设置，在些设置不会生效，得在总的网站上设置一下，然后覆盖所有子结点 
Sub AddGrantIP2All(strGrantIp) 
	On Error Resume Next 
	Set SecObj = GetObject("IIS://LocalHost/W3SVC") 
	Set MyIPSec = SecObj.IPSecurity 
	'默认是允许所有ip访问，所以设置为TRUE ，如果默认是拒绝所有ip，则设置为False
	MyIPSec.GrantByDefault = True
	IPList = MyIPSec.IPGrant 
	i = UBound(IPList) + 1 
	ReDim Preserve IPList(i) 
	IPList(i) = strGrantIp 
	MyIPSec.IPGrant = IPList 
	SecObj.IPSecurity = MyIPSec 
	SecObj.Setinfo 
End Sub 


'显示IIS公共配置里禁止访问的IP 
Sub ListDenyIP() 
	Set SecObj = GetObject("IIS://LocalHost/W3SVC") 
	Set MyIPSec = SecObj.IPSecurity 
	IPList = MyIPSec.IPDeny 'IPGrant/IPDeny 
	WScript.Echo Join(IPList, vbCrLf) 
	For i = 0 To UBound(IPList) 
		WScript.Echo i + 1 & "-->" & IPList(i) 
	Next 
End Sub 



Sub BatchAddDenyIP2All(readfilepath) 
	On Error Resume Next 
	Set fs = CreateObject("Scripting.FileSystemObject") 
	Set file = fs.OpenTextFile(readfilepath, 1, false)
	Do While file.AtEndOfLine <> True 
		ip = file.ReadLine
		AddDenyIP2All(ip)
	loop
	file.close 
	set fs=nothing 
end Sub


'AddDenyIP2All "192.168.1.106,255.255.255.0" 
'AddDenyIP2All "192.168.163.1" 

'172.18.30.164,255.255.255.0
'192.168.62.1
'192.168.163.1


BatchAddDenyIP2All("ip.txt")

ListDenyIP() 