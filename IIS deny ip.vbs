'/*========================================================================= 
' * Intro VBScriptʹ��ADSIΪIIS����������λ�������ʵ�IP 
' * FileName VBScript-ADSI-IIS-Add-Deny-Grant-IP-Change-MetaBase.xml.vbs 
' *==========================================================================*/ 


'���Ҫ���ε�IP��һ����������IIS�������ã���Ӧ�õ�����վ�� 
'���֮ǰ����Щվ�㵥����������IP���ã���Щ���ò�����Ч�������ܵ���վ������һ�£�Ȼ�󸲸������ӽ�� 
Sub AddDenyIP2All(strDenyIp) 
	On Error Resume Next 
	Set SecObj = GetObject("IIS://LocalHost/W3SVC") 
	Set MyIPSec = SecObj.IPSecurity 
	'Ĭ������������ip���ʣ���������ΪTRUE �����Ĭ���Ǿܾ�����ip��������ΪFalse
	MyIPSec.GrantByDefault = True 
	IPList = MyIPSec.IPDeny 
	i = UBound(IPList) + 1 
	ReDim Preserve IPList(i) 
	IPList(i) = strDenyIp 
	MyIPSec.IPDeny = IPList 
	SecObj.IPSecurity = MyIPSec 
	SecObj.Setinfo 
End Sub 


'��������IP��һ����������IIS�������ã���Ӧ�õ�����վ�� 
'���֮ǰ����Щվ�㵥����������IP���ã���Щ���ò�����Ч�������ܵ���վ������һ�£�Ȼ�󸲸������ӽ�� 
Sub AddGrantIP2All(strGrantIp) 
	On Error Resume Next 
	Set SecObj = GetObject("IIS://LocalHost/W3SVC") 
	Set MyIPSec = SecObj.IPSecurity 
	'Ĭ������������ip���ʣ���������ΪTRUE �����Ĭ���Ǿܾ�����ip��������ΪFalse
	MyIPSec.GrantByDefault = True
	IPList = MyIPSec.IPGrant 
	i = UBound(IPList) + 1 
	ReDim Preserve IPList(i) 
	IPList(i) = strGrantIp 
	MyIPSec.IPGrant = IPList 
	SecObj.IPSecurity = MyIPSec 
	SecObj.Setinfo 
End Sub 


'��ʾIIS�����������ֹ���ʵ�IP 
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