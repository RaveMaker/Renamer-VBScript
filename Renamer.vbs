' Script to disable IPv6, Change Computer Name, Join Domain & Place in Specified OU
'
' by RaveMaker - http://ravemaker.net
'
' ---------- Body
RegResultStr = readfromRegistry("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\DisabledComponents", "Blank")
If Not (RegResultStr = "-1") Then 
	disableIPv6
	joinADDomain
Else
	joinADDomain
End if

' ---------- Functions
' Read IPv6 Registry Status
Function readFromRegistry (strRegistryKey, strDefault)
	Dim WSHShell, value
	On Error Resume Next
	Set WSHShell = CreateObject("WScript.Shell")
	value = WSHShell.RegRead(strRegistryKey)
	if err.number <> 0 then
		readFromRegistry = strDefault
	else
		readFromRegistry = value
	end if
	set WSHShell = nothing
End function

' Disable IPv6 if needed
Function disableIPv6
		Dim OperationRegistry
		Set OperationRegistry=WScript.CreateObject("WScript.Shell")
		OperationRegistry.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\DisabledComponents",-1, "REG_DWORD"
		Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
End function

' Join Domain
Function joinADDomain
	' This code renames a computer in AD and on the host itself.
	strDomainUser="user@domain.com"
	strDomainPasswd="password"
	strDomainName="DOMAIN"
	VLAN1="228"
	VLAN2="226"
	VLAN1OU="Faculty"
	VLAN2OU="Students"
	VLAN1strOU = "OU=Faculty,OU=Workstations,DC=DOMAIN,DC=COM"
	VLAN2strOU = "OU=Faculty,OU=Workstations,DC=DOMAIN,DC=COM"
	Const JOIN_DOMAIN = 1
	Const ACCT_CREATE = 2
	Const ACCT_DOMAIN_JOIN_IF_JOINED = 32

	' Connect to Computer
	Set objNet = CreateObject("WScript.network")
	strComputerName = objNet.computerName
	strComputer = strComputerName
	Set objWMILocator = CreateObject("WbemScripting.SWbemLocator")
	objWMILocator.Security_.AuthenticationLevel = 6
	Set objWMIComputer = objWMILocator.ConnectServer(strComputerName,"root\cimv2")
	Set objWMIComputerSystem = objWMIComputer.Get("Win32_ComputerSystem.Name='" & strComputer & "'")

	'now gonna try to get ip and create a name with it, checking all IP addresses
	strcomputer = "."
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	for each objitem in colItems
		strIPAddress = Join(objitem.IPAddress, ",")
		IP = strIPAddress
		if len(IP) < "16" then 
			IPTemp = split(IP,".",4)
			VLAN = IPTemp(2)
			IPFinal = IPTemp(3)
			'Add Zeros (2 Digits - 00x , 0xx) to IP Address
			If IPFinal < 100 Then
				If IPFinal < 10 Then
					IPFinal = "0" & IPFinal
				End If
				IPFinal = "0" & IPFinal
			End If
			
			'Change Computer Name By IP And Select OU
			Select Case VLAN
			Case VLAN1
				strNewName = VLAN1OU & IPFinal
				strOU = VLAN1strOU
				Exit For
			Case VLAN2
				strNewName = VLAN2OU & IPFinal
				strOU = VLAN2strOU
				Exit For
			Case Else
				msgbox ("Wrong VLAN - Contact Administrator")
				WScript.Quit
			End Select
		End If
	next

	'Set New Computer Name
	strNewComputer = strNewName
	'Change Name to capital letters
	strNewComputer = UCase(strNewComputer)
	'Rename Computer if needed
	If Not strComputerName = strNewComputer Then
		rc = objWMIComputerSystem.Rename(strNewComputer, strDomainPasswd, strDomainUser)
		'Successfully renamed
		If rc = 0 Then
			Set OpSysSet = GetObject("winmgmts:{(Shutdown)}!\\.\root\cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
			For Each OpSys In OpSysSet
				OpSys.Reboot()
				WScript.Quit
			Next
		End If
	Else
		'Already Named Right
		Dim strComputer, objNetwork, objWMIService
		Dim colItems, objItem
		Set objNetwork = WScript.CreateObject("WScript.Network")
		strComputer = objNetwork.ComputerName
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
		For Each objItem In colItems
			strTmpDomain = objItem.domain
		Next
		'if domain is not joined - join domain
		If Not strTmpDomain = strDomainName Then
			Set objNetwork = CreateObject("WScript.Network")
			strHostName = objNetwork.ComputerName
			Set objNetwork = Nothing
			Set objWMIComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & strHostName & "\root\cimv2:Win32_ComputerSystem.Name='" & strHostName & "'")
			'Joining computer to domain
			varWMIJoinReturnValue = objWMIComputer.JoinDomainOrWorkGroup(strDomainName, strDomainPasswd, strDomainUser, strOU, JOIN_DOMAIN + ACCT_CREATE)
			If varWMIjoinreturnvalue = 0 Then
				Set OpSysSet = GetObject("winmgmts:{(Shutdown)}!\\.\root\cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")
				For Each OpSys In OpSysSet
					OpSys.Reboot()
					WScript.Quit
				Next
			End If
		End if
	End if
End function
