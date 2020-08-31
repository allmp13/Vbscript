OsName = GetShortOsName()
If OsName = "Windows 10" Then 
	MsgBox "Exécution avec " & OsName, 64, "Système d'exploitation"
	' Votre code pour Windows 10
    strComputer = "."
 
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" _
 & strComputer & "\root\cimv2")
 
Set colChassis = objWMIService.ExecQuery _
 ("Select * from Win32_ComputerSystem")
For Each objItem In colChassis
 strChassis = objItem.SystemFamily
 WScript.Echo "Computer SystemFamily: " & strChassis
Next

Else
	MsgBox "Exécution avec " & OsName, 64, "Système d'exploitation"
	' Votre code pour les autres systèmes
    Dim wmiObject
Set wmiObject = GetObject( _
 "WINMGMTS:\\.\ROOT\WMI:" + _
 "MS_SystemInformation.InstanceName=""ROOT\\mssmbios\\0000_0""")
Wscript.Echo wmiObject.SystemFamily 'or other property name, see properties


End If

Function GetShortOsName()
	' https://msdn.microsoft.com/fr-fr/library/windows/desktop/ms724832(v=vs.85).aspx
	' https://msdn.microsoft.com/en-us/library/aa394239(v=vs.85).aspx
	' http://www.nogeekleftbehind.com/2013/09/10/updated-list-of-os-version-queries-for-wmi-filters/
	
	For Each objItem in GetObject("winmgmts://./root/cimv2").ExecQuery("Select * from Win32_OperatingSystem",,48)
		version = objItem.Version
		ProductType = objItem.ProductType
	Next

	Select Case Left(version, Instr(version, ".") + 1)
	Case "10.0"
		If (ProductType = "1") Then
			GetShortOsName = "Windows 10"
		Else
			GetShortOsName = "Windows Server 2016"
		End If
	Case "6.3"
		If (ProductType = "1") Then
			GetShortOsName = "Windows 8.1"
		Else
			GetShortOsName = "Windows Server 2012 R2"
		End If
	Case "6.2"
		If (ProductType = "1") Then
			GetShortOsName = "Windows 8"
		Else
			GetShortOsName = "Windows Server 2012"
		End If
	Case "6.1"
		If (ProductType = "1") Then
			GetShortOsName = "Windows 7"
		Else
			GetShortOsName = "Windows Server 2008 R2"
		End If
	Case "6.0"
		If (ProductType = "1") Then
			GetShortOsName = "Windows Vista"
		Else
			GetShortOsName = "Windows Server 2008"
		End If
	Case "5.2"
		If (ProductType = "1") Then
			GetShortOsName = "Windows XP 64-Bit Edition"
		ElseIf (Left(Version, 5) = "5.2.3") Then
			GetShortOsName = "Windows Server 2003 R2"
		Else
			GetShortOsName = "Windows Server 2003"
		End If
	Case "5.1"
		GetShortOsName = "Windows XP"
	Case "5.0"
		GetShortOsName = "Windows 2000"
	End Select
End Function

