strComputer = "."
 
Set objWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate}!\\" _
 & strComputer & "\root\cimv2")
 
Set colChassis = objWMIService.ExecQuery _
 ("Select * from Win32_SystemEnclosure")
 
For Each objItem In colChassis
 strChassis = Join(objItem.ChassisTypes, ",")
 
 Select Case strChassis
  Case 0
   strCaseType = "Other"
  Case 1
   strCaseType = "Unknown"
  Case 3
   strCaseType = "Desktop"
  Case 4
   strCaseType = "Low Profile Desktop"
  Case 5
   strCaseType = "Pizza Box"
  Case 6
   strCaseType = "Mini Tower"
  Case 7
   strCaseType = "Tower"
  Case 8
   strCaseType = "Portable"
  Case 9
   strCaseType = "Laptop"
  Case 10
   strCaseType = "Notebook"
  Case 11
   strCaseType = "Hand-held"
  Case 12
   strCaseType = "Docking Station"
  Case 13
   strCaseType = "All-in-one"
  Case 14
   strCaseType = "Sub notebook"
  Case 15
   strCaseType = "Space-saving"
  Case 16
   strCaseType = "Lunch Box"
  Case 17
   strCaseType = "Main System Chassis"
  Case 18
   strCaseType = "Expansion chassis"
  Case 19
   strCaseType = "Sub chassis"
  Case 20
   strCaseType = "Bus Expansion Chassis"
  Case 21
   strCaseType = "Peripheral Chassis"
  Case 22
   strCaseType = "Storage chassis"
  Case 23
   strCaseType = "Rack mount chassis"
  Case 24
   strCaseType = "Sealed-case PC"
  Case 25
   strCaseType ="Multi-system chassis"
  Case 26
   strCaseType ="Compact PCI"
  Case 27
   strCaseType ="Advanced TCA"
  Case 28
   strCaseType ="Blade"
  Case 29
   strCaseType ="Blade Enclosure"
  Case 30
   strCaseType ="Tablet"
  Case 31
   strCaseType ="Convertible"
  Case 32
   strCaseType ="Detachable"
  Case 33
   strCaseType ="IoT Gateway"
  Case 34
   strCaseType ="Embedded PC"
  Case 35
   strCaseType ="Mini PC"
  Case 36
   strCaseType ="Stick PC"
  Case Else
   strCaseType = "Undefined"
 End Select
 WScript.Echo "Computer chassis type: " & strCaseType
Next