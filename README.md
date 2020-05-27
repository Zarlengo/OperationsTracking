** Repository for all of the contract cutting scripts **

 KNOWN ISSUES:
 
 * If an error happens on line of:    .Run
	* Happens in Windows 10. Windows Defender has a security block for all of these scripts using VBScript
	* Contact IT and provide the computer name so they can add it to the exclusion list in it's security profile
 * Error connecting to the SQL database
	* User is not on the permission list for the SQL database
	* Contact IT to add the account to the access list for PRODSQLAPP01\PRODSQLAPP01 | CMM_Repository
 * HTA errors
	* Script needs to be run in administrator mode
		* "HKLM\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1\"
		* New DWORD: 1406
		* Value: 0
 
*For any questions contact Chris Zarlengo*

---
# Operations Scripts #

## [All_Operations.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/All_Operations.vbs)
Code summary: [All Operations.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/All%20Operations.md)

VBScript file that provides a single source access to every file used in production. Accessed by the blue plus on operation scripts.

---
## [Operation_00.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_00.vbs)
Code summary: [Operation_00.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_00.md)

VBScript for incoming inspection to scan each of the slugs into the SQL database.

Verifies each slugs is already in the system.


*In Work: Sends an automated email to Kelli Oliver that the slugs are received for her to transact them into AX. Still needs account email access for WKG_Cutting

---
## [Operation_10.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_10.vbs)
Code summary: [Operation_10.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_10.md)

VBScript for part marking. Reads the slug serial number from the slug and outputs the blade serial number to the Dapra marker.
Admin features allow:

1. Duplicate marking (remarking when the part is loaded wrong in the machine), and
2. Cross-out mode which is when the serial number is lined out and the correct one is placed above/below the old one

Traveler scanning was added to ensure the correct paperwork follows the blades during processing

---
## [Operation_20.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_20.vbs)
Code summary: [Operation_20.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_20.md)

VBScript for the tablets located at each WaterJet Machine. This takes the fixture serial number, operator name, blade serial number and load it to the SQL database.

New feature shows the usage levels of abrasive, mixing tube, and orifice. Also shows CMM history of the fixture.

---
## [Operation_30.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_30.vbs)
Code summary: [Operation_30.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_30.md)

VBScript for offset recording. Has the ability to load the results from the offset calculations excel document.

---
## [Operation_40.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_40.vbs)
Code summary: [Operation_40.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_40.md)

VBScript for MRB activities. Create and edit e-tags, and track the part location in the MRB cage

---
## [Operation_50.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_50.vbs)
Code summary: [Operation_50.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_50.md)

VBScript for Final inspection. Loads the results for the blade to the SQL database.

Can also add e-tag information for the blade. Verifies every process has been completed to this point.

---
## [Operation_60.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_60.vbs)
Code summary: [Operation_60.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Operations/Operation_60.md)

VBScript for loading shipping data to the SQL database. Allows tracking for blades into pallet and box.

Verifies all operations are completed and no open e-tag exist. Also checks e-tags to make sure none are marked scrap.

---
# Admin Scripts #


## [CC_Load_AE.xlsm](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/CC_Load_AE.xlsm)
Code summary: [CC_Load_AE.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/CC_Load_AE.md)

Excel macro to upload the serial numbers provided by AeroEdge to the SQL server

---
## [Monitor_All_CMM.hta](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Monitor_All_CMM.hta)
Code summary: [Monitor_All_CMM.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Monitor_All_CMM.md)

HTA script which is running on the television over the CMM station. Reviews all the CMM files and displays real-time information for all of the operators.

---
## [Offset_Calculations.xlsx](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Offset_Calculations.xlsx)
Program summary: [Offset_Calculations.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Offset_Calculations.md)

Excel worksheets to sort the last 2 weeks of CMM data by fixture. This data is sorted to provide offset suggestions for the fixture in question.
The Solver will average the previous 2 weeks of cutting (approximately 3 mixing tubes worth of data) and find the minimum deviation from nominal for the program.

---
## [Search_Files.ps1](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Search_Files.ps1)
Code summary: [Search_Files.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Search_Files.md)

PowerShell file that runs on a Server 2012 machine (sstgkendba10.shapetechnologies.com).

This is set up as a repeating task occurring every 5 minutes. Looks through all the CMM files and uploads new ones into the SQL database.

---
# Extra Scripts #

## [DataBase_Test.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Extra%20Scripts/DataBase_Test.vbs)
Code summary: [DataBase_Test.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Extra%20Scripts/DataBase_Test.md)

VBScript file to test if Windows Defender is blocking the script, it also checks the connection to the SQL database.

---
## [SQL_Test.ps1](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Extra%20Scripts/SQL_Test.ps1)
Code summary: [SQL_Test.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Extra%20Scripts/SQL_Test.md)

Used to test the connection to the AX database and send email option

---
## [TestAXProd.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Extra%20Scripts/TestAXProd.vbs)
Code summary: [TestAXProd.md](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Extra%20Scripts/TestAXProd.md)

Used to test the connection to the AX database and send email option

---
# Obsolete Scripts #

## [Load_CMM.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Obsolete%20Scripts/Load_CMM.vbs)

VBScript version of [Search_Files.ps1](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Search_Files.ps1). Was replaced with PowerShell version to allow script to run without console output.
Allows system to run as a recurring task without needing to be logged in.  Was able to reduce run time to under 20 seconds.

## [Monitor_All_CMM.vbs](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Obsolete%20Scripts/Monitor_All_CMM.vbs)

VBScript version of [Monitor_All_CMM.hta](https://gitlab.com/FlowCorp_CC/contract_cutting/blob/master/Admin/Monitor_All_CMM.hta).
Was replaced with HTA version to allow graphing options, also improves processing time and allows integration of JS libraries.