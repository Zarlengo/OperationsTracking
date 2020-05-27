Option Explicit
 '****** CHANGE THESE SETTINGS *********
 Const debugMode = false	
 Const dateHistory = 20	
 Dim tolName : tolName = Array("Dim 1.1",	"Dim 1.2",	"Dim 2.1",	"Dim 2.2",	"Dim 3.1",	"Dim 3.2",	"Dim 4.1",	"Dim 4.2",	"Dim 5.1",	"Dim 5.2",	"Dim 9.1",	"Dim 9.2",	"Dim 10.1",	"Dim 10.2",	"Dim 11 Max",	"Dim 11 Min",	"Dim 12 Max",	"Dim 12 Min")
 Dim minTol : minTol =   Array(40.8, 		40.8, 		155.3, 		155.3, 		168.2, 		168.2,  	155.3, 		155.3, 		26.9, 		26.9, 		16.75,  	16.75,  	32.25,  	32.25,  	-0.5,  	 		-0.5,  			-0.5,  			-0.5)
 Dim maxTol : maxTol =   Array(41.8, 		41.8, 		156.3, 		156.3,  	169.2,  	169.2,  	156.3,  	156.3,  	27.9,  		27.9,  		17.75,  	17.75,  	99.99,  	99.99,   	 0.5,   		 0.5,   		 0.5,  			 0.5)	
 '***************** Database Settings *******************
 Const dataSource = "PRODSQLAPP01\PRODSQLAPP01"
 Const initialCatalog = "CMM_Repository"								'Initial database
 Const MSNL = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Master Serial Number Listing-AeroEdge.xlsx"
 Const MRBL = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\MRB Review\MRB Inventory.xlsx"
  '**************** INITIAL PARAMETERS *******************
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
 Dim recTotal : recTotal = 0
 '**************** DATABASE CONSTANTS *******************
 Const adOpenStatic			= 3	 '// Uses a static cursor.		
 Const adStateOpen			= 1  '// The object is open		
 '*********************************************************

 'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
 Dim wmi : Set wmi = GetObject("winmgmts:root\cimv2") 
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
 Dim cProcesses, oProcess : For Each oProcess in cProcesses
	oProcess.Terminate()
 Next
		
 'Function to check for access connection and load info from database
 If Load_Access = true Then
	Dim Eff
	Dim objCmd, rs : Set objCmd = GetNewConnection
	
	'Call CMMSearch
	Const tableHeadCnt = 9
	ReDim CCArray(dateHistory + 1, tableHeadCnt)
	Call CCHistory
	Dim tableString : tableString = "<head><style>table, th, td {border: 1px solid black;border-collapse: collapse;text-align: center;}</style></head><body><table><tr>" & _ 
									"<th>Date</th>" & _ 
									"<th>&nbsp;Received&nbsp;<br>Blades</th>" & _ 
									"<th>&nbsp;Unprocessed&nbsp;<br>Blades</th>" & _ 
									"<th>&nbsp;Processed&nbsp;<br>Blades</th>" & _ 
									"<th>&nbsp;Passed&nbsp;<br>Blades</th>" & _ 
									"<th>&nbsp;Failed&nbsp;<br>Blades</th>" & _ 
									"<th>&nbsp;Unshipped&nbsp;<br>Blades</th>" & _ 
									"<th>&nbsp;Shipped&nbsp;<br>Blades</th>" & _ 
									"<th>&nbsp;Yield&nbsp;<br>Rate</th>" & _ 
									"<th>&nbsp;Machine&nbsp;<br>Efficiency</th>"
	
	Dim sqlString : sqlString = "SELECT Count (DISTINCT [MachineName]) FROM [30_Fixtures] WHERE [ActiveFixture] is not null;"
	Set rs = objCmd.Execute(sqlString)
	Dim MachineArray() : ReDim MachineArray(rs(0).value)
	sqlString = "SELECT DISTINCT [MACHINENAME], [MachNomen] FROM [30_Fixtures] WHERE [ActiveFixture] is not NULL ORDER BY [MachNomen];"
	Set rs = objCmd.Execute(sqlString)
	Dim MachineCnt : MachineCnt = -1
	DO WHILE NOT rs.EOF
		tableString = tableString & "<th>&nbsp;" & rs.Fields(1) & "&nbsp;</th>"
		MachineCnt = MachineCnt + 1
		MachineArray(MachineCnt) = rs.Fields(0)
		rs.MoveNext
	Loop
	tableString = tableString & "</tr>"
	ReDim CCMachineArray(dateHistory, MachineCnt)
	Call MachineHistory
	Dim cellColor, dayCnt
	Dim a, b : For a = 0 to dateHistory + 1
		tableString = tableString & "<tr>"
		For b = 0 to tableHeadCnt
			If b = 0 Then
				If a > dateHistory Then
					tableString = tableString & "<td style='text-align:left;'>&nbsp;1 Week Running Total:&nbsp;</td>"
				Else
					tableString = tableString & "<td style='text-align:left;'>&nbsp;" & WeekdayName(Weekday(CCArray(a, 0))) & " " & CCArray(a, b) & "&nbsp;</td>"
				End If
			ElseIf a <= dateHistory or b = 1 or b = 3 or b = 4 or b = 5 or b = 7 Then
				tableString = tableString & "<td>" & CCArray(a, b) & "</td>"
			ElseIf a > dateHistory and b = 8 Then
				tableString = tableString & "<td>" & FormatPercent(CCArray(a, 4) / CCArray(a, 3), 1) & "</td>"
			Else
				tableString = tableString & "<td></td>"
			End If
		Next
		If a <= dateHistory Then
			For b = 0 to MachineCnt
				If Weekday(CCArray(a, 0), 2) < 5 Then
					dayCnt = 40
				Else
					dayCnt = 24
				End If
				Eff = FormatPercent(CCMachineArray(a, b) / dayCnt, 0)
				If CCMachineArray(a, b) / dayCnt > .65 Then
					cellColor = "limegreen"
				ElseIf CCMachineArray(a, b) / dayCnt > .32 Then
					cellColor = "yellow"
				Else
					cellColor = "red"
				End If
				
				tableString = tableString & "<td style='background-color: " & cellColor & ";'>" & Eff & " (" & CCMachineArray(a, b) & ")</td>"
			Next
		End If
		tableString = tableString & "</tr>"
	Next
	tableString = tableString & "</table></body>"
	
	'5) Number of parts in MRB - TODAY ONLY
	sqlString = "SELECT COUNT (*) FROM [40_E-tags] WHERE [Status] = 'Open';"
	Set rs = objCmd.Execute(sqlString)
	Call Send_Email(tableString, rs(0).value)
	
	objCmd.Close
	Set objCmd = Nothing
 End If
 ServerClose()																	'Function to close open connections and return settings back to original	
 Wscript.Quit
 
Sub CMMSearch()
	Dim fileName, sqlUpdate, j, rsUpdate
	Dim Total : Total = 0
	Dim sqlQuery : sqlQuery = "SELECT [40_CMM_LPT5].[File Name], " _
							& "[40_CMM_LPT5].[Dim 1_1], [40_CMM_LPT5].[Dim 1_2], [40_CMM_LPT5].[Dim 2_1], [40_CMM_LPT5].[Dim 2_2], [40_CMM_LPT5].[Dim 3_1], [40_CMM_LPT5].[Dim 3_2], " _
							& "[40_CMM_LPT5].[Dim 4_1], [40_CMM_LPT5].[Dim 4_2], [40_CMM_LPT5].[Dim 5_1], [40_CMM_LPT5].[Dim 5_2], [40_CMM_LPT5].[Dim 9_1], [40_CMM_LPT5].[Dim 9_2], " _
							& "[40_CMM_LPT5].[Dim 10_1], [40_CMM_LPT5].[Dim 10_2], [40_CMM_LPT5].[Dim 11 Max], [40_CMM_LPT5].[Dim 11 Min], [40_CMM_LPT5].[Dim 12 Max], [40_CMM_LPT5].[Dim 12 Min]" _
							& " FROM [40_CMM_LPT5] WHERE [Failures] IS NULL;"
	
	Set rs = objCmd.Execute(sqlQuery)
	DO WHILE NOT rs.EOF
		ReDim toleranceArray(UBound(tolName))
		Total = Total + 1
		fileName = rs.Fields(0) '"Blade"
		For j = 0 to 17
			If Not IsNull(rs.Fields(j + 1)) Then toleranceArray(j) = rs.Fields(j + 1)
		Next
		sqlUpdate = "UPDATE [40_CMM_LPT5] SET [Failures]='" & toleranceCheck(toleranceArray) & "' WHERE  [File Name]='" & fileName & "';"
		'msgbox sqlUpdate
		set rsUpdate = objCmd.Execute(sqlUpdate)
		rs.MoveNext
	Loop	
	Set rs = Nothing
 End Sub

Sub CCHistory()
	Dim sqlString(6), total, yield, colX, CurDate, NextDate, MaxHours
	Dim dateX : For dateX = dateHistory to 0 step -1
		CurDate = Date - dateX - 1
		NextDate = Date - dateX
		CCArray(dateX, 0) = CurDate
		'0) Number of blades scanned in 00_Initial
		sqlString(0) = "SELECT COUNT(*) FROM [00_Initial] WHERE [Slug Inspection Date] >= '" & CurDate + 0.15 & "' AND [Slug Inspection Date] < '" & NextDate + 0.15 & "';"
		'1) Number of blades not cut
		sqlString(1) =  "SELECT COUNT(*) " & _
						"FROM ([00_Initial] RIGHT JOIN [00_AE_SN_Control] ON [00_Initial].[Slug S/N] = [00_AE_SN_Control].[Slug Serial Number]) " & _
						"LEFT JOIN [50_Final] ON [50_Final].[Blade S/N] = [00_AE_SN_Control].[Blade Serial Number] " & _
						"WHERE ([00_Initial].[Slug Inspection Date] >= GetDate() - 60 and [00_Initial].[Slug Inspection Date] <= '" & NextDate + 0.15 & "') and " & _
						"([50_Final].[Blade S/N] IS NULL or [50_Final].[Blade Inspected Date] >= '" & CurDate + 0.15 & "');"
		'2) Number of blades in CMM
		sqlString(2) = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Date] >= '" & CurDate + 0.15 & "' AND [Date] < '" & NextDate + 0.15 & "';"
		'3) Number of blades in CMM Pass
		sqlString(3) = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Date] >= '" & CurDate + 0.15 & "' AND [Date] < '" & NextDate + 0.15 & "' AND FAILURES = 0;"
		'4) Number of blades in CMM Fail
		sqlString(4) = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Date] >= '" & CurDate + 0.15 & "' AND [Date] < '" & NextDate + 0.15 & "' AND FAILURES > 0;"
		'5) Number of blades unshipped
		sqlString(5) =  "SELECT COUNT(*) " & _
						"FROM [50_Final] " & _
						"LEFT JOIN [60_Shipping] ON [50_Final].[Blade S/N] = [60_Shipping].[Blade Serial Number] " & _
						"WHERE ([60_Shipping].[Date Shipped] >= '" & CurDate + 0.15 & "' or [60_Shipping].[Blade Serial Number] IS NULL) " & _
						"and ([50_Final].[Blade Inspected Date] >= GetDate() - 30 and [50_Final].[Blade Inspected Date] <= '" & CurDate + 0.15 & "' and [50_Final].[Accepted Y/N] = 'Y');"
		'6) Number of blades shipped
		sqlString(6) = "SELECT COUNT(*) FROM [60_Shipping] WHERE [Date Shipped] = '" & CurDate & "';"
		'X)E-Tags Cleared?
		'sqlString(5) = "SELECT COUNT(*) FROM [40_E-Tags] WHERE [Close Date] >= '" & CurDate & "' AND [Close Date] >= '" & NextDate & "';"
		For colX = 0 to UBound(sqlString)
			Set rs = objCmd.Execute(sqlString(colX))
			If colX = 0 Then
				CCArray(dateX, colX + 1) = rs(0).value * 2
				CCArray(dateHistory + 1, colX + 1) = CCArray(dateHistory + 1, colX + 1)  + rs(0).value * 2
			Else
				CCArray(dateX, colX + 1) = rs(0).value
				CCArray(dateHistory + 1, colX + 1) = CCArray(dateHistory + 1, colX + 1)  + rs(0).value
			End If
			If colX = 0 Then
				If Weekday(CurDate, 2) < 5 Then
					MaxHours = 20 * 2 * 5
				Else
					MaxHours = 12 * 2 * 5
				End If
			End If
			If colX = 2 Then
				total = rs(0).value
				CCArray(dateX, 9) = FormatPercent(total / MaxHours, 0)
			End If
			If colX = 3 and total <> 0 Then
				CCArray(dateX, 8) = FormatPercent(rs(0).value / total, 0)
			ElseIf colX = 3 Then
				CCArray(dateX, 8) = FormatPercent(0, 0)
			End If
		Next
	Next
 End Sub
	 
Function getInvoiceCount()
	getInvoiceCount = "<table><tr><th>&nbsp;PO Name&nbsp;</th><th>&nbsp;PO Quantity&nbsp;</th><th>&nbsp;Shipped Quantity&nbsp;</th></tr>"
	Dim sqlString : sqlString = "SELECT [PONumber], [POQuantity] FROM [60_PO] WHERE [POFilled] = 0;"
	Dim rsCnt, sqlInvoiceString
	Set rs = objCmd.Execute(sqlString)
	Do While not rs.EOF
		getInvoiceCount = getInvoiceCount & "<tr><td>&nbsp;" & rs(0).value & "&nbsp;</td><td>&nbsp;" & rs(1).value & "&nbsp;</td><td>&nbsp;"
		sqlInvoiceString = "SELECT COUNT(*) FROM [60_Shipping] WHERE [AE PO Number] = '" & rs(0).value & "';"
		Set rsCnt = objCmd.Execute(sqlInvoiceString)
		getInvoiceCount = getInvoiceCount & rsCnt(0).value & "&nbsp;</td></tr>"
		rs.MoveNext
	Loop
	getInvoiceCount = getInvoiceCount & "</table>"
 End Function

Function toleranceCheck(toleranceArray)
	toleranceCheck = 0
	Dim n : For n = lbound(toleranceArray) to ubound(toleranceArray)
		If IsNull(toleranceArray(n)) or IsEmpty(toleranceArray(n)) Then
		ElseIf toleranceArray(n) < minTol(n) or toleranceArray(n) > maxTol(n) Then
			toleranceCheck = toleranceCheck + 1
		End If
	Next
 End Function

Sub MachineHistory
	Dim FixtureID, dateX, sqlFixtureString, rsFixture, a
	Dim sqlString : sqlString = "SELECT [MachineName], [FixtureID] FROM [30_Fixtures] WHERE [ActiveFixture] is not null;"
	Set rs = objCmd.Execute(sqlString)
	Do While not rs.EOF
		FixtureID = rs.Fields(1)
		For a = 0 to UBound(MachineArray)
			if MachineArray(a) = rs.Fields(0) Then Exit For
		Next
		For dateX = dateHistory to 0 step -1
			sqlFixtureString = "SELECT COUNT(*) FROM [20_LPT5] WHERE [Fixture Location] = '" & FixtureID & "' and [Cut Date] >= '" & Date - dateX - 1 & "' and [Cut Date] < '" & Date - dateX & "';"
			rsFixture = objCmd.Execute(sqlFixtureString)
			CCMachineArray(dateX, a) = CCMachineArray(dateX, a) + rsFixture(0).value * 2
		Next
		rs.MoveNext
	Loop
 End Sub
 
Sub Send_Email(Message, MRBCount)
	' Exit Sub
	Dim MyEmail : Set MyEmail=CreateObject("CDO.Message")
	Dim bodyPre : bodyPre = "<p><span style='font-size:12pt; color:red'>This is an automatically generated daily email, please do not reply to sender. Email <a href=""mailto:CZarlengo@flowcorp.com"">Chris Zarlengo</a> if you have any issues.</span></p><br>"
	Dim body : body =  Message & _
		 "<br><p><span style='font-size:12pt;'>There are currently " & MRBCount & " blades in MRB awaiting disposition.</span>"
	Dim invoiceTable : invoiceTable = getInvoiceCount()
	Dim Signature : Signature = "<footer><div>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Chris Zarlengo</span><span style='color:#1F497D'></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>Manufacturing Engineer</span><span style='color:#1F497D'></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Flow International Corporation | <a href=""http://www.flowwaterjet.com/"">http://www.FlowWaterjet.com/</a></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>23500 64th Ave. S. | Kent | Washington | 98032 | USA</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>253-246-3741 | <a href=""mailto:CZarlengo@flowcorp.com"">CZarlengo@flowcorp.com</a><br>" _
		& "</div></footer>"
	
	MyEmail.Subject="Contract cutting daily summary"
	MyEmail.From="czarlengo@flowcorp.com"
	'MyEmail.To="dhaensel@shapetechnologies.com;HFonda@flowcorp.com;Corentin.ryo@aquarese.fr"
	MyEmail.BCC="czarlengo@flowcorp.com"
	MyEmail.HTMLBody = bodyPre & body & invoiceTable & Signature

	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

	'SMTP Server
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="SKENEXC60.flowcorp.com"

	'SMTP Port
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

	'SMTP Auth (For Windows Auth set this to 2)
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=2

	MyEmail.Configuration.Fields.Update
	MyEmail.Send

	set MyEmail=nothing


 End Sub
  
Function GetNewConnection()
	Dim objCmd : Set objCmd = CreateObject("ADODB.Connection")
	Dim sConnection : sConnection = "Data Source=" & dataSource & ";Initial Catalog=" & initialCatalog & ";Integrated Security=SSPI;"
	Dim sProvider : sProvider = "SQLOLEDB.1;"
	
	
	objCmd.ConnectionString	= sConnection	'Contains the information used to establish a connection to a data store.
	objCmd.Provider = sProvider				'Indicates the name of the provider used by the connection.
	objCmd.CursorLocation = adOpenStatic	'Sets or returns a value determining who provides cursor functionality.
	If debugMode = False Then On Error Resume Next
    objCmd.Open
	On Error GoTo 0 
	If objCmd.State = adStateOpen Then  
        Set GetNewConnection = objCmd  
	Else
        Set GetNewConnection = Nothing
    End If  
 End Function

Function Load_Access()
	Dim objCmd : set objCmd = GetNewConnection
	If objCmd is Nothing Then Load_Access = false : Exit Function
	objCmd.Close
	Set objCmd = Nothing
	Load_Access = true
 End Function

'// EXIT SCRIPT
Sub ServerClose()
	If debugMode = False Then On Error Resume Next

	WScript.Sleep 1000  '// REQUIRED OR ERRORS
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"
		
	On Error GoTo 0
    Wscript.Quit
 End Sub
 