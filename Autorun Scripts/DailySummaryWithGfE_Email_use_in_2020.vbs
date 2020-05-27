Option Explicit
 '****** REVISION HISTORY **************
 '	1.0		2/28/2019	Initial Release
 '	2.0		7/9/2019	Updated for GfE separation from ATEP
 '****** CHANGE THESE SETTINGS *********
 Const debugMode = false	
 Const dateHistory = 6
 Const dateStep = 1
 '***************** Database Settings *******************
 Const dataSource = "PRODSQLAPP01\PRODSQLAPP01"
 Const initialCatalog = "CMM_Repository"								'Initial database
  '**************** INITIAL PARAMETERS *******************
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
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
	Dim objCmd, rs, sqlString, MachineArray(), CCMachineArray
	Set objCmd = GetNewConnection
	
	Const tableHeadCnt = 10
	ReDim CCArray(dateHistory + 1, tableHeadCnt)
	Dim emailString : emailString = "<head><style>table, th, td {border: 1px solid black;border-collapse: collapse;text-align: center;}</style></head><body>"
	Call CCHistory("060052-1")
	emailString = emailString & "<span style='font-size:160%;font-weight: bold;'>ATEP Material:</span><table>" & createTable("060052-1") & "</table><br><br>"
	Call CCHistory("062084-1")
	emailString = emailString & "<span style='font-size:160%;font-weight: bold;'>GfE Material:</span><table>" & createTable("062084-1") & "</table></body>"
	
	'5) Number of parts in MRB - TODAY ONLY
	sqlString = "SELECT COUNT (*) FROM [40_E-tags] WHERE [Status] = 'Open';"
	Set rs = objCmd.Execute(sqlString)
	Call Send_Email(emailString, rs(0).value)
	
	objCmd.Close
	Set objCmd = Nothing
 End If
 ServerClose()																	'Function to close open connections and return settings back to original	
 Wscript.Quit
 
Function createTable(slugPN)
	Dim tableString : tableString = "<tr>" & _ 
										"<th>Date</th>" & _ 
										"<th>&nbsp;Received&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;Unprocessed&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;Processed&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;Passed&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;Failed&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;Unshipped&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;Shipped&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;In MRB&nbsp;</th>" & _ 
										"<th>&nbsp;Yield&nbsp;<br>Rate</th>" & _ 
										"<th>&nbsp;Machine&nbsp;<br>Efficiency</th>"
		
	sqlString = "SELECT Count (DISTINCT [MachineName]) FROM [30_Fixtures] WHERE [ActiveFixture] is not null;"
	Set rs = objCmd.Execute(sqlString)
	ReDim MachineArray(rs(0).value)
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
	Call MachineHistory(slugPN)
	Dim cellColor, dayCnt
	Dim a, b : For a = 0 to dateHistory + 1 Step dateStep
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
					dayCnt = 1128 / 7 / 11 * 2
				Else
					dayCnt = 1128 / 7 / 11
				End If
				'dayCnt = 1128 / MachineCnt
				'1128 / 7 EA WK
				Eff = FormatPercent(CCMachineArray(a, b) / dayCnt, 0)
				If CCMachineArray(a, b) / dayCnt > .66 Then
					cellColor = "limegreen"
				ElseIf CCMachineArray(a, b) / dayCnt > .33 Then
					cellColor = "yellow"
				Else
					cellColor = "red"
				End If
				
				tableString = tableString & "<td style='background-color: " & cellColor & ";'>" & Eff & " (" & CCMachineArray(a, b) & ")</td>"
			Next
		End If
		tableString = tableString & "</tr>"
	Next
	createTable = tableString
 End Function

Sub CCHistory(slugPN)
	Dim sqlString(7), total, yield, colX, CurDate, NextDate, MaxHours
	Dim dateX : For dateX = dateHistory to 0 step -dateStep
		CurDate = Date - dateX - dateStep -1
		NextDate = Date - dateX
		CCArray(dateX, 0) = CurDate
		'0) Number of blades scanned in 00_Initial
		sqlString(0) = "SELECT COUNT(*) FROM [00_Initial] LEFT JOIN [00_AE_SN_Control] ON [00_Initial].[Slug S/N] = [00_AE_SN_Control].[Slug Serial Number] WHERE [Slug Inspection Date] >= '" & CurDate + 0.15 & "' AND [Slug Inspection Date] < '" & NextDate + 0.15 & "' AND [FIC Slug Part Number] = '" & slugPN & "';"
		'1) Number of blades not cut
		sqlString(1) =  "SELECT COUNT(*) " & _
						"FROM ([00_Initial] RIGHT JOIN [00_AE_SN_Control] ON [00_Initial].[Slug S/N] = [00_AE_SN_Control].[Slug Serial Number]) " & _
						"LEFT JOIN [50_Final] ON [50_Final].[Blade S/N] = [00_AE_SN_Control].[Blade Serial Number] " & _
						"WHERE ([00_Initial].[Slug Inspection Date] >= GetDate() - 60 and [00_Initial].[Slug Inspection Date] <= '" & NextDate + 0.15 & "') and " & _
						"([50_Final].[Blade S/N] IS NULL or [50_Final].[Blade Inspected Date] >= '" & CurDate + 0.15 & "') AND [FIC Slug Part Number] = '" & slugPN & "';"
		'2) Number of blades in CMM
		sqlString(2) = "SELECT COUNT(*) FROM [40_CMM_LPT5] LEFT JOIN [00_AE_SN_Control] ON [40_CMM_LPT5].[Serial Number] = [00_AE_SN_Control].[Blade Serial Number] WHERE [Date] >= '" & CurDate + 0.15 & "' AND [Date] < '" & NextDate + 0.15 & "' AND [FIC Slug Part Number] = '" & slugPN & "';"
		'3) Number of blades in CMM Pass
		sqlString(3) = "SELECT COUNT(*) FROM [40_CMM_LPT5] LEFT JOIN [00_AE_SN_Control] ON [40_CMM_LPT5].[Serial Number] = [00_AE_SN_Control].[Blade Serial Number] WHERE [Date] >= '" & CurDate + 0.15 & "' AND [Date] < '" & NextDate + 0.15 & "' AND FAILURES = 0 AND [FIC Slug Part Number] = '" & slugPN & "';"
		'4) Number of blades in CMM Fail
		sqlString(4) = "SELECT COUNT(*) FROM [40_CMM_LPT5] LEFT JOIN [00_AE_SN_Control] ON [40_CMM_LPT5].[Serial Number] = [00_AE_SN_Control].[Blade Serial Number] WHERE [Date] >= '" & CurDate + 0.15 & "' AND [Date] < '" & NextDate + 0.15 & "' AND FAILURES > 0 AND [FIC Slug Part Number] = '" & slugPN & "';"
		'5) Number of blades unshipped
		sqlString(5) =  "SELECT COUNT(*) " & _
						"FROM [50_Final] " & _
						"LEFT JOIN [60_Shipping] ON [50_Final].[Blade S/N] = [60_Shipping].[Blade Serial Number] " & _
						"LEFT JOIN [00_AE_SN_Control] ON [50_Final].[Blade S/N] = [00_AE_SN_Control].[Blade Serial Number] " & _
						"WHERE ([60_Shipping].[Date Shipped] >= '" & CurDate + 0.15 & "' or [60_Shipping].[Blade Serial Number] IS NULL) " & _
						"and ([50_Final].[Blade Inspected Date] >= GetDate() - 30 and [50_Final].[Blade Inspected Date] <= '" & CurDate + 0.15 & "' and [50_Final].[Accepted Y/N] = 'Y' AND [FIC Slug Part Number] = '" & slugPN & "');"
		'6) Number of blades shipped
		sqlString(6) = "SELECT COUNT(*) FROM [60_Shipping] LEFT JOIN [00_AE_SN_Control] ON [60_Shipping].[Blade Serial Number] = [00_AE_SN_Control].[Blade Serial Number] WHERE [Date Shipped] = '" & CurDate & "' AND [FIC Slug Part Number] = '" & slugPN & "';"
		'7)Parts waiting on E-Tags
		sqlString(7) = "SELECT COUNT(*) " & _
						"FROM [40_E-tags] " & _
						"LEFT JOIN [40_Rejections] ON LEFT([40_Rejections].[Tag Numbers], 6) = [40_E-tags].[Tag Number] " & _
						"LEFT JOIN [00_AE_SN_Control] ON [00_AE_SN_Control].[Blade Serial Number] = [Serial Number]" & _
						"WHERE [40_E-tags].[Status] = 'Open' AND [FIC Slug Part Number] = '" & slugPN & "';"
		For colX = 0 to UBound(sqlString)
			'msgbox sqlString(colX)
			Set rs = objCmd.Execute(sqlString(colX))
			If colX = 0 Then
				CCArray(dateX, colX + 1) = rs(0).value * 2
				CCArray(dateHistory + 1, colX + 1) = CCArray(dateHistory + 1, colX + 1)  + rs(0).value * 2
			Else
				CCArray(dateX, colX + 1) = rs(0).value
				CCArray(dateHistory + 1, colX + 1) = CCArray(dateHistory + 1, colX + 1)  + rs(0).value
			End If
			If colX = 0 Then
				MaxHours = 1128
			End If
			If colX = 2 Then
				total = rs(0).value
				CCArray(dateX, 10) = FormatPercent(total / MaxHours, 0)
			End If
			If colX = 3 and total <> 0 Then
				CCArray(dateX, 9) = FormatPercent(rs(0).value / total, 0)
			ElseIf colX = 3 Then
				CCArray(dateX, 9) = FormatPercent(0, 0)
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


Sub MachineHistory(slugPN)
	Dim FixtureID, dateX, sqlFixtureString, rsFixture, a
	Dim sqlString : sqlString = "SELECT [MachineName], [FixtureID] FROM [30_Fixtures] WHERE [ActiveFixture] is not null;"
	Set rs = objCmd.Execute(sqlString)
	Do While not rs.EOF
		FixtureID = rs.Fields(1)
		For a = 0 to UBound(MachineArray)
			if MachineArray(a) = rs.Fields(0) Then Exit For
		Next
		For dateX = dateHistory to 0 step -1
			sqlFixtureString = "SELECT COUNT(*) FROM [20_LPT5] LEFT JOIN [00_AE_SN_Control] ON [20_LPT5].[Blade SN Dash 1] = [00_AE_SN_Control].[Blade Serial Number] WHERE [Fixture Location] = '" & FixtureID & "' and [Cut Date] >= '" & Date - dateX - dateStep & "' and [Cut Date] < '" & Date - dateX & "' AND [FIC Slug Part Number] = '" & slugPN & "';"
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
 