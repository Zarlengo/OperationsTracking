Option Explicit
 'Run only on Monday
 If Weekday(Now, 2) <> 1 Then Wscript.Quit

 '****** REVISION HISTORY **************
 '	1.0		2/28/2019	Initial Release
 '	2.0		7/9/2019	Updated for GfE separation from ATEP
 '****** CHANGE THESE SETTINGS *********
 Const debugMode = false	
 Const dateHistory = 7
 Const dateStep = 6
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
	ReDim CCArray(dateHistory, tableHeadCnt)
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
										"<th>&nbsp;E-Tags&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;In Scrap&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;Shipped&nbsp;<br>Blades</th>" & _ 
										"<th>&nbsp;On Hand&nbsp;</th>"
	tableString = tableString & "</tr>"
	Dim a, b : For a = 0 to dateHistory
		tableString = tableString & "<tr>"
		tableString = tableString & "<td style='text-align:left;'>&nbsp;" & CCArray(a, 0) & " to " & CCArray(a, 0) + dateStep & "&nbsp;</td>"
		For b = 0 to 4
			tableString = tableString & "<td>" & CCArray(a, b + 1) & "</td>"
		Next
		tableString = tableString & "</tr>"
	Next
	createTable = tableString
 End Function

Sub CCHistory(slugPN)
	Dim sqlString(4), total, yield, colX, CurDate, NextDate, MaxHours
	Dim dateX : For dateX = 0 to dateHistory
		NextDate = Date - Weekday(Date, 1) - (dateHistory - dateX) * 7
		CurDate = NextDate - dateStep
		CCArray(dateX, 0) = CurDate
		'0) Number of blades scanned in 00_Initial
		sqlString(0) = "SELECT COUNT(*) FROM [00_Initial] LEFT JOIN [00_AE_SN_Control] ON [00_Initial].[Slug S/N] = [00_AE_SN_Control].[Slug Serial Number] WHERE [Slug Inspection Date] >= '" & CurDate & "' AND [Slug Inspection Date] < '" & NextDate + 1 & "' AND [FIC Slug Part Number] = '" & slugPN & "';"
		'1) Parts waiting on E-Tags
		sqlString(1) = "SELECT COUNT(*) " & _
						"FROM [40_E-tags] " & _
						"LEFT JOIN [40_Rejections] ON LEFT([40_Rejections].[Tag Numbers], 6) = [40_E-tags].[Tag Number] " & _
						"LEFT JOIN [00_AE_SN_Control] ON [00_AE_SN_Control].[Blade Serial Number] = [Serial Number]" & _
						"WHERE ([40_E-tags].[Status] = 'Open' or ([40_E-tags].[Status] = 'Closed' AND [Close Date] >= '" & CurDate & "')) AND [FIC Slug Part Number] = '" & slugPN & "' AND [Open Date] < '" & NextDate + 1 & "';"
		'2) Parts in Scrap
		sqlString(2) = "SELECT COUNT(*) " & _
						"FROM [40_E-tags] " & _
						"LEFT JOIN [40_Rejections] ON LEFT([40_Rejections].[Tag Numbers], 6) = [40_E-tags].[Tag Number] " & _
						"LEFT JOIN [00_AE_SN_Control] ON [00_AE_SN_Control].[Blade Serial Number] = [Serial Number]" & _
						"LEFT JOIN [60_Shipping] ON [60_Shipping].[Blade Serial Number] = [Serial Number]" & _
						"WHERE [40_E-tags].[Status] = 'Closed' AND [40_E-tags].[Disposition] = 'Scrap' AND [FIC Slug Part Number] = '" & slugPN & "' AND [Open Date] < '" & NextDate + 1 & "' AND [Close Date] < '" & NextDate + 1 & "' AND [Date Shipped] IS NULL;"
		'3) Number of blades shipped
		sqlString(3) = "SELECT COUNT(*) FROM [60_Shipping] LEFT JOIN [00_AE_SN_Control] ON [60_Shipping].[Blade Serial Number] = [00_AE_SN_Control].[Blade Serial Number] WHERE [Date Shipped] >= '" & CurDate & "' AND [Date Shipped] < '" & NextDate + 1 & "' AND [FIC Slug Part Number] = '" & slugPN & "';"
		'4) Number of blades on hand
		sqlString(4) =  "SELECT COUNT(*) " & _
						"FROM ([00_AE_SN_Control] " & _
						"LEFT JOIN [00_Initial] ON [00_Initial].[Slug S/N] = [00_AE_SN_Control].[Slug Serial Number]) " & _
						"LEFT JOIN [60_Shipping] ON [60_Shipping].[Blade Serial Number] = [00_AE_SN_Control].[Blade Serial Number] " & _
						"WHERE [00_Initial].[Slug Inspection Date] < '" & CurDate & "' AND ([Date Shipped] IS NULL OR [Date Shipped] >= '" & CurDate & "') AND [FIC Slug Part Number] = '" & slugPN & "';"
		For colX = 0 to UBound(sqlString)
			'msgbox sqlString(colX)
			Set rs = objCmd.Execute(sqlString(colX))
			CCArray(dateX, colX + 1) = rs(0).value 
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
 