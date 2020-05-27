Option Explicit
	'****** Version History *********
	'1.0 -	1/3/2019	Initial Release
	
	'***** CHANGE THESE SETTINGS *****
	Const debugMode = false							'True - displays connection failures to the user
	Const CMMHistory = 7							'[Days] How far back to display data
	Const emailList = ""							'TO email list
	Const BCC_List = "czarlengo@flowcorp.com"		'BCC email list
	
	'***** DATABASE SETTINGS *****
	Const dataSource = "PRODSQLAPP01\PRODSQLAPP01"	'Server location for the database
	
	'***** DATABASE CONSTANTS *****
	Const adOpenDynamic			= 2	 				'// Uses a dynamic cursor.
	Const adOpenForwardOnly		= 0	 				'// Default.
	Const adOpenKeyset			= 1	 				'// Uses a keyset cursor.
	Const adOpenStatic			= 3	 				'// Uses a static cursor.
	Const adOpenUnspecified		= -1 				'// Does not specify the type of cursor.

	Const adStateClosed			= 0  				'// The object is closed
	Const adStateOpen			= 1  				'// The object is open
	Const adStateConnecting		= 2  				'// The object is connecting
	Const adStateExecuting		= 4  				'// The object is executing a command
	Const adStateFetching		= 8  				'// The rows of the object are being retrieved
	
	'***** VARIABLE DEFINITION *****
	Dim ColCount : ColCount = 0						'Number of active fixtures, updated in InitialParameters
	Dim FixtureArray()								'Array for each fixtureID
	Dim machineNameArray()							'Array for machine name for each fixtureID
	'*********************************************************
	
	If Not WScript.Arguments.Count = 0 Then
		Dim sArg : sArg = ""
		Dim Arg : For Each Arg In Wscript.Arguments
			  sArg = sArg & " " & Arg
		Next
	End If
sArg = "daily"
	Dim AccessResult : AccessResult = Load_Access	'Function to check for SQL connection
	Call InitialParameters							'Loads initial variables into the script
	Call CMMLoop									'Runs through the loop for gathering all the data
	
Function Load_Access()								'Function to check for SQL connection
	Dim objCmd : set objCmd = GetNewConnection		'Creates the connection object
	If objCmd is Nothing Then Load_Access = false : Exit Function	'If the connection fails exits function with a return of false
	objCmd.Close									'Closes the connection
	Set objCmd = Nothing							'Empties the connection variable
	Load_Access = true								'Returns true
 End Function										'Ends the function
 
Function InitialParameters()
	Dim objCmd : Set objCmd = GetNewConnection		'Creates the connection object
	Dim sqlQuery(2), rs								'Defines initial variables
	
	'SQL query string to get the number of active fixtures
	sqlQuery(0) = "Select COUNT(*) From [30_Fixtures] WHERE ((([ActiveFixture])= 1) and ([ProgramName] = 'Cut1'));"
	set rs = objCmd.Execute(sqlQuery(0))			'Sends the query to the SQL server
	If rs(0).value <> 0 Then						'If the result is more than 0
		ColCount = rs(0).value - 1					'Updates the total fixture count value
	End If
	Set rs = Nothing								'Empties the query variable
	ReDim FixtureArray(ColCount)					'Re-Dimensions the fixture array size
	ReDim machineNameArray(ColCount)				'Re-Dimensions the machine name array size
	
	'SQL query to get all of the fixture ID's and the machine names
	sqlQuery(1) = "SELECT [FixtureID], [Nomenclature], [ProgramName]" _
				& " FROM [30_Fixtures]" _
				& " WHERE ((([ActiveFixture])= 1) and ([ProgramName] = 'Cut1'))" _
				& " ORDER BY [Nomenclature] ASC;"
	Dim a : a = 0									'Initializes counter variable
	set rs = objCmd.Execute(sqlQuery(1))			'Sends the query to the SQL server
	DO WHILE NOT rs.EOF								'Loops through each result of the query
		FixtureArray(a) = rs.Fields(0)				'Stores the fixtureID
		machineNameArray(a) = rs.Fields(1)			'Stores the machine name
		rs.MoveNext									'Moves to the next record
		a = a + 1									'Increments the counter
	Loop											'Returns to the beginning of the while loop
	Set rs = Nothing								'Empties the query variable
 End Function										'Ends the function
 
Function CMMLoop()
	Dim sqlQuery
	Dim tableBody
	
	Dim objCmd : Set objCmd = GetNewConnection
	Dim tableHead : tableHead = "<head><style>" & _
									"table, th, td {  border: 1px solid black;  border-collapse: collapse; text-align: center; vertical-align: top;}" & _
									".box {  height: 10px; width: 10px;}" & _
								"</style></head><table><tr><th></th>"
	
	Dim a : For a = 0 to UBound(FixtureArray)
		tableHead = tableHead & "<th>&nbsp;" & machineNameArray(a) & "&nbsp;</th>"	
	Next
	tableHead = tableHead & "</tr>"
	
	Dim bStart
	If InStr(1, sArg, "daily") > 0 Then
		bStart = 0
	Else
		bStart = 1
	End If
	
	Dim b : For b = bStart to CMMHistory + 1
		tableBody = tableBody & "<tr>"
		tableBody = tableBody & "<td>&nbsp;" & (CDate(FormatDateTime(Now, vbShortDate)) - b) & "&nbsp;</td>"
		For a = 0 to UBound(FixtureArray)
			tableBody = tableBody & "<td>"
			tableBody = tableBody & CMMSearch(FixtureArray(a), (CDate(FormatDateTime(Now, vbShortDate)) - b), objCmd)
			tableBody = tableBody & "</td>"
		Next
		tableBody = tableBody & "</tr>"
	Next
	
	tableBody = tableBody & "</tr></table>"
	objCmd.Close
	Set objCmd = Nothing
	Send_Email(tableHead & tableBody)
 End Function

Function CMMSearch(FixtureID, dateStart, objCmd)
	Dim rs, dash1color, dash2color
	Dim doesExist : doesExist = False
	Dim sqlQuery : sqlQuery = "SELECT [Blade SN Dash 1], [CMM1].[Failures] AS [Fail1], [Blade SN Dash 2], [CMM2].[Failures] AS [Fail2]" _
							& " FROM ([20_LPT5]" _
							& " LEFT JOIN [40_CMM_LPT5] AS [CMM1] on [CMM1].[Serial Number] = [20_LPT5].[Blade SN Dash 1])" _
							& " LEFT JOIN [40_CMM_LPT5] AS [CMM2] on [CMM2].[Serial Number] = [20_LPT5].[Blade SN Dash 2]" _
							& " WHERE [Fixture Location] = '" & FixtureID & "' AND [Cut Date] >= '" & dateStart & "' AND [Cut Date] < '" & dateStart + 1 & "'" _
							& " ORDER BY [Cut Date] DESC;"
	Set rs = objCmd.Execute(sqlQuery)
	CMMSearch = "<table style='width:100%'>"
	DO WHILE NOT rs.EOF
		doesExist = True
		If rs.Fields(1) > 0 Then
			dash1color = "red"
		ElseIf rs.Fields(1) = 0 Then
			dash1color = "limegreen"
		ElseIf IsNull(rs.Fields(1)) Then
			dash1color = "gray"
		Else
			dash1color = "blue"
		End If
		
		If rs.Fields(3) > 0 Then
			dash2color = "red"
		ElseIf rs.Fields(3) = 0 Then
			dash2color = "limegreen"
		ElseIf IsNull(rs.Fields(3)) Then
			dash2color = "gray"
		Else
			dash2color = "blue"
		End If
		CMMSearch = CMMSearch & "<tr style='height:40px;'><td class='box' style='background-color: " & dash1color & ";'></td><td class='box' style='background-color: " & dash2color & ";'></td></tr>"
		rs.MoveNext
	Loop	
	
	Set rs = Nothing
	
	
	
	CMMSearch = CMMSearch & "</table>"
	If doesExist = False Then CMMSearch = ""
	
 End Function
	
Function GetNewConnection()
	Dim objCmd : Set objCmd = CreateObject("ADODB.Connection")
	Dim sConnection : sConnection = "Data Source=" & dataSource & ";Initial Catalog=CMM_Repository;Integrated Security=SSPI;"
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

Sub Send_Email(Message)
	Dim MyEmail : Set MyEmail=CreateObject("CDO.Message")
	Dim bodyPre : bodyPre = "<p><span style='font-size:12pt; color:red'>This is an automatically generated daily email, please do not reply to sender. Email <a href=""mailto:CZarlengo@flowcorp.com"">Chris Zarlengo</a> if you have any issues.</span></p><br>"
	Dim body : body =  Message
	Dim Signature : Signature = "<footer><div>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Chris Zarlengo</span><span style='color:#1F497D'></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>Manufacturing Manager</span><span style='color:#1F497D'></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Flow International Corporation | <a href=""http://www.flowwaterjet.com/"">http://www.FlowWaterjet.com/</a></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>23500 64th Ave. S. | Kent | Washington | 98032 | USA</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>253-246-3741 | <a href=""mailto:CZarlengo@flowcorp.com"">CZarlengo@flowcorp.com</a><br>" _
		& "</div></footer>"
	
	MyEmail.Subject="Contract cutting daily details"
	MyEmail.From="czarlengo@flowcorp.com"
	MyEmail.To=emailList
	MyEmail.BCC=BCC_List
	MyEmail.HTMLBody = bodyPre & body & Signature

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
 