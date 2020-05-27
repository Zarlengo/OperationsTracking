Dim THIS_IS_A_PLACEHOLDER 'OPTION EXPLICIT

 '******  CHANGE THESE SETTINGS *********
 Dim adminMode : adminMode = false
 Dim debugMode : debugMode = false
 '***************** Database Settings *******************
 Dim dataSource : dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"
 Dim initialCatalog : initialCatalog = "CMM_Repository"								'Initial database
 Dim tabletPassword : tabletPassword = "Fl0wSh0p17"
 Dim computerPassword : computerPassword = "Snowball18!"
 '***************************************
 Const adOpenStatic			= 3	 '// Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
 Const adStateOpen			= 1  '// The object is open
 '***************************************
 
 Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
 Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")
 Dim manualSave : manualSave = False
 Const searchTime = 10000
 Dim errorFound : errorFound = false
 Dim emailBody

 If Load_Access Then
	Call Check_00_SN
	Call Check_30_Offset
	Call Check_30_Reason
	Call Check_40_CMM
	Call Check_50_Final
	If errorFound = true Then
		Const Subject = "AMP SQL Audit"
		Const EmailList = "czarlengo@flowcorp.com"
		Dim messageBody : messageBody = "<body><p><span style='font-size:12pt; color:red'>This is an automatically generated email.</span></p><br>" _
			& "<p><span>" & emailBody & "</span></p>"
		Call Send_Email(messageBody, subject, EmailList, "")
	End If
 End If
 Wscript.Quit
 
Sub Check_00_SN()
	Set objCmd = GetNewConnection									'Creates the connection object to the database
	If objCmd is Nothing Then Exit Sub
	Dim msgString : msgString = ""
	Dim sqlString : sqlString = "SELECT [Invoice Number] FROM [00_AE_SN_CONTROL] WHERE [Invoice Number] NOT LIKE 'AEFL%' AND [Invoice Number] NOT LIKE 'A2017%' AND [Invoice Number] NOT LIKE 'NOT-PROVIDED';"
	Dim rs : set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF 
		errorFound = true
		msgString = msgString & "Invoice Number: " & rs.Fields(0) & " is in the wrong format.<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [Slug Serial Number], [Invoice Number] FROM [00_AE_SN_CONTROL] WHERE [Slug Serial Number] LIKE '%Material%';"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "Slug Serial Number: " & rs.Fields(0) & " is in the wrong format for ID: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [Blade Serial Number], [Invoice Number] FROM [00_AE_SN_CONTROL] WHERE [Blade Serial Number] LIKE '%S/N%';"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "Blade Serial Number: " & rs.Fields(0) & " is in the wrong format for ID: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing	
	If msgString <> "" Then
		emailBody = emailBody & "<br><span style='font-size:12pt; color:red'>00_AE_SN_Control has error(s).</span><br>" & msgString
	Else
		emailBody = emailBody & "<br><span style='font-size:12pt; color:limegreen'>00_AE_SN_Control has passed error checks.</span><br>" & msgString
	End If
 End Sub
 
Sub Check_30_Offset()
	Set objCmd = GetNewConnection									'Creates the connection object to the database
	If objCmd is Nothing Then Exit Sub
	Dim msgString : msgString = ""
	Dim sqlString : sqlString = "SELECT [MachineNumber], [ID] FROM [30_OFFSET] WHERE [MachineNumber] NOT LIKE 'WJM%';"
	Dim rs : set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF 
		errorFound = true
		msgString = msgString & "MachineNumber: " & rs.Fields(0) & " is in the wrong format for ID: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [FileName], [ID] FROM [30_OFFSET] WHERE [FileName] NOT LIKE 'Fixture_';"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "FileName: " & rs.Fields(0) & " is in the wrong format for ID: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [PersonName], [ID] FROM [30_OFFSET] WHERE [PersonName] IS NULL;"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "PersonName: " & rs.Fields(0) & " is missing for ID: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing	
	If msgString <> "" Then
		emailBody = emailBody & "<br><span style='font-size:12pt; color:red'>30_Offset has error(s).</span><br>" & msgString
	Else
		emailBody = emailBody & "<br><span style='font-size:12pt; color:limegreen'>30_Offset has passed error checks.</span><br>" & msgString
	End If
 End Sub
 
Sub Check_30_Reason()
	Set objCmd = GetNewConnection									'Creates the connection object to the database
	If objCmd is Nothing Then Exit Sub
	Dim msgString : msgString = ""
	Dim sqlString : sqlString = "SELECT [Reason], [ID] FROM [30_REASON] WHERE [Reason] IS NULL;"
	Dim rs : set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF 
		errorFound = true
		msgString = msgString & "Reason: " & rs.Fields(0) & " is missing for ID: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [SNs], [ID] FROM [30_REASON] WHERE [SNs] LIKE '%Dim%';"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "SNs: " & rs.Fields(0) & " is in the wrong format for ID: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing	
	If msgString <> "" Then
		emailBody = emailBody & "<br><span style='font-size:12pt; color:red'>30_Reason has error(s).</span><br>" & msgString
	Else
		emailBody = emailBody & "<br><span style='font-size:12pt; color:limegreen'>30_Reason has passed error checks.</span><br>" & msgString
	End If
 End Sub
 
Sub Check_40_CMM()
	Set objCmd = GetNewConnection									'Creates the connection object to the database
	If objCmd is Nothing Then Exit Sub
	Dim msgString : msgString = ""
	Dim sqlString : sqlString = "SELECT [Serial Number], [File Name] FROM [40_CMM_LPT5] WHERE [Serial Number] NOT LIKE 'H_______-_';"
	Dim rs : set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF 
		errorFound = true
		msgString = msgString & "Serial Number: " & rs.Fields(0) & " is in the wrong format for File Name: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [Part Number], [File Name] FROM [40_CMM_LPT5] WHERE [Part Number] <> '060053-1' AND [Part Number] <> '060053-2';"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "Part Number: " & rs.Fields(0) & " is in the wrong format for File Name: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [Operator], [File Name] FROM [40_CMM_LPT5] WHERE [Operator] LIKE '%>%';"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "Operator: " & rs.Fields(0) & " is in the wrong format for File Name: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	sqlString = "SELECT [Date], [File Name] FROM [40_CMM_LPT5] WHERE [Date] = 0 OR [Date] IS NULL;"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		errorFound = true
		msgString = msgString & "Date: " & rs.Fields(0) & " is invalid for File Name: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing	
	If msgString <> "" Then
		emailBody = emailBody & "<br><span style='font-size:12pt; color:red'>40_CMM has error(s).</span><br>" & msgString
	Else
		emailBody = emailBody & "<br><span style='font-size:12pt; color:limegreen'>40_CMM has passed error checks.</span><br>" & msgString
	End If
 End Sub
 
Sub Check_50_Final()
	Set objCmd = GetNewConnection									'Creates the connection object to the database
	If objCmd is Nothing Then Exit Sub
	Dim msgString : msgString = ""
	Dim sqlString : sqlString = "SELECT [Accepted Y/N], [Blade S/N] FROM [50_FINAL] WHERE [Accepted Y/N] NOT LIKE '_';"
	Dim rs : set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF 
		errorFound = true
		msgString = msgString & "Accepted Y/N: " & rs.Fields(0) & " is in the wrong format for Serial Number: " & rs.Fields(1) & ".<br>"
		rs.MoveNext
	Loop
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing	
	If msgString <> "" Then
		emailBody = emailBody & "<br><span style='font-size:12pt; color:red'>50_Final has error(s).</span><br>" & msgString
	Else
		emailBody = emailBody & "<br><span style='font-size:12pt; color:limegreen'>50_Final has passed error checks.</span><br>" & msgString
	End If
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
 
 
Function Send_Email(Message, subject, EmailTo, EmailBCC)
' exit function
	Dim MyEmail : Set MyEmail=CreateObject("CDO.Message")
	
	Dim Signature : Signature = "<footer><div>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Chris Zarlengo</span><span style='color:#1F497D'></span><br>" _
		& "</div></footer>"
	
	MyEmail.Subject = subject
	MyEmail.From="czarlengo@flowcorp.com"
	MyEmail.To = EmailTo
	MyEmail.BCC = EmailBCC
	MyEmail.HTMLBody = Message & Signature

	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

	'SMTP Server
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="SKENEXC60.flowcorp.com"

	'SMTP Port
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

	'SMTP Auth (For Windows Auth set this to 2)
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=2

	MyEmail.Configuration.Fields.Update
	MyEmail.Send

	set MyEmail = nothing


 End Function