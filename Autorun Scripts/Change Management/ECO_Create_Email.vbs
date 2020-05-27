Option Explicit
 '****** CHANGE THESE SETTINGS *********
 Const debugMode = false	
 Const dateHistory = 2	
 
 'Only run on Monday's
 If Weekday(Now, 2) <> 1 Then Wscript.Quit
 
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
 If Load_SQL = true Then
	Dim Eff
	Dim objCmd, rs : Set objCmd = GetNewConnection
	Call LateECO
	'Call ExpiredECO
	objCmd.Close
	Set objCmd = Nothing
 End If																'Function to close open connections and return settings back to original	
 Wscript.Quit
 
Function LateECO()
	Dim sqlString : sqlString = "SELECT [ECO_Number], [ECO_Desc], [ECO_Reason], [ECO_Notes], [ECO_State], [ECO_Orig], [ECO_Aging] FROM [ECOCreate] WHERE [ECO_Aging] > 14 order by [ECO_Aging] Asc;"
	Set rs = objCmd.Execute(sqlString)
	
	
	Dim OriginatorList : Set OriginatorList = CreateObject("Scripting.Dictionary")
	Dim ECOList, DateList, name, Originator, Description, Reason, Notes, State
	Dim ApproverArray, approver, ECONumber, ECOAge, DictArray, rowString
	Const tablePre = "<table border='1'><tr><th>ECO</th><th>Description</th><th>Reason for Change</th><th>Notes</th><th>State</th><th>ECO Age</th></tr>"
	Const tableSuf = "</table>"
  
	Do While not rs.EOF
		ECONumber = rs.Fields(0)
		Description = TrimString(rs.Fields(1))
		Reason = TrimString(rs.Fields(2))
		Notes = TrimString(rs.Fields(3))
		State = rs.Fields(4)
		Originator = TrimString(rs.Fields(5))
		ECOAge = rs.Fields(6)
		rowString = "<tr><td>" & ECONumber & "</td><td>" & Description & "</td><td>" & Reason & "</td><td>" & Notes & "</td><td>" & State & "</td><td>" & ECOAge & "</td></tr>"
		If OriginatorList.Exists(Originator) Then
			rowString = OriginatorList.Item(Originator) & rowString
			OriginatorList.Remove(Originator)
		End If
		OriginatorList.Add Originator, rowString
		rs.MoveNext
	Loop
	Set rs = nothing
	Dim arrayKeys : arrayKeys = OriginatorList.Keys
	Dim i, s, j, msgString, email
	for i = 0 to OriginatorList.Count - 1 
		Originator = arrayKeys(i) 
		msgString = tablePre & OriginatorList.Item(Originator) & tableSuf
		Call Send_Email(Originator, msgString, False)
	next
 End Function
 
Function ExpiredECO()
	Dim sqlString : sqlString = "SELECT [ECO_Number], [ECO_Desc], [ECO_Reason], [ECO_Notes], [ECO_State], [ECO_Orig], [ECO_Aging] FROM [ECOCreate] WHERE [ECO_Aging] > 30 order by [ECO_Aging] Asc;"
	Set rs = objCmd.Execute(sqlString)
	
	
	Dim OriginatorList : Set OriginatorList = CreateObject("Scripting.Dictionary")
	Dim ECOList, DateList, name, Originator, Description, Reason, Notes, State
	Dim ApproverArray, approver, ECONumber, ECOAge, DictArray, rowString
	Const tablePre = "<table border='1'><tr><th>ECO</th><th>Description</th><th>Reason for Change</th><th>Notes</th><th>State</th><th>ECO Age</th></tr>"
	Const tableSuf = "</table>"
  
	Do While not rs.EOF
		ECONumber = rs.Fields(0)
		Description = TrimString(rs.Fields(1))
		Reason = TrimString(rs.Fields(2))
		Notes = TrimString(rs.Fields(3))
		State = rs.Fields(4)
		Originator = TrimString(rs.Fields(5))
		ECOAge = rs.Fields(6)
		rowString = "<tr><td>" & ECONumber & "</td><td>" & Description & "</td><td>" & Reason & "</td><td>" & Notes & "</td><td>" & State & "</td><td>" & ECOAge & "</td></tr>"
		If OriginatorList.Exists(Originator) Then
			rowString = OriginatorList.Item(Originator) & rowString
			OriginatorList.Remove(Originator)
		End If
		OriginatorList.Add Originator, rowString
		rs.MoveNext
	Loop
	Set rs = nothing
	Dim arrayKeys : arrayKeys = OriginatorList.Keys
	Dim i, s, j, msgString, email
	for i = 0 to OriginatorList.Count - 1 
		Originator = arrayKeys(i) 
		msgString = tablePre & OriginatorList.Item(Originator) & tableSuf
		Call Send_Email(Originator & ";czarlengo@flowcorp.com", msgString, True)
	next
 End Function
 
Function TrimString(ByVal VarIn)
	VarIn = Trim(VarIn)   
	If Len(VarIn) > 0 Then
		Do While AscW(Right(VarIn, 1)) = 10 or AscW(Right(VarIn, 1)) = 13
			VarIn = Left(VarIn, Len(VarIn) - 1)
		Loop
	End If
	VarIn = Replace(VarIn, ";", "|")
	TrimString = Trim(VarIn)
End Function

Sub Send_Email(email, Message, expired)
	Dim MyEmail : Set MyEmail=CreateObject("CDO.Message")
	Dim bodyPre : bodyPre = "<p><span style='font-size:12pt; color:red'>This is an automatically generated weekly email</span></p>"
	Dim body : If expired = False Then
		body =  _
		 "<br><p><span style='font-size:12pt;'>Hello, you have ECO(s) which have been idle for 2 weeks. Please process these to CCB or cancel the ECO and create a new one when the change is ready to be processed. </span>" & _
		 "<br><p><span style='font-size:12pt;'>The ECO(s) below will automatically be canceled at 30 days, to have these place on a 2 week exclusion list please reply back with a completion date and reason for delay.<br>" & Message & "</span>"
	Else
		body =  _
		 "<br><p><span style='font-size:12pt;'>Hello, you have ECO(s) which have exceeded the allotted duration to reach CCB (30 days).</span>" & _
		 "<br><p><span style='font-size:12pt;'>The ECO(s) below will now automatically be canceled. Please create a new ECO when the change is ready to be processed.<br>" & Message & "</span>"
	End If
	
	MyEmail.Subject="Canceled ECO Summary"
	MyEmail.From="czarlengo@flowcorp.com"
	MyEmail.To = email
	MyEmail.HTMLBody = bodyPre & body
	
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

Function Load_SQL()
	Dim objCmd : set objCmd = GetNewConnection
	If objCmd is Nothing Then Load_SQL = false : Exit Function
	objCmd.Close
	Set objCmd = Nothing
	Load_SQL = true
 End Function
 