Option Explicit
 '****** CHANGE THESE SETTINGS *********
 Const debugMode = false	
 Const dateHistory = 2	
 
 'Only run on Monday and Wednesday
 If Weekday(Now, 2) <> 1 and Weekday(Now, 2) <> 3 Then Wscript.Quit
 
 '***************** Database Settings *******************
 Const dataSource = "AXPRODSQL02\AXPRODSQL02"
 Const initialCatalog = "WJH_DB_Warehouse"								'Initial database
 '**************** INITIAL PARAMETERS *******************
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
 '**************** DATABASE CONSTANTS *******************
 Const adOpenStatic			= 3	 '// Uses a static cursor.		
 Const adStateOpen			= 1  '// The object is open		
 '*********************************************************
 '***************** EMAIL LIST ****************************
 Dim eMailList : Set eMailList = CreateObject("Scripting.Dictionary")
 Dim managerList : Set managerList = CreateObject("Scripting.Dictionary")
 eMailList.Add		"Albertson, Shawn", "joe.kraus@kmtwaterjet.com"
 managerList.Add	"Albertson, Shawn", "joe.kraus@kmtwaterjet.com"
 
 eMailList.Add		"Arena, Mike", "marena@shapetechnologies.com"
 managerList.Add	"Arena, Mike", "slowery@shapetechnologies.com"
 
 eMailList.Add		"Burkey, Sean", "sburkey@flowcorp.com"
 managerList.Add	"Burkey, Sean", "akeating@shapetechnologies.com"
 
 eMailList.Add		"Chen, Cindy", "cindy@flowcorp.com"
 managerList.Add	"Chen, Cindy", "swong@flowcorp.com"
 
 eMailList.Add		"Cho, Jinwoo", "jinwoo.cho@tops21.com"
 managerList.Add	"Cho, Jinwoo", "JWaite@shapetechnologies.com"
 
 eMailList.Add		"Christianson, Leigh", "lchristianson@flowcorp.com"
 managerList.Add	"Christianson, Leigh", "pvarney@flowcorp.com"
 
 eMailList.Add		"Durand, Mike", "mdurand@flowcorp.com"
 managerList.Add	"Durand, Mike", "JWaite@shapetechnologies.com"
 
 eMailList.Add		"Edwards, Tim", "tedwards@flowcorp.com"
 managerList.Add	"Edwards, Tim", "JWaite@shapetechnologies.com"
 
 eMailList.Add		"Falany, Jeff", "JFalany@flowcorp.com"
 managerList.Add	"Falany, Jeff", "pvarney@flowcorp.com"
 
 eMailList.Add		"Final: Durand, Mike", "mdurand@flowcorp.com"
 managerList.Add	"Final: Durand, Mike", "JWaite@shapetechnologies.com"
 
 eMailList.Add		"Fujitani, Kazumi", "fujitani@flowcorp.com"
 managerList.Add	"Fujitani, Kazumi", "sugahara@flowcorp.com"

 eMailList.Add		"Hall, Gary", "Gary.Hall@kmtwaterJet.com"
 managerList.Add	"Hall, Gary", "joe.kraus@kmtwaterjet.com"
 
 eMailList.Add		"Jeon, Young Kwang", "ykjun@tops21.com"
 managerList.Add	"Jeon, Young Kwang", "twoh@tops21.com"
 
 eMailList.Add		"Jung, Seung Youl", "seungyoul.jung@tops21.com"
 managerList.Add	"Jung, Seung Youl", "twoh@tops21.com"
 
 eMailList.Add		"Kang, Sung Soo", "sskang@tops21.com"
 managerList.Add	"Kang, Sung Soo", "twoh@tops21.com "
 
 eMailList.Add		"Kraus, Joe", "joe.kraus@kmtwaterjet.com"
 managerList.Add	"Kraus, Joe", "joe.kraus@kmtwaterjet.com"

 eMailList.Add		"Kowalczyk, Dave", "DKowalczyk@Flowcorp.com"
 managerList.Add	"Kowalczyk, Dave", "ERomanoff@flowcorp.com"

 eMailList.Add		"Kwon, Bo Sang", "bskwon@tops21.com"
 managerList.Add	"Kwon, Bo Sang", "jskim@tops21.com"
 
 eMailList.Add		"Lee, Seung Han", "seunghan.lee@tops21.com"
 managerList.Add	"Lee, Seung Han", "ticho@tops21.com"
 
 eMailList.Add		"Lo, Clark", "clark@flowcorp.com"
 managerList.Add	"Lo, Clark", "richard@flowcorp.com"
 
 eMailList.Add		"McLane, Amy", "AMclane@flowcorp.com"
 managerList.Add	"McLane, Amy", "amclane@flowcorp.com"
 
 eMailList.Add		"Moto, Sugahara", "sugahara@flowcorp.com"
 managerList.Add	"Moto, Sugahara", "Masuko@flowcorp.com"
 
 eMailList.Add		"Mueller, Kurt", "kmueller@shapetechnologies.com"
 managerList.Add	"Mueller, Kurt", "dcrewe@shapetechnologies.com"
 
 eMailList.Add		"Onari, Tsutomu", "onari@flowcorp.com"
 managerList.Add	"Onari, Tsutomu", "watanabe@flowcorp.com"
 
 eMailList.Add		"Romanoff, Ethan", "ERomanoff@flowcorp.com"
 managerList.Add	"Romanoff, Ethan", "Cwakefield@flowcorp.com"
 
 eMailList.Add		"Schramm, Sean", "sschramm@shapetechnologies.com"
 managerList.Add	"Schramm, Sean", "steve.harris@kmtwaterjet.com"

 eMailList.Add		"Schultz, Larry", "Larry.Schultz@shapetechnologies.com"
 managerList.Add	"Schultz, Larry", "Michael.Barber@kmtwaterjet.com"
 
 eMailList.Add		"Shetty, Randhir", "randhir.shetty@flowcorp.com"
 managerList.Add	"Shetty, Randhir", "jjenson@shapetechnologies.com"
 
 eMailList.Add		"Smith, Jean", "msmith@flowcorp.com"
 managerList.Add	"Smith, Jean", "pvarney@flowcorp.com"
 
 eMailList.Add		"Van Sickle, Paul", "pvansickle@shapetechnologies.com"
 managerList.Add	"Van Sickle, Paul", "akeating@shapetechnologies.com"
 
 eMailList.Add		"Varney, Pat", "pvarney@flowcorp.com"
 managerList.Add	"Varney, Pat", "Hfonda@flowcorp.com"
 
 eMailList.Add		"Vaughan, Sean", "svaughan@flowcorp.com"
 managerList.Add	"Vaughan, Sean", "Cwakefield@flowcorp.com"
 
 eMailList.Add		"Waite, John", "JWaite@shapetechnologies.com"
 managerList.Add	"Waite, John", "akeating@shapetechnologies.com"
 
 eMailList.Add 		"Warren, Joseph", "jwarren@flowcorp.com"
 managerList.Add	"Warren, Joseph", "cwakefield@flowcorp.com"
 
 eMailList.Add		"Wakefield, Charles", "cwakefield@flowcorp.com"
 managerList.Add	"Wakefield, Charles", "jjenson@shapetechnologies.com"
 
 eMailList.Add		"Watanabe, Hiroshi", "watanabe@flowcorp.com"
 managerList.Add	"Watanabe, Hiroshi", "Masuko@flowcorp.com"
 
 eMailList.Add		"Wyman, Karri", "kwyman@flowcorp.com"
 managerList.Add	"Wyman, Karri", "pvarney@flowcorp.com"
 
 eMailList.Add		"Zarlengo, Christopher", "CZarlengo@flowcorp.com"
 managerList.Add	"Zarlengo, Christopher", "Hfonda@flowcorp.com"
 '*********************************************************
 eMailList.Add		"WORKFLOW, BAXTER", "Will.Lambeth@kmtwaterjet.com"
 managerList.Add	"WORKFLOW, BAXTER", "joe.kraus@kmtwaterjet.com"
 
 eMailList.Add		"WORKFLOW, ZOLLNER_TAICANG", "jinwoo.cho@tops21.com"
 managerList.Add	"WORKFLOW, ZOLLNER_TAICANG", "JWaite@shapetechnologies.com"
 
 eMailList.Add		"WORKFLOW, ZOLLNER_VAC", "JWaite@shapetechnologies.com"
 managerList.Add	"WORKFLOW, ZOLLNER_VAC", "akeating@shapetechnologies.com"
 '*********************************************************
 eMailList.Add		"McQueen, Colleen", "pvarney@flowcorp.com"
 managerList.Add	"McQueen, Colleen", "Hfonda@flowcorp.com"
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
	
	
	'Call CCHistory
	Dim sqlString : sqlString = "SELECT [ECO_Number], [Current_Approver], [Approver_Aging] FROM ECO.ECOApprovers WHERE [Approver_Aging] > 2 order by [Approver_Aging] Asc;"
	Set rs = objCmd.Execute(sqlString)
	
	
	Dim ApproverList : Set ApproverList = CreateObject("Scripting.Dictionary")
	Dim ECOList, DateList, name, maxDays, mEmail
	Dim ApproverArray, approver, ECONumber, ECOAge, DictArray
	
	
	Do While not rs.EOF
		ApproverArray = Split(rs.Fields(1), ";")
		For Each approver In ApproverArray
			name = TrimString(approver)
			ECONumber = rs.Fields(0)
			ECOAge = rs.Fields(2)
			If ApproverList.Exists(name) Then
				DictArray = ApproverList.Item(name)
				ApproverList.Remove(name)
				ECONumber = ECONumber & ";" & DictArray(0)
				ECOAge = ECOAge & ";" & DictArray(1)
			End If
			ApproverList.Add name, array(ECONumber, ECOAge)
		Next
		rs.MoveNext
	Loop
	Set rs = nothing
	Dim arrayKeys : arrayKeys = ApproverList.Keys
	Dim i, s, j, msgString, email
	for i = 0 to ApproverList.Count - 1 
		approver = arrayKeys(i)
		msgString = ""
		maxDays = 0
		If InStr(1, ApproverList.Item(approver)(0), ";") <> 0 Then
			ECOList = Split(ApproverList.Item(approver)(0), ";")
			DateList = Split(ApproverList.Item(approver)(1), ";")
			For j = 0 To UBound(ECOList)
				s = ECOList(j)
				msgString = msgString & "<br>" & ECOList(j) & " at " & DateList(j) & " days"
				maxDays = DateList(j)
			Next
		Else
			msgString = msgString & "<br>" & ApproverList.Item(approver)(0) & " at " & ApproverList.Item(approver)(1) & " days"
			maxDays = ApproverList.Item(approver)(1)
		End IF
		If eMailList.Exists(approver) Then
			email = eMailList.Item(approver)
			mEmail = managerList.Item(approver)
		ElseIf InStr(1, approver, "Rejected") Then
			email = "mdurand@flowcorp.com;czarlengo@flowcorp.com;tedwards@flowcorp.com"
		Else
			email = "czarlengo@flowcorp.com;tedwards@flowcorp.com"
		End If
		If CInt(maxDays) >= 30 Then
			msgString = "You have an employee with ECO(s) that have exceeded 30 days and their task has not been completed. Please have this addressed or reassign as appropriate.<br><br>" & msgString
			email = mEmail & ";" & email & ";czarlengo@flowcorp.com;jwaite@flowcorp.com;tedwards@flowcorp.com"
		Else
			msgString = "<br><p><span style='font-size:12pt;'>" & approver & "</span>" & _
		 "<br><p><span style='font-size:12pt;'>The following ECO's are delinquent and awaiting your approval:<br>" & msgString & "</span>"
		End If
		Call Send_Email(approver, email, msgString)
	next
	objCmd.Close
	Set objCmd = Nothing
 End If																'Function to close open connections and return settings back to original	
 Wscript.Quit
 
Function TrimString(ByVal VarIn)
	VarIn = Trim(VarIn)   
	If Len(VarIn) > 0 Then
		Do While AscW(Right(VarIn, 1)) = 10 or AscW(Right(VarIn, 1)) = 13
			VarIn = Left(VarIn, Len(VarIn) - 1)
		Loop
	End If
	TrimString = Trim(VarIn)
End Function

Sub Send_Email(recipient, email, Message)
	Dim MyEmail : Set MyEmail=CreateObject("CDO.Message")
	Dim bodyPre : bodyPre = "<p><span style='font-size:12pt; color:red'>This is an automatically generated daily email</span></p>"
	
	MyEmail.Subject="ECO Implementation Daily Summary"
	MyEmail.From="czarlengo@flowcorp.com"
	MyEmail.To = email
	MyEmail.HTMLBody = bodyPre & Message
	
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
 