Option Explicit
 '****** Version History *********
 '1.0 -	Initial release to production
 '1.1 -	Updated to include M500
 '1.2 -	Added maintenance counter
 '   	Records maintenance resets
 '		Small code updates for efficiency (match other OPs)
 '
 '2.0 -	Update to allow usage of script on M500's
 '		Checks for duplicate scans and updates existing entries
 '		Counters do not increment on duplicates
 '2.1 - Added USB (COM) connection option
 '2.2 - Moved initialization of AXSerial and OSWinsock to allow script to run if the drivers are not installed
 '
 '3.0 - Added Operator timeout, added GfE options, added search for PID before switching windows
 '***************************************
 
 Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
 Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")
 Dim oShell : Set oShell = CreateObject("WScript.Shell")
 Dim shl : set shl = createobject("wscript.shell")

 Const allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\All Operations.vbs"
 Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """" & sArg
 Const dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"
 Const adminPassword = "FLOW288"
 Const tabletPassword = "Fl0wSh0p17"
 Const computerPassword = "Snowball18!"
 Const ProgramVersion = "Revision B"

 Dim closeWindow : closeWindow = false
 Dim errorWindow : errorWindow = false
 Const adminMode = false
 Const debugMode = false

 Const AbrasiveLimit = 30
 Const MixLimit = 40
 Const OrificeLimit = 1000
 Const RowCount = 8
 Const WaitMS = 50
 Const TimeOut = 5000
 Const timerCount = 30000
 Dim logoutTime : logoutTime = 45 / 60 / 24

 Dim tolName : tolName = Array("Dim 1.1",	"Dim 1.2",	"Dim 2.1",	"Dim 2.2",	"Dim 3.1",	"Dim 3.2",	"Dim 4.1",	"Dim 4.2",	"Dim 5.1",	"Dim 5.2",	"Dim 9.1",	"Dim 9.2",	"Dim 10.1",	"Dim 10.2",	"Dim 11 Max",	"Dim 11 Min",	"Dim 12 Max",	"Dim 12 Min")
 Dim tolID : tolID = 	 Array("1_1",		"1_2",		"2_1",		"2_2",		"3_1",		"3_2",		"4_1",		"4_2",		"5_1",		"5_2",		"9_1",		"9-2",		"10_1",		"10_2",		"11_Max",		"11_Min",		"12_Max",		"12_Min")
 Dim minTol : minTol =   Array(40.8, 		40.8, 		155.3, 		155.3, 		168.2, 		168.2,  	155.3, 		155.3, 		26.9, 		26.9, 		16.75,  	16.75,  	32.25,  	32.25,  	-0.5,  	 		-0.5,  			-0.5,  			-0.5)
 Dim maxTol : maxTol =   Array(41.8, 		41.8, 		156.3, 		156.3,  	169.2,  	169.2,  	156.3,  	156.3,  	27.9,  		27.9,  		17.75,  	17.75,  	99.99,  	99.99,   	 0.5,   		 0.5,   		 0.5,  			 0.5)

 Dim strData, windowBox, AccessArray, AccessResult, blankArray
 Dim SendData, RecieveData, cProcesses, oProcess
 Dim machineBox, strSelection, RemoteHost, RemotePort
 Dim timerCnt, TimeCnt
 
 ReDim blankArray(RowCount, 7)

 '***************************************
 Const sckClosed             = 0  '// Default. Closed 
 Const sckOpen               = 1  '// Open 
 Const sckListening          = 2  '// Listening 
 Const sckConnectionPending  = 3  '// Connection pending 
 Const sckResolvingHost      = 4  '// Resolving host 
 Const sckHostResolved       = 5  '// Host resolved 
 Const sckConnecting         = 6  '// Connecting 
 Const sckConnected          = 7  '// Connected 
 Const sckClosing            = 8  '// Peer is closing the connection 
 Const sckError              = 9  '// Error 

 Const adOpenDynamic		= 2	 '// Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
 Const adOpenForwardOnly	= 0	 '// Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
 Const adOpenKeyset			= 1	 '// Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
 Const adOpenStatic			= 3	 '// Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
 Const adOpenUnspecified	= -1 '// Does not specify the type of cursor.

 Const adLockBatchOptimistic= 4	 '// Indicates optimistic batch updates. Required for batch update mode.
 Const adLockOptimistic		= 3	 '// Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method.
 Const adLockPessimistic	= 2	 '// Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing.
 Const adLockReadOnly		= 1	 '// Indicates read-only records. You cannot alter the data.
 Const adLockUnspecified	= -1 '// Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.

 Const adStateClosed		= 0  '// The object is closed
 Const adStateOpen			= 1  '// The object is open
 Const adStateConnecting	= 2  '// The object is connecting
 Const adStateExecuting		= 4  '// The object is executing a command
 Const adStateFetching		= 8  '// The rows of the object are being retrieved

 '*********************************************************
 
 ' Am I running 64-bit version of WScript.exe/Cscript.exe? So, call script again in x86 script host and then exit.
 If InStr(LCase(WScript.FullName), LCase(oProcEnv("windir") & "\System32\")) And oProcEnv("PROCESSOR_ARCHITECTURE") = "AMD64" Then
	If Not WScript.Arguments.Count = 0 Then
        Dim sArg, Arg
        sArg = ""
        For Each Arg In Wscript.Arguments
              sArg = sArg & " " & """" & Arg & """"
        Next
    End If
    Dim sCmd : sCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & WScript.ScriptFullName & """" & sArg
    objShell.Run sCmd
    WScript.Quit
 End If

 If debugMode = False Then On Error Resume Next
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
	'objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "30", "REG_DWORD"	'Changes TCP timeout settings if needing to restart program w/in 5 minutes
 On Error Goto 0

 'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
 Dim wmi : Set wmi = GetObject("winmgmts:root\cimv2") 
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
 For Each oProcess in cProcesses
	oProcess.Terminate()
 Next

 Dim FlowCUT_ID : FlowCUT_ID = getFlowPID()
 If Not WScript.Arguments.Count = 0 Then
	sArg = ""
	For Each Arg In Wscript.Arguments
		  sArg = Arg
	Next
 End If

 'Function to check for access connection and load info from database
 AccessResult = Load_Access

 Dim machineString : If sArg = "" Then
	machineString = "Manual"
 Else
	machineString = sArg
 End If


 '// CREATE IP & USB connection ports
 If Left(machineString, 3) = "COM" Then
	Dim objComport : Set objComport = CreateObject( "AxSerial.ComPort" )
	objComport.Clear()
	objComport.LicenseKey = "FD2C1-DC93A-6BFBF"
	objComport.Device = machineString
	objComport.BaudRate  = 112500
	objComport.ComTimeout = 1000  ' Timeout after 1000msecs 
 ElseIf Left(machineString, 3) = "WJM" Then
	Dim winsock0 : Set winsock0 = Wscript.CreateObject("OSWINSCK.Winsock", "winsock0_")
	'// CREATE WINSOCK: 0 - QA Scanner
	If Err.Number <> 0 Then
		MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
		WScript.Quit
	End If
	Load_IP
 End If
 
 'Calls function to create ie window
 '780, 900
 Dim logoutStart
 set windowBox = HTABox("white", 800, 1280, 0, 0) : with windowBox
	Call checkAccess
	Call connect2Scanner
	do until closeWindow = true													'Run loop until conditions are met
		logoutStart = now
		timerCnt = timerCount
		do until .done.value = "cancel" or .done.value = "access" or .done.value = "scanner" or .submitButton.value = "true" or .done.value = "allOps"
			If Left(machineString, 3) <> "COM" Then wsh.sleep WaitMS
			On Error Resume Next
			If .done.value = true Then
				wsh.quit
			End If
			On Error GoTo 0
			If Left(machineString, 3) = "COM" Then
				ReadResponse(objComport)	
				timerCnt = timerCnt - WaitMS * 20
			Else
				timerCnt = timerCnt - WaitMS
			End If
			If timerCnt = 0 Then 
				TimeCnt = 0
				FlowCUT_ID = getFlowPID()
				Do
					wsh.sleep WaitMS
					TimeCnt = TimeCnt + WaitMS
				Loop While shl.AppActivate(FlowCUT_ID) = False and TimeCnt < TimeOut	'"FlowCUT"
				'oShell.SendKeys "% x"
				.timerText.innerText = ""
			ElseIf timerCnt > 0 Then
				.timerText.innerText = "Returning to FlowCUT in " & int(timerCnt/1000)
			End If
			If logoutStart + logoutTime < now Then
				windowBox.operatorID.innerText = ""
				windowbox.errorDiv.style.background = "red"
				windowBox.errorString.innerText = "Logged out"
				logoutStart = now
			End If
		loop
		if .done.value = "cancel" then											'If the x button is clicked
			closeWindow = true													'Variable to end loop
		ElseIf .done.value = "access" then
			.done.value = false
			windowBox.accessText.innerText = "Retrying connection."
			windowBox.accessButton.style.backgroundcolor = "orange"
			If windowbox.FixtureID.innerText <> "" and windowbox.blade1String.innerText <> "" and windowbox.blade2String.innerText <> "" and windowbox.operatorID.innerText <> "" Then
				LoadSNtoAccess
			Else
				AccessResult = Load_Access
				checkAccess
			End If
		ElseIf .done.value = "scanner" then
			.done.value = false
			connect2Scanner
		ElseIf .submitButton.value = "true" Then
			.submitButton.value = false
			Check_String(windowbox.submitText.value)
			.returnToHTA.click()
		ElseIf .done.value = "allOps" Then
			objShell.Run sOPsCmd
			WScript.Quit	
		End If 
	loop
	.close																		'Closes the window
 end with
 ServerClose()																	'Function to close open connections and return settings back to original	
 Wscript.Quit

Function HTABox(sBgColor, h, w, l, t) 
	Dim IE, nRnd
	randomize : nRnd = Int(1000000 * rnd) 
	sCmd = "mshta.exe ""javascript:{new " _ 
		& "ActiveXObject(""InternetExplorer.Application"")" _ 
		& ".PutProperty('" & nRnd & "',window);" _ 
		& "window.moveTo(" & l & ", " & t & ");    " _
		& "window.resizeTo(" & w & "," & h & ")}""" 
	with CreateObject("WScript.Shell")
		.Run sCmd, 1, False 
		do until .AppActivate("javascript:{new ") : WSH.sleep 10 : loop 
	end with
	For Each IE In CreateObject("Shell.Application").windows 
		If IsObject(IE.GetProperty(nRnd)) Then 
			set HTABox = IE.GetProperty(nRnd)
			IE.Quit 
			HTABox.document.write LoadHTML(sBgColor)
			HTABox.document.title = "Operation 20"
			oShell.SendKeys "% x"
			Exit Function 
		End If 
	Next 
	MsgBox "HTA window not found." 
	wsh.quit
 End Function

Function connect2Scanner()
	Dim secs : secs = 0
	If machineString <> "Manual" and machineString <> "" Then
		windowBox.partMarkText.innerText = "Connect to " & machineString
		windowBox.partMarkButton.style.backgroundcolor = "orange"
		windowBox.partMarkButton.disabled = true
		windowBox.errorString.innerText = ""
	End If
	
	'Stores variable if connected to part marker
	If left(machineString, 3) = "COM" Then
		' loads port settings into winsock
		objComport.Open
		If( objComport.LastError <> 0 ) Then
			windowBox.partMarkText.innerText = "Error: " & machineString
			windowBox.errorString.innerText = objComport.LastError & " (" & objComport.GetErrorDescription( objComport.LastError ) & ")"
			windowBox.partMarkButton.style.backgroundcolor = "red"
			windowBox.partMarkButton.disabled = false
		Else
			windowBox.partMarkText.innerText = "Connected to " & machineString
			windowBox.partMarkButton.style.backgroundcolor = "limegreen"
			windowBox.partMarkButton.disabled = true
		End If
	ElseIf machineString = "Manual" Then
		windowBox.partMarkText.innerText = "Manual scanner mode"
		windowBox.partMarkButton.style.backgroundcolor = "limegreen"
		windowBox.partMarkButton.disabled = true
		windowBox.errorString.innerText = ""
		windowBox.manualSerialNumber.style.backgroundColor = "DimGrey"
		windowBox.SerialNumberText.style.visibility = "hidden"
		windowBox.inputFormDiv.style.visibility = "visible"
		windowBox.inputForm.disabled = false
		windowBox.inputForm.stringInput.disabled = false
		windowBox.inputForm.stringInput.focus
	ElseIf Left(machineString, 3) = "WJM" Then
		If winsock0.state <> sckClosed Then winsock0.Disconnect
		If RemoteHost <> "" and RemotePort <> "" Then 
			winsock0.RemoteHost = RemoteHost
			winsock0.RemotePort = RemotePort
			'Connects to the scanner
			On Error Resume Next
			winsock0.Connect    
			On Error GoTo 0
			'// MAIN DELAY - WAITS FOR CONNECTED STATE
			'// SOCKET ERROR RAISES WINSOCK ERROR SUB
			while winsock0.State <> sckError And winsock0.state <> sckConnected And winsock0.state <> sckClosing And secs < 25
				WScript.Sleep 1000  '// 1 sec delay in loop
				secs = secs + 1     '// wait 25 secs max
			Wend
		End If
		If winsock0.state = sckConnected Then 
			windowBox.partMarkText.innerText = "Connected to " & machineString
			windowBox.partMarkButton.style.backgroundcolor = "limegreen"
			windowBox.partMarkButton.disabled = true
		Else
			windowBox.partMarkText.innerText = "Error: " & machineString
			windowBox.partMarkButton.style.backgroundcolor = "red"
			windowBox.partMarkButton.disabled = false
		End If
	End If
 End Function

Function checkAccess()
	If AccessResult = false Then
		windowBox.accessText.innerText = "Access database not loaded"
		windowBox.accessButton.style.backgroundcolor = "red"
	Else
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
	End If
 End Function

Function adminSettings()
	windowBox.SerialNumberInput.value = ""
	duplicatePrefix = ""
	windowBox.errorString.innerText = "ADMIN ACCESS GRANTED"
	windowBox.duplicateButton.disabled = false
	windowBox.adminText.style.visibility = "visible"
	windowBox.adminButton.style.visibility = "visible"
	windowBox.adminString.style.visibility = "visible"
	windowBox.logoutButton.style.visibility = "visible"
 End Function

Function TrimString(ByVal VarIn)
	VarIn = Trim(VarIn)   
	If Len(VarIn) > 0 Then
		Do While AscW(Right(VarIn, 1)) = 10 or AscW(Right(VarIn, 1)) = 13
			VarIn = Left(VarIn, Len(VarIn) - 1)
		Loop
	End If
	TrimString = Trim(VarIn)
 End Function

Function Check_String(stringFromScanner)
	Dim SN1_Found : SN1_Found = false
	Dim SN2_Found : SN2_Found = false
	Dim stringObj, inputString
														'Run loop until conditions are met
	timerCnt = timerCount 
	inputString = TrimString(stringFromScanner)
	windowbox.errorDiv.style.background = ""
	windowBox.errorString.innerText = ""
	windowbox.submitText.value = ""
	If inputString = tabletPassword or inputString = computerPassword Then
		Exit Function
	ElseIf inputString = "Logout" Then
		Logout
		Exit Function
	ElseIf inputString = "Reset" Then
		CleanUpScreen
		windowBox.errorString.innerText = "Fields Reset"
		Exit Function
	ElseIf inputString = "AccessRetry" Then
		windowBox.done.value = "access"
		Exit Function
	ElseIf inputString = "Cancel" Then
		windowBox.done.value = "cancel"
		Exit Function
	ElseIf Left(inputString, 2) = "TF" Then
		windowbox.FixtureID.innerText = inputString
		windowbox.modalFixture.innerText = inputString
		MaintReset "", windowbox.FixtureID.innerText
	ElseIf Left(inputString, 3) = "CMM" Then
		If windowbox.commentModal.style.visibility = "visible" Then
			windowbox.commentModal.style.visibility = "hidden"
		Else
			windowbox.commentModal.style.visibility = "visible"
		End IF
	ElseIf Left(inputString, 4) = "WJM_" Then
		machineString = inputString
		sArg = """" & inputString & """"
		RemoteHost = ""
		RemotePort = ""
		Load_IP
		connect2Scanner
	ElseIf inputString = "Abrasive" or inputString = "Orifice" or inputString = "Mixing Tube" Then
		If windowbox.FixtureID.innerText = "" or windowbox.operatorID.innerText = "" Then
			windowbox.errorDiv.style.background = "red"
			windowBox.errorString.innerText = "Maintenance scan requires a fixture and operator:" & Chr(13) & inputString
		Else
			MaintReset inputString, windowbox.FixtureID.innerText
		End If
		Exit Function
	ElseIF Len(inputString) = 10 and Left(inputString, 1) = "H" and Mid(inputString, 9, 1) = "-" Then
		Load_SN(inputString)
	Else
		windowbox.operatorID.innerText = inputString
	End If	
	If windowbox.FixtureID.innerText <> "" and windowbox.blade1String.innerText <> "" and windowbox.blade2String.innerText <> "" and windowbox.operatorID.innerText <> "" Then
		LoadSNtoAccess
	End if
 End Function

Function Load_SN(serial_number)
	Dim objCmd : Set objCmd = GetNewConnection
	If objCmd is Nothing Then
		windowBox.errorString.innerText = "Error connecting to database, data not sent"
		windowBox.accessText.innerText = "Connection failed, click to retry."
		windowBox.accessButton.style.backgroundcolor = "red"
		windowBox.accessButton.disabled = false
		windowbox.errorDiv.style.background = "red"
		Exit Function
	ElseIf windowBox.accessButton.style.backgroundcolor <> "limegreen" Then
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
		windowBox.errorString.innerText = ""
		windowbox.errorDiv.style.background = ""
	End If
	
	Dim sqlQuery : sqlQuery = "SELECT [Slug Serial Number] FROM [00_AE_SN_Control] WHERE [Blade Serial Number] = '" & serial_number & "';"
	Dim rs : Set rs = objCmd.Execute(sqlQuery)
	Dim Slug_SN : DO WHILE NOT rs.EOF
		Slug_SN = rs.Fields(0)
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If Slug_SN = "" Then
		windowBox.errorString.innerText = "Serial Number not found"
		windowbox.errorDiv.style.background = "red"
		Exit Function
	End If
	
	sqlQuery = "SELECT [Blade Serial Number], [FIC Blade Part Number] FROM [00_AE_SN_Control] WHERE [Slug Serial Number] = '" & Slug_SN & "';"
	Set rs = objCmd.Execute(sqlQuery)
	Dim blade1SN, blade2SN : DO WHILE NOT rs.EOF
		If Right(rs.Fields(1), 1) = "1" Then
			blade1SN = rs.Fields(0)
		Elseif Right(rs.Fields(1), 1) = "2" Then
			blade2SN = rs.Fields(0)
		End If
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If blade1SN = "" or blade2SN = "" Then
		windowBox.errorString.innerText = "Second Serial Number not found"
		windowbox.errorDiv.style.background = "red"
		Exit Function
	End If
	
	sqlQuery = "SELECT COUNT(*) FROM [20_LPT5] WHERE [Blade SN Dash 1] = '" & blade1SN & "';"
	Set rs = objCmd.Execute(sqlQuery)
	If rs(0).value <> 0 Then
		windowBox.errorString.innerText = "Serial number already scanned" & chr(10) & "Scan fixture to overwrite data"
		windowbox.errorDiv.style.background = "red"
		windowbox.blade1Button.style.backgroundcolor = "red"
		windowbox.blade2Button.style.backgroundcolor = "red"
		windowBox.FixtureID.innerText = "" 
	Else
		windowbox.blade1Button.style.backgroundcolor = "limegreen"
		windowbox.blade2Button.style.backgroundcolor = "limegreen"
	End IF
	windowbox.blade1String.innerText = blade1SN
	windowbox.blade2String.innerText = blade2SN
 End Function

Function Logout()
	windowBox.operatorID.innerText = ""
	windowBox.errorString.innerText = "Logged Out"
	windowBox.accessButton.disabled = true
	CleanUpScreen
 End Function

Function MaintReset(inputString, FixtureID)
	Dim sqlString, sqlString2, MachineID, LocationID
	Dim rs, AbrasiveCnt, MixCnt, OrificeCnt, MaintCnt
	Dim CurrentTime : CurrentTime = Now
	Dim OprID : OprID = windowbox.operatorID.innerText
	
	Dim objCmd : Set objCmd = GetNewConnection
	If objCmd is Nothing Then
		windowBox.errorString.innerText = "Error connecting to database, data not sent"
		windowBox.accessText.innerText = "Connection failed, click to retry."
		windowBox.accessButton.style.backgroundcolor = "red"
		windowBox.accessButton.disabled = false
		windowbox.errorDiv.style.background = "red"
		Exit Function
	ElseIf windowBox.accessButton.style.backgroundcolor <> "limegreen" Then
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
		windowBox.errorString.innerText = ""
		windowbox.errorDiv.style.background = ""
	End If
	sqlString = "SELECT TOP 1 [MachineName], [Location] FROM [30_Fixtures] WHERE [FixtureID] = '" & FixtureID & "';"
	set rs = objCmd.Execute(sqlString)		
	DO WHILE NOT rs.EOF
		MachineID = rs.Fields(0)
		LocationID = rs.Fields(1)
		rs.MoveNext
	Loop	
	Set rs = Nothing
		
	sqlString = "SELECT TOP 1 [AbrasiveCnt], [MixCnt], [OrificeCnt] FROM [20_Counters] WHERE [MachineID] = '" & MachineID & "';"
	set rs = objCmd.Execute(sqlString)		
	DO WHILE NOT rs.EOF
		AbrasiveCnt = rs.Fields(0)
		MixCnt = rs.Fields(1)
		OrificeCnt = rs.Fields(2)
		rs.MoveNext
	Loop
	
	Select Case inputString
		Case"Abrasive"
			sqlString = "UPDATE [20_Counters] Set [AbrasiveCnt] = 0 WHERE [MachineID] = '" & MachineID & "';"
			MaintCnt = AbrasiveCnt
			AbrasiveCnt = 0
		Case "Orifice"
			sqlString = "UPDATE [20_Counters] Set [OrificeCnt] = 0 WHERE [MachineID] = '" & MachineID & "';"
			MaintCnt = OrificeCnt
			OrificeCnt = 0
		Case "Mixing Tube"
			sqlString = "UPDATE [20_Counters] Set [MixCnt] = 0 WHERE [MachineID] = '" & MachineID & "';"
			MaintCnt = MixCnt
			MixCnt = 0
	End Select
	If inputString <> "" Then
		objCmd.Execute(sqlString)
		sqlString2 = "INSERT INTO [20_Maint_History] ([MachineID], [MaintDate], [OperatorID], [MaintType], [Counter]) VALUES ('" & MachineID & "', '" & CurrentTime & "', '" & OprID & "', '" & inputString & "', " & MaintCnt & ");"
		objCmd.Execute(sqlString2)
	Else
		'CMMSearch FixtureID, LocationID, objCmd
	End If
	Set rs = Nothing
		objCmd.Close
		Set objCmd = Nothing
	
	windowbox.AbrasiveCount.innerText = AbrasiveCnt
	If AbrasiveCnt >= AbrasiveLimit Then windowbox.AbrasiveCount.Style.BackgroundColor = "red" Else windowbox.AbrasiveCount.Style.BackgroundColor = ""
	windowbox.MixCount.innerText = MixCnt
	If MixCnt >= MixLimit Then windowbox.MixCount.Style.BackgroundColor = "red" Else windowbox.MixCount.Style.BackgroundColor = ""
	windowbox.OrificeCount.innerText = OrificeCnt
	If OrificeCnt >= OrificeLimit Then windowbox.OrificeCount.Style.BackgroundColor = "red" Else windowbox.OrificeCount.Style.BackgroundColor = ""
	windowbox.MachineID.innerText = MachineID
	If inputString <> "" Then CleanUpScreen
 End Function

Function CMMSearch(FixtureID, LocationID, objCmd)
	Dim cParts1 : Set cParts1 = CreateObject("Scripting.Dictionary")
	Dim cParts2 : Set cParts2 = CreateObject("Scripting.Dictionary")
	Dim cDates : Set cDates = CreateObject("System.Collections.ArrayList")
	Dim cBlade(5)
	Dim sqlQuery(2)
	Dim toleranceArray(17)
	Dim FixtureID1, FixtureID2, rs, n, dateValue, i, slugArray, DateString, DateSerial, a, b, tolResult
	ReDim slugArray(RowCount + 1, 7)
	
	If LocationID mod 2 = 0 Then
		FixtureID1 = Left(FixtureID, Len(FixtureID) - 1) & LocationID - 1
		FixtureID2 = FixtureID
	Else
		FixtureID1 = FixtureID
		FixtureID2 = Left(FixtureID, Len(FixtureID) - 1) & LocationID + 1
	End If
	sqlQuery(0) = "SELECT [40_CMM_LPT5].[Serial Number], [40_CMM_LPT5].Date, [40_CMM_LPT5].[Part Number], [00_AE_SN_Control].[Slug Serial Number], "
	sqlQuery(0) = sqlQuery(0) & "[40_CMM_LPT5].[Dim 1_1], [40_CMM_LPT5].[Dim 1_2], [40_CMM_LPT5].[Dim 2_1], [40_CMM_LPT5].[Dim 2_2], [40_CMM_LPT5].[Dim 3_1], [40_CMM_LPT5].[Dim 3_2], [40_CMM_LPT5].[Dim 4_1], [40_CMM_LPT5].[Dim 4_2], [40_CMM_LPT5].[Dim 5_1], [40_CMM_LPT5].[Dim 5_2], "
	sqlQuery(0) = sqlQuery(0) & "[40_CMM_LPT5].[Dim 9_1], [40_CMM_LPT5].[Dim 9_2], [40_CMM_LPT5].[Dim 10_1], [40_CMM_LPT5].[Dim 10_2], [40_CMM_LPT5].[Dim 11 Max], [40_CMM_LPT5].[Dim 11 Min], [40_CMM_LPT5].[Dim 12 Max], [40_CMM_LPT5].[Dim 12 Min]"
	sqlQuery(0) = sqlQuery(0) & " FROM ([40_CMM_LPT5] INNER JOIN [20_LPT5_Cut_1] ON [40_CMM_LPT5].[Serial Number] = [20_LPT5_Cut_1].[Blade Serial Number]) INNER JOIN [00_AE_SN_Control] ON [40_CMM_LPT5].[Serial Number] = [00_AE_SN_Control].[Blade Serial Number]"
	sqlQuery(0) = sqlQuery(0) & " WHERE ((([20_LPT5_Cut_1].[Fixture Location])='" & FixtureID1 & "') and ([40_CMM_LPT5].Date >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - 7) & "'))"
	sqlQuery(0) = sqlQuery(0) & " ORDER BY [40_CMM_LPT5].Date DESC;"
	
	sqlQuery(1) = "SELECT [40_CMM_LPT5].[Serial Number], [40_CMM_LPT5].Date, [40_CMM_LPT5].[Part Number], [00_AE_SN_Control].[Slug Serial Number], "
	sqlQuery(1) = sqlQuery(1) & "[40_CMM_LPT5].[Dim 1_1], [40_CMM_LPT5].[Dim 1_2], [40_CMM_LPT5].[Dim 2_1], [40_CMM_LPT5].[Dim 2_2], [40_CMM_LPT5].[Dim 3_1], [40_CMM_LPT5].[Dim 3_2], [40_CMM_LPT5].[Dim 4_1], [40_CMM_LPT5].[Dim 4_2], [40_CMM_LPT5].[Dim 5_1], [40_CMM_LPT5].[Dim 5_2], "
	sqlQuery(1) = sqlQuery(1) & "[40_CMM_LPT5].[Dim 9_1], [40_CMM_LPT5].[Dim 9_2], [40_CMM_LPT5].[Dim 10_1], [40_CMM_LPT5].[Dim 10_2], [40_CMM_LPT5].[Dim 11 Max], [40_CMM_LPT5].[Dim 11 Min], [40_CMM_LPT5].[Dim 12 Max], [40_CMM_LPT5].[Dim 12 Min]"
	sqlQuery(1) = sqlQuery(1) & " FROM ([40_CMM_LPT5] INNER JOIN [20_LPT5_Cut_5] ON [40_CMM_LPT5].[Serial Number] = [20_LPT5_Cut_5].[Blade Serial Number]) INNER JOIN [00_AE_SN_Control] ON [40_CMM_LPT5].[Serial Number] = [00_AE_SN_Control].[Blade Serial Number]"
	sqlQuery(1) = sqlQuery(1) & " WHERE ((([20_LPT5_Cut_5].[Fixture Location])='" & FixtureID2 & "') and ([40_CMM_LPT5].Date >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - 7) & "'))"
	sqlQuery(1) = sqlQuery(1) & " ORDER BY [40_CMM_LPT5].Date DESC;"

	For n = 0 to 1
		set rs = objCmd.Execute(sqlQuery(n))
		DO WHILE NOT rs.EOF
			cBlade(0) = rs.Fields(0) '"Blade"
			cBlade(1) = rs.Fields(1) '"Date"
			cBlade(2) = rs.Fields(2) '"PartNumber"
			cBlade(3) = rs.Fields(3) ' "Slug"
			For j = 0 to 17
				If Not IsNull(rs.Fields(j + 4)) Then toleranceArray(j) = rs.Fields(j + 4)
			Next
			tolResult = toleranceCheck(toleranceArray)
			cBlade(4) = tolResult(0)
			cBlade(5) = tolResult(1)
			Erase toleranceArray
			DateString = Split(rs.Fields(1), " ")
			If UBound(DateString) = 2 Then
				DateSerial = CDbl(CDate(DateString(0))) + CDbl(CDate(DateString(1) & " " & DateString(2)))
			Else				
				DateSerial = CDbl(CDate(DateString(0))) + CDbl(CDate(DateString(1)))
			End If
			If Not cDates.Contains(DateSerial) Then
				cDates.Add DateSerial
			End If
			If Right(rs.Fields(2), 1) = "1" Then
				If Not cParts1.Exists(DateSerial) Then
					cParts1.Add DateSerial, cBlade
				End If
			ElseIf Right(rs.Fields(2), 1) = "2" Then
				If Not cParts2.Exists(DateSerial) Then
					cParts2.Add DateSerial, cBlade
				End If
			Else
				msgbox("INVALID PART NUMBER: " & rs.Fields(0) & " " & rs.Fields(2))
			End If
			rs.MoveNext
		Loop	
		Set rs = Nothing
	Next
	cDates.Sort
	Dim arrayKeys, j, SlugRow
	Dim RowX : RowX = 0
	For n = cDates.Count - 1 To 0 Step -1
		SlugRow = RowX
		If cParts1.Exists(cDates(n)) Then
			For j = 0 to RowCount - 1
				If slugArray(j, 0) = cParts1.Item(cDates(n))(3) Then
					SlugRow = j
				End If
			Next
			slugArray(SlugRow, 0) = cParts1.Item(cDates(n))(3)
			slugArray(SlugRow, 1) = cDates(n)
			slugArray(SlugRow, 2) = cParts1.Item(cDates(n))(0)
			slugArray(SlugRow, 3) = cParts1.Item(cDates(n))(4)
			slugArray(SlugRow, 4) = cParts1.Item(cDates(n))(5)
		End If
		If cParts2.Exists(cDates(n)) Then
			For j = 0 to RowCount - 1
				If slugArray(j, 0) = cParts2.Item(cDates(n))(3) Then
					SlugRow = j
				End If
			Next
			slugArray(SlugRow, 0) = cParts2.Item(cDates(n))(3)
			slugArray(SlugRow, 1) = cDates(n)
			slugArray(SlugRow, 5) = cParts2.Item(cDates(n))(0)
			slugArray(SlugRow, 6) = cParts2.Item(cDates(n))(4)
			slugArray(SlugRow, 7) = cParts2.Item(cDates(n))(5)
		End If
		If SlugRow = RowX Then RowX = RowX + 1
		If RowX > RowCount + 1 Then n = 0
	Next
	HTASlug blankArray
	HTASlug slugArray
 End Function

Function HTASlug(slugArray)
	Dim a, b, j, TextID, failArray1, failArray2, failSplit
	'FixtureID = "TF-19-0003:0002:5"
	For a = 1 to RowCount
		windowBox.Document.getElementByID("SN1_" & a & "Text").innerHTML = ""
		windowBox.Document.getElementByID("SN2_" & a & "Text").innerHTML = ""
		For b = 0 to Ubound(tolName)
			windowBox.Document.getElementByID("CMM" & a & "_" & b & "_1Button").style.backgroundcolor = ""
			windowBox.Document.getElementByID("CMM" & a & "_" & b & "_1Button").title = ""
			windowBox.Document.getElementByID("CMM" & a & "_" & b & "_2Button").style.backgroundcolor = ""
			windowBox.Document.getElementByID("CMM" & a & "_" & b & "_2Button").title = ""			
		Next
	Next
	For a = 1 to RowCount
		If slugArray(a - 1, 4) <> "" Then
			failArray1 = Split(slugArray(a - 1, 4), ", ")
		Else
			failArray1 = Array()
		End If
		If slugArray(a - 1, 7) <> "" Then
			failArray2 = Split(slugArray(a - 1, 7), ", ")
		Else
			failArray2 = Array()
		End If
		windowBox.Document.getElementByID("Part" & a & "_1Button").title = slugArray(a - 1, 2)
		If slugArray(a - 1, 4) <> "" Then
			windowBox.Document.getElementByID("Part" & a & "_1Button").title = windowBox.Document.getElementByID("Part" & a & "_1Button").title & chr(10) & Replace(slugArray(a - 1, 4), ", " , chr(10))
		End If
		windowBox.Document.getElementByID("Part" & a & "_2Button").title = slugArray(a - 1, 5)
		If slugArray(a - 1, 7) <> "" Then
			windowBox.Document.getElementByID("Part" & a & "_2Button").title = windowBox.Document.getElementByID("Part" & a & "_2Button").title & chr(10) & Replace(slugArray(a - 1, 7), ", " , chr(10))
		End If
		If slugArray(a - 1, 1) <> 0 Then
			windowBox.Document.getElementByID("Part" & a & "Text").innerHTML = CDate(slugArray(a - 1, 1))
			windowBox.Document.getElementByID("CMM" & a & "Text").innerHTML = Replace(CDate(slugArray(a - 1, 1)), chr(32), "<br>", 1, 1)
			windowBox.Document.getElementByID("SN1_" & a & "Text").innerHTML = slugArray(a - 1, 2)
			windowBox.Document.getElementByID("SN2_" & a & "Text").innerHTML = slugArray(a - 1, 5)
		Else
			windowBox.Document.getElementByID("Part" & a & "Text").innerHTML = ""
			windowBox.Document.getElementByID("CMM" & a & "Text").innerHTML = ""
		End If
		If slugArray(a - 1, 3) = "Pass" Then
			windowBox.Document.getElementByID("Part" & a & "_1Button").style.backgroundcolor = "limegreen"
			For b = 0 to Ubound(tolName)
				windowBox.Document.getElementByID("CMM" & a & "_" & b & "_1Button").style.backgroundcolor = "limegreen"
			Next
		ElseIF slugArray(a - 1, 3) = "Fail" Then
			windowBox.Document.getElementByID("Part" & a & "_1Button").style.backgroundcolor = "red"
			For b = 0 to Ubound(tolName)
				windowBox.Document.getElementByID("CMM" & a & "_" & b & "_1Button").style.backgroundcolor = "limegreen"
				For j = 0 to Ubound(failArray1)
					failSplit = Split(failArray1(j), ":")
					If failSplit(0) = tolName(b) Then
						windowBox.Document.getElementByID("CMM" & a & "_" & b & "_1Button").style.backgroundcolor = "red"
						windowBox.Document.getElementByID("CMM" & a & "_" & b & "_1Button").title = failSplit(1)
					End If
				Next
			Next
		Else
			windowBox.Document.getElementByID("Part" & a & "_1Button").style.backgroundcolor = ""
		End If
		If slugArray(a - 1, 6) = "Pass" Then
			windowBox.Document.getElementByID("Part" & a & "_2Button").style.backgroundcolor = "limegreen"
			For b = 0 to Ubound(tolName)
				windowBox.Document.getElementByID("CMM" & a & "_" & b & "_2Button").style.backgroundcolor = "limegreen"
			Next
		ElseIF slugArray(a - 1, 6) = "Fail" Then
			windowBox.Document.getElementByID("Part" & a & "_2Button").style.backgroundcolor = "red"
			For b = 0 to Ubound(tolName)
				windowBox.Document.getElementByID("CMM" & a & "_" & b & "_2Button").style.backgroundcolor = "limegreen"
				For j = 0 to Ubound(failArray2)
					failSplit = Split(failArray2(j), ":")
					If failSplit(0) = tolName(b) Then
						windowBox.Document.getElementByID("CMM" & a & "_" & b & "_2Button").style.backgroundcolor = "red"
						windowBox.Document.getElementByID("CMM" & a & "_" & b & "_2Button").title = failSplit(1)
					End If
				Next
			Next
		Else
			windowBox.Document.getElementByID("Part" & a & "_2Button").style.backgroundcolor = ""
		End If
	Next
 End Function

Function toleranceCheck(toleranceArray)
	Dim n
	Dim failString : failString = ""
	Dim tolName : tolName = Array("Dim 1.1",	"Dim 1.2",	"Dim 2.1",	"Dim 2.2",	"Dim 3.1",	"Dim 3.2",	"Dim 4.1",	"Dim 4.2",	"Dim 5.1",	"Dim 5.2",	"Dim 9.1",	"Dim 9.2",	"Dim 10.1",	"Dim 10.2",	"Dim 11 Max",	"Dim 11 Min",	"Dim 12 Max",	"Dim 12 Min")
	Dim minTol : minTol =   Array(40.8, 		40.8, 		155.3, 		155.3, 		168.2, 		168.2,  	155.3, 		155.3, 		26.9, 		26.9, 		16.75,  	16.75,  	32.25,  	32.25,  	-0.5,  	 		-0.5,  			-0.5,  			-0.5)
	Dim maxTol : maxTol =   Array(41.8, 		41.8, 		156.3, 		156.3,  	169.2,  	169.2,  	156.3,  	156.3,  	27.9,  		27.9,  		17.75,  	17.75,  	99.99,  	99.99,   	 0.5,   		 0.5,   		 0.5,  			 0.5)
	toleranceCheck = "Pass"
	For n = lbound(toleranceArray) to ubound(toleranceArray)
		If IsNull(toleranceArray(n)) or IsEmpty(toleranceArray(n)) Then
		ElseIf toleranceArray(n) < minTol(n) or toleranceArray(n) > maxTol(n) Then
			toleranceCheck = "Fail"
			failString = failString & tolName(n) & ": " & toleranceArray(n) & " (" &  minTol(n) & " to " & maxTol(n) & "), "
		End If
	Next
	If toleranceCheck = "Fail" Then failString = Left(failString, len(failString) - 2)
	toleranceCheck = Array(toleranceCheck, failString)
 End Function

Function RemainderAccess()
	For j=LBound(AccessArray, 1) to UBound(AccessArray, 1)	
		If AccessArray(0, j) = Serial_Number and Right(AccessArray(2, j), 1) = "1" Then
			SN1_Found = AccessArray(1, j)
		ElseIf AccessArray(0, j) = Serial_Number and Right(AccessArray(2, j), 1) = "2" Then
			SN2_Found = AccessArray(1, j)
		ElseIf AccessArray(0, j) = Serial_Number Then
			SN1_Found = AccessArray(1, j)
		End If
		If SN1_Found <> false and SN2_Found <> false Then
			Exit For
		End If
	Next
	windowBox.duplicateButton.style.visibility = "hidden"
	If InStr(PMString, SN1_Found) <> 0 or InStr(PMString, SN2_Found) <> 0 Then
		windowBox.errorString.innerText = duplicatePrefix & "Serial Number Already Marked: " & Serial_Number
		windowBox.blade1string.innerText = SN1_Found
		windowBox.blade2string.innerText = SN2_Found
		windowBox.duplicateButton.style.visibility = "visible"
	ElseIf SN1_Found <> false and SN2_Found <> false Then
		windowBox.errorString.innerText = ""
		windowBox.blade1string.innerText = SN1_Found
		windowBox.blade2string.innerText = SN2_Found
		windowBox.accessNewEntry.Value = True
		Load_SN_to_PM ""
	ElseIf SN1_Found = false and SN2_Found = false Then
		windowBox.errorString.innerText = "Serial number not found: " & Serial_Number
		windowBox.blade2string.innerText = "Waiting..."
		windowBox.blade2string.innerText = "Waiting..."
	ElseIf SN1_Found <> false or SN2_Found <> false Then
		windowBox.errorString.innerText = "Missing blade serial number: " & Serial_Number
		windowBox.blade2string.innerText = "Waiting..."
		windowBox.blade2string.innerText = "Waiting..."
	End If
 End Function

Function LoadSNtoAccess()
	Dim Serial1, Serial2, sqlString, strQuery1, strQuery2, Operator, FixtureID, strQueryPre, Comments
	Dim rs, MachineID, AbrasiveCnt, MixCnt, OrificeCnt
	Dim updateCount : updateCount = False
	Dim CurrentTime : CurrentTime = Now
	
	Dim objCmd : Set objCmd = GetNewConnection
	If objCmd is Nothing Then
		windowBox.errorString.innerText = "Error connecting to database, data not sent"
		windowBox.accessText.innerText = "Connection failed, click to retry."
		windowBox.accessButton.style.backgroundcolor = "red"
		windowBox.accessButton.disabled = false
		windowbox.errorDiv.style.background = "red"
		Exit Function
	ElseIf windowBox.accessButton.style.backgroundcolor <> "limegreen" Then
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
		windowBox.errorString.innerText = ""
		windowbox.errorDiv.style.background = ""
	End If
	
	Operator = windowBox.operatorID.innerText
	FixtureID = windowBox.FixtureID.innerText
	Serial1 = windowBox.blade1string.innerText
	Serial2 = windowBox.blade2string.innerText
	If windowbox.blade1Button.style.backgroundcolor = "limegreen" Then
		sqlString = "INSERT INTO [20_LPT5] ([Blade SN Dash 1], [Blade SN Dash 2], [Cut Date], [Operator], [Program Version], [Fixture Location]) " _
				  & "VALUES ('" & Serial1 & "', '" & Serial2 & "', '" & CurrentTime & "', '" & Operator & "', '" & ProgramVersion & "', '" & FixtureID & "'); "
		Set rs = objCmd.Execute(sqlString)
		Set rs = Nothing
		sqlString = "SELECT TOP 1 [MachineName] FROM [30_Fixtures] WHERE [FixtureID] = '" & FixtureID & "';"
		set rs = objCmd.Execute(sqlString)		
		DO WHILE NOT rs.EOF
			MachineID = rs.Fields(0)
			rs.MoveNext
		Loop	
		Set rs = Nothing
		sqlString = "SELECT TOP 1 [AbrasiveCnt], [MixCnt], [OrificeCnt] FROM [20_Counters] WHERE [MachineID] = '" & MachineID & "';"
		set rs = objCmd.Execute(sqlString)		
		DO WHILE NOT rs.EOF
			AbrasiveCnt = rs.Fields(0) + 1
			MixCnt = rs.Fields(1) + 1
			OrificeCnt = rs.Fields(2) + 1
			rs.MoveNext
		Loop	
		Set rs = Nothing
		sqlString = "UPDATE [20_Counters] Set [AbrasiveCnt] = " & AbrasiveCnt & ", [MixCnt] = " & MixCnt & ", [OrificeCnt] = " & OrificeCnt & "  WHERE [MachineID] = '" & MachineID & "';"
		objCmd.Execute(sqlString)
		windowbox.AbrasiveCount.innerText = AbrasiveCnt
		If AbrasiveCnt >= AbrasiveLimit Then windowbox.AbrasiveCount.Style.BackgroundColor = "red" Else windowbox.AbrasiveCount.Style.BackgroundColor = ""
		windowbox.MixCount.innerText = MixCnt
		If MixCnt >= MixLimit Then windowbox.MixCount.Style.BackgroundColor = "red" Else windowbox.MixCount.Style.BackgroundColor = ""
		windowbox.OrificeCount.innerText = OrificeCnt
		If OrificeCnt >= OrificeLimit Then windowbox.OrificeCount.Style.BackgroundColor = "red" Else windowbox.OrificeCount.Style.BackgroundColor = ""
		windowbox.MachineID.innerText = MachineID
	Else
		sqlString = "SELECT [Cut Date], [Operator], [Program Version], [Fixture Location], [Comments] FROM [20_LPT5] WHERE [Blade SN Dash 1] = '" & Serial1 & "';"
		Set rs = objCmd.Execute(sqlString)		
		DO WHILE NOT rs.EOF
			Comments = "Changed " & CurrentTime & ";" & rs.Fields(0) & ";" & rs.Fields(1) & ";" & rs.Fields(2) & ";" & rs.Fields(3) & "|" & rs.Fields(4)
			rs.MoveNext
		Loop	
		Set rs = Nothing
		sqlString = "UPDATE [20_LPT5] SET [Operator]='" & Operator & "', [Program Version]='" & ProgramVersion & "', " _
				  & "[Fixture Location]='" & FixtureID & "', [Comments]='" & Comments & "' WHERE  [Blade SN Dash 1]='" & Serial1 & "';"
		Set rs = objCmd.Execute(sqlString)
		Set rs = Nothing
	End If
		
	objCmd.Close
	Set objCmd = Nothing
	windowbox.errorDiv.style.background = "limegreen"
	windowBox.errorString.innerText = "Successful"
	CleanUpScreen

 End Function

Function CleanUpScreen()
	windowBox.FixtureID.innerText = ""
	windowBox.blade1string.innerText = ""
	windowBox.blade2string.innerText = ""
	windowbox.blade1Button.style.backgroundcolor = ""
	windowbox.blade2Button.style.backgroundcolor = ""
 End Function

Function GetNewConnection()
	Dim objCmd : Set objCmd = CreateObject("ADODB.Connection")
	Dim sConnection : sConnection = "Data Source=" & dataSource & ";Initial Catalog=CMM_Repository;Integrated Security=SSPI;"
	Dim sProvider : sProvider = "SQLOLEDB.1;"
	
	
	objCmd.ConnectionString	= sConnection	'Contains the information used to establish a connection to a data store.
	'objCmd.ConnectionTimeout				'Indicates how long to wait while establishing a connection before terminating the attempt and generating an error.
	'objCmd.CommandTimeout					'Indicates how long to wait while executing a command before terminating the attempt and generating an error.
	'objCmd.State							'Indicates whether a connection is currently open, closed, or connecting.
	objCmd.Provider = sProvider				'Indicates the name of the provider used by the connection.
	'objCmd.Version							'Indicates the ADO version number.
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
	Dim sqlString, SN_Size, rs
	Dim SN_String : SN_String = "Slug_SN;"
	Dim objCmd : set objCmd = GetNewConnection
	
	If objCmd is Nothing Then Load_Access = false : Exit Function
	sqlString = "Select [FixtureID], [ProgramID] From [30_Fixtures]"
	set rs = objCmd.Execute(sqlString)
	ReDim AccessArray(1,0)
	AccessArray(0, 0) = "Fixture_ID"
	AccessArray(1, 0) = "ProgramID"
	
	DO WHILE NOT rs.EOF
		SN_Size = UBound(AccessArray, 1) + 1
		ReDim Preserve AccessArray(1, SN_Size)
		AccessArray(0, SN_Size) = rs.Fields(0)
		AccessArray(1, SN_Size) = rs.Fields(1)
		rs.MoveNext
	Loop
	
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
	Load_Access = true
 End Function

Function Load_IP()
	Dim sqlString, SN_Size, rs
	Dim SN_String : SN_String = "Slug_SN;"
	Dim objCmd : set objCmd = GetNewConnection
	On Error GoTo 0
	If objCmd is Nothing Then Exit Function
	sqlString = "Select [IPAddress], [Port] From [00_Machine_IP] WHERE [DeviceType] = 'CognexBTHandheld' AND [MachineName] = '" & machineString & "'"
	If machineString <> "Manual" Then
		set rs = objCmd.Execute(sqlString)		
		DO WHILE NOT rs.EOF
			RemoteHost = rs.Fields(0)
			RemotePort = rs.Fields(1)
			rs.MoveNext
		Loop	
	End If
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
	
 End Function

'// WINSOCK DATA ARRIVES
Sub winsock0_OnDataArrival(bytesTotal)
    winsock0.GetData strData, vbString
	TimeCnt = 0
	Do
		wsh.sleep WaitMS
		TimeCnt = TimeCnt + WaitMS
	Loop While shl.AppActivate("Operation 20") = False and TimeCnt < TimeOut
	oShell.SendKeys "% x"
    WScript.Sleep 1000
	Check_String strData
 End Sub


'// WINSOCK ERROR
Sub winsock0_OnError(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
	windowBox.partMarkText.innerText = "Error: " & machineString
	windowBox.partMarkButton.style.backgroundcolor = "red"
	windowBox.partMarkButton.disabled = false
    windowBox.errorString.innerText = "Scanner Error: " & Number & vbCrLf & Description
 End Sub

Sub ReadResponse(ByVal objComport)
  Dim str : str = "notempty"
  Dim TimeCnt 
  objComport.Sleep(waitMS)
  While (str <> "")
    str = objComport.ReadString()
    If (str <> "") Then
	TimeCnt = 0
	Do
		wsh.sleep WaitMS
		TimeCnt = TimeCnt + WaitMS
	Loop While shl.AppActivate("Operation 20") = False and TimeCnt < TimeOut
	oShell.SendKeys "% x"
		Call Check_String(str)
    End If

  WEnd
 End Sub
 
'// EXIT SCRIPT
Sub ServerClose()
	If debugMode = False Then On Error Resume Next

	WScript.Sleep 1000  '// REQUIRED OR ERRORS
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	'objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"

	objComport.Close()
	objComport.Clear()
	
	If winsock0.state <> sckClosed Then winsock0.Disconnect
    winsock0.CloseWinsock
    Set winsock0 = Nothing
	
	windowBox.close
	
	On Error GoTo 0
    Wscript.Quit
 End Sub

 Function getFlowPID()
	 Dim objWMIService, objItem, colItems
	 Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	 Set colItems = objWMIService.ExecQuery("Select * From Win32_Process where name='FlowCUT.exe'")
	 For Each objItem in colItems
		getFlowPID = objItem.ProcessID
	 Next
  End Function

'Function to create all of the JS and HTML code for the window
Function LoadHTML(sBgColor)
	Dim a, b
	
	'HTA String
	LoadHTML = "<HTA:Application border=thin />"	
	
	'CSS String
	LoadHTML = LoadHTML _	
		& "<head><style>" _
		& "body {" _
			& "background-color: " & sBgColor & ";" _
			& "font:normal 28px Tahoma;" _
			& "border-Style:outset" _
			& "border-Width:3px" _
			& "}" _
		& ".CMM {" _
			& "font:normal 20px Tahoma;" _
			& "}" _
		& ".dimension {" _
			& "font:normal 14px Tahoma;" _
			& "}" _
		& ".HTAButton {" _
			& "border-top-left-radius: 50%;" _
			& "border-radius: 12px;" _
			& "}" _
		& ".unselectable {" _
			& "-moz-user-select: -moz-none;" _
			& "-khtml-user-select: none;" _
			& "-webkit-user-select: none;" _
			& "-o-user-select: none;" _
			& "user-select: none;" _
			& "}" _
		& ".closeButton {" _
			& "background-color: red;" _
			& "height: 30px;" _
			& "width: 30px;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
			& "}" _
		& ".modal {" _
			& "background-color: red;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
			& "}" _
		& "#commentModal, #nameModal {" _
			& "font:normal 30px Tahoma;" _
			& "background-color = 'grey';"
	If adminMode <> true Then
		LoadHTML = LoadHTML _
			& "visibility: hidden;" 
	End If
		LoadHTML = LoadHTML _
			& "}" _
		& ".verticaltext {" _
			& "writing-mode: tb-rl;" _
			& "" _
			& "}" _
		& ".opButton {" _
			& "background-color: blue;" _
			& "height: 30px;" _
			& "width: 30px;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
			& "color: white;" _
			& "}" _
		& "div {" _
			& "position:absolute;"
	If adminMode = true Then
		LoadHTML = LoadHTML _
			& "border-style: solid;" _
			& "border-Width:1px;"
	End If
	LoadHTML = LoadHTML _
			& "}" _
		& "</style>"
			
	'JS String
	LoadHTML = LoadHTML _
		& "<script language='javascript'>" _
		& "function manualButton() {" _
			& "if (document.getElementById('manualSerialNumber').style.backgroundColor == 'dimgrey') {" _
				& "document.getElementById('manualSerialNumber').style.backgroundColor = '';" _
				& "document.getElementById('SerialNumberText').style.visibility = 'visible';" _
				& "document.getElementById('stringInput').value = '';" _
				& "document.getElementById('inputFormDiv').style.visibility = 'hidden';" _
				& "document.getElementById('inputForm').disabled = true;" _
				& "document.getElementById('stringInput').disabled = true;" _
				& "document.getElementById('errorString').innerText = '';" _
				& "document.getElementById('manualSerialNumber').disabled = true;" _
				& "document.getElementById('manualSerialNumber').disabled = false;" _
			& "} else {" _
				& "document.getElementById('manualSerialNumber').style.backgroundColor = 'DimGrey';" _
				& "document.getElementById('SerialNumberText').style.visibility = 'hidden';" _
				& "document.getElementById('inputFormDiv').style.visibility = 'visible';" _
				& "document.getElementById('inputForm').disabled = false;" _
				& "document.getElementById('stringInput').disabled = false;" _
				& "document.getElementById('manualSerialNumber').disabled = true;" _
				& "document.getElementById('manualSerialNumber').disabled = false;" _
			& "}" _
		& "}" _
		& "function crossOutButton() {" _
			& "if (document.getElementById('crossOutMode').value == 'true') {" _
				& "document.getElementById('crossOutMode').value = false;" _
				& "document.getElementById('adminString').innerText = 'Click to cross out part mark';" _
				& "document.getElementById('adminButton').style.backgroundColor = '';" _
			& "} else {" _
				& "document.getElementById('crossOutMode').value = true;" _
				& "document.getElementById('adminString').innerText = 'Click to disable cross out mode';" _
				& "document.getElementById('adminButton').style.backgroundColor = 'DimGrey';" _
			& "}" _
			& "document.getElementById('crossOutClick').value = true;" _
		& "}" _
		& "function logoutAdmin() {" _
			& "document.getElementById('errorString').innerText = 'LOGGED OUT';" _
			& "document.getElementById('duplicateButton').disabled = true;" _
			& "document.getElementById('adminText').style.visibility = 'hidden';" _
			& "document.getElementById('adminButton').style.visibility = 'hidden';" _
			& "document.getElementById('adminString').style.visibility = 'hidden';" _
			& "document.getElementById('logoutButton').style.visibility = 'hidden';" _
		& "}" _
		& "function inputComplete() {" _
			& "document.getElementById('submitText').value = document.getElementById('stringInput').value;" _
			& "document.getElementById('submitButton').value = true;" _
			& "event.cancelBubble = true;" _
			& "event.returnValue = false;" _
			& "return false;" _
		& "}" _
		& "function HTAReturn() {" _
			& "document.getElementById('stringInput').value = '';" _
			& "" _
		& "}" _
		& "function logoutFunction() {" _
			& "document.getElementById('operatorID').innerText = '';" _
			& "document.getElementById('FixtureID').innerText = '';" _
			& "document.getElementById('blade1String').innerText = '';" _
			& "document.getElementById('blade2String').innerText = '';" _
			& "document.getElementById('errorDiv').style.background = '';" _
			& "document.getElementById('accessButton').disabled = true;" _
			& "document.getElementById('errorString').innerText = 'Logged Out';" _
			& "document.getElementById('logoutButton').disabled = true;" _
			& "document.getElementById('logoutButton').disabled = false;" _
		& "}" _
		& "function resetFunction() {" _
			& "document.getElementById('FixtureID').innerText = '';" _
			& "document.getElementById('blade1String').innerText = '';" _
			& "document.getElementById('blade2String').innerText = '';" _
			& "document.getElementById('errorDiv').style.background = '';" _
			& "document.getElementById('accessButton').disabled = true;" _
			& "document.getElementById('errorString').innerText = 'Fields Reset';" _
			& "document.getElementById('resetButton').disabled = true;" _
			& "document.getElementById('resetButton').disabled = false;" _
		& "}" _
		& "function CMMFunction(e) {" _
			& "e = e || window.event;" _
			& "e = e.target || e.srcElement;" _
			& "if (e.nodeName === 'BUTTON') {" _
				& "document.getElementById('commentModal').style.visibility = 'visible';" _
			& "}" _
		& "}" _
		& "function dimFunction(e) {" _
			& "e = e || window.event;" _
			& "e = e.target || e.srcElement;" _
			& "if (e.nodeName === 'BUTTON' && e.title != '') {alert(e.title);}" _
		& "}" _
		& "function cancelComment() {" _
			& "document.getElementById('commentModal').style.visibility = 'hidden';" _
		& "};" _
		& "</script></head>"
		
	'Body Start String							
	LoadHTML = LoadHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
	
	'Access Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=accessButton style='height: 30px; width: 30px; text-align: center;background-color:orange;' disabled onclick='done.value=""access""'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 60px; height: 30px; width: 380px; text-align: left;' id='accessText'>Waiting for Access connection&nbsp;</div>"
		
	'Scanner Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 60px; left: 25px;height: 30px; width: 30px;'>" _
		& "<button id=partMarkButton style='height: 30px; width: 30px;background-color:orange;' disabled onclick='done.value=""scanner""'><span>&nbsp;</span></button></div>" _
		& "<div id=partMarkText unselectable='on' class='unselectable' style='top: 60px; left: 60px;height: 30px; width: 380px;'><span>Waiting for scanner connection&nbsp;</span></div>" 
		
	'Reset Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id=resetButton style='height: 30px; width: 30px;' onclick='resetFunction()'><span>&nbsp;</span></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 60px;height: 30px; width: 380px;'><span>Click to reset fields&nbsp;</span></div>" 
		
	'Logout Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id=logoutButton style='height: 30px; width: 30px;' onclick='logoutFunction()'><span>&nbsp;</span></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 60px;height: 30px; width: 380px;'><span>Click to Logout&nbsp;</span></div>" 
		
	'Input String
	LoadHTML = LoadHTML _
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id='manualSerialNumber' style='height: 30px; width: 30px;' onclick='manualButton()'><span>&nbsp;</span></button></div>" _
		& "<div id='SerialNumberText' unselectable='on' class='unselectable' style='top: 165px; left: 60px;height: 30px; width: 380px;'><span>Click to enter data manually&nbsp;</span></div>" _
		& "<div id='inputFormDiv' style='top: 165px; left: 60px; height: 30px; width: 380px;visibility:hidden;'>" _
			& "<form id=inputForm onsubmit='inputComplete();' disabled>" _
				& "<input id=stringInput style='top: 0px; left: 0px; height: 30px; width: 380px;' value='' disabled /></form></div>"
	
	'Timer String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 525px; height: 30px; width: 400px; text-align: left;' id='timerText'></div>"
		
	'Maintenance String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 200px; left: 525px; height: 30px; width: 350px; text-align: center;'>Maintenance Counters&nbsp;</div>"
		
	'Operator String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 230px; left: 25px; height: 30px; width: 175px; text-align: right;'>Operator:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 230px; left: 200px; height: 30px; width: 275px; text-align: center;' id=OperatorID></div>"
		
	'Machine String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 230px; left: 525px; height: 30px; width: 175px; text-align: right;'>Machine ID:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 230px; left: 700px; height: 30px; width: 175px; text-align: center;' id=MachineID></div>"
		
	'Fixture String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 260px; left: 25px; height: 30px; width: 175px; text-align: right;'>Fixture:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 260px; left: 200px; height: 30px; width: 275px; text-align: center;' id=FixtureID></div>"
	
	'Abrasive String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 260px; left: 525px; height: 30px; width: 175px; text-align: right;'>Abrasive:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 260px; left: 700px; height: 30px; width: 175px; text-align: center;' id=AbrasiveCount></div>"
	
	'Blade 1 SN String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 290px; left: 25px; height: 30px; width: 175px; text-align: right;'>Blade 1:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 290px; left: 200px; height: 30px; width: 250px; text-align: center;' id=blade1String></div>" _	
		& "<div unselectable='on' class='unselectable' style='top: 292px; left: 450px;height: 30px; width: 30px;'>" _
			& "<button id=blade1Button style='height: 30px; width: 30px;' disabled><span>&nbsp;</span></button></div>"
	
	'Mixing Tube String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 290px; left: 525px; height: 30px; width: 175px; text-align: right;'>Mixing Tube:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 290px; left: 700px; height: 30px; width: 175px; text-align: center;' id=MixCount></div>"
	
	'Blade 2 SN String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 320px; left: 25px; height: 30px; width: 175px; text-align: right;'>Blade 2:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 320px; left: 200px; height: 30px; width: 250px; text-align: center;' id=blade2String></div>" _	
		& "<div unselectable='on' class='unselectable' style='top: 322px; left: 450px;height: 30px; width: 30px;'>" _
			& "<button id=blade2Button style='height: 30px; width: 30px;' disabled><span>&nbsp;</span></button></div>"
	
	'Orifice String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 320px; left: 525px; height: 30px; width: 175px; text-align: right;'>Orifice:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 320px; left: 700px; height: 30px; width: 175px; text-align: center;' id=OrificeCount></div>"
			
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div id=errorDiv style='top: 385px; left: 0px; height: 355px; width: 500px;'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 410px; left: 50px; height: 300px; width: 400px; text-align: center;' id=errorString></div>"
		
		
		
	'CMM History String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 385px; left: 525px; height: 30px; width: 225px; text-align: center;' id=LocationText>CMM Date&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 385px; left: 750px; height: 30px; width: 50px; text-align: center;'>-1</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 385px; left: 800px; height: 30px; width: 50px; text-align: center;'>-2</div>"
		
	For a = 1 to RowCount
		LoadHTML = LoadHTML _	
			& "<div unselectable='on' class='unselectable CMM' style='top: " & a * 35 + 385 & "px; left: 525px; height: 30px; width: 225px; text-align: center;' id=Part" & a & "Text>&nbsp;</div>" _
			& "<div unselectable='on' class='unselectable' style='top: " & a * 35 + 385 & "px; left: 760px;height: 30px; width: 30px;'>" _
				& "<button style='height: 30px; width: 30px;' title='No serial number' onclick='CMMFunction()' id=Part" & a & "_1Button></button></div>" _
			& "<div unselectable='on' class='unselectable' style='top: " & a * 35 + 385 & "px; left: 810px;height: 30px; width: 30spx;'>" _
				& "<button style='height: 30px; width: 30px;' title='No serial number' onclick='CMMFunction()' id=Part" & a & "_2Button></button></div>"
	Next
	
	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div style='top: 0px; left: -100px;'><button type=hidden id=returnToHTA 		style='visibility:hidden;' value=false onclick='HTAReturn()'><center><span>&nbsp;</span></button></div>" _
		& "<div style='top: 0px; left: -100px;'><input type=hidden id=done 				style='visibility:hidden;' value=false><center><span>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: -100px;'><input type=hidden id=submitButton 		style='visibility:hidden;' value=false><center><span>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: -100px;'><input type=hidden id=submitText 		style='visibility:hidden;' value=false><center><span>&nbsp;</span></div>" 
		
	'Modal Comment Div String
	LoadHTML = LoadHTML _
		& "<div id='commentModal' style='top: 1px; left: 1px; height: 773px; width: 893px;'>" _
		& "<div unselectable='on' class='unselectable' style='top: 50px; left: 50px; height: 40px; width: 350px;' id=modalFixture>No fixture scanned</div>"
		
	'CMM History String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 110px; left: 50px; height: 30px; width: 175px; text-align: center;' id=LocationText>Date</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 110px; left: 225px; height: 30px; width: 175px; text-align: center;' id=LocationText>Blade S/N</div>"
	For a = 0 to Ubound(tolName)
		LoadHTML = LoadHTML _
			& "<div unselectable='on' class='unselectable dimension verticaltext' style='top: 50px; left: " & 400 + a * 25 & "px; height: 90px; width: 25px; text-align: right;'>" & tolName(a) & "</div>"
	Next
	For a = 1 to RowCount
		LoadHTML = LoadHTML _	
			& "<div unselectable='on' class='unselectable CMM' style='top: " & a * 70 + 75 & "px; left: 50px; height: 50px; width: 175px; text-align: center;' id=CMM" & a & "Text>&nbsp;</div>" _
			& "<div unselectable='on' class='unselectable CMM' style='top: " & a * 70 + 75 & "px; left: 225px; height: 25px; width: 175px; text-align: center;' id=SN1_" & a & "Text>&nbsp;</div>" _
			& "<div unselectable='on' class='unselectable CMM' style='top: " & a * 70 + 100 & "px; left: 225px; height: 25px; width: 175px; text-align: center;' id=SN2_" & a & "Text>&nbsp;</div>"
		
		For b = 0 to Ubound(tolName)
			LoadHTML = LoadHTML _	
				& "<div unselectable='on' class='unselectable' style='top: " & a * 70 + 75 & "px; left: " & b * 25 + 400 & "px;height: 25px; width: 25px;'>" _
					& "<button style='height: 25px; width: 25px;' onclick='dimFunction()' id=CMM" & a & "_" & b & "_1Button></button></div>" _
				& "<div unselectable='on' class='unselectable' style='top: " & a * 70 + 100 & "px; left: " & b * 25 + 400 & "px;height: 25px; width: 25px;'>" _
					& "<button style='height: 25px; width: 25px;' onclick='dimFunction()' id=CMM" & a & "_" & b & "_2Button></button></div>"
		Next
	Next
	
	LoadHTML = LoadHTML _
		& "<div style='top: 713px; left: 743px; height: 50px; width: 100px;'><input type=button value='Close' style='height: 48px; width: 100px;' onclick='cancelComment()'></div>" _
		& "</div>"
		
	'End Body String
	LoadHTML = LoadHTML _
		& "</body>"

	'Footer String
	LoadHTML = LoadHTML _
		& "<footer><script language='javascript'>" _
			& "document.getElementById('stringInput').focus();" _
		& "</script></footer>"

 End Function
 