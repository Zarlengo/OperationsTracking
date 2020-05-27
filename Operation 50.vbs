Option Explicit
 '****** Version History *********
 '1.0	10/15/2018	- Initial release to production
 '1.1	4/9/2019	- Updated "In MRB" search to ignore Null. MRB location table tracking was updated to record in and out from the cage, null is for when the part is no longer in the cage [shows red in the final inspection, but didn't error in shipping]
 '***************************************
	Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")

	Dim allOPSsource : allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\All Operations.vbs"
	Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """" & sArg
	Dim dataSource : dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"
	Dim adminPassword : adminPassword = "FLOW288"
	Dim tabletPassword : tabletPassword = "Fl0wSh0p17"
	Dim computerPassword : computerPassword = "Snowball18!"

	Dim closeWindow : closeWindow = false
	Dim errorWindow : errorWindow = false
	Dim adminMode : adminMode = false
	Dim debugMode : debugMode = false
	Dim POChange : POChange = false
	Dim boxSize : boxSize = 12
	Dim notFoundCount : notFoundCount = 0
	Dim prodCount : prodCount = -1
	Dim prodArray()
	Dim fieldArray(2)

	Dim MRBArray : MRBArray = Array("MRB Staging", "Scrap", "Scrap Box 1", "Scrap Box 2", _
									"A1", "A2", "A3", "A4", "A5", "A6", _
									"B1", "B2", "B3", "B4", "B5", "B6", _
									"C1", "C2", "C4", "C5", "C5", "C6", _
									"D1", "D2", "D3", "D4", "D5", "D6", _
									"E1", "E2", "E3", "E4", "E5", "E6", _
									"F1", "F2", "F3", "F4", "F5", "F6", _
									"G1", "G2", "G3", "G4", "G5", "G6", _
									"H1", "H2", "H3", "H4", "H5", "H6", _
									"I1", "I2", "I3", "I4", "I5", "I6", _
									"J1", "J2", "J3", "J4", "J5", "J6", _
									"K1", "K2", "K3", "K4", "K5", "K6", _
									"L1", "L2", "L3", "L4", "L5", "L6")
									
	Dim DispositionArray : DispositionArray = Array("In MRB", "Return to Customer", "Scrap", "Supplier Rework/Remake", "Use As Is", "Void")
	Dim StatusArray : StatusArray = Array("Need to be created", "Created", "Closed")
									
									

	Dim strData, fieldsBad
	Dim SendData, RecieveData, wmi, cProcesses, oProcess
	Dim machineBox, strSelection, RemoteHost, RemotePort

	'************** TO DO *****************



	'****** CHANGE THESE SETTINGS *********



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

	Const adOpenDynamic			= 2	 '// Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
	Const adOpenForwardOnly		= 0	 '// Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
	Const adOpenKeyset			= 1	 '// Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
	Const adOpenStatic			= 3	 '// Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
	Const adOpenUnspecified		= -1 '// Does not specify the type of cursor.

	Const adLockBatchOptimistic	= 4	 '// Indicates optimistic batch updates. Required for batch update mode.
	Const adLockOptimistic		= 3	 '// Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method.
	Const adLockPessimistic		= 2	 '// Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing.
	Const adLockReadOnly		= 1	 '// Indicates read-only records. You cannot alter the data.
	Const adLockUnspecified		= -1 '// Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.

	Const adStateClosed			= 0  '// The object is closed
	Const adStateOpen			= 1  '// The object is open
	Const adStateConnecting		= 2  '// The object is connecting
	Const adStateExecuting		= 4  '// The object is executing a command
	Const adStateFetching		= 8  '// The rows of the object are being retrieved

'*********************************************************

' Am I running 64-bit version of WScript.exe/Cscript.exe? So, call script again in x86 script host and then exit.
 If InStr(LCase(WScript.FullName), LCase(oProcEnv("windir") & "\System32\")) And oProcEnv("PROCESSOR_ARCHITECTURE") = "AMD64" Then
	If Not WScript.Arguments.Count = 0 Then
        Dim sArg, Arg : sArg = ""
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
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "30", "REG_DWORD"	'Changes TCP timeout settings if needing to restart program w/in 5 minutes
	On Error Goto 0

	'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
	Set wmi = GetObject("winmgmts:root\cimv2") 
	Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
	For Each oProcess in cProcesses
		oProcess.Terminate()
	Next

 
	If Not WScript.Arguments.Count = 0 Then
		sArg = ""
		For Each Arg In Wscript.Arguments
			  sArg = sArg & " " & """" & Arg & """"
		Next
	End If

	 Dim machineString : machineString = sArg
	 If sArg <> "" Then
		Do While AscW(Right(machineString, 1)) = 34 or AscW(Right(machineString, 1)) = 32
			machineString = Left(machineString, Len(machineString) - 1)
		Loop
		Do While AscW(Left(machineString, 1)) = 34 or AscW(Left(machineString, 1)) = 32
			machineString = Right(machineString, Len(machineString) - 1)
		Loop
	 Else
		machineString = "Manual"
	 End If

	 If Left(machineString, 3) = "COM" Then
		Dim objComport : Set objComport = CreateObject( "AxSerial.ComPort" )
		objComport.Clear()
		objComport.LicenseKey = "FD2C1-DC93A-6BFBF"
		objComport.Device = machineString
		objComport.BaudRate  = 112500
		objComport.ComTimeout = 1000  ' Timeout after 1000msecs 
	 ElseIf Left(machineString, 4) = "SHIP" or Left(machineString, 2) = "QA" Then
		'// CREATE WINSOCK: 0 - QA Scabben
		Dim winsock : Set winsock = Wscript.CreateObject("OSWINSCK.Winsock", "winsock_")
		'// CREATE WINSOCK: 0 - QA Scanner
		If Err.Number <> 0 Then
			MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
			WScript.Quit
		End If
		Load_IP
	 End If
	
	Dim leftX : leftX = 0
	Dim topY : topY = 0
	If Left(MachineString, 2) = "QA" Then
		leftX = 0 '2200
		topY = 20
	End If

	'Calls function to create ie window
	Dim windowBox : set windowBox = HTABox("white", 830, 1000, leftX , topY) : with windowBox
		.document.title = "Operation 50"
		'Function to check for access connection and load info from database
		Dim AccessResult : AccessResult = Load_Access
		Call checkDatabase
		'Connects to the scanner
		Call connect2Scanner
				
		'.document.accessText.focus
		'.document.accessText.select
		do until closeWindow = true													'Run loop until conditions are met
			do until .done.value = "cancel" or .done.value = "access" or .done.value = "scanner" or .submitButton.value = "true" _
				  or .done.value = "allOps" or .done.value = "SQLSubmit" or .done.value = "addETag" or .done.value = "removeETag" _
				  or .done.value = "eTagChange" or .done.value = "cancelMRB" or .done.value = "okMRB" or .done.value = "Reset"
				wsh.sleep 50
				On Error Resume Next
				If .done.value = true Then
					wsh.quit
				End If
				On Error GoTo 0
				If Left(machineString, 3) = "COM" Then ReadResponse(objComport)
			loop
			if .done.value = "cancel" then											'If the x button is clicked
				closeWindow = true													'Variable to end loop
			ElseIf .done.value = "access" then
				.done.value = false
				windowBox.accessText.innerText = "Retrying connection."
				windowBox.accessButton.style.backgroundcolor = "orange"
				If FieldsCheckEmpty(windowBox.bladeID.innerText) Then
					fieldArray(0) = windowBox.bladeID.innerText
					fieldArray(1) = windowBox.POID.innerText
					fieldArray(2) = windowBox.shipDate.value
					fieldArray(3) = windowBox.PalletID.innerText
					fieldArray(4) = windowBox.BoxID.innerText
					fieldsBad = False
					For n = 0 to ubound(fieldArray)
						If FieldsCheckEmpty(fieldArray(n)) Then	fieldsBad = True
					Next
					If fieldsBad = False Then LoadSNtoAccess
				Else
					AccessResult = Load_Access
					checkDatabase
				End If
			ElseIf .done.value = "scanner" then
				.done.value = false
				connect2Scanner
			ElseIf .done.value = "SQLSubmit" then
				.done.value = false
				windowBox.SubmitSQLButton.disabled = false
				LoadSNtoAccess
			ElseIf .submitButton.value = "true" Then
				.submitButton.value = false
				Check_String(windowBox.submitText.value)
				.returnToHTA.click()
			ElseIf .done.value = "allOps" Then
				objShell.Run sOPsCmd
				WScript.Quit	
			ElseIf .done.value = "addETag" Then
				.done.value = false
				.addMRBButton.disabled = true
				SQL_ETag("Add")
				.addMRBButton.disabled = false
			ElseIf .done.value = "removeETag" Then
				.done.value = false
				.removeMRBButton.disabled = true
				SQL_ETag("Remove")
				.removeMRBButton.disabled = false
				AccessCheck
			ElseIf  .done.value = "eTagChange" Then
				.done.value = false
				SQL_ETag("Change")
			ElseIf  .done.value = "okMRB" Then
				.done.value = false
				.addMRBButton.number = false
				.addMRBButton.disabled = false
				saveMRBChange(.addMRBButton.number)
				AccessCheck
			ElseIf .done.value = "cancelMRB" Then
				.done.value = false
				.addMRBButton.disabled = false
				AccessCheck
			ElseIf .done.value = "Reset" Then
				.done.value = false
				CleanUpScreen
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
			HTABox.document.title = "HTABox" 
			HTABox.document.write LoadHTML(sBgColor)
			Exit Function 
		End If 
	Next 
	MsgBox "HTA window not found." 
	wsh.quit
	End Function

Function connect2Scanner()
	Dim secs : secs = 0
	If machineString <> "Manual" and machineString <> "" Then
		windowBox.scannerText.innerText = "Connect to " & machineString
		windowBox.scannerButton.style.backgroundcolor = "orange"
		windowBox.scannerButton.disabled = true
		windowBox.errorString.innerText = ""
	End If
	'Stores variable if connected to part marker
	If machineString = "Manual" Then
		windowBox.scannerText.innerText = "Manual scanner mode"
		windowBox.scannerButton.style.backgroundcolor = "limegreen"
		windowBox.scannerButton.disabled = true
		windowBox.errorString.innerText = ""
		windowBox.manualSerialNumber.style.backgroundColor = "DimGrey"
		windowBox.SerialNumberText.style.visibility = "hidden"
		windowBox.inputFormDiv.style.visibility = "visible"
		windowBox.inputForm.disabled = false
		windowBox.inputForm.stringInput.disabled = false
		windowBox.inputForm.stringInput.focus
	ElseIf Left(machineString, 3) = "COM" Then
		objComport.Open
		If( objComport.LastError <> 0 ) Then
			windowBox.scannerText.innerText = "Error: " & machineString
			windowBox.errorString.innerText = objComport.LastError & " (" & objComport.GetErrorDescription( objComport.LastError ) & ")"
			windowBox.scannerButton.style.backgroundcolor = "red"
			windowBox.scannerButton.disabled = false
		Else
			windowBox.scannerText.innerText = "Connected to " & machineString
			windowBox.scannerButton.style.backgroundcolor = "limegreen"
			windowBox.scannerButton.disabled = true
		End If
	Else
		
		' loads port settings into winsock
		If winsock.state <> sckClosed Then winsock.Disconnect
		If RemoteHost <> "" and RemotePort <> "" Then 
			winsock.RemoteHost = RemoteHost
			winsock.RemotePort = RemotePort
			'Connects to the scanner
			On Error Resume Next
			winsock.Connect    
			On Error GoTo 0
			'// MAIN DELAY - WAITS FOR CONNECTED STATE
			'// SOCKET ERROR RAISES WINSOCK ERROR SUB
			while winsock.State <> sckError And winsock.state <> sckConnected And winsock.state <> sckClosing And secs < 25
				WScript.Sleep 1000  '// 1 sec delay in loop
				secs = secs + 1     '// wait 25 secs max
			Wend
		End If
		If winsock.state = sckConnected Then 
			windowBox.scannerText.innerText = "Connected to " & machineString
			windowBox.scannerButton.style.backgroundcolor = "limegreen"
			windowBox.scannerButton.disabled = true
		Else
			windowBox.scannerText.innerText = "Error: " & machineString
			windowBox.scannerButton.style.backgroundcolor = "red"
			windowBox.scannerButton.disabled = false
		End If
	End If
	End Function

Function checkDatabase()
	If AccessResult = false Then
		windowBox.accessText.innerText = "SQL Database not loaded"
		windowBox.accessButton.style.backgroundcolor = "red"
	Else
		windowBox.accessText.innerText = "SQL Database connection successful"
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
	End Function

Function TrimString(ByVal VarIn)
	VarIn = Trim(VarIn)   
	If Len(VarIn) > 0 Then
		Do While AscW(Right(VarIn, 1)) = 10 or AscW(Right(VarIn, 1)) = 13
			VarIn = Left(VarIn, Len(VarIn) - 1)
		Loop
		Do While AscW(Left(VarIn, 1)) = 10 or AscW(Left(VarIn, 1)) = 13
			VarIn = Right(VarIn, Len(VarIn) - 1)
		Loop
	End If
	TrimString = Trim(VarIn)
	End Function

Function Check_String(stringFromScanner)
	Dim inputString, n
	
	windowBox.submitText.value = ""
	inputString = TrimString(stringFromScanner)
	windowBox.errorDiv.style.background = ""
	windowBox.errorString.innerText = ""
	windowBox.SubmitSQLButton.disabled = true
	If inputString = tabletPassword or inputString = computerPassword Then
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
	ElseIf Left(inputString, 5) = "QA_" Then
		machineString = inputString
		sArg = """" & inputString & """"
		RemoteHost = ""
		RemotePort = ""
		Load_IP
		connect2Scanner
	ElseIF Len(inputString) = 10 and Mid(inputString, 9, 1) = "-" and (Left(inputString, 1) = "D" or Left(inputString, 1) = "H") Then
		CleanUpScreen
		windowBox.bladeID.innerText = inputString
		AccessCheck
	End If	
	
	fieldArray(0) = windowBox.bladeID.innerText
	fieldArray(1) = windowBox.inspectorID.innerText
	fieldArray(2) = windowBox.CMMID.value
	'H0000000-0
	For n = 0 to ubound(fieldArray)
		If FieldsCheckEmpty(fieldArray(n)) Then	Exit Function
	Next
	windowBox.SubmitSQLButton.disabled = false
	End Function

Function FieldsCheckEmpty(VarIN)
	FieldsCheckEmpty = False
	If VarIN = false Then
		FieldsCheckEmpty = True
	ElseIf VarIN = "false" Then
		FieldsCheckEmpty = True
	ElseIf VarIN = "" Then
		FieldsCheckEmpty = True
	ElseIf AscW(Left(VarIN,1)) = 32 Then
		FieldsCheckEmpty = True
	ElseIf AscW(Left(VarIN,1)) = 160 Then
		FieldsCheckEmpty = True
	ElseIf VarIN = "NO OPEN PO'S FOUND" Then
		FieldsCheckEmpty = True
	End If
	End Function

Function LoadSNtoAccess()
	'Dim CurrentTime, Operator, strQueryPre, sqlString, rs, Duplicate, SCFound, POID, ShipDate, PalletID, BoxID, strQuery, accept
	
	Dim objCmd : set objCmd = GetNewConnection
	Dim bladeID : bladeID = windowBox.bladeID.innerText
	Dim accept: If windowBox.CMMID.value = 1 Then
		accept = "'N'"
	ElseIf windowBox.CMMID.value = 2 Then
		accept = "'Y'"
	Else
		accept = "null"
	End If
	Dim comment : comment = windowBox.commentNote.value
	
	Dim strQuery : strQuery = "UPDATE [50_Final] SET "
	strQuery = strQuery & "[Accepted Y/N] = " & accept & ", "
	strQuery = strQuery & "[Comments] = '" & comment & "' "
	strQuery = strQuery & "WHERE [Blade S/N] = '" & bladeID & "';"
	
	On Error GoTo 0
	If objCmd is Nothing Then
		windowBox.errorString.innerText = "Error connecting to database, data not sent"
		windowBox.accessText.innerText = "Connection failed, click to retry."
		windowBox.accessButton.style.backgroundcolor = "red"
		windowBox.accessButton.disabled = false
		windowBox.errorDiv.style.background = "red"
		Exit Function
	ElseIf windowBox.accessButton.style.backgroundcolor <> "limegreen" Then
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
		windowBox.errorDiv.style.background = ""
	End If
	Dim rs : Set rs = objCmd.Execute(strQuery)	
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
	windowBox.errorDiv.style.background = "limegreen"
	windowBox.errorString.innerText = bladeID & " updated " & accept & "."
	CleanUpScreen
	End Function

Function CleanUpScreen()
	Dim oOption : Set oOption = windowBox.Document.CreateElement("OPTION")
	
	windowBox.bladeID.innerText = ""
	windowBox.PartNumber.innerHTML = ""
	windowBox.inspectorID.innerText = ""
	windowBox.prodID.innerHTML = ""
	windowBox.inspectDate.innerHTML = ""
	windowBox.CMMID.value = 0
	windowBox.commentNote.value = ""
	windowBox.MRBLoc.innerHTML = ""	

	windowBox.AEButton.style.backgroundcolor = ""
	windowBox.InitialButton.style.backgroundcolor = ""
	windowBox.PartMarkButton.style.backgroundcolor = ""
	windowBox.FixtureButton.style.backgroundcolor = ""
	windowBox.CMMButton.style.backgroundcolor = ""
	windowBox.HasTagButton.style.backgroundcolor = ""
	windowBox.MRBLocButton.style.backgroundColor = ""
		
	windowBox.SubmitSQLButton.disabled = true
	windowBox.ETagID.innerHTML = ""
	oOption.text = "TBD"
	oOption.Value = 0
	windowBox.ETagID.add(oOption)
	Set oOption = Nothing
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

Function AccessCheck()
	Dim objCmd : set objCmd = GetNewConnection
	Dim SlugSN : SlugSN = false
	Dim sqlString, rs, a
	Dim ETagIDS, ETag, oOption
	
	If objCmd is Nothing Then Exit Function
	sqlString = "SELECT COUNT([Blade S/N]) FROM [50_Final] WHERE [Blade S/N]='" & windowBox.bladeID.innerText & "';"
	Set rs = objCmd.Execute(sqlString)
	If rs(0).value = 0 Then
		windowBox.errorDiv.style.background = "red"
		windowBox.errorString.innerText = "Final Inspection is missing"
		Exit Function
	End If
	Set rs = Nothing
	sqlString = "SELECT TOP 1 [Blade S/N], [Blade Inspected Date], [Accepted Y/N], [Final Insp Inspector Last Name], [Comments], [ProdID] FROM [50_Final] WHERE [Blade S/N]='" & windowBox.bladeID.innerText & "';"
	Set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		windowBox.bladeID.innerHTML = rs.Fields(0)
		windowBox.inspectDate.innerHTML = rs.Fields(1)
		If UCase(rs.Fields(2)) = "Y" Then
			windowBox.CMMID.value = 2
		ElseIf UCase(rs.Fields(2)) = "N" Then
			windowBox.CMMID.value = 1
		Else
			windowBox.CMMID.value = 0
		End IF
		If rs.Fields(3) <> "" Then
			windowBox.inspectorID.innerHTML = rs.Fields(3)
		Else
			windowBox.inspectorID.innerHTML = ""
		End If
		If rs.Fields(4) <> "" Then
			windowBox.commentNote.value = rs.Fields(4)
		End If
		If rs.Fields(5) <> "" Then
			windowBox.prodID.innerHTML = rs.Fields(5)
		End If
		rs.MoveNext
	Loop	
	Set rs = Nothing
	
	
	sqlString = "SELECT TOP 1 [FIC Blade Part Number], [Slug Serial Number]  FROM [00_AE_SN_Control] WHERE [Blade Serial Number]='" & windowBox.bladeID.innerText & "';"
	Set rs = objCmd.Execute(sqlString)
	windowBox.AEButton.style.backgroundcolor = "red"
	DO WHILE NOT rs.EOF
		windowBox.PartNumber.innerHTML = rs.Fields(0)
		SlugSN = rs.Fields(1)
		windowBox.AEButton.style.backgroundcolor = "limegreen"
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If SlugSN = false Then
		windowBox.AEButton.style.backgroundcolor = "red"
		windowBox.InitialButton.style.backgroundcolor = "red"
	Else
		sqlString = "SELECT COUNT(*) FROM [00_Initial] WHERE [Slug S/N]='" & SlugSN & "';"
		set rs = objCmd.Execute(sqlString)
		If rs(0).value = 0 Then	
			windowBox.InitialButton.style.backgroundcolor = "red"
		Else
			windowBox.InitialButton.style.backgroundcolor = "limegreen"
		End If
	End If
	Set rs = Nothing
	
	sqlString = "SELECT COUNT(*) FROM [10_Part_Marking] WHERE [Blade Serial Number]='" & windowBox.bladeID.innerText & "';"
	Set rs = objCmd.Execute(sqlString)
	If rs(0).value <> 0 Then
		windowBox.PartMarkButton.style.backgroundColor = "limegreen"
	Else
		windowBox.PartMarkButton.style.backgroundColor = "Red"
	End If
	Set rs = Nothing
	
	sqlString = "SELECT COUNT(*) FROM [20_LPT5] WHERE [Blade SN Dash 1]='" & windowBox.bladeID.innerText & "' or [Blade SN Dash 2]='" & windowBox.bladeID.innerText & "';"
	Set rs = objCmd.Execute(sqlString)
	windowBox.FixtureButton.style.backgroundColor = "red"
	If rs(0).value <> 0 Then
		windowBox.FixtureButton.style.backgroundColor = "limegreen"
	End If
	Set rs = Nothing
	
	sqlString = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Serial Number]='" & windowBox.bladeID.innerText & "';"
	Set rs = objCmd.Execute(sqlString)
	If rs(0).value <> 0 Then
		windowBox.CMMButton.style.backgroundColor = "limegreen"
	Else
		windowBox.CMMButton.style.backgroundColor = "Red"
	End If
	Set rs = Nothing
	
	Dim rsFound : rsFound = false
	sqlString = "SELECT TOP 1 [Serial Number], [Location] FROM [40_MRB] WHERE [Serial Number]='" & windowBox.bladeID.innerText & "' AND [Location] IS NOT NULL;"
	Set rs = objCmd.Execute(sqlString)
	windowBox.MRBLocButton.style.backgroundColor = "limegreen"
	DO WHILE NOT rs.EOF
		windowBox.MRBLocButton.style.backgroundColor = "red"
		windowBox.MRBLoc.innerHTML = rs.Fields(1)
		rs.MoveNext
	Loop	
	Set rs = Nothing
	
	
	

	
	rsFound = false
	sqlString = "SELECT TOP 1 [Serial Number], [Tag Numbers], [Dispositions], [Status], [Summary Disposition], [Summary Status] FROM [40_Rejections] WHERE [Serial Number]='" & windowBox.bladeID.innerText & "';"
	set rs = objCmd.Execute(sqlString)
	windowBox.HasTagButton.style.backgroundColor = "limegreen"
	DO WHILE NOT rs.EOF
		windowBox.HasTagButton.style.backgroundColor = "red"
		windowBox.ETagID.innerHTML = ""
		ETagIDS = Split(rs.Fields(1), ";")
		a = 0
		For Each ETag in ETagIDS
			If ETag <> "" Then
				Set oOption = windowBox.Document.CreateElement("OPTION")
				oOption.text = TrimString(ETag)
				oOption.Value = a
				windowBox.ETagID.add(oOption)
				Set oOption = Nothing
				a = a + 1
			End If
		Next
		If rs.Fields(5) = "Closed" and (rs.Fields(4) = "Use As Is" or rs.Fields(4) = "Void" or rs.Fields(4) = "Return to Sender") Then
			windowBox.HasTagButton.style.backgroundColor = "limegreen"
		End If
		rs.MoveNext
	Loop	
	Set rs = Nothing
	'SQL_ETag("Change")
	objCmd.Close
	Set objCmd = Nothing
	End Function

Function SQL_ETag(Method)
	Dim objCmd : set objCmd = GetNewConnection
	Dim SlugSN : SlugSN = false
	Dim sqlString, rs, a
	Dim ETagNumber, oOption, objOption, objOptions
	'msgbox Method
	If objCmd is Nothing Then Exit Function
	If Method = "Change" Then
		'Dim DispositionArray : DispositionArray = Array("In MRB", "Return to Customer", "Scrap", "Supplier Rework/Remake", "Use As Is", "Void")
		'Dim StatusArray : StatusArray = Array("Need to be created", "Created", "Closed")
		
		sqlString = "SELECT TOP 1 [Disposition], [Status], [Created By], [Open Date] FROM [40_E-Tags] WHERE [Tag Number]='" & windowBox.ETagID(int(windowBox.ETagID.value)).innerText & "';"
		set rs = objCmd.Execute(sqlString)
		DO WHILE NOT rs.EOF
			For a = 0 to UBound(DispositionArray)
				If InStr(1, rs.Fields(0), DispositionArray(a)) <> 0 Then Exit For
			Next
			If a > UBound(DispositionArray) Then a = 0
			windowBox.eTagDispositionSelect.value = a
			
			For a = 0 to UBound(StatusArray)
				If StatusArray(a) = rs.Fields(1) Then Exit For
			Next
			If a > UBound(StatusArray) Then a = 0
			windowBox.eTagStatusSelect.value = a
			windowBox.originatorID.innerHTML = rs.Fields(2)
			windowBox.ETagDate.innerHTML = rs.Fields(3)
			rs.MoveNext
		Loop
		Set rs = Nothing
		
		sqlString = "SELECT [Serial Number] FROM [40_Rejections] WHERE [Tag Numbers] LIKE '%" & windowBox.ETagID(int(windowBox.ETagID.value)).innerText & "%';"
		windowBox.SNIDs.innerHTML = ""
		set rs = objCmd.Execute(sqlString)
		DO WHILE NOT rs.EOF
			If windowBox.SNIDs.innerHTML = "" Then
				windowBox.SNIDs.innerHTML = rs.Fields(0)
			Else
				windowBox.SNIDs.innerHTML = windowBox.SNIDs.innerHTML & chr(10) & rs.Fields(0)
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
	ElseIf Method = "Add" Then
		ETagNumber = inputBox("Enter a Tag Number")
		If ETagNumber <> "" Then
			windowBox.addMRBButton.disabled = true
			windowBox.addMRBButton.number = ETagNumber
			If windowBox.ETagID(0).text = "TBD" Then windowBox.ETagID.innerHTML = ""
			Set objOptions = windowBox.document.getElementById("ETagID")
			a = 0
			For Each objOption in objOptions
				a = a + 1
			Next
			Set objOptions = Nothing
			Set oOption = windowBox.Document.CreateElement("OPTION")
			oOption.text = TrimString(ETagNumber)
			oOption.Value = a
			windowBox.ETagID.add(oOption)
			Set oOption = Nothing
			windowBox.ETagID.value = a
			
			sqlString = "SELECT COUNT(*) FROM [40_E-Tags] WHERE [Tag Number]='" & ETagNumber & "';"
			set rs = objCmd.Execute(sqlString)
			
			If rs(0).value = 0 Then
				windowBox.ETagDate.innerHTML = Now
				windowBox.eTagDispositionSelect.value = 0
				windowBox.eTagStatusSelect.value = 0
				windowBox.statusSummary.innerHTML = "New Tag"
				windowBox.originatorID.innerHTML = windowBox.InspectorID.innerHTML
				windowBox.SNIDs.innerHTML = windowBox.bladeID.innerHTML
			Else
				msgbox windowBox.ETagID.value
				windowBox.statusSummary.innerHTML = "New SN, Save"
				SQL_ETag("Change")
			End If
			Set rs = Nothing
			
		End If
	Else
		windowBox.addMRBButton.disabled = true
		msgbox "Remove!"
	End If
	objCmd.Close
	Set objCmd = Nothing
	End Function

Function saveMRBChange(TagNumber)
	Dim objCmd : set objCmd = GetNewConnection
	Dim sqlString, rs
	
	sqlString = "SELECT COUNT(*) FROM [40_Rejections] WHERE [Serial Number] = '" & windowBox.bladeID.innerHTML & "';" 
	set rs = objCmd.Execute(sqlString)
	If rs(0).value = 0 Then
		sqlString = "INSERT INTO [40_Rejections] ([Serial Number],  				  [Tag Numbers],   [Dispositions], 	[Status],  [Summary Disposition], [Summary Status]) " _
					& " VALUES ('" & windowBox.bladeID.innerHTML & "', '" & TagNumber & "', 'New Tag', 'New Tag', 'New Tag', 'New Tag');"
	Else
		Set rs = Nothing
		sqlString = "SELECT TOP 1 [Tag Numbers] FROM [40_Rejections] WHERE [Serial Number] = '" & windowBox.bladeID.innerHTML & "';" 
		Set rs = objCmd.Execute(sqlString)
		sqlString = ""
		DO WHILE NOT rs.EOF
			If InStr(1, rs.Fields(0), TagNumber) <> 0 Then
				Exit Do
			Else
				msgbox rs.Fields(0)
				sqlString = "UPDATE [40_Rejections] Set [Tag Numbers] = '" & rs.Fields(0) & chr(10) & TagNumber & ";' "_
											   & "WHERE [Serial Number] = '" & windowBox.bladeID.innerHTML & "';" 
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
		msgbox sqlString
		If sqlString <> "" Then
			Set rs = objCmd.Execute(sqlString)
			Set rs = Nothing
		End If
	End If
	
	
	msgbox sqlString

	
	
	
	sqlString = "SELECT COUNT(*) FROM [40_MRB] WHERE [Serial Number] = '" & windowBox.bladeID.innerHTML & "';" 
	set rs = objCmd.Execute(sqlString)
	If rs(0).value = 0 Then	
		sqlString = "INSERT INTO [40_MRB] ([Serial Number], [Location]) " _
					& " VALUES ('" & windowBox.bladeID.innerHTML & "', '" & windowBox.MRBLocSelect(Int(windowBox.MRBLocSelect.Value)).Text & "');" 
	Else
		sqlString = "UPDATE [40_MRB] Set [Location] = '" & windowBox.MRBLocSelect(Int(windowBox.MRBLocSelect.Value)).Text _
					& "' WHERE [Serial Number] = '" & windowBox.bladeID.innerHTML & "';" 
	End If
	Set rs = objCmd.Execute(sqlString)
	Set rs = Nothing

	objCmd.Close
	Set objCmd = Nothing
	End Function
	
Function Load_Access()
	Dim objCmd : set objCmd = GetNewConnection
	If objCmd is Nothing Then Load_Access = false : Exit Function
	objCmd.Close
	Set objCmd = Nothing
	Load_Access = true
	End Function

Function Load_IP()
	Dim sqlString, rs
	Dim objCmd : set objCmd = GetNewConnection
	On Error GoTo 0
	If objCmd is Nothing Then Exit Function
	sqlString = "Select [IPAddress], [Port] From [00_Machine_IP] WHERE [DeviceType] = 'CognexBTHandheld' AND [MachineName] = '" & machineString & "';"
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

Function Send_Email(Message)
 Exit Function
	Dim MyEmail : Set MyEmail=CreateObject("CDO.Message")
	Dim bodyPre : bodyPre = "<body><p><span style='font-size:12pt; color:red'>This is an automatically generated email, please reply to sender if you have any issues</span></p><br>" _
		& "<p><span>Please close the following work orders:</span></p>"
	
	Dim Signature : Signature = "<footer><div>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Chris Zarlengo</span><span style='color:#1F497D'></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>Manufacturing Engineer</span><span style='color:#1F497D'></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Flow International Corporation | <a href=""http://www.flowwaterjet.com/"">http://www.FlowWaterjet.com/</a></span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>23500 64th Ave. S. | Kent | Washington | 98032 | USA</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>253-246-3741 | <a href=""mailto:CZarlengo@flowcorp.com"">CZarlengo@flowcorp.com</a><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span style='font-size:8.0pt;font-family:""Franklin Gothic Medium"",sans-serif; color:gray'>" _
			& "This electronic message contains information from and is the property of Flow International Corporation (Flow). " _
			& "The contents of this electronic message may be privileged and confidential and are for the use of the intended addressee(s) only. " _
			& "If you are not an intended addressee, note that any disclosure, copying, distribution, or use of the contents of this message is prohibited. " _
			& "If you have received this message in error, please contact the sender or call Flow immediately at (800) 962-8576.</span>" _
		& "</div></footer>"
	
	MyEmail.Subject="Contract cutting work orders need closing"
	MyEmail.From="czarlengo@flowcorp.com"
	MyEmail.To="czarlengo@flowcorp.com"
	MyEmail.BCC="czarlengo@flowcorp.com"
	MyEmail.HTMLBody = bodyPre & Message & Signature

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


 End Function

Sub winsock_OnDataArrival(bytesTotal)
    winsock.GetData strData, vbString
    WScript.Sleep 1000
	Check_String strData
	End Sub

'// WINSOCK ERROR
Sub winsock_OnError(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
	windowBox.scannerText.innerText = "Error: " & machineString
	windowBox.scannerButton.style.backgroundcolor = "red"
	windowBox.scannerButton.disabled = false
    windowBox.errorString.innerText = "Scanner Error: " & Number & vbCrLf & Description
	End Sub

	
Sub ReadResponse(ByVal objComport)
  Dim str : str = "notempty"
  objComport.Sleep(200)
  While (str <> "")
    str = objComport.ReadString()
    If (str <> "") Then
		Call Check_String(str)
    End If

  WEnd
 End Sub
 
'// EXIT SCRIPT
Sub ServerClose()
	If debugMode = False Then On Error Resume Next

	WScript.Sleep 1000  '// REQUIRED OR ERRORS
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"

	objComport.Close()
	objComport.Clear()
	Set objComport = Nothing
	
	If winsock.state <> sckClosed Then winsock.Disconnect
    winsock.CloseWinsock
    Set winsock = Nothing
	
	windowBox.close
	
	On Error GoTo 0
    Wscript.Quit
	End Sub

'Function to create all of the JS and HTML code for the window
Function LoadHTML(sBgColor)
	Dim a, Status, Disposition
	Dim bodyTop : bodyTop = 250
	Dim footerTop : footerTop = 25
	
	'HTA String
	LoadHTML = "<HTA:Application contextMenu=no border=thin minimizebutton=no maximizebutton=no sysmenu=no />"
	
	'CSS String
	LoadHTML = LoadHTML _	
		& "<head><style>" _
		& "body {" _
			& "background-color: " & sBgColor & ";" _
			& "font:normal 28px Tahoma;" _
			& "border-Style:outset" _
			& "border-Width:3px" _
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
			& "color: white;" _
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
		& "#SubmitSQLButton {" _
			& "font:normal 30px Tahoma;" _
			& "}" _
		& ".prodCount, #MRBLoc {" _
			& "font:normal 20px Tahoma;" _
			& "}" _
		& "#ETagButton {" _
			& "font:normal 30px Tahoma;" _
			& "}" _
		& "#commentModal {" _
			& "font:normal 30px Tahoma;" _
			& "background-color = 'grey';" _
			& "visibility: hidden;" _
			& "}" _
		& "#MRBModal {" _
			& "font:normal 20px Tahoma;" _
			& "background-color = 'grey';"
	If adminMode <> true Then
		LoadHTML = LoadHTML _
			& "visibility: hidden;"
	End If
	LoadHTML = LoadHTML _
			& "}" _
		& "#table_wrapper {" _
			& "width:100%;" _
			& "height:100%;" _
			& "overflow:auto; " _
			& "}" _
		& "#table_wrapper table {" _
			& "margin-right: 20px;" _
			& "border-collapse: collapse;" _
			& "}" _
		& "tr, th, td {" _
			& "border-bottom: 1px solid black;" _
			& "}" _
		& "#SNIDs, #commentNote {" _
			& "overflow-y: scroll;" _
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
		& "function logoutAdmin() {" _
			& "document.getElementById('errorString').innerText = 'LOGGED OUT';" _
			& "document.getElementById('duplicateButton').disabled = true;" _
			& "document.getElementById('adminText').style.visibility = 'hidden';" _
			& "document.getElementById('adminButton').style.visibility = 'hidden';" _
			& "document.getElementById('adminString').style.visibility = 'hidden';" _
			& "document.getElementById('logoutButton').style.visibility = 'hidden';" _
		& "}" _
		& "function tableClick(clickObj) {" _
			& "document.getElementById('POID').innerText = clickObj;" _
			& "document.getElementById('commentModal').style.visibility = 'hidden';" _
			& "document.getElementById('commentModal').style.left = '1000px';" _
			& "document.getElementById('done').value = 'reloadPO';" _
			& "event.cancelBubble = true;" _
			& "event.returnValue = false;" _
			& "return false;" _
		& "}" _
		& "function inputComplete() {" _
			& "document.getElementById('submitText').value = document.getElementById('stringInput').value;" _
			& "document.getElementById('stringInput').value = '';" _
			& "document.getElementById('submitButton').value = true;" _
			& "event.cancelBubble = true;" _
			& "event.returnValue = false;" _
			& "return false;" _
		& "}" _
		& "function HTAReturn() {" _
			& "document.getElementById('stringInput').value = '';" _
			& "" _
		& "}" _
		& "function resetFunction() {" _
			& "document.getElementById('done').value = 'Reset';" _
		& "}" _
		& "function commentFunction() {" _
			& "document.getElementById('commentModal').style.visibility = 'visible';" _
			& "document.getElementById('commentModal').style.left = '1px';" _
		& "};" _
		& "function cancelComment() {" _
			& "document.getElementById('commentModal').style.visibility = 'hidden';" _
			& "document.getElementById('commentModal').style.left = '1000px';" _
		& "};" _
		& "function okMRB() {" _
			& "document.getElementById('MRBModal').style.visibility = 'hidden';" _
			& "document.getElementById('MRBModal').style.left = '1000px';" _
			& "document.getElementById('MRBLocButton').style.backgroundColor = 'Blue';" _
			& "document.getElementById('done').value = 'okMRB';" _
		& "};" _
		& "function cancelMRB() {" _
			& "document.getElementById('MRBModal').style.visibility = 'hidden';" _
			& "document.getElementById('MRBModal').style.left = '1000px';" _
			& "document.getElementById('done').value = 'cancelMRB';" _
		& "};" _
		& "function okComment() {" _
			& "document.getElementById('commentModal').style.visibility = 'hidden';" _
			& "commentValue = document.getElementById('commentText').firstChild.value.replace('\'', '\'\'');" _
			& "document.getElementById('commentTextSave').value = commentValue;" _
			& "if (commentValue == '') {" _ 
				& "document.getElementById('commentButton').style.backgroundColor  = '';" _ 
			& "} else {" _
				& "document.getElementById('commentButton').style.backgroundColor  = 'limegreen';" _ 
			& "}" _
		& "};" _
		& "function MRBComplete(e) {" _
			& "document.getElementById('ETagButton').style.visibility = 'visible';" _
			& "document.getElementById('ETagButton').disabled = false;" _
			& "event.cancelBubble = true;" _
			& "event.returnValue = false;" _
			& "return false;" _
		& "};" _
		& "</script></head>"

	'Body Start String							
	LoadHTML = LoadHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
	
	'Access Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=accessButton style='height: 30px; width: 30px; text-align: center;background-color:orange;' disabled onclick='done.value=""access""'></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 60px; height: 30px; width: 480px; text-align: left;' id='accessText'>Waiting for SQL database connection</div>"
		
	'Scanner Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 25px;height: 30px; width: 30px;'>" _
		& "<button id=scannerButton style='height: 30px; width: 30px;background-color:orange;' disabled onclick='done.value=""scanner""'></button></div>" _
		& "<div id=scannerText unselectable='on' class='unselectable' style='top: 95px; left: 60px;height: 30px; width: 480px;'>Waiting for scanner connection</div>" 
		
	'Reset Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id=resetButton style='height: 30px; width: 30px;' onclick='resetFunction()'></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 60px;height: 30px; width: 480px;'>Click to reset fields</div>" 
			
	'Input String
	LoadHTML = LoadHTML _
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id='manualSerialNumber' style='height: 30px; width: 30px;' onclick='manualButton()'></button></div>" _
		& "<div id='SerialNumberText' unselectable='on' class='unselectable' style='top: 165px; left: 60px;height: 30px; width: 480px;'>Click to enter data manually</div>" _
		& "<div id='inputFormDiv' style='top: 165px; left: 60px; height: 30px; width: 480px;visibility:hidden;'>" _
			& "<form id=inputForm onsubmit='inputComplete();' disabled>" _
				& "<input id=stringInput style='top: 0px; left: 0px; height: 30px; width: 480px;' value='' disabled /></form></div>"
	
	'Serial Number String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop & "px; left: 25px;height: 30px; width: 275px;text-align: right;'>Blade Serial Number:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop & "px; left: 310px;height: 30px; width: 340px;' id=bladeID></div>" 
							 
	'Part Number String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop +  30 & "px; left: 25px;height: 30px; width: 275px; text-align: right;'>Part Number:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop +  30 & "px; left: 310px;height: 30px; width: 340px;' id=PartNumber></div>" 
				
	'Inspector Name String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop +  60 & "px; left: 25px; height: 30px; width: 275px; text-align: right;'>Inspector Name:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop +  60 & "px; left: 310px; height: 30px; width: 340px;' id=inspectorID></div>"
	
	'Work Order String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop +  90 & "px; left: 25px; height: 30px; width: 275px; text-align: right;'>Work Order:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop +  90 & "px; left: 310px; height: 30px; width: 340px;' id=prodID></div>"
		
	'Inspection Date String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop + 120 & "px; left: 25px; height: 30px; width: 275px; text-align: right;'>Inspection Date:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop + 120 & "px; left: 310px; height: 30px; width: 340px;' id=inspectDate></div>"
			
	'CMM Result String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop + 180 & "px; left: 25px; height: 30px; width: 275px; text-align: right;'>Accepted:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop + 158 & "px; left: 310px; height: 130px; width: 340px;'>" _
		& "<select size='1' id=CMMID style='height: 100px; width: 300px;font: 60px;'>" _
			& "<option value='0' disabled selected>Choose:</option>" _
			& "<option value='1'>No</option>" _
			& "<option value='2'>Yes</option>" _
			& "</select></div>"
					 
	'Comment String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop + 280 & "px; left: 25px;height: 30px; width: 275px; text-align: right;'>Comments:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop + 280 & "px; left: 310px;height: 100px; width: 330px;'>" _
			& "<input id=commentNote type='text' style='top: 1px; left: 1px; height: 98px; width: 328px;'></div>" 
		
	'Submit Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & bodyTop + 205 & "px; left: 665px;height: 60px; width: 255px;'>" _
			& "<button style='height: 60px; width: 255px;' disabled id=SubmitSQLButton value=false onclick='done.value=""SQLSubmit""'>Submit</button></div>" _
		
	'AE Paperwork Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop & "px; left: 665px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=AEButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop & "px; left: 700px;height: 30px; width: 220px;'>AE Paperwork</div>" 
		
	'Initial Inspection Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 40 & "px; left: 665px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=InitialButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 40 & "px; left: 700px;height: 30px; width: 220px;'>Initial Inspection</div>" 
				
	'Part Mark Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 80 & "px; left: 665px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=PartMarkButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 80 & "px; left: 700px;height: 30px; width: 220px;'>Part Marking</div>" 
		
	'Fixture Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 120 & "px; left: 665px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=FixtureButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 120 & "px; left: 700px;height: 30px; width: 220px;'>Fixture Scan</div>" 
		
	'CMM Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 160 & "px; left: 665px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=CMMButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 160 & "px; left: 700px;height: 30px; width: 220px;'>CMM File</div>" 
		
	'ETag Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 200 & "px; left: 665px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=HasTagButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 200 & "px; left: 700px;height: 30px; width: 220px;'>Open E-Tag</div>" 
		
	'E-Tag String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' id='MRBinputFormDiv' style='top: " & footerTop + 240 & "px; left: 705px; height: 30px; width: 175px;'>" _
			& "<select size='1' id=ETagID style='height: 30px; width: 175px;' onChange='done.value=""eTagChange""'>" _
			& "<option value='0'>TBD</option></select></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 240 & "px; left: 900px; height: 30px; width: 70px;'>" _
			& "<button style='height: 25px; width: 70px;' onclick='done.value=""addETag""' id=addMRBButton number=false>Add</button></div>"
			
	'MRB Location Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 280 & "px; left: 665px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=MRBLocButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 280 & "px; left: 700px;height: 30px; width: 100px;'>In MRB</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 283 & "px; left: 805px; height: 30px; width: 150px;' id=MRBLoc></div>"
		
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div id=errorDiv style='top: " & footerTop + 610 & "px; left: 0px; height: 165px; width: 600px;'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & footerTop + 620 & "px; left: 50px; height: 140px; width: 500px; text-align: center;' id=errorString></div>"
		
	'All Op String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 905px;height: 30px; width: 30px;'><button class='opButton' style='height: 30px; width: 30px;' onclick='done.value=""allOps""'>&#10010;</button></div>"
		
	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 945px;height: 30px; width: 30px;'><button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'>&#10006;</button></div>" _
		& "<div style='top: 0px; left: 1000px;'><button type=hidden id=returnToHTA 		style='visibility:hidden;' value=false onclick='HTAReturn()'><center></button></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=done 				style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=SubmitButton		style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=submitText 		style='visibility:hidden;' value=false><center></div>"  _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=shipDate 			style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=isDuplicate 		style='visibility:hidden;' value=false><center></div>"
		
	'End Body String
	LoadHTML = LoadHTML _
		& "</body>"

	'Footer String
	LoadHTML = LoadHTML _
		& "<footer><script language='javascript'>" _
			& "document.getElementById('stringInput').focus();" _
		& "</script></footer>"
 End Function
 