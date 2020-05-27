Option Explicit
 '********* VERSION HISTORY ************
 ' 1.0	8/12/2018	Initial Release for production
 ' 1.1	10/5/2018	Added COM scanner option
 '					Added auto change for boxes
 ' 1.2	10/8/2018	Added scanning of paperwork (scans Slug SN, so not 100% foolproof)
 '					Added auto-selection for ship date
 '					Added auto-selection for box
 ' 1.3	10/28/2018	Added a visual color change for a new box
 '					Added CMM file verification
 '					Added crate option
 ' 2.0	2/1/2018	Added Operator ID scan & time code to SQL
 '					Removed Job Traveler scanning
 '************** TO DO *****************
 ' Edit mode for existing data
 ' 	Edit scanned blades
 ' 	Change ship date
 ' Popup Calendar
 ' Edit E-Tags
 ' Edit In MRB
 ' Slug ignore Final
 ' Ability to change existing ship dates
 ' Ability to add PO's

 '****** CHANGE THESE SETTINGS *********
 Dim adminMode : adminMode = false
 Dim debugMode : debugMode = false
 Dim boxSize : boxSize = 12
 Dim shipDay : shipDay = vbMonday
 Dim shipOff : shipOff = vbFriday
 Dim tabletPassword : tabletPassword = "Fl0wSh0p17"
 Dim computerPassword : computerPassword = "Snowball18!"
 Dim documentTitle : documentTitle = "Operation 60"								'IE window title
 Dim logoutTime : logoutTime = 45 / 60 / 24
 
 '***************** Database Settings *******************
 Dim dataSource : dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"
 Dim initialCatalog : initialCatalog = "CMM_Repository"								'Initial database
 
 '*********************************************************************************************************************************************
 '**************** INITIAL PARAMETERS *******************
 Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
 Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")

 Dim allOPSsource : allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\All Operations.vbs"
 Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """" & sArg

 Dim closeWindow : closeWindow = false
 Dim waitLoop : waitLoop = true
 Dim errorWindow : errorWindow = false
 Dim POChange : POChange = false
 Dim isCrate : isCrate = false
 Dim notFoundCount : notFoundCount = 0

 Dim strData, AccessResult, fieldArray(4), fieldsBad
 Dim RemoteHost, RemotePort

 '**************** DATABASE CONSTANTS *******************

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

 Const adOpenDynamic		 = 2  '// Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
 Const adOpenForwardOnly	 = 0  '// Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
 Const adOpenKeyset			 = 1  '// Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
 Const adOpenStatic			 = 3  '// Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
 Const adOpenUnspecified	 = -1 '// Does not specify the type of cursor.

 Const adLockBatchOptimistic = 4  '// Indicates optimistic batch updates. Required for batch update mode.
 Const adLockOptimistic		 = 3  '// Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method.
 Const adLockPessimistic	 = 2  '// Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing.
 Const adLockReadOnly		 = 1  '// Indicates read-only records. You cannot alter the data.
 Const adLockUnspecified	 = -1 '// Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.

 Const adStateClosed		 = 0  '// The object is closed
 Const adStateOpen			 = 1  '// The object is open
 Const adStateConnecting	 = 2  '// The object is connecting
 Const adStateExecuting		 = 4  '// The object is executing a command
 Const adStateFetching		 = 8  '// The rows of the object are being retrieved

 '*********************************************************

 
' Need this to run OSWINSCK, 32-bit required
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
 Dim cProcesses, oProcess : Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
 For Each oProcess in cProcesses
	oProcess.Terminate()
 Next

 
Dim bypassMode : bypassMode = false
 If Not WScript.Arguments.Count = 0 Then
	sArg = ""
	For Each Arg In Wscript.Arguments
		If InStr(Arg, "BYPASS") > 0 Then
			bypassMode = true
		Else
			sArg = Arg
		End If
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
 ElseIf Left(machineString, 4) = "SHIP" or Left(machineString, 2) = "QA" or Left(machineString, 2) = "PM" Then
	Dim winsock0 : Set winsock0 = Wscript.CreateObject("OSWINSCK.Winsock", "winsock0_")
	'// CREATE WINSOCK: 0 - QA Scanner
	If Err.Number <> 0 Then
		MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
		WScript.Quit
	End If
	Load_IP
 End If


'Calls function to create ie window
Dim windowBox : set windowBox = HTABox("white", 780, 900, 30, 30) : with windowBox	
	'Function to check for access connection and load info from database
	AccessResult = Load_Access
	Call checkAccess
	Call connect2Scanner				'Connects to the scanner
	Call Load_POs
	
	If bypassMode = true Then windowBox.bypassText.style.visibility = "visible"
	'.document.accessText.focus
	'.document.accessText.select
	Dim logoutStart
	do until closeWindow = true													'Run loop until conditions are met
		logoutStart = now
		do until waitLoop = False
			On Error Resume Next
			wsh.sleep 50
			If .done.value = true Then
				waitLoop = False
				wsh.quit
			ElseIf .done.value = "cancel" or .done.value = "access" or .done.value = "scanner" or .done.value = "reloadPO" or .submitButton.value = "true" or .done.value = "allOps"  or .done.value = "update" Then
				waitLoop = False
			End If
			On Error GoTo 0
			If Left(machineString, 3) = "COM" Then ReadResponse(objComport)
			If logoutStart + logoutTime < now Then
				windowBox.OperID.innerText = ""
				windowbox.errorDiv.style.background = "red"
				windowBox.errorString.innerText = "Logged out"
				logoutStart = now
			End If
		loop
		if .done.value = "cancel" then											'If the x button is clicked
			closeWindow = true													'Variable to end loop
		ElseIf .done.value = "access" then
			.done.value = false
			waitLoop = true
			windowBox.accessText.innerText = "Retrying connection."
			windowBox.accessButton.style.backgroundcolor = "orange"
			AccessResult = Load_Access
			Call checkAccess
		ElseIf .done.value = "scanner" then
			.done.value = false
			waitLoop = true
			Call connect2Scanner
		ElseIf .submitButton.value = "true" Then
			.submitButton.value = false
			waitLoop = true
			Call Check_String(windowbox.submitText.value)
			.returnToHTA.click()
		ElseIf .done.value = "reloadPO" Then
			.done.value = false
			waitLoop = true
			Call CheckShipQTY
		ElseIf .done.value = "allOps" Then
			objShell.Run sOPsCmd
			WScript.Quit	
		ElseIf .done.value = "update" Then
			.done.value = false
			waitLoop = true
			Call UpdateBlade(windowbox.updateValue.value)
		End If 
	loop
	.close																		'Closes the window
 end with
 ServerClose()																	'Function to close open connections and return settings back to original	
 Wscript.Quit

Function HTABox(sBgColor, h, w, l, t) 
	Dim IE, nRnd
	randomize : nRnd = Int(1000000 * rnd)
	
	Dim sCmd : sCmd = "mshta.exe ""javascript:{new " _ 
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
			HTABox.document.title = documentTitle									'Changes the window's title
			Exit Function 
		End If 
	Next 
	MsgBox "HTA window not found." 
	wsh.quit
 End Function

Sub connect2Scanner()
	Dim secs : secs = 0
	If machineString <> "Manual" and machineString <> "" Then
		windowBox.scannerText.innerText = "Connect to " & machineString
		windowBox.scannerButton.style.backgroundcolor = "orange"
		windowBox.scannerButton.disabled = true
		windowBox.errorString.innerText = ""
	End If
	' loads port settings into winsock
	If left(machineString, 3) = "COM" Then
		
	Else
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
			windowBox.scannerText.innerText = "Connected to " & machineString
			windowBox.scannerButton.style.backgroundcolor = "limegreen"
			windowBox.scannerButton.disabled = true
		Else
			windowBox.scannerText.innerText = "Error: " & machineString
			windowBox.scannerButton.style.backgroundcolor = "red"
			windowBox.scannerButton.disabled = false
		End If
	End If
 End Sub

Sub checkAccess()
	If AccessResult = false Then
		windowBox.accessText.innerText = "Database not loaded"
		windowBox.accessButton.style.backgroundcolor = "red"
	Else
		windowBox.accessText.innerText = "Database connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
	End If
 End Sub

Function TrimString(ByVal VarIn)
	VarIn = Trim(VarIn)   
	If Len(VarIn) > 0 Then
		Do While AscW(Right(VarIn, 1)) = 10 or AscW(Right(VarIn, 1)) = 13
			VarIn = Left(VarIn, Len(VarIn) - 1)
		Loop
	End If
	TrimString = Trim(VarIn)
 End Function

Sub Check_String(stringFromScanner)	
	If machineString <> "Manual" and windowbox.manualSerialNumber.style.backgroundColor = "dimgrey" Then
		windowbox.manualSerialNumber.disabled = true
		windowbox.manualSerialNumber.disabled = false
		windowbox.manualSerialNumber.style.backgroundColor = ""
		windowbox.SerialNumberText.style.visibility = "visible"
		windowbox.inputFormDiv.style.visibility = "hidden"
		windowbox.inputForm.disabled = true
		windowBox.inputForm.stringInput.disabled = false
	End If
	
	Dim inputString : inputString = TrimString(stringFromScanner)
	windowbox.errorDiv.style.background = ""
	windowBox.errorString.innerText = ""
	windowbox.submitText.value = ""
	windowbox.AEButton.style.backgroundcolor = ""
	windowbox.InitialButton.style.backgroundcolor = ""
	windowbox.FinalButton.style.backgroundcolor = ""
	windowbox.CMMButton.style.backgroundcolor = ""
	windowbox.ETagButton.style.backgroundcolor = ""
	windowbox.MRBButton.style.backgroundcolor = ""
	If inputString = tabletPassword or inputString = computerPassword Then
	ElseIf inputString = "Reset" Then
		Call CleanUpScreen
		windowBox.errorString.innerText = "Fields Reset"
	ElseIf inputString = "AccessRetry" Then
		windowBox.done.value = "access"
	ElseIf inputString = "Cancel" Then
		windowBox.done.value = "cancel"
	ElseIf Left(inputString, 5) = "SHIP_" Then
		machineString = inputString
		sArg = """" & inputString & """"
		RemoteHost = ""
		RemotePort = ""
		Call Load_IP
		Call connect2Scanner
	ElseIf Left(inputString, 4) = "AEFL" Then
		windowBox.POID.innerText = inputString
		windowBox.POID.style.backgroundcolor = ""
		Call CheckShipQTY
	ElseIf IsDate(inputString) Then
		windowbox.shipText.innerText = WeekdayName(Weekday(DateValue(inputString)),False) & " " & DateValue(inputString)
		windowbox.shipDate.value = DateValue(inputString)
		Call CheckShipQTY
	ElseIf Left(inputString, 7) = "Pallet " Then
		windowbox.PalletID.innerText = Split(inputString)(1)
		Call CheckShipQTY
	ElseIf Left(inputString, 4) = "BOX " or Left(inputString, 4) = "Box "Then
		windowbox.BoxID.innerText = Split(inputString)(1)
		Call CheckShipQTY
	ElseIF Len(inputString) = 10 and Mid(inputString, 9, 1) = "-" and Left(inputString, 1) = "H" Then
		If windowBox.loadMode.value = true or windowBox.loadMode.value = "true" Then
			Call loadBladeData(inputString)
		ElseIf windowbox.POID.innerText = "" or windowbox.shipDate.value = "" or windowbox.PalletID.innerText = "" or windowbox.BoxID.innerText = ""  or windowbox.OperID.innerText = "" Then
			windowBox.errorString.innerText = "Please scan shipping fields first"
		Else
			If CheckIfDuplicate(inputString) Then Exit Sub
			fieldArray(0) = windowbox.OperID.innerText
			fieldArray(1) = windowbox.POID.innerText
			fieldArray(2) = windowbox.shipDate.value
			fieldArray(3) = windowbox.PalletID.innerText
			fieldArray(4) = windowbox.BoxID.innerText
			Dim n : For n = 0 to ubound(fieldArray)
				If FieldsCheckEmpty(fieldArray(n)) Then	Exit Sub
			Next
			Call LoadSNtoAccess(inputString)
		End If
	ElseIf inputString = "Crate" Then
		isCrate = true
		windowbox.BoxID.innerText = "Crate"
		Call CheckShipQTY
	ElseIf inputString = "Box" Then
		isCrate = false
		windowbox.BoxID.innerText = ""
		Call CheckShipQTY
	ElseIf IsNumeric(inputString) Then
		Dim sqlString : sqlString = "SELECT TOP 1 [LastName] FROM [00_Personnel] WHERE [UserID]=" & inputString & ";"
		Dim objCmd : Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkAccess : Exit Sub
		Dim rs : set rs = objCmd.Execute(sqlString)	
		DO WHILE NOT rs.EOF
			windowBox.OperID.innerText = rs.Fields(0)
			rs.MoveNext
		Loop	
	Else
		Dim inputArray : inputArray = Split(inputString)
		If UBound(inputArray) > 0 Then
			windowBox.OperID.innerText = inputArray(1)
		Else
			windowBox.OperID.innerText = inputArray(0)
		End If
	End If
 End Sub

Function FieldsCheckEmpty(VarIN)
	FieldsCheckEmpty = False
	If VarIN = false Then
		FieldsCheckEmpty = True
	ElseIf VarIN = "false" Then
		FieldsCheckEmpty = True
	ElseIf VarIN = "False" Then
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

Function CheckIfDuplicate(bladeID)
	On Error GoTo 0
	Dim objCmd : Set objCmd = GetNewConnection : If objCmd is Nothing Then
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
		windowbox.errorDiv.style.background = ""
	End If
	
	Dim sqlString : sqlString = "Select COUNT(*) From [60_Shipping] WHERE [Blade Serial Number] = '" & bladeID & "';"
	Dim rs : Set rs = objCmd.Execute(sqlString)	
	If rs(0).value <> 0 Then
		windowBox.errorString.innerText = "Serial number already scanned: " & bladeID
		windowbox.errorDiv.style.background = "red"
		Call CleanUpScreen
	End If
		
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
 End Function

Sub LoadSNtoAccess(bladeID)
	Dim SCFound
	Dim ErrorFound : ErrorFound = False
	Dim errorNote : errorNote = ""
	Dim POID : POID = windowbox.POID.innerText
	Dim OperID : OperID = windowbox.OperID.innerText
	Dim ShipDate : ShipDate = windowbox.shipDate.value
	Dim PalletID : PalletID = windowbox.PalletID.innerText
	Dim BoxID : BoxID = windowbox.BoxID.innerText
	If BoxID = "Crate" Then BoxID = 0
	
	On Error GoTo 0
	Dim objCmd : Set objCmd = GetNewConnection : If objCmd is Nothing Then
		windowBox.errorString.innerText = "Error connecting to database, data not sent"
		windowBox.accessText.innerText = "Connection failed, click to retry."
		windowBox.accessButton.style.backgroundcolor = "red"
		windowBox.accessButton.disabled = false
		windowbox.errorDiv.style.background = "red"
		Exit Sub
	ElseIf windowBox.accessButton.style.backgroundcolor <> "limegreen" Then
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
		windowbox.errorDiv.style.background = ""
	End If
	
	Dim sqlString : sqlString = "Select COUNT(*) From [60_Shipping] WHERE [Blade Serial Number] = '" & bladeID & "';"
	Dim rs : Set rs = objCmd.Execute(sqlString)	
	If rs(0).value <> 0 Then
		windowBox.errorString.innerText = "Serial number already scanned: " & bladeID
		windowbox.errorDiv.style.background = "red"
		objCmd.Close
		Set objCmd = Nothing
		Call CleanUpScreen
		Exit Sub
	End If
	
	sqlString = "SELECT TOP 1 [Slug Serial Number] FROM [00_AE_SN_Control] WHERE [Blade Serial Number]='" & bladeID & "';"
	Set rs = objCmd.Execute(sqlString)		
	DO WHILE NOT rs.EOF
		SCFound = rs.Fields(0)
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If SCFound = "" Then
		windowbox.AEButton.style.background = "red"
		windowbox.InitialButton.style.background = "red"
		If ErrorFound = False Then ErrorFound = True
	Else
		windowbox.AEButton.style.background = "limegreen"
		sqlString = "SELECT COUNT(*) FROM [00_Initial] WHERE [Slug S/N]='" & SCFound & "';"
		Set rs = objCmd.Execute(sqlString)
		If rs(0).value = 0 Then	
			windowbox.InitialButton.style.background = "red"
			If ErrorFound = False Then ErrorFound = True
		Else
			windowbox.InitialButton.style.background = "limegreen"
		End If
	End If
	
	sqlString = "SELECT [Accepted Y/N] FROM [50_Final] WHERE [Blade S/N]='" & bladeID & "';"
	Set rs = objCmd.Execute(sqlString)	
	SCFound = ""
	DO WHILE NOT rs.EOF
		If UCase(rs.Fields(0)) = "Y" Then
			SCFound = rs.Fields(0)
		End If
		rs.MoveNext
	Loop	
	If SCFound = "" Then
		windowbox.FinalButton.style.background = "red"
		If ErrorFound = False Then ErrorFound = True
	ElseIf SCFound = "N" Then
		windowbox.FinalButton.style.background = "orange"
		If ErrorFound = False Then ErrorFound = True
	Else
		windowbox.FinalButton.style.background = "limegreen"
	End If
	
	sqlString = "SELECT Count(*) FROM [40_CMM_LPT5] WHERE [Serial Number]='" & bladeID & "';"
	set rs = objCmd.Execute(sqlString)	
	If rs(0).value = 0 Then 
		windowbox.CMMButton.style.background = "red"
		If ErrorFound = False Then ErrorFound = True
	ElseIf rs(0).value > 1 Then
		windowbox.CMMButton.style.background = "orange"
		errorNote = errorNote & chr(10) & "Multiple CMM Files Found"
	Else
		windowbox.CMMButton.style.background = "limegreen"
	End If
	Set rs = Nothing
	
	sqlString = "SELECT COUNT(*) FROM [40_Rejections] WHERE [Serial Number]='" & bladeID & "';"
	Set rs = objCmd.Execute(sqlString)
	If rs(0).value <> 0 Then
		Set rs = Nothing
		sqlString = "SELECT [Summary Status], [Summary Disposition] FROM [40_Rejections] WHERE [Serial Number]='" & bladeID & "';"
		Set rs = objCmd.Execute(sqlString)
		SCFound = ""
		DO WHILE NOT rs.EOF
			If (rs.Fields(1) = "Use As Is" or rs.Fields(1) = "Void" or rs.Fields(1) = "Return to Customer") and (rs.Fields(0) = "Closed") Then
			Else
				SCFound = rs.Fields(0)
				errorNote = errorNote & chr(10) & "E-Tag not resolved"
				If ErrorFound = False Then ErrorFound = True
			End If
			rs.MoveNext
		Loop
		If SCFound <> "" Then
			windowbox.ETagButton.style.background = "red"
			If ErrorFound = False Then ErrorFound = True
		Else
			windowbox.ETagButton.style.background = "limegreen"
		End If
	Else
		Set rs = Nothing
		Dim failCount : failCount = 0
		sqlString = "SELECT [Failures] FROM [40_CMM_LPT5] WHERE [Serial Number]='" & bladeID & "';"
		Set rs = objCmd.Execute(sqlString)
		DO WHILE NOT rs.EOF
			failCount = failCount + rs.Fields(0)
			rs.MoveNext
		Loop	
		If failCount > 0 Then
			windowbox.ETagButton.style.background = "red"
			errorNote = errorNote & chr(10) & "E-Tag missing"
			If ErrorFound = False Then ErrorFound = True
		Else
			windowbox.ETagButton.style.background = "limegreen"
		End If
	End If
	Set rs = Nothing
	
	sqlString = "SELECT TOP 1 [Location] FROM [40_MRB] WHERE [Serial Number]='" & bladeID & "';"
	set rs = objCmd.Execute(sqlString)
	SCFound = ""
	DO WHILE NOT rs.EOF
		SCFound = rs.Fields(0)
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If SCFound <> "" Then
		windowbox.MRBButton.style.background = "red"
		errorNote = errorNote & chr(10) & "Part is still located in MRB"
		If ErrorFound = False Then ErrorFound = True
	Else
		windowbox.MRBButton.style.background = "limegreen"
	End If
On Error GoTo 0
	Dim strQueryPre : strQueryPre = "INSERT INTO [60_Shipping] ([Blade Serial Number], [Date Shipped], [Pallet], [Box ID], [AE PO Number], [OperID], [ScanDate]) "
	Dim strQuery
	If ErrorFound = True and bypassMode = True Then
		notFoundCount = notFoundCount + 1
		windowBox.notFoundCnt.InnerHTML = notFoundCount
		windowBox.notFoundText.style.visibility = "visible"
		strQuery = strQueryPre & "VALUES ('" & bladeID & "', '" & ShipDate & "', " & PalletID & ", " & BoxID & ", '" & POID & "', '" & OperID & "', '" & now & "'); "
		objCmd.Execute(strQuery)
		windowBox.errorString.innerText = "Errors found for " & bladeID & ", bypass mode enabled: data still sent" & chr(10) & "Please ensure part is corrected before shipping: " & errorNote
		windowbox.errorDiv.style.background = "limegreen"
	ElseIf ErrorFound = True Then
		notFoundCount = notFoundCount + 1
		windowBox.notFoundCnt.InnerHTML = notFoundCount
		windowBox.notFoundText.style.visibility = "visible"
		windowBox.errorString.innerText = "Errors found for " & bladeID & ", data not send" & chr(10) & "Please correct" & errorNote
		windowbox.errorDiv.style.background = "red"
	Else
		strQuery = strQueryPre & "VALUES ('" & bladeID & "', '" & ShipDate & "', " & PalletID & ", " & BoxID & ", '" & POID & "', '" & OperID & "', '" & now & "'); "
		objCmd.Execute(strQuery)
		windowBox.errorString.innerText = "S.N. scan successful: " & bladeID
		windowbox.errorDiv.style.background = "limegreen"
	End If
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
	CheckShipQTY
	If POChange = True Then
		windowbox.POID.innerText = ""
		windowBox.POID.style.backgroundcolor = ""
		Load_Access
		POChange = False
	End If
	CleanUpScreen

 End Sub

Sub UpdateBlade(bladeID)
	Dim SCFound, Comments
	Dim ErrorFound : ErrorFound = False
	Dim errorNote : errorNote = ""
	Dim POID : POID = windowbox.POID.innerText
	Dim OperID : OperID = windowbox.OperID.innerText
	Dim ShipDate : ShipDate = windowbox.shipDate.value
	Dim PalletID : PalletID = windowbox.PalletID.innerText
	Dim BoxID : BoxID = windowbox.BoxID.innerText
	If BoxID = "Crate" Then BoxID = 0
	
	On Error GoTo 0
	Dim objCmd : Set objCmd = GetNewConnection : If objCmd is Nothing Then
		windowBox.errorString.innerText = "Error connecting to database, data not sent"
		windowBox.accessText.innerText = "Connection failed, click to retry."
		windowBox.accessButton.style.backgroundcolor = "red"
		windowBox.accessButton.disabled = false
		windowbox.errorDiv.style.background = "red"
		Exit Sub
	ElseIf windowBox.accessButton.style.backgroundcolor <> "limegreen" Then
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
		windowbox.errorDiv.style.background = ""
	End If
	
	Dim sqlString : sqlString = "Select COUNT(*) From [60_Shipping] WHERE [Blade Serial Number] = '" & bladeID & "';"
	Dim rs : Set rs = objCmd.Execute(sqlString)	
	If rs(0).value = 0 Then
		windowBox.errorString.innerText = "Serial number not found: " & bladeID
		windowbox.errorDiv.style.background = "red"
		objCmd.Close
		Set objCmd = Nothing
		Call CleanUpScreen
		Exit Sub
	End If
	Set rs = Nothing
	
	sqlString = "Select [Comments] From [60_Shipping] WHERE [Blade Serial Number] = '" & bladeID & "';"
	Set rs = objCmd.Execute(sqlString)	
	DO WHILE NOT rs.EOF
		Comments = Comments & rs.Fields(0)
		windowbox.errorDiv.style.background = "red"
		rs.MoveNext
	Loop
	Set rs = Nothing
	Comments = Comments & " Updated " & now & ";"
	sqlString = "Update [60_Shipping] Set [Date Shipped] = '" & ShipDate & "', [Pallet] = '" & PalletID & "', [Box ID] = '" & BoxID & "', [AE PO Number] = '" & POID & "', [OperID] = '" & OperID & "', [Comments] = '" & Comments & "' WHERE [Blade Serial Number] = '" & bladeID & "';"
	objCmd.Execute(sqlString)
	windowBox.errorString.innerText = "Update successful: " & bladeID
	windowbox.errorDiv.style.background = "limegreen"
	objCmd.Close
	Set objCmd = Nothing
	
	windowbox.POID.innerText = ""
	windowbox.POCount.innerText = ""
	
	windowbox.shipDate.value = false
	windowbox.shipText.innerText = ""
	windowbox.ShipCount.innerText = 0
	
	windowbox.PalletID.innerText = ""
	windowbox.PalletCount.innerText = 0
	
	windowbox.BoxID.innerText = ""
	windowbox.BoxCount.innerText = 0
	
	windowbox.OperID.innerText = ""
	
	windowBox.loadMode.value = false
	windowBox.loadButton.style.backgroundcolor = ""
	
	windowBox.updateButton.style.visibility = "hidden"
	windowbox.updateValue.value = false
	
	windowbox.AEButton.style.backgroundcolor = ""
	windowbox.InitialButton.style.backgroundcolor = ""
	windowbox.FinalButton.style.backgroundcolor = ""
	windowbox.CMMButton.style.backgroundcolor = ""
	windowbox.ETagButton.style.backgroundcolor = ""
	windowbox.MRBButton.style.backgroundcolor = ""
	
	windowBox.done.value = "access"
	CleanUpScreen
 End Sub
 
Sub loadBladeData(bladeID)
	On Error GoTo 0
	Dim objCmd : Set objCmd = GetNewConnection : If objCmd is Nothing Then
		windowBox.errorString.innerText = "Error connecting to database, data not sent"
		windowBox.accessText.innerText = "Connection failed, click to retry."
		windowBox.accessButton.style.backgroundcolor = "red"
		windowBox.accessButton.disabled = false
		windowbox.errorDiv.style.background = "red"
		Exit Sub
	ElseIf windowBox.accessButton.style.backgroundcolor <> "limegreen" Then
		windowBox.accessText.innerText = "Access connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
		windowbox.errorDiv.style.background = ""
	End If
	
	Dim sqlString : sqlString = "Select [Date Shipped], [Pallet], [Box ID], [AE PO Number], [OperID] From [60_Shipping] WHERE [Blade Serial Number] = '" & bladeID & "';"
	Dim rs : Set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		windowbox.shipDate.value = rs.Fields(0)
		windowbox.shipText.innerText = WeekdayName(Weekday(DateValue(rs.Fields(0))), False) & " " & DateValue(rs.Fields(0))
		windowbox.PalletID.innerText = rs.Fields(1)
		windowbox.BoxID.innerText = rs.Fields(2)
		windowbox.POID.innerText = rs.Fields(3)
		windowbox.updateValue.value = bladeID
		If IsNull(rs.Fields(4)) Then windowBox.OperID.innerText = "" Else windowBox.OperID.innerText = rs.Fields(4)
		rs.MoveNext
	Loop
	
	Dim SCFound : SCFound = ""
	Dim errorNote : errorNote = ""
	sqlString = "SELECT TOP 1 [Slug Serial Number] FROM [00_AE_SN_Control] WHERE [Blade Serial Number]='" & bladeID & "';"
	Set rs = objCmd.Execute(sqlString)		
	DO WHILE NOT rs.EOF
		SCFound = rs.Fields(0)
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If SCFound = "" Then
		windowbox.AEButton.style.background = "red"
		windowbox.InitialButton.style.background = "red"
		errorNote = errorNote & chr(10) & "AE information missing"
	Else
		windowbox.AEButton.style.background = "limegreen"
		sqlString = "SELECT COUNT(*) FROM [00_Initial] WHERE [Slug S/N]='" & SCFound & "';"
		Set rs = objCmd.Execute(sqlString)
		If rs(0).value = 0 Then	
			windowbox.InitialButton.style.background = "red"
			errorNote = errorNote & chr(10) & "Initial inspection missing"
		Else
			windowbox.InitialButton.style.background = "limegreen"
		End If
	End If
	
	sqlString = "SELECT [Accepted Y/N] FROM [50_Final] WHERE [Blade S/N]='" & bladeID & "';"
	Set rs = objCmd.Execute(sqlString)	
	SCFound = ""
	DO WHILE NOT rs.EOF
		If UCase(rs.Fields(0)) = "Y" Then
			SCFound = rs.Fields(0)
		End If
		rs.MoveNext
	Loop	
	If SCFound = "" Then
		windowbox.FinalButton.style.background = "red"
		errorNote = errorNote & chr(10) & "Final inspection missing"
	ElseIf SCFound = "N" Then
		windowbox.FinalButton.style.background = "orange"
		errorNote = errorNote & chr(10) & "Final inspection result is 'N'"
	Else
		windowbox.FinalButton.style.background = "limegreen"
	End If
	
	sqlString = "SELECT Count(*) FROM [40_CMM_LPT5] WHERE [Serial Number]='" & bladeID & "';"
	set rs = objCmd.Execute(sqlString)	
	If rs(0).value = 0 Then 
		windowbox.CMMButton.style.background = "red"
		errorNote = errorNote & chr(10) & "CMM file missing"
	ElseIf rs(0).value > 1 Then
		windowbox.CMMButton.style.background = "orange"
		errorNote = errorNote & chr(10) & "Multiple CMM Files Found"
	Else
		windowbox.CMMButton.style.background = "limegreen"
	End If
	Set rs = Nothing
	
	sqlString = "SELECT COUNT(*) FROM [40_Rejections] WHERE [Serial Number]='" & bladeID & "';"
	Set rs = objCmd.Execute(sqlString)
	If rs(0).value <> 0 Then
		Set rs = Nothing
		sqlString = "SELECT [Summary Status], [Summary Disposition] FROM [40_Rejections] WHERE [Serial Number]='" & bladeID & "';"
		Set rs = objCmd.Execute(sqlString)
		SCFound = ""
		DO WHILE NOT rs.EOF
			If (rs.Fields(1) = "Use As Is" or rs.Fields(1) = "Void" or rs.Fields(1) = "Return to Customer") and (rs.Fields(0) = "Closed") Then
			Else
				SCFound = rs.Fields(0)
				errorNote = errorNote & chr(10) & "E-Tag not resolved"
			End If
			rs.MoveNext
		Loop
		If SCFound <> "" Then
			windowbox.ETagButton.style.background = "red"
		Else
			windowbox.ETagButton.style.background = "limegreen"
		End If
	Else
		Set rs = Nothing
		Dim failCount : failCount = 0
		sqlString = "SELECT [Failures] FROM [40_CMM_LPT5] WHERE [Serial Number]='" & bladeID & "';"
		Set rs = objCmd.Execute(sqlString)
		DO WHILE NOT rs.EOF
			failCount = failCount + rs.Fields(0)
			rs.MoveNext
		Loop	
		If failCount > 0 Then
			windowbox.ETagButton.style.background = "red"
			errorNote = errorNote & chr(10) & "E-Tag missing"
		Else
			windowbox.ETagButton.style.background = "limegreen"
		End If
	End If
	Set rs = Nothing
	
	sqlString = "SELECT TOP 1 [Location] FROM [40_MRB] WHERE [Serial Number]='" & bladeID & "';"
	set rs = objCmd.Execute(sqlString)
	SCFound = ""
	DO WHILE NOT rs.EOF
		SCFound = rs.Fields(0)
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If SCFound <> "" Then
		windowbox.MRBButton.style.background = "red"
		errorNote = errorNote & chr(10) & "Part is still located in MRB"
	Else
		windowbox.MRBButton.style.background = "limegreen"
	End If
	If errorNote <> "" Then windowBox.errorString.innerText = windowBox.errorString.innerText & chr(10) & errorNote
 End Sub
 
Sub CleanUpScreen()
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
	Dim PONum
	Dim PONumber : PONumber = "AEFL99999999"
	Dim POFound : POFound = False
	Dim objCmd : set objCmd = GetNewConnection : If objCmd is Nothing Then Load_Access = false : Exit Function
	Dim sqlString : sqlString = "Select [PONumber] From [60_PO] WHERE [POFilled] = 'False';"
	Dim rs : Set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		POFound = True
		PONum = rs.Fields(0)
		If PONum < PONumber Then PONumber = PONum
		rs.MoveNext
	Loop	
	If POFound = True Then
		windowBox.POID.innerText = PONumber
		windowBox.POID.style.backgroundcolor = ""
		CheckShipQTY
	Else
		windowBox.POID.innerText = "NO OPEN PO'S FOUND"
		windowBox.POID.style.backgroundcolor = "red"
	End If
	objCmd.Close
	Set objCmd = Nothing
	Load_Access = true
 End Function

Sub Load_POs()
	Dim PONum, POFilled, tableString
	Dim tableStringPre : tableStringPre = "<table id='POTable' visibility:hidden;><thead><tr><th><span class='text'>PO Number</span></th><th><span class='text'>PO Filled</span></th></tr></thead><tbody>"
	Dim tableStringSuf : tableStringSuf = "</tbody></table>"
	Dim objCmd : set objCmd = GetNewConnection : If objCmd is Nothing Then
		windowbox.table_wrapper.innerHTML = tableStringPre & "<tr><td>ERROR LOADING</td><td></td></tr>" & tableStringSuf
		Exit Sub
	End If
	Dim sqlString : sqlString = "Select [PONumber], [POFilled] From [60_PO];"
	Dim rs : Set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		PONum = rs.Fields(0)
		POFilled = rs.Fields(1)
		tableString = tableString & "<tr onclick='tableClick(this.firstChild.innerHTML);'><td>" & PONum & "</td><td style='text-align: center;'>" & POFilled & "</td></tr>"
		rs.MoveNext
	Loop
	objCmd.Close
	Set objCmd = Nothing
	windowbox.table_wrapper.innerHTML = tableStringPre & tableString & tableStringSuf
 End Sub

Sub CheckShipQTY()
	Dim sqlBoxString, boxNum, BoxID
	Dim rs, POTotal, loopRun
	Dim newShip : newShip = False
	On Error GoTo 0
	Dim objCmd : set objCmd = GetNewConnection : If objCmd is Nothing Then Exit Sub

	If FieldsCheckEmpty(windowbox.POID.innerText) = false Then
		Dim sqlPOString : sqlPOString = "Select COUNT(*) From [60_Shipping] WHERE [AE PO Number] = '" & windowbox.POID.innerText & "';"
		Set rs = objCmd.Execute(sqlPOString)
		Dim POCount : POCount = rs(0).value
		Dim sqlPOTotalString : sqlPOTotalString = "Select [POQuantity], [POFilled] From [60_PO] WHERE [PONumber] = '" & windowbox.POID.innerText & "';"
		Set rs = objCmd.Execute(sqlPOTotalString)
		If rs.EOF Then
			POTotal = "MISSING"
			windowbox.POCount.style.background = "red"
		Else
			POTotal = rs(0).value
			If rs(1).value = True Then
				windowbox.POCount.style.background = "limegreen"
			ElseIf POCount = POTotal Then
				POChange = True
				Dim sqlPOEndString : sqlPOEndString = "UPDATE [60_PO] SET [POFilled] = 1 WHERE [PONumber] = '" & windowbox.POID.innerText & "';"
				Set rs = objCmd.Execute(sqlPOEndString)
				windowbox.POCount.style.background = "limegreen"
			ElseIf POCount > POTotal Then
				windowbox.POCount.style.background = "red"
			Else
				windowbox.POCount.style.background = ""
			End If
		End If
		windowbox.POCount.innerText = POCount & " of " & POTotal
	End If
	If FieldsCheckEmpty(windowbox.shipDate.value) = false Then
		If InStr(windowbox.shipDate.value, "/") <> 3 Then windowbox.shipDate.value = 0 & windowbox.shipDate.value
		Dim sqlShipString : sqlShipString = "Select COUNT(*) From [60_Shipping] WHERE [Date Shipped] = '" & windowbox.shipDate.value & "';"
		Set rs = objCmd.Execute(sqlShipString)
		windowbox.ShipCount.innerText = rs(0).value
		If rs(0).value = 0 Then newShip = True
		If FieldsCheckEmpty(windowbox.PalletID.innerText) = false Then
			Dim sqlPalletString : sqlPalletString = "Select COUNT(*) From [60_Shipping] WHERE [Pallet] = " & windowbox.PalletID.innerText & " and [Date Shipped] = '" & windowbox.shipDate.value & "';"
			Set rs = objCmd.Execute(sqlPalletString)
			windowbox.PalletCount.innerText = rs(0).value
			If rs(0).value = 0 Then newShip = True
			If isCrate = true Then
				sqlBoxString = "Select COUNT(*) From [60_Shipping] WHERE [Box ID] = 0 and [Pallet] = " & windowbox.PalletID.innerText & " and [Date Shipped] = '" & windowbox.shipDate.value & "';"
				set rs = objCmd.Execute(sqlBoxString)
				windowbox.BoxCount.innerText = rs(0).value
			ElseIf FieldsCheckEmpty(windowbox.BoxID.innerText) = false Then
				Do
					loopRun = false
					sqlBoxString = "Select COUNT(*) From [60_Shipping] WHERE [Box ID] = " & windowbox.BoxID.innerText & " and [Pallet] = " & windowbox.PalletID.innerText & " and [Date Shipped] = '" & windowbox.shipDate.value & "';"
					set rs = objCmd.Execute(sqlBoxString)
					windowbox.BoxCount.innerText = rs(0).value & " of " & boxSize
					If rs(0).value > boxSize Then
						windowbox.BoxCount.style.background = "red"
					ElseIf rs(0).value = boxSize Then
						windowbox.BoxID.innerText = CInt(windowbox.BoxID.innerText) + 1
						loopRun = true
						windowBox.errorString.innerHTML = windowBox.errorString.innerText & "<br><br>New box loaded<br><bold><font size='+20'>BOX: " & windowbox.BoxID.innerText & "</font></bold>"
						windowbox.errorDiv.style.background = "cyan"
					Else
						windowbox.BoxCount.style.background = ""
					End If
				Loop While loopRun = true
			Else
				sqlBoxString = "Select TOP 1 [Box ID] From [60_Shipping] ORDER BY [Box ID] DESC;"
				set rs = objCmd.Execute(sqlBoxString)
				If NOT rs.EOF Then
					boxNum = rs(0).value
					sqlBoxString = "Select COUNT(*) From [60_Shipping] WHERE [Box ID] = " & boxNum & ";"
					set rs = objCmd.Execute(sqlBoxString)
					If NOT rs.EOF Then
						If rs(0).value >= boxSize or newShip = True Then
							windowbox.BoxID.innerText = CInt(boxNum + 1)
							windowbox.BoxCount.innerText = "0 of " & boxSize
							windowBox.errorString.innerHTML = windowBox.errorString.innerText & "<br><br>New box loaded<br><bold><font size='+20'>BOX: " & windowbox.BoxID.innerText & "</font></bold>"
							windowbox.errorDiv.style.background = "cyan"
						Else
							windowbox.BoxID.innerText = CInt(boxNum)
							windowbox.BoxCount.innerText = rs(0).value & " of " & boxSize
						End If
					End If
				End If
			End If
		End If
	Else
		Dim tempDate : tempDate = CDate("5/11/2019")
		' NEED TO FIX IT SO FRIDAY SWITCHES TO NEW DATE
		Dim shipDate : shipDate = date + 8 - Weekday(date, shipDay)
		If Weekday(date, shipDay) = 7 Then
			windowbox.shipDate.value = shipDate
			windowbox.shipText.innerText = WeekdayName(Weekday(DateValue(shipDate)), False) & " " & DateValue(shipDate)
		Else
			windowbox.shipDate.value = shipDate
			windowbox.shipText.innerText = WeekdayName(Weekday(DateValue(shipDate)), False) & " " & DateValue(shipDate)
		End If
	End If
	
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
 End Sub

Sub Load_IP()
	On Error GoTo 0
	Dim objCmd : set objCmd = GetNewConnection : If objCmd Is Nothing Then Exit Sub
	Dim sqlString : sqlString = "Select [IPAddress], [Port] From [00_Machine_IP] WHERE [DeviceType] = 'CognexBTHandheld' AND [MachineName] = '" & sArg & "';"

	If machineString <> "Manual" Then
		Dim rs : Set rs = objCmd.Execute(sqlString)		
		DO WHILE NOT rs.EOF
			RemoteHost = rs.Fields(0)
			RemotePort = rs.Fields(1)
			rs.MoveNext
		Loop
		Set rs = Nothing
	End If
	objCmd.Close
	Set objCmd = Nothing
 End Sub

'// WINSOCK DATA ARRIVES
Sub winsock0_OnDataArrival(bytesTotal)
    winsock0.GetData strData, vbString
    WScript.Sleep 1000
	Call Check_String(strData)
 End Sub

'// WINSOCK ERROR
Sub winsock0_OnError(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
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
	
	If winsock0.state <> sckClosed Then winsock0.Disconnect
    winsock0.CloseWinsock
    Set winsock0 = Nothing
	
	windowBox.close
	
	On Error GoTo 0
    Wscript.Quit
 End Sub


'Function to create all of the JS and HTML code for the window
Function LoadHTML(sBgColor)
	Dim midStart : midStart = 300
	Dim botStart : botStart = 500

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
			& "background-color = 'grey';" _
			& "visibility: hidden;" _
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
				& "document.getElementById('stringInput').focus();" _
				& "document.getElementById('manualSerialNumber').disabled = true;" _
				& "document.getElementById('manualSerialNumber').disabled = false;" _
			& "}" _
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
			& "document.getElementById('submitButton').value = true;" _
			& "event.cancelBubble = true;" _
			& "event.returnValue = false;" _
			& "return false;" _
		& "}" _
		& "function HTAReturn() {" _
			& "document.getElementById('stringInput').value = '';" _
			& "" _
		& "}" _
		& "function LoadFunction() {" _
			& "if (document.getElementById('loadButton').style.backgroundColor == '') {" _
				& "resetFunction();" _
				& "document.getElementById('POID').innerText = '';" _
				& "document.getElementById('loadButton').style.backgroundColor = 'limegreen';" _
				& "document.getElementById('loadMode').value = true;" _
				& "document.getElementById('updateButton').style.visibility = 'visible';" _
				& "document.getElementById('errorString').innerText = 'Edit mode: Please scan a blade';" _
			& "} else {" _
				& "document.getElementById('loadButton').style.backgroundColor = '';" _
				& "document.getElementById('loadMode').value = false;" _
				& "document.getElementById('updateButton').style.visibility = 'hidden';" _
				& "resetFunction();" _
				& "document.getElementById('POID').innerText = '';" _
				& "document.getElementById('done').value = 'access';" _
			& "}" _
			& "" _
		& "}" _
		& "function resetFunction() {" _
			& "document.getElementById('accessButton').disabled = true;" _
			& "document.getElementById('shipDate').value = false;" _
			& "document.getElementById('shipText').innerText = '';" _
			& "document.getElementById('shipCount').innerText = '0';" _
			& "document.getElementById('palletID').innerText = '';" _
			& "document.getElementById('palletCount').innerText = '0';" _
			& "document.getElementById('boxID').innerText = '';" _
			& "document.getElementById('boxCount').innerText = '0';" _
			& "document.getElementById('OperID').innerText = '';" _
			& "document.getElementById('errorDiv').style.background = '';" _
			& "document.getElementById('errorString').innerText = 'Fields Reset';" _
			& "document.getElementById('AEButton').style.backgroundColor = '';" _
			& "document.getElementById('InitialButton').style.backgroundColor = '';" _
			& "document.getElementById('FinalButton').style.backgroundColor = '';" _
			& "document.getElementById('CMMButton').style.backgroundColor = '';" _
			& "document.getElementById('ETagButton').style.backgroundColor = '';" _
			& "document.getElementById('resetButton').disabled = true;" _
			& "document.getElementById('resetButton').disabled = false;" _
		& "}" _
		& "function commentFunction() {" _
			& "document.getElementById('commentModal').style.visibility = 'visible';" _
			& "document.getElementById('commentModal').style.left = '1px';" _
			& "document.getElementById('POTable').style.visibility = 'visible';" _
		& "};" _
		& "function cancelComment() {" _
			& "document.getElementById('commentModal').style.visibility = 'hidden';" _
			& "document.getElementById('commentModal').style.left = '1000px';" _
			& "document.getElementById('POTable').style.visibility = 'hidden';" _
		& "};" _
		& "</script></head>"

	'Body Start String							
	LoadHTML = LoadHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
	
	'Access Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=accessButton style='height: 30px; width: 30px; text-align: center;background-color:orange;' disabled onclick='done.value=""access""'></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 60px; height: 30px; width: 480px; text-align: left;' id='accessText'>Waiting for database connection</div>"
		
	'Scanner Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 60px; left: 25px;height: 30px; width: 30px;'>" _
		& "<button id=scannerButton style='height: 30px; width: 30px;background-color:orange;' disabled onclick='done.value=""scanner""'></button></div>" _
		& "<div id=scannerText unselectable='on' class='unselectable' style='top: 60px; left: 60px;height: 30px; width: 480px;'>Waiting for scanner connection</div>" 
		
	'BYPASS String
	LoadHTML = LoadHTML _	
		& "<div id=bypassText unselectable='on' class='unselectable' style='top: 60px; left: 575px;height: 30px; width: 300px; background-color:red; text-align: center; visibility:hidden;'>BYPASS mode enabled</div>" 
		
	'Reset Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id=resetButton style='height: 30px; width: 30px;' onclick='resetFunction()'></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 60px;height: 30px; width: 480px;'>Click to reset fields</div>" 
		
	'PO Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id='commentButton' style='height: 30px; width: 30px;' onclick='commentFunction()'></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 60px;height: 30px; width: 480px;'>Click to change PO"  _
		& "<input type=hidden id='commentTextSave' value=''>" _
		& "<input type=hidden id='keepCommentSave' value='False'></div>"
				
	'Edit Blade String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id=loadButton style='height: 30px; width: 30px;' onclick='LoadFunction()'></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 60px;height: 30px; width: 480px;'>Click to load blade data</div>" 
		
	'Change Ship Date String
	'LoadHTML = LoadHTML _	
	'	& "<div unselectable='on' class='unselectable' style='top: 200px; left: 25px;height: 30px; width: 30px;'>" _
	'		& "<button id=dateButton style='height: 30px; width: 30px;' onclick='resetFunction()'></button></div>" _
	'	& "<div unselectable='on' class='unselectable' style='top: 200px; left: 60px;height: 30px; width: 480px;'>Click to change ship date</div>" 
		
	'Input String
	LoadHTML = LoadHTML _
		& "<div unselectable='on' class='unselectable' style='top: 200px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id='manualSerialNumber' style='height: 30px; width: 30px;' onclick='manualButton()'></button></div>" _
		& "<div id='SerialNumberText' unselectable='on' class='unselectable' style='top: 200px; left: 60px;height: 30px; width: 480px;'>Click to enter data manually</div>" _
		& "<div id='inputFormDiv' style='top: 200px; left: 60px; height: 30px; width: 480px;visibility:hidden;'>" _
			& "<form id=inputForm onsubmit='inputComplete();' disabled>" _
				& "<input id=stringInput style='top: 0px; left: 0px; height: 30px; width: 480px;' value='' disabled /></form></div>"
	
	'PO Number String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +   0 & "px; left: 25px;height: 30px; width: 175px;text-align: right;'>PO Number:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +   0 & "px; left: 200px;height: 30px; width: 340px;' id=POID></div>" 
		
	'PO Count String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +   0 & "px; left: 540px;height: 30px; width: 100px;text-align: right;'>Count:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +   0 & "px; left: 640px;height: 30px; width: 225px;' id=POCount>0</div>" 
		
	'Ship Text String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  30 & "px; left: 25px; height: 30px; width: 175px; text-align: right;'>Ship Date:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  30 & "px; left: 200px; height: 30px; width: 340px;' id=shipText></div>"
	
	'Ship Count String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  30 & "px; left: 540px;height: 30px; width: 100px;text-align: right;'>Count:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  30 & "px; left: 640px;height: 30px; width: 225px;' id=shipCount>0</div>" 
		
	'Pallet Number String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  60 & "px; left: 25px; height: 30px; width: 175px; text-align: right;'>Pallet #:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  60 & "px; left: 200px; height: 30px; width: 340px;' id=palletID></div>" _
		
	'Pallet Count String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  60 & "px; left: 540px;height: 30px; width: 100px;text-align: right;'>Count:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  60 & "px; left: 640px;height: 30px; width: 225px;' id=palletCount>0</div>" 
		
	'Box ID String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  90 & "px; left: 25px; height: 30px; width: 175px; text-align: right;'>Box ID:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  90 & "px; left: 200px; height: 30px; width: 340px;' id=boxID></div>" _
		
	'Box Count String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  90 & "px; left: 540px;height: 30px; width: 100px;text-align: right;'>Count:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart +  90 & "px; left: 640px;height: 30px; width: 225px;' id=boxCount>0</div>" 
		
	'Operator String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart + 120 & "px; left: 25px; height: 30px; width: 175px; text-align: right;'>Operator:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & midStart + 120 & "px; left: 200px; height: 30px; width: 340px;' id=OperID></div>" _
		
	'Update Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & midStart + 140 & "px; left: 615px;height: 50px; width: 200px;'>" _
			& "<button style='height: 50px; width: 200px;visibility:hidden;' id=updateButton onclick='done.value=""update""'>Update</button></div>"
		
	'AE Paperwork Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & botStart +   0 & "px; left: 615px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=AEButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart +   0 & "px; left: 650px;height: 30px; width: 220px;'>AE Paperwork</div>" 
		
	'Initial Inspection Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & botStart +  35 & "px; left: 615px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=InitialButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart +  35 & "px; left: 650px;height: 30px; width: 220px;'>Initial Inspection</div>" 
		
	'Final Inspection Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & botStart +  70 & "px; left: 615px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=FinalButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart +  70 & "px; left: 650px;height: 30px; width: 220px;'>Final Inspection</div>" 
		
	'CMM Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 105 & "px; left: 615px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=CMMButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 105 & "px; left: 650px;height: 30px; width: 220px;'>CMM File</div>" 
		
	'E-Tag Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 140 & "px; left: 615px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=ETagButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 140 & "px; left: 650px;height: 30px; width: 220px;'>E-Tags</div>" 
		
	'MRB Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 175 & "px; left: 615px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' disabled id=MRBButton></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 175 & "px; left: 650px;height: 30px; width: 220px;'>In MRB</div>" 
		
	'Bad Scan String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 205 & "px; left: 615px; height: 30px; width: 160px;visibility:hidden;' id=notFoundText>Error count:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart + 205 & "px; left: 775px; height: 30px; width: 95px;' id=notFoundCnt></div>" _
	
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div id=errorDiv style='top: " & botStart +   0 & "px; left: 0px; height: 240px; width: 600px;'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: " & botStart +  10 & "px; left: 50px; height: 215px; width: 500px; text-align: center;' id=errorString></div>"
		
	'All Op String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 805px;height: 30px; width: 30px;'><button class='opButton' style='height: 30px; width: 30px;' onclick='done.value=""allOps""'>&#10010;</button></div>"
		
	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 845px;height: 30px; width: 30px;'>" _
			& "<button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'>&#10006;</button></div>" _
		& "<div style='top: 0px; left: 1000px;'><button type=hidden id=returnToHTA 		style='visibility:hidden;' value=false onclick='HTAReturn()'><center></button></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=done 				style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=submitButton 		style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=submitText 		style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=loadMode 			style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=shipDate 			style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: 1000px;'><input type=hidden id=updateValue		style='visibility:hidden;' value=false><center></div>"
		
	'Modal Comment Div String
	LoadHTML = LoadHTML _
		& "<div id='commentModal' style='top: 1px; left: 1000px; height: 778px; width: 898px;'>" _
		& "<div style='top: 50px; left: 50px; height: 550px; width: 330px;' id='table_wrapper'></div>" _
		& "<div style='top: 650px; left: 450px; height: 48px; width: 100px;'><input type=button value='Close' style='height: 48px; width: 100px;' onclick='cancelComment()'></div>" _
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
 