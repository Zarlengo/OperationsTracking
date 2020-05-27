Option Explicit
	Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")

	Dim allOPSsource : allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\All Operations.vbs"
	Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """" & sArg
	Dim dataSource : dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"
	Dim adminPassword : adminPassword = "FLOW288"
	Dim tabletPassword : tabletPassword = "Fl0wSh0p17"
	Dim computerPassword : computerPassword = "Snowball18!"
	Dim duplicatePrefix : duplicatePrefix = "ADMIN ACCESS REQUIRED." & Chr(13)
	Dim SerialMode : SerialMode = "Automatic"
	Dim PMPre : PMPre = Chr(2)
	Dim PMSuf : PMSuf = Chr(3)

	Dim PMFile : PMFile = false
	Dim closeWindow : closeWindow = false
	Dim connect0 : connect0 = false
	Dim connect1 : connect1 = false
	Dim connect2 : connect2 = false
	Dim connect3 : connect3 = false
	Dim listen1 : listen1 = false
	Dim listen2 : listen2 = false
	Dim errorWindow : errorWindow = false
	Dim adminMode : adminMode = false
	Dim debugMode : debugMode = false

	Dim winsock0, winsock1, winsock2, winsock3, strAnswer
	Dim strData, windowBox, SNArray, PMString, AccessResult, machineString
	Dim SendData, RecieveData, secs, wmi, cProcesses, oProcess, RemoteHost3, RemotePort3
	'****** REVISION HISTORY **************
	'V1.0 -	Initial Release
	'V1.1 - Added requirement to code to scan each blade paperwork before part marking
	'		Update code to match improvements from other OP scrips
	'V1.2 - Switched to SQL
	'V1.3 - Changed admin login to a prompt for duplicate
	'		Changed admin login process
	'		Added SN wipe to part marker each time slug is scanned
	'		Added Handheld to type of scanning for SQL data
	'V1.4	Added ability to line out without needing to have a new serial number
	'		Fixed comments not being added to database when duplicating / cross-out mode
	'		Need to do:
	'		Fix bug which stops listening to left camera
	'		N/A per Sean...
	'
	'V2.0 - Added Operator ID
	'		Removed Job Traveler Scanning
	'		Added automatic logout
	'V2.1 - Bug fix
	'		Duplicate marking not working
	'		Reset command not being sent to  marker
	'		Added check for AX Work Orders
	'		Code clean up
	'V2.2 - Added Work order check
	'V2.3 - Added Static IPs for Markers - Ben Horbul
	'****** CHANGE THESE SETTINGS *********

	Const LocalPort1           = 3000
	Const LocalPort2           = 3001

	'Const RemoteHost0          = "10.2.105.217" 	' M4 Inline Part Marker - MAIN   MAC# 00:50:C2:80:73:DA
	'	Const RemoteHost0          = "10.2.105.218"		' M4 Inline Part Marker - SPARE  MAC# 00:50:C2:80:73:2F
	Const RemoteHost0          = "10.2.101.28" 	' M4 Inline Part Marker - MAIN   MAC# 00:50:C2:80:73:DA

	
	Const RemotePort0          = 6120

	Dim searchTime : searchTime = 10 / 60 / 60 / 24
	Dim logoutTime : logoutTime = 45 / 60 / 24

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

	On Error Resume Next
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "30", "REG_DWORD"	'Changes TCP timeout settings if needing to restart program w/in 5 minutes
	On Error Goto 0

	'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
	Set wmi = GetObject("winmgmts:root\cimv2") 
	Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
	For Each oProcess in cProcesses
		oProcess.Terminate()
	Next
	
	'// CREATE WINSOCK: 0 - Part marker; 1 - Left camers; 2 - Right camera, 3 - Scanner
	Set winsock0 = Wscript.CreateObject("MSWINSOCK.Winsock", "winsock0_")
	If Err.Number <> 0 Then
		MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
		WScript.Quit
	End If
	Set winsock1 = Wscript.CreateObject("MSWINSOCK.Winsock", "winsock1_")
	If Err.Number <> 0 Then
		MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
		WScript.Quit
	End If
	Set winsock2 = Wscript.CreateObject("MSWINSOCK.Winsock", "winsock2_")
	If Err.Number <> 0 Then
		MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
		WScript.Quit
	End If
	Set winsock3 = Wscript.CreateObject("OSWINSCK.Winsock", "winsock3_")
	If Err.Number <> 0 Then
		MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
		WScript.Quit
	End If
	
	If Not WScript.Arguments.Count = 0 Then
		sArg = ""
		For Each Arg In Wscript.Arguments
			  sArg = sArg & " " & """" & Arg & """"
		Next
	End If

	machineString = sArg
	If sArg <> "" Then
		Do While AscW(Right(machineString, 1)) = 34 or AscW(Right(machineString, 1)) = 32
			machineString = Left(machineString, Len(machineString) - 1)
		Loop
		Do While AscW(Left(machineString, 1)) = 34 or AscW(Left(machineString, 1)) = 32
			machineString = Right(machineString, Len(machineString) - 1)
		Loop
		Load_IP
	Else
		machineString = "Manual"
	End If
	
	' loads port settings into winsock
	winsock0.RemoteHost = RemoteHost0
	winsock0.RemotePort = RemotePort0
	winsock1.LocalPort = LocalPort1
	winsock2.LocalPort = LocalPort2

	'Function to check for access connection and load info from database
	AccessResult = Load_Access

	'Calls function to create ie window
	set windowBox = HTABox("white", 940, 850, 30, 30) 
	
	'Turns on port listening for camera 1 and 2
	Server1Listen
	Server2Listen

	 Dim startTime, logoutStart 
with windowBox
	.document.title = "Offset Change Recording"
	'Connects to the part marker
	winsock0.Connect       
	'// MAIN DELAY - WAITS FOR CONNECTED STATE
	'// SOCKET ERROR RAISES WINSOCK ERROR SUB
	while winsock0.State <> sckError And winsock0.state <> sckConnected And winsock0.state <> sckClosing And secs <> 25
		WScript.Sleep 1000  '// 1 sec delay in loop
		secs = secs + 1     '// wait 25 secs max
	Wend
	'// CONNECTION TIMED OUT
	If secs > 24 Then
		MsgBox "Timed Out"
		ServerClose()
	End If
	'Stores variable if connected to part marker
	IF winsock0.state = sckConnected Then connect0 = true
	IF winsock3.state = sckConnected Then connect3 = true

	'Function to verify all connections are open
	checkConnections
	If adminMode = true Then adminSettings
	do until closeWindow = true													'Run loop until conditions are met
		startTime = now
		logoutStart = now
		do until .done.value = "cancel" or .adminCheck.value = "true" or .crossOutClick.value = "true" or .done.value = "allOps"
			wsh.sleep 50	
			On Error Resume Next
			If .done.value = true Then
				ServerClose()
				wsh.quit
			End If
			On Error GoTo 0
			If logoutStart + logoutTime < now Then
				windowBox.operator.innerText = ""
				windowBox.errorString.innerText = "Logged out"
				logoutStart = now
			End If
		loop
		If .done.value = "allOps" and .adminValue.value = 3 then											'If the x button is clicked
			.OPSButton.disabled = true
			.OPSButton.disabled = false
			.done.value = false
			strAnswer = InputBox("Please enter the Administrator password:")
			If strAnswer = adminPassword Then
				adminSettings
			Else
				.adminValue.value = 0
			End If
		elseif .done.value = "cancel" then											'If the x button is clicked
			closeWindow = true													'Variable to end loop
		elseif .crossOutClick.value = "true" then								'(ADMIN ONLY) If the cross-out option is selected
			.crossOutClick.value = "false"										'Change cross-out variable back to false		
			Load_PM_File .crossOutMode.value									'Function to change what file is loaded on the part marker
		ElseIf .done.value = "allOps" Then
			objShell.Run sOPsCmd
			ServerClose()	
			WScript.Quit	
		ElseIf .adminCheck.value = "true"  Then
			.adminCheck.value = false
			If adminMode <> true Then strAnswer = InputBox("Please enter the Administrator password:")
			If strAnswer = adminPassword or adminMode = true Then
				windowBox.allowDuplicate.value = true
				windowBox.errorString.innerText = "Access granted." & chr(10) & "Please rescan Slug"
			Else
				msgbox "Wrong password"
			End If
		ElseIf .done.value = "allOps" Then
			objShell.Run sOPsCmd
			WScript.Quit	
		End If 
	loop
	.close																		'Closes the window
 end with

 ServerClose()																	'Function to close open connections and return settings back to original	
 wsh.quit

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

Function checkConnections()
	Dim testInput
	Dim secs : secs = 0
	If AccessResult = false Then
		windowBox.errorString.innerText = "Access database not loaded"
		If winsock0.state <> sckClosed Then winsock0.Close
		If winsock1.State <> sckClosed Then winsock1.Close
		If winsock2.State <> sckClosed Then	winsock2.Close
		If winsock3.State <> sckClosed Then	winsock3.Close
		connect0 = false
		listen1 = false
		listen2 = false
		connect3 = false
	End if
	If connect0 = true Then
		Load_PM_File false
		windowBox.partMarkText.innerText = "Connected to Part Marker"
		windowBox.partMarkButton.style.backgroundcolor = "limegreen"
		windowBox.partMarkButton.disabled = true
		If listen1 = true and listen2 = true Then
			windowBox.connectionText.innerText = "Listening for Cameras Left and Right"
			windowBox.connectButton.style.backgroundcolor = "yellow"
			windowBox.connectButton.disabled = true
		ElseIf listen1 = true Then
			windowBox.connectionText.innerText = "Listening for Camera 1"
			windowBox.connectButton.style.backgroundcolor = "orange"
		ElseIf listen2 = true then
			windowBox.connectionText.innerText = "Listening for Camera 2"
			windowBox.connectButton.style.backgroundcolor = "orange"
		Else
			windowBox.connectionText.innerText = "Not Listening for any Camera"
			windowBox.connectButton.style.backgroundcolor = "red"
		End If
	Else
		windowBox.partMarkText.innerText = "Not Connected to the Part Market"
		windowBox.partMarkButton.style.backgroundcolor = "red"
		If winsock1.State <> sckClosed Then winsock1.Close
		If winsock2.State <> sckClosed Then	winsock2.Close
		windowBox.connectionText.innerText = "Disconnected"
		windowBox.connectButton.style.backgroundcolor = "red"
		windowBox.connectButton.disabled = true
	End If
	
	If machineString <> "Manual" and machineString <> "" Then
		windowBox.scannerText.innerText = "Connect to " & machineString
		windowBox.scannerButton.style.backgroundcolor = "orange"
		windowBox.scannerButton.disabled = true
		windowBox.errorString.innerText = ""
	End If
	' loads port settings into winsock
	If RemoteHost3 <> "" and RemotePort3 <> "" Then 
		winsock3.RemoteHost = RemoteHost3
		winsock3.RemotePort = RemotePort3
		'Connects to the scanner
		On Error Resume Next
		winsock3.Connect    
		On Error GoTo 0
		'// MAIN DELAY - WAITS FOR CONNECTED STATE
		'// SOCKET ERROR RAISES WINSOCK ERROR SUB
		while winsock3.State <> sckError And winsock3.state <> sckConnected And winsock3.state <> sckClosing And secs < 25
			WScript.Sleep 1000  '// 1 sec delay in loop
			secs = secs + 1     '// wait 25 secs max
		Wend
	End If
	'Stores variable if connected to part marker
	If machineString = "Manual" Then
		windowBox.scannerText.innerText = "Manual scanner mode"
		windowBox.scannerButton.style.backgroundcolor = "limegreen"
		windowBox.scannerButton.disabled = true
		windowBox.errorString.innerText = ""
	ElseIf winsock3.state = sckConnected Then 
		windowBox.scannerText.innerText = "Connected to " & machineString
		windowBox.scannerButton.style.backgroundcolor = "limegreen"
		windowBox.scannerButton.disabled = true
	Else
		windowBox.scannerText.innerText = "Error: " & machineString
		windowBox.scannerButton.style.backgroundcolor = "red"
		windowBox.scannerButton.disabled = false
	End If
 End Function

Function HTA_Data(DataString, PortNumber)
	Dim dataArray, dateString, timeString, SerialNumber, Width, Height, Angle, Length
	On Error Resume Next
	dataArray = Split(TrimString(DataString), ";")
	dateString = dataArray(0)
	timeString = dataArray(1)
	SerialNumber = dataArray(2)
	Width = dataArray(3)
	Height = dataArray(4)
	Angle = dataArray(5)
	Length = dataArray(6)
	On Error Goto 0
	If PortNumber = 3001 then
		windowBox.camera1date.innerText = dateString
		windowBox.camera1time.innerText = timeString
		windowBox.camera1SN.innerText = TrimString(SerialNumber)
		windowBox.camera1width.innerText = Width
		windowBox.camera1height.innerText = Height
		windowBox.camera1angle.innerText = Angle
	ElseIf PortNumber = 3000 then
		windowBox.camera2date.innerText = dateString
		windowBox.camera2time.innerText = timeString
		windowBox.camera2SN.innerText = TrimString(SerialNumber)
		windowBox.camera2width.innerText = Width
		windowBox.camera2height.innerText = Height
		windowBox.camera2angle.innerText = Angle
		windowBox.camera2length.innerText = Length
	End If
	If InStr(SerialNumber, "#ERR") = 0 Then
		SerialMode = "Automatic"
		Find_SN SerialNumber
	ElseIf windowBox.SerialNumberInput.value <> "" Then
		If windowBox.SerialNumberInput.value = adminPassword Then
			windowBox.manualSerialNumber.style.backgroundColor = ""
			windowBox.SerialNumberText.style.visibility = "visible"
			windowBox.SerialNumberInput.value = ""
			windowBox.SerialNumberInput.style.visibility = "hidden"
			Exit Function
		End If
		SerialNumber = windowBox.SerialNumberInput.value
		windowBox.camera1SN.innerText = TrimString(SerialNumber)
		SerialMode = "Manual"
		Find_SN SerialNumber
	End If
 End Function

Function adminSettings()
	windowBox.SerialNumberInput.value = ""
	windowBox.operator.innerText = ""
	windowBox.errorString.innerText = "ADMIN ACCESS GRANTED"
	windowBox.duplicateButton.disabled = false
	windowBox.adminText.style.visibility = "visible"
	windowBox.adminButton.style.visibility = "visible"
	windowBox.adminString.style.visibility = "visible"
	windowBox.logoutButton.style.visibility = "visible"
	adminMode = true
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

Function Find_SN(Serial_Number)
	Dim j, objCmd , rs, userName, sqlString
	Dim SN1_Found : SN1_Found = false
	Dim SN2_Found : SN2_Found = false
	windowBox.manualSerialNumber.style.backgroundColor = ""
	windowBox.SerialNumberText.style.visibility = "visible"
	windowBox.SerialNumberInput.value = ""
	windowBox.SerialNumberInput.style.visibility = "hidden"
	If Serial_Number = computerPassword or Serial_Number = tabletPassword Then
		Exit Function
	ElseIf UCase(Left(Serial_Number, 2)) = "AE" Or UCase(Left(Serial_Number, 4)) = "M500" Then
		windowBox.errorString.innerText = "Test Part"
		windowBox.blade1string.innerText = "1" & UCase(Serial_Number) & " " & Right(FormatNumber(Now(), 3), 7)
		windowBox.blade2string.innerText = "2" & UCase(Serial_Number) & " " & Right(FormatNumber(Now(), 3), 7)
		windowBox.accessNewEntry.Value = True
		Load_SN_to_PM ""
		Exit Function
	ElseIf Serial_Number = "WRONG_PM" Then
		windowBox.errorString.innerText = "Wrong Part Mark"
		windowBox.blade1string.innerText = "WRONG_PM"
		windowBox.blade2string.innerText = "WRONG_PM"
		windowBox.accessNewEntry.Value = True
		Load_SN_to_PM ""
		Exit Function
	ElseIf Left(Serial_Number, 1) = "D" Then 'and Len(Serial_Number) = 10 and Mid(Serial_Number, 9, 1) = "-" Then
		windowBox.waiting4PM.value = False
		windowBox.blade1string.innerText = "Waiting..."
		windowBox.blade2string.innerText = "Waiting..."
		Load_SN_to_PM "close"
		If windowBox.operator.innerText = "" or windowBox.operator.innerText = "Not Authorized" Then
			windowBox.errorString.innerText = "Missing Operator"
			Exit Function
		ElseIf SerialMode = "Handheld" Then
			windowBox.handheld.innerText = Serial_Number
		End If
		
		Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkConnections : Exit Function
		sqlString = "SELECT DISTINCT [SlugProdID], [Dash1ProdID], [Dash2ProdID] FROM [00_Invoice] LEFT JOIN [00_AE_SN_Control] ON [00_AE_SN_Control].[Invoice Number] = [00_Invoice].[Invoice Number] WHERE [00_AE_SN_Control].[Slug Serial Number] = '" & Serial_Number & "';"
		set rs = objCmd.Execute(sqlString)
		DO WHILE NOT rs.EOF
			If IsNull(rs.Fields(0)) or IsNull(rs.Fields(1)) or IsNull(rs.Fields(2)) Then
				windowBox.errorString.innerText = "Missing Work Orders."
				Exit Function
			End If
			rs.MoveNext
		Loop	
		Set rs = Nothing
		
		For j=LBound(SNArray, 2) to UBound(SNArray, 2)	
			If SNArray(0, j) = Serial_Number and Right(SNArray(2, j), 1) = "1" Then
				SN1_Found = SNArray(1, j)
			ElseIf SNArray(0, j) = Serial_Number and Right(SNArray(2, j), 1) = "2" Then
				SN2_Found = SNArray(1, j)
			ElseIf SNArray(0, j) = Serial_Number Then
				SN1_Found = SNArray(1, j)
			End If
			If SN1_Found <> false and SN2_Found <> false Then
				Exit For
			End If
		Next
		windowBox.duplicateButton.style.visibility = "hidden"
		If InStr(PMString, SN1_Found) <> 0 or InStr(PMString, SN2_Found) <> 0 Then
			If (windowBox.duplicateSlug.value = "True" or windowBox.duplicateSlug.value = "true") and (windowBox.allowDuplicate.value = "False" or windowBox.allowDuplicate.value = "false") Then
				windowBox.errorString.innerText = "Please approve duplicate marking"
				windowBox.duplicateButton.style.visibility = "visible"
				Exit Function
			ElseIf (windowBox.duplicateSlug.value = "True" or windowBox.duplicateSlug.value = "true") and (windowBox.allowDuplicate.value = "True" or windowBox.allowDuplicate.value = "true") Then
				windowBox.allowDuplicate.value = False
				windowBox.accessNewEntry.Value = False
				windowBox.duplicateSlug.value = False
				windowBox.blade1string.innerText = SN1_Found
				windowBox.blade2string.innerText = SN2_Found
				Load_SN_to_PM ""
				Exit Function
			End If
			windowBox.errorString.innerText = duplicatePrefix & "Serial Number Already Marked: " & Serial_Number
			windowBox.blade1string.innerText = SN1_Found
			windowBox.blade2string.innerText = SN2_Found
			windowBox.duplicateButton.style.visibility = "visible"
			windowBox.allowDuplicate.value = False
			windowBox.accessNewEntry.Value = False
			windowBox.duplicateSlug.value = True
		ElseIf SN1_Found <> false and SN2_Found <> false Then
			windowBox.blade1string.innerText = SN1_Found
			windowBox.blade2string.innerText = SN2_Found
			windowBox.allowDuplicate.value = False
			windowBox.accessNewEntry.Value = True
			windowBox.duplicateSlug.value = False
			Load_SN_to_PM ""
		ElseIf SN1_Found = false and SN2_Found = false Then
			windowBox.errorString.innerText = "Serial number not found: " & Serial_Number
			windowBox.blade1string.innerText = "Waiting..."
			windowBox.blade2string.innerText = "Waiting..."
			windowBox.allowDuplicate.value = False
			windowBox.accessNewEntry.Value = False
		ElseIf SN1_Found <> false or SN2_Found <> false Then
			windowBox.errorString.innerText = "Missing blade serial number: " & Serial_Number
			windowBox.blade1string.innerText = "Waiting..."
			windowBox.blade2string.innerText = "Waiting..."
			windowBox.allowDuplicate.value = False
			windowBox.accessNewEntry.Value = False
		End If
	ElseIf Left(Serial_Number, 3) = "_NO" Then
		windowBox.waiting4PM.value = False
		Load_SN_to_PM "close"
		For j=LBound(SNArray, 2) to UBound(SNArray, 2)	
			If SNArray(0, j) = Serial_Number and SNArray(2, j) = "060053-1" Then
				SN1_Found = SNArray(1, j)
			ElseIf SNArray(0, j) = Serial_Number and SNArray(2, j) = "060053-2" Then
				SN2_Found = SNArray(1, j)
			ElseIf SNArray(0, j) = Serial_Number Then
				SN1_Found = SNArray(1, j)
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
			windowBox.allowDuplicate.value = False
			windowBox.accessNewEntry.Value = False
			windowBox.duplicateSlug.value = True
		Else
			windowBox.blade1string.innerText = SN1_Found
			windowBox.blade2string.innerText = SN2_Found
			windowBox.allowDuplicate.value = False
			windowBox.accessNewEntry.Value = True
			windowBox.duplicateSlug.value = False
		End If
		If (windowBox.duplicateSlug.value = "True" or windowBox.duplicateSlug.value = "true") and (windowBox.allowDuplicate.value = "False" or windowBox.allowDuplicate.value = "false") Then
			windowBox.errorString.innerText = "Please login to admin mode"
			Exit Function
		Else
			Load_SN_to_PM ""
		End If
	ElseIf IsNumeric(Serial_Number) Then
		Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkConnections : Exit Function
		sqlString = "SELECT TOP 1 [LastName] FROM [00_Personnel] WHERE [USERID]='" & (Serial_Number) & "';"
		set rs = objCmd.Execute(sqlString)	
		userName = ""
		DO WHILE NOT rs.EOF
			userName = rs.Fields(0)
			rs.MoveNext
		Loop	
		Set rs = Nothing
		objCmd.Close
		Set objCmd = Nothing
		If userName <> "" Then
			windowBox.operator.innerText = userName
		Else
			windowBox.operator.innerText = "Not Authorized"
		End If
	Else
		Dim inputArray : inputArray = Split(Serial_Number)
		If UBound(inputArray) > 0 Then
			sqlString = "SELECT TOP 1 [LastName] FROM [00_Personnel] WHERE [FirstName]='" & inputArray(0) & "' and [LastName] ='" & inputArray(1) & "';"
		Else
			sqlString = "SELECT TOP 1 [LastName] FROM [00_Personnel] WHERE [LastName] ='" & inputArray(0) & "';"
		End If
		Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkConnections : Exit Function
		set rs = objCmd.Execute(sqlString)	
		userName = ""
		DO WHILE NOT rs.EOF
			userName = rs.Fields(0)
			rs.MoveNext
		Loop	
		Set rs = Nothing
		objCmd.Close
		Set objCmd = Nothing
		If userName <> "" Then
			windowBox.operator.innerText = userName
		Else
			windowBox.operator.innerText = "Not Authorized"
		End If
	End If
 End Function

'********* PARK MARK INFORMATION | CHAPTER 4 OF M4 INTERFACING INSTRUCTION *******************
 ' String to send to the Part Marker
 '  Chr(2)									[STX] String Start Code
 '  Chr(3)									[ETX] String End Code
 '  Chr(13)									[CR] End of Command
 '  [STX]COMMAND[CR][ETX]					Part Mark String 
 '  Chr(2) & "COMMAND" & Chr(13) & Chr(3)	String in ASCII Format
 '
 ' Standard Variables for the Part Marker
 '  A & "FILENAME"							Command to change the file on the part marker
 '	B & "VARIABLENAME"						Command to send a variables value to the marker
 '  G										Disable response to the computer (used when changing files)
 '	R0										Return code inactive (will not send signal to computer)
 '	R1										Return code active (will send signal to computer when complete)
 '											Return signal is: [STX]Marking End[ETX] and [STX]Cycle End[ETX]
 '
 ' Output state change - P00E
 ' 00										Number of the output
 ' E										0 for inactive, 1 for active
 ' P01										//NOT USED// Marking
 ' P02										//NOT USED// Marking End
 ' P03										Ready - Triggers the system to activate the safety control system
 ' P04										Marking - Used to clear the marking process
 ' P05										//NOT USED// Fault
 ' P06										//NOT USED// Marking End
 ' P07										//NOT USED// Origine
 ' P08										//NOT USED// Maintenance: Revision
 ' P09										//NOT USED// Pause
 '
 ' Custom Variable for the Part Marker
 '  BBLADESN1="SERIALNUMBER"				Variable for the first serial number
 '  BBLADESN2="SERIALNUMBER"				Variable for the second serial number
 '
 ' File Names:
 '  "LPT5_FROM_COMP"						File used for normal part marking, includes data matrix
 '  "LPT5_CROSS_OUT"						File used to line out existing serial numbers, also will write a new serial number above previous ones
 
	'Dim StringIO : StringIO = "G" & Chr(13) & "R0" & Chr(13) & "P031" & Chr(13) & "P040" & Chr(13) & "A"
	'Dim blankSN : blankSN = "BBLADESN1=  " & Chr(13) &  "BBLADESN2=  " & Chr(13)
	'StringToSend = PMPre & StringIO & PMString & blankSN & PMSuf
	
Function Load_PM_File(CrossOut)
	Dim StringToSend
	Dim StringIO : StringIO = "G" & Chr(13) & "A"
	Dim PMString : PMString = "LPT5_FROM_COMP" & Chr(13)
	Dim CrossOutStringSuf : CrossOutStringSuf = "BBLADESN1=  " & Chr(13) &  "BBLADESN2=  " & Chr(13) & "R0" & Chr(13) & "P031" & Chr(13) & "P040" & Chr(13)
	'Get current file from controller
	'	StringToSend = PMPre & "X" & StringIO & PMSuf
	If CrossOut = "true" or CrossOut = "True" Then
		PMString = "LPT5_CROSS_OUT" & Chr(13) & CrossOutStringSuf
	End If
	StringToSend = PMPre & StringIO & PMString & PMSuf
	winsock0.SendData StringToSend                ' Send string to machine
 End Function

Function Load_SN_to_PM(state)
	'state = "close" for clear all, "" for all other
	Dim Serial1, Serial2, SNString1, SNString2, StringToSend
	Dim StringIO : StringIO = "P031" & Chr(13) & "P040" & Chr(13) & "R1" & Chr(13)
	
	If state = "close" Then
		SNString1 = "BBLADESN1=" & Chr(13)
		SNString2 = "BBLADESN2=" & Chr(13)
		StringIO = "P030" & Chr(13) & "P040" & Chr(13)
	Else
		Serial1 = windowBox.blade1string.innerText
		Serial2 = windowBox.blade2string.innerText
		SNString1 = "BBLADESN1=" & Serial1 & Chr(13)
		SNString2 = "BBLADESN2=" & Serial2 & Chr(13)
		windowBox.errorString.innerText = "Data sent, waiting for completion." 
		windowBox.waiting4PM.value = True
	End If
	StringToSend = PMPre & SNString1 & SNString2 & StringIO & PMSuf
	winsock0.SendData StringToSend                ' Send string to machine
 End Function

Function LoadSNtoAccess(receiveString)
	Dim Serial1, Serial2, strQuery1, strQuery2, CurrentTime, objCmd, Comment, Operator
	Dim addString : addString = ""

	If InStr(1, receiveString, "Cycle End") <> 0 And windowBox.waiting4PM.value = "True" Then
		windowBox.waiting4PM.value = False
		Serial1 = windowBox.blade1string.innerText
		Serial2 = windowBox.blade2string.innerText
		Operator = windowBox.operator.innerText
		If windowBox.crossOutMode.value = "true" or windowBox.crossOutMode.value = "True" Then
			If windowBox.accessNewEntry.Value = "False" Then
				Comment = "Cross-Out Duplicate"
				addString = ", [Duplicate] = 1"
			Else
				Comment = "Cross-Out"
			End If
		ElseIf windowBox.accessNewEntry.Value = "False" Then
			Comment = "Duplicate Part Mark by:" & Operator
			addString = ", [Duplicate] = 1"
		End If
		CurrentTime = Now
		If windowBox.accessNewEntry.Value = "True" Then
			strQuery1 = "INSERT INTO [10_Part_Marking] ([Blade Serial Number], [Date SN Marked], [Mode], [Operator]) " & "VALUES ('" & Serial1 & "', '" & CurrentTime & "', '" & SerialMode & "', '" & Operator & "'); "
			strQuery2 = "INSERT INTO [10_Part_Marking] ([Blade Serial Number], [Date SN Marked], [Mode], [Operator]) " & "VALUES ('" & Serial2 & "', '" & CurrentTime & "', '" & SerialMode & "', '" & Operator & "'); "
		Else
			strQuery1 = "UPDATE [10_Part_Marking] SET [Date SN Marked] = '" & CurrentTime & "', [Mode] = '" & SerialMode & "', [Comments] = '" & Comment & "' " & addString & " WHERE [Blade Serial Number] = '" & Serial1 & "'; "
			strQuery2 = "UPDATE [10_Part_Marking] SET [Date SN Marked] = '" & CurrentTime & "', [Mode] = '" & SerialMode & "', [Comments] = '" & Comment & "' " & addString & " WHERE [Blade Serial Number] = '" & Serial2 & "'; "
		End If
		
		set objCmd = GetNewConnection
		If objCmd is Nothing Then
			windowBox.errorString.innerText = "Error connecting to database, data not sent"
			Exit Function
		End If
		objCmd.Execute(strQuery1)
		objCmd.Execute(strQuery2)
		
		PMString = PMString & Serial1 & ";" & Serial2 & ";"
		objCmd.Close
		Set objCmd = Nothing
		windowBox.errorString.innerText = "Part mark complete." & Chr(13) & "060053-1: " & Serial1 & Chr(13) & "060053-2: " & Serial2
		CleanUpScreen
	End If
 End Function

Function Load_IP()
	Dim sqlString, rs
	Dim objCmd : set objCmd = GetNewConnection
	On Error GoTo 0
	If objCmd is Nothing Then Exit Function
	sqlString = "Select [IPAddress], [Port] From [00_Machine_IP] WHERE [DeviceType] = 'CognexBTHandheld' AND [MachineName] = '" & machineString & "'"
	If machineString <> "Manual" Then
		set rs = objCmd.Execute(sqlString)		
		DO WHILE NOT rs.EOF
			RemoteHost3 = rs.Fields(0)
			RemotePort3 = rs.Fields(1)
			rs.MoveNext
		Loop	
	End If
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
	
 End Function

Function CleanUpScreen()
	windowBox.allowDuplicate.value = False
	windowBox.waiting4PM.value = False
	windowBox.waiting4PM.value = False
	windowBox.accessNewEntry.Value = False
	windowBox.SerialNumberInput.Value = ""	
	windowBox.camera1date.innerText = ""
	windowBox.camera1time.innerText = ""
	windowBox.camera1SN.innerText = ""
	windowBox.camera2date.innerText = ""
	windowBox.camera2time.innerText = ""
	windowBox.camera2SN.innerText = ""
	windowBox.camera1width.innerText = ""
	windowBox.camera1height.innerText = ""
	windowBox.camera1angle.innerText = ""
	windowBox.camera2width.innerText = ""
	windowBox.camera2height.innerText = ""
	windowBox.camera2angle.innerText = ""
	windowBox.camera2length.innerText = ""
	windowBox.duplicateButton.style.visibility = "hidden"
	windowBox.duplicateSlug.value = False
	windowBox.blade1string.innerText = "Waiting..."
	windowBox.blade2string.innerText = "Waiting..."
	windowBox.handheld.innerText = ""
	windowBox.manualSerialNumber.style.backgroundColor = ""
	windowBox.SerialNumberText.style.visibility = "visible"
	windowBox.SerialNumberInput.value = ""
	windowBox.SerialNumberInput.style.visibility = "hidden"
	SerialMode = "Automatic"
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
	sqlString = "Select [Slug Serial Number], [Blade Serial Number], [FIC Blade Part Number] From [00_AE_SN_Control]"
	set rs = objCmd.Execute(SQLString)
	ReDim SNArray(2,0)
	SNArray(0, 0) = "Slug_SN"
	SNArray(1, 0) = "BladeSN"
	SNArray(2, 0) = "PartNumber"
	
	DO WHILE NOT rs.EOF
		SN_Size = UBound(SNArray, 2) + 1
		ReDim Preserve SNArray(2, SN_Size)
		SNArray(0, SN_Size) = rs.Fields(0)
		SNArray(1, SN_Size) = rs.Fields(1)
		SNArray(2, SN_Size) = rs.Fields(2)
		If (Right(rs.Fields(2), 1) = "1") Then
			SN_String = SN_String & rs.Fields(0) & ";"
		End If
		rs.MoveNext
	Loop
	
	Set rs = Nothing
	
	sqlString = "Select [Blade Serial Number] From [10_Part_Marking]"
	set rs = objCmd.Execute(SQLString)
	
	DO WHILE NOT rs.EOF
		PMString = PMString & rs.Fields(0) & ";"
		rs.MoveNext
	Loop
	
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
	Load_Access = true
 End Function

'// WINSOCK CONNECT REQUEST
Sub winsock1_ConnectionRequest(requestID)
    If winsock1.State <> sckClosed Then
        winsock1.Close
    End If
    winsock1.Accept requestID
    winsock1.SendData "Server Received okay"
	If connect2 = true Then
		windowBox.connectionText.innerText = "Connected to Cameras Left and Right"
	Else
		windowBox.connectionText.innerText = "Connected to Camera Right"
	End If
	windowBox.connectButton.style.backgroundcolor = "limegreen"
	windowBox.connectButton.disabled = true
	connect1 = true
    WScript.Sleep 1000  '// REQUIRED OR ERRORS
 End Sub
Sub winsock2_ConnectionRequest(requestID)
    If winsock2.State <> sckClosed Then
        winsock2.Close
    End If
    winsock2.Accept requestID
    winsock2.SendData "Server Received okay"
	If connect1 = true Then
		windowBox.connectionText.innerText = "Connected to Cameras Left and Right"
	Else
		windowBox.connectionText.innerText = "Connected to Camera Left"
	End If
	windowBox.connectButton.style.backgroundcolor = "limegreen"
	windowBox.connectButton.disabled = true
	connect2 = true
    WScript.Sleep 1000  '// REQUIRED OR ERRORS
 End Sub

'// WINSOCK DATA ARRIVES
Sub winsock0_dataArrival(bytesTotal)
	logoutStart = now
    winsock0.GetData strData, vbString
    WScript.Sleep 1000
	LoadSNtoAccess strData
 End Sub
Sub winsock1_dataArrival(bytesTotal)
	logoutStart = now
    winsock1.GetData strData, vbString
	HTA_Data strData, LocalPort1
    WScript.Sleep 2000  '// REQUIRED OR ERRORS
    Server1Listen()
 End Sub
Sub winsock2_dataArrival(bytesTotal)
	logoutStart = now
    winsock2.GetData strData, vbString
	HTA_Data strData, LocalPort2
    WScript.Sleep 2000  '// REQUIRED OR ERRORS
    Server2Listen()
 End Sub
Sub winsock3_OnDataArrival(bytesTotal)
	Dim dataString
	logoutStart = now
    winsock3.GetData strData, vbString
    WScript.Sleep 1000
	dataString = TrimString(strData)
	If Left(dataString, 1) = "D" and Len(dataString) = 10 and Mid(dataString, 9, 1) = "-" Then SerialMode = "Handheld"
	Find_SN dataString
 End Sub

'// WINSOCK ERROR
Sub winsock0_Error(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
    MsgBox "Part marker Client Error: " & Number & vbCrLf & Description
    ServerClose()
 End Sub
Sub winsock1_Error(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
    MsgBox "Left camera Server Error " & Number & vbCrLf & Description
    ServerClose()
 End Sub
Sub winsock2_Error(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
    MsgBox "Right camera Server Error " & Number & vbCrLf & Description
    ServerClose()
 End Sub
Sub winsock3_OnError(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
	windowBox.scannerText.innerText = "Error: " & machineString
	windowBox.scannerButton.style.backgroundcolor = "red"
	windowBox.scannerButton.disabled = false
    windowBox.errorString.innerText = "Scanner Error: " & Number & vbCrLf & Description
 End Sub

'// LISTEN FOR REQUEST
Sub Server1Listen()
    If winsock1.State <> sckClosed Then
        winsock1.Close
    End If
    winsock1.Listen
	listen1 = true
 End Sub
Sub Server2Listen()
    If winsock2.State <> sckClosed Then
        winsock2.Close
    End If
    winsock2.Listen
	listen2 = true
 End Sub

'// EXIT SCRIPT
Sub ServerClose()
	ON ERROR RESUME NEXT
	Load_SN_to_PM "close"
	Load_PM_File false
	windowBox.accessNewEntry.Value = False
    WScript.Sleep 1000  '// REQUIRED OR ERRORS
	windowBox.crossOutClick.value = "false"
    If winsock0.state <> sckClosed Then winsock0.Close
    If winsock1.state <> sckClosed Then winsock1.Close
	If winsock2.state <> sckClosed Then winsock2.Close
	If winsock3.state <> sckClosed Then winsock3.Close
    Set winsock0 = Nothing
    Set winsock1 = Nothing
    Set winsock2 = Nothing
    Set winsock3 = Nothing
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"

	windowBox.close
	On Error GoTo 0
    Wscript.Quit
 End Sub

'Function to create all of the JS and HTML code for the window
Function LoadHTML(sBgColor)
	'HTA String
	LoadHTML = "<HTA:Application contextMenu=no border=thin caption=no minimizebutton=yes maximizebutton=no sysmenu=yes />"
	
	'CSS String
	LoadHTML = LoadHTML _	
		& "<head><style>" _
		& "body {" _
			& "background-color: " & sBgColor & ";" _
			& "font:normal 22px Tahoma;" _
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
			& "document.getElementById('manualSerialNumber').disabled=true;" _
			& "document.getElementById('manualSerialNumber').disabled=false;" _
			& "if (document.getElementById('manualSerialNumber').style.backgroundColor == 'dimgrey') {" _
				& "document.getElementById('manualSerialNumber').style.backgroundColor = '';" _
				& "document.getElementById('SerialNumberText').style.visibility = 'visible';" _
				& "document.getElementById('SerialNumberInput').value = '';" _
				& "document.getElementById('SerialNumberInput').style.visibility = 'hidden';" _
				& "document.getElementById('errorString').innerText = '';" _
				& "document.getElementById('blade1string').innerText = 'Waiting...';" _
				& "document.getElementById('blade2string').innerText = 'Waiting...';" _
				& "document.getElementById('camera1SN').innerText = '';" _
				& "document.getElementById('camera2SN').innerText = '';" _
			& "} else {" _
				& "document.getElementById('manualSerialNumber').style.backgroundColor = 'DimGrey';" _
				& "document.getElementById('SerialNumberText').style.visibility = 'hidden';" _
				& "document.getElementById('SerialNumberInput').style.visibility = 'visible';" _
			& "}" _
		& "}" _
		& "function crossOutButton() {" _
			& "document.getElementById('adminButton').disabled = true;" _
			& "document.getElementById('adminButton').disabled = false;" _
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
		& "function adminLogin() {" _
			& "document.getElementById('duplicateButton').disabled=true;" _
			& "document.getElementById('duplicateButton').disabled=false;" _
			& "document.getElementById('adminCheck').value=true;" _
		& "}" _
		& "function adminMode(divID) {" _
			& "if (document.getElementById('adminValue').value == 0 && divID == 1) {" _
				& "document.getElementById('errorString').innerText = 'LOGGED OUT';" _
				& "document.getElementById('operator').innerText = '';" _
				& "document.getElementById('adminValue').value = 1;" _
			& "} else if (document.getElementById('adminValue').value == 1 && divID == 2) {" _
				& "document.getElementById('adminValue').value = 2;" _
			& "} else if (document.getElementById('adminValue').value == 2 && divID == 3) {" _
				& "document.getElementById('adminValue').value = 3;" _
			& "} else if (divID == 0) {" _
				& "document.getElementById('adminValue').value = 0;" _
				& "document.getElementById('errorString').innerText = 'LOGGED OUT';" _
				& "document.getElementById('adminText').style.visibility = 'hidden';" _
				& "document.getElementById('adminButton').style.visibility = 'hidden';" _
				& "document.getElementById('adminString').style.visibility = 'hidden';" _
				& "document.getElementById('operator').innerText = '';" _
			& "}" _
		& "}" _
		& "</script></head>"

	'Body Start String							
	LoadHTML = LoadHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
	
	'ADMIN Mode String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 60px; left: 550px; height: 30px; width: 200px; text-align: left;'>" _
		& "<span unselectable='on' id=adminText class='unselectable' style='visibility:hidden;text-align: center;'>ADMIN MODE</span></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 550px; height: 30px; width: 200px; text-align: left;'>" _
		& "<button id=logoutButton class=HTAButton style='height: 30px; width: 200px; text-align: center;' onclick='adminMode(0)'>Logout&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button id=adminButton class=HTAButton style='height: 30px; width: 30px; text-align: center;visibility:hidden;' onclick='crossOutButton()'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 60px; height: 30px; width: 440px; text-align: left;'>" _
		& "<span id=adminString unselectable='on' class='unselectable' style='visibility:hidden;'>Click to cross out part mark</span></div>"
		
	'Camera Listen String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=connectButton style='height: 30px; width: 30px; text-align: center;'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 60px; height: 30px; width: 440px; text-align: left;' id='connectionText'>&nbsp;</div>"
			
	'Part Mark Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 60px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=partMarkButton style='height: 30px; width: 30px; text-align: center;'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 60px; left: 60px; height: 30px; width: 440px; text-align: left;' id='partMarkText'>&nbsp;</div>"
		
	'Scanner Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=scannerButton style='height: 30px; width: 30px; text-align: center;'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 60px; height: 30px; width: 440px; text-align: left;' id='scannerText'>&nbsp;</div>"
		
	'Serial Number String
	LoadHTML = LoadHTML _
		& "<div style='top: 130px; left: 25px; height: 30px; width: 30px;'>" _
		& "<button class=HTAButton id=manualSerialNumber style='height: 30px; width: 30px; text-align: center;' onclick='manualButton()'>&nbsp;</button></div>" _
		& "<div style='top: 130px; left: 60px; height: 30px; width: 440px;' unselectable='on' class='unselectable'><span unselectable='on' class='unselectable' style='top: 0px; left: 0px; height: 30px; width: 440px;' id=SerialNumberText>Click button to type in serial number</span></div>" _
		& "<div style='top: 130px; left: 60px; height: 30px; width: 440px;'><input style='top: 0px; left: 0px; height: 30px; width: 440px; visibility:hidden;' value='' id=SerialNumberInput>&nbsp;</div>"
	
	'Camera Output String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 230px; left: 25px; height: 30px; width: 175px; text-align: right;'>Date&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 260px; left: 25px; height: 30px; width: 175px; text-align: right;'>Time&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 290px; left: 25px; height: 30px; width: 175px; text-align: right;'>Serial Number&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 320px; left: 25px; height: 30px; width: 175px; text-align: right;'>Width&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 350px; left: 25px; height: 30px; width: 175px; text-align: right;'>Height&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 380px; left: 25px; height: 30px; width: 175px; text-align: right;'>Angle&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 410px; left: 25px; height: 30px; width: 175px; text-align: right;'>Length&nbsp;</div>"
		
	'Camera 1 Output String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 200px; left: 200px; height: 30px; width: 270px; text-align: center;'>Left Camera</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 230px; left: 200px; height: 30px; width: 270px; text-align: center;' id=camera1date></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 260px; left: 200px; height: 30px; width: 270px; text-align: center;' id=camera1time></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 290px; left: 200px; height: 30px; width: 270px; text-align: center;' id=camera1SN></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 320px; left: 200px; height: 30px; width: 270px; text-align: center;' id=camera1width></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 350px; left: 200px; height: 30px; width: 270px; text-align: center;' id=camera1height></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 380px; left: 200px; height: 30px; width: 270px; text-align: center;' id=camera1angle></div>"
		
	'Camera 2 Output String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 200px; left: 475px; height: 30px; width: 270px; text-align: center;'>Right Camera</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 230px; left: 475px; height: 30px; width: 270px; text-align: center;' id=camera2date ></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 260px; left: 475px; height: 30px; width: 270px; text-align: center;' id=camera2time></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 290px; left: 475px; height: 30px; width: 270px; text-align: center;' id=camera2SN></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 320px; left: 475px; height: 30px; width: 270px; text-align: center;' id=camera2width></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 350px; left: 475px; height: 30px; width: 270px; text-align: center;' id=camera2height></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 380px; left: 475px; height: 30px; width: 270px; text-align: center;' id=camera2angle></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 410px; left: 475px; height: 30px; width: 270px; text-align: center;' id=camera2length></div>"
		
	'Operator String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 440px; left: 25px; height: 30px; width: 175px; text-align: right;'>Operator:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 440px; left: 200px; height: 30px; width: 270px; text-align: center;' id=operator></div>"
		
	'Handheld Output String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 470px; left: 25px; height: 30px; width: 175px; text-align: right;'>Handheld&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 470px; left: 200px; height: 30px; width: 270px; text-align: center;' id=handheld>&nbsp;</div>"
		
	'Blade Output String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 500px; left: 230px; height: 30px; width: 240px; text-align: center;'>060053-1:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 530px; left: 230px; height: 30px; width: 240px; text-align: center;' id=blade1string>Waiting...</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 500px; left: 505px; height: 30px; width: 240px; text-align: center;'>060053-2:</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 530px; left: 505px; height: 30px; width: 240px; text-align: center;' id=blade2string>Waiting...</div>"
		
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 570px; left: 25px; height: 100px; width: 775px; text-align: center;' id=errorString></div>" _
		& "<div unselectable='on' id=duplicateSlug value=false class='unselectable' style='top: 700; left: 25px; height: 150px; width: 300px; text-align: center;'>" _
		& "<button id=duplicateButton style='height: 150px; width: 300px;visibility:hidden;text-align: center;font:normal 28px Tahoma;' onclick='adminLogin()'><span unselectable='on' class='unselectable'>Load duplicate<br/>serial numbers<br/>to part marker</span></button></div>"

	'Port Parameters String
	LoadHTML = LoadHTML _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=portListen value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=portNumber><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>"
		
	'All Op String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 760px;height: 30px; width: 30px;'><button class='opButton' style='height: 30px; width: 30px;' onclick='done.value=""allOps""' id='OPSButton'>&#10010;</button></div>"
		
	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 795px;height: 30px; width: 30px;'><button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'><span unselectable='on' class='unselectable'>X</span></button></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=done value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=waiting4PM value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=accessNewEntry value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=allowDuplicate value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=crossOutMode value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=crossOutClick value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div style='top: 0px; left: 900px;'><input type=hidden id=adminCheck value=false><center><span unselectable='on' class='unselectable'>&nbsp;</span></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 0px; left: 0px;height: 30px; width: 30px;' onclick='adminMode(1)'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 905px; left: 0px;height: 30px; width: 30px;' onclick='adminMode(2)'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 905px; left: 815px;height: 30px; width: 30px;' onclick='adminMode(3)'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 0px; left: 900px;'><input type=hidden id=adminValue 		style='visibility:hidden;' value=0><center>&nbsp;</div>"
	
	'End Body String
	LoadHTML = LoadHTML _
		& "</body>"

 End Function
 