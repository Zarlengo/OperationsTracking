Option Explicit
 '********** VERSION HISTORY ************
 ' 1.0	2/15/2019	Initial Release
 '
 '*************** TO DO *****************

 '******* CHANGE THESE SETTINGS *********
 Const adminMode = false
 Const debugMode = false
 Const tabletPassword = "Fl0wSh0p17"
 Const computerPassword = "Snowball18!"
 '********* DATABASE SETTINGS ***********
 Const dataSource = "PRODSQLAPP01\PRODSQLAPP01"
 Const initialCatalog = "CMM_Repository" 
 '***************************************
 Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
 Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")

 Dim allOPSsource : allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\All Operations.vbs"
 Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """" & sArg
 

 Dim closeWindow : closeWindow = false
 Dim errorWindow : errorWindow = false
 Dim BypassMode : BypassMode = false
 
 Dim winsock0
 Dim strData, windowBox, AccessArray, AccessResult, fieldArray(4), fieldsBad
 Dim SendData, RecieveData, wmi, cProcesses, oProcess
 Dim machineBox, strSelection, RemoteHost, RemotePort, machineString

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

 If debugMode = False Then On Error Resume Next
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
	'objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "30", "REG_DWORD"	'Changes TCP timeout settings if needing to restart program w/in 5 minutes
 On Error Goto 0

 'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
 Set wmi = GetObject("winmgmts:root\cimv2") 
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
 For Each oProcess in cProcesses
	oProcess.Terminate()
 Next

    

 '// CREATE WINSOCK: 0 - QA Scabber
 Set winsock0 = Wscript.CreateObject("OSWINSCK.Winsock", "winsock0_")
 If Err.Number <> 0 Then
    MsgBox "Winsock Object Error!" & vbCrLf & "Script will exit now."
    WScript.Quit
 End If

 If Not WScript.Arguments.Count = 0 Then
	sArg = ""
	For Each Arg In Wscript.Arguments
		If Arg = "BYPASS" Then
			BypassMode = True
		Else
			sArg = Arg
		End If
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

 'Calls function to create ie window
 set windowBox = HTABox("white", 770, 1280, 0, 0) 

 with windowBox
	.document.title = "MRB Inventory Location"
	
	'Function to check for access connection and load info from database
	AccessResult = Load_Access
	checkAccess
	'Connects to the scanner
	connect2Scanner
	Load_Location
	
	do until closeWindow = true													'Run loop until conditions are met
		On Error Resume Next
		do until .done.value = "cancel" or .done.value = "access" or .done.value = "scanner" or .done.value = "MRB" or .submitButton.value = "true" or .done.value = "allOps"
			wsh.sleep 50
			If .done.value = true Then
				wsh.quit
			End If
		loop
		On Error GoTo 0
		if .done.value = "cancel" then											'If the x button is clicked
			closeWindow = true													'Variable to end loop
		ElseIf .done.value = "access" then
			.done.value = false
			windowBox.accessText.innerText = "Retrying connection."
			windowBox.accessButton.style.backgroundcolor = "orange"
			AccessResult = Load_Access
			checkAccess
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
			HTABox.document.title = "HTABox" 
			HTABox.document.write LoadHTML(sBgColor)
			Exit Function 
		End If 
	Next 
	MsgBox "HTA window not found." 
	wsh.quit
 End Function

Function connect2Scanner()
	Dim testInput
	Dim secs : secs = 0
	If machineString <> "Manual" and machineString <> "" Then
		windowBox.scannerText.innerText = "Connect to " & machineString
		windowBox.scannerButton.style.backgroundcolor = "orange"
		windowBox.scannerButton.disabled = true
		windowBox.errorString.innerText = ""
	End If
	' loads port settings into winsock
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
	ElseIf winsock0.state = sckConnected Then 
		windowBox.scannerText.innerText = "Connected to " & machineString
		windowBox.scannerButton.style.backgroundcolor = "limegreen"
		windowBox.scannerButton.disabled = true
	Else
		windowBox.scannerText.innerText = "Error: " & machineString
		windowBox.scannerButton.style.backgroundcolor = "red"
		windowBox.scannerButton.disabled = false
	End If
 End Function

Function checkAccess()
	If AccessResult = false Then
		windowBox.accessText.innerText = "Database not loaded"
		windowBox.accessButton.style.backgroundcolor = "red"
	Else
		windowBox.accessText.innerText = "Database connection successful"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
	End If
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
	Dim inputString : inputString = TrimString(stringFromScanner)
	windowbox.errorDiv.style.background = ""
	windowBox.errorString.innerText = ""
	If inputString = "" or inputString = tabletPassword or inputString = computerPassword Then
	ElseIf inputString = "AccessRetry" Then
		windowBox.done.value = "access"
	ElseIf inputString = "Cancel" Then
		windowBox.done.value = "cancel"
	ElseIf Left(inputString, 4) = "COMP" Then
		machineString = inputString
		sArg = """" & inputString & """"
		RemoteHost = ""
		RemotePort = ""
		Load_IP
		connect2Scanner
	ElseIF Len(inputString) = 10 and Mid(inputString, 9, 1) = "-" and (Left(inputString, 1) = "D" or Left(inputString, 1) = "H") Then
		If windowBox.LocationDiv.innerText = "" Then
			windowBox.errorString.innerText = "Missing location, please scan the shelf ID"
			windowBox.errorDiv.style.backgroundColor = "red"
			Exit Function
		End If
		LoadSNtoAccess(inputString)
	Else
		windowBox.LocationDiv.innerText = inputString
		Dim sqlString : sqlString = "Select COUNT(*) From [40_MRB] WHERE [Location] = '" & inputString & "';"
		Dim objCmd : set objCmd = GetNewConnection
		On Error GoTo 0
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
			windowbox.errorDiv.style.background = ""
		End If
		Dim rs : Set rs = objCmd.Execute(sqlString)	
		If rs(0).value <> 0 Then
			windowBox.locationCount.innerText = rs(0).value
		End If
	End If
 End Function

Function LoadSNtoAccess(serialNumber)
	Dim strQuery, CurrentTime, Operator, strQueryPre, sqlString, rs, Duplicate, SCFound, POID, ShipDate, PalletID, BoxID, alreadyInSQL
	Dim objCmd : set objCmd = GetNewConnection
	Dim ErrorFound : ErrorFound = False
	Dim location : location = windowBox.LocationDiv.innerText
	
	On Error GoTo 0
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
		windowbox.errorDiv.style.background = ""
	End If
	
	If location = "Remove" Then
		sqlString = "Select COUNT(*) From [40_MRB] WHERE [Serial Number] = '" & serialNumber & "';"
		set rs = objCmd.Execute(sqlString)	
		If rs(0).value = 0 Then
			windowBox.errorString.innerText = "Serial number not in MRB: " & serialNumber
			windowbox.errorDiv.style.background = "red"
			Exit Function
		End If
		alreadyInSQL = true
	ElseIf windowBox.saveBlade.value <> serialNumber Then
		windowBox.saveBlade.value = false
		If BypassMode <> True Then
			sqlString = "Select [Location] From [40_MRB] WHERE [Serial Number] = '" & serialNumber & "' and [Location] IS NOT NULL;"
			set rs = objCmd.Execute(sqlString)		
			DO WHILE NOT rs.EOF
				windowBox.errorString.innerText = serialNumber & " already in " & rs.Fields(0) & chr(10) & "Scan again to overwrite location"
				windowbox.errorDiv.style.background = "red"
				windowBox.saveBlade.value = serialNumber
				objCmd.Close
				Set objCmd = Nothing
				Exit Function
				rs.MoveNext
			Loop	
		End If
		sqlString = "Select COUNT(*) From [40_MRB] WHERE [Serial Number] = '" & serialNumber & "';"
		set rs = objCmd.Execute(sqlString)	
		If rs(0).value <> 0 Then
			alreadyInSQL = true
		End If
	Else
		windowBox.saveBlade.value = false
		alreadyInSQL = true
	End If
	
	If alreadyInSQL = true and location = "Remove" then
		strQuery = "UPDATE [40_MRB] SET [Location] = NULL, [DateRemoved] ='" & now & "' WHERE [Serial Number]= '" & serialNumber & "';"
	ElseIf alreadyInSQL = true then
		strQuery = "UPDATE [40_MRB] SET [Location] = '" & location & "' WHERE  [Serial Number]='" & serialNumber & "';"
	Else
		strQueryPre = "INSERT INTO [40_MRB] ([Serial Number], [Location], [DateAdded]) "
		strQuery = strQueryPre & "VALUES ('" & serialNumber & "', '" & location & "', '" & now & "');"
	End If
	objCmd.Execute(strQuery)
	windowBox.errorString.innerText = "S.N. scan successful: " & serialNumber
	windowbox.errorDiv.style.background = "limegreen"
	
	sqlString = "Select COUNT(*) From [40_MRB] WHERE [Location] = '" & location & "';"
	set rs = objCmd.Execute(sqlString)	
	If rs(0).value <> 0 Then
		windowBox.locationCount.innerText = rs(0).value
	End If
	
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
 End Function

Function GetNewConnection()
	Dim objCmd : Set objCmd = CreateObject("ADODB.Connection")
	Dim sConnection : sConnection = "Data Source=" & dataSource & ";Initial Catalog=" & InitialCatalog & ";Integrated Security=SSPI;"
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
	Dim objCmd : set objCmd = GetNewConnection
	If objCmd is Nothing Then Load_Access = false : Exit Function
	objCmd.Close
	Set objCmd = Nothing
	Load_Access = true
 End Function

Function Load_Location()
	Dim objCmd : set objCmd = GetNewConnection
	Dim sqlString, rs, SN, Location, tableStringPre, tableString, tableStringSuf, removeButtonString
	tableStringPre = "<table id='POTable'><thead><tr><th><span class='text'>&nbsp; Serial Number &nbsp;</span></th><th><span class='text'>&nbsp; Location &nbsp;</span></th></tr></thead><tbody>"
	tableStringSuf = "</tbody></table>"
	removeButtonString = ""
	If objCmd is Nothing Then
		windowbox.table_wrapper.innerHTML = tableStringPre & "<tr><td>ERROR LOADING</td><td></td><td></td></tr>" & tableStringSuf
		Exit Function
	End If
	sqlString = "Select [Serial Number], [Location] From [40_MRB] Where [Location] IS NOT NULL;"
	set rs = objCmd.Execute(sqlString)
	rs.Sort="[Location], [Serial Number]"
	DO WHILE NOT rs.EOF
		SN = rs.Fields(0)
		Location = rs.Fields(1)
		If InStr(1, UCase(Location), "SCRAP") <> 0 Then
			tableString = tableString & "<tr class='Scrap' style='display:none;'>"
		Else
			tableString = tableString & "<tr class='MRB'>"
		End If
		tableString = tableString & "<td>" & SN & "</td><td style='text-align: center;'>" & Location & "</td></tr>"
		rs.MoveNext
	Loop
	objCmd.Close
	Set objCmd = Nothing
	windowbox.table_wrapper.innerHTML = tableStringPre & tableString & tableStringSuf
 End Function

Function Load_IP()
	Dim sqlString, rs
	Dim objCmd : set objCmd = GetNewConnection
	On Error GoTo 0
	If objCmd is Nothing Then Exit Function
	sqlString = "Select [IPAddress], [Port] From [00_Machine_IP] WHERE [DeviceType] = 'CognexBTHandheld' AND [MachineName] = '" & sArg & "';"
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
    WScript.Sleep 1000
	Check_String strData
 End Sub


'// WINSOCK ERROR
Sub winsock0_OnError(Number, Description, SCode, Source, HelpFile, HelpContext, CancelDisplay)
	windowBox.scannerText.innerText = "Error: " & machineString
	windowBox.scannerButton.style.backgroundcolor = "red"
	windowBox.scannerButton.disabled = false
    windowBox.errorString.innerText = "Scanner Error: " & Number & vbCrLf & Description
 End Sub

'// EXIT SCRIPT
Sub ServerClose()
	If debugMode = False Then On Error Resume Next

	WScript.Sleep 1000  '// REQUIRED OR ERRORS
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	'objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"

	If winsock0.state <> sckClosed Then winsock0.Disconnect
    winsock0.CloseWinsock
    Set winsock0 = Nothing
	
	windowBox.close
	
	On Error GoTo 0
    Wscript.Quit
 End Sub


'Function to create all of the JS and HTML code for the window
Function LoadHTML(sBgColor)
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
		& ".locationText {" _
			& "font: 50px;" _
			& "}" _
		& "#commentModal, #nameModal {" _
			& "font:normal 30px Tahoma;" _
			& "background-color = 'grey';" _
			& "}" _
		& "#eTagModal {" _
			& "font:normal 30px Tahoma;" _
			& "background-color = 'blue';" _
			& "}" _
		& "#table_wrapper, #e_table_wrapper {" _
			& "width:100%;" _
			& "height:100%;" _
			& "overflow:auto; " _
			& "}" _
		& "#table_wrapper table, #e_table_wrapper table {" _
			& "margin-right: 20px;" _
			& "border-collapse: collapse;" _
			& "}" _
		& "#SNIDs {" _
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
				& "document.getElementById('errorString').innerText = '';" _
			& "} else {" _
				& "document.getElementById('manualSerialNumber').style.backgroundColor = 'DimGrey';" _
				& "document.getElementById('SerialNumberText').style.visibility = 'hidden';" _
				& "document.getElementById('inputFormDiv').style.visibility = 'visible';" _
				& "document.getElementById('inputForm').disabled = false;" _
				& "document.getElementById('stringInput').disabled = false;" _
			& "}" _
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
		& "function showScrap() {" _ 
			& "if (document.getElementById('scrapText').innerText == 'Click to show scrap') {" _
				& "document.getElementById('scrapText').innerText = 'Click to show MRB';" _
				& "toggle('Scrap', 'inline');" _
				& "toggle('MRB', 'none');" _
			& "} else {" _
				& "document.getElementById('scrapText').innerText = 'Click to show scrap';" _
				& "toggle('Scrap', 'none');" _
				& "toggle('MRB', 'inline');" _
			& "}" _
		& "}" _
		& "function toggle(className, displayState){" _
			& "var elements = getElementsByClassName(document.body, className);" _
			& "for (var i = 0; i < elements.length; i++){" _
				& "elements[i].style.display = displayState;" _
			& "}" _
		& "};" _
		& "function getElementsByClassName(node, classname) {" _
			& "var a = [];" _
			& "var re = new RegExp('(^| )'+classname+'( |$)');" _
			& "var els = node.getElementsByTagName('*');" _
			& "for(var i=0,j=els.length; i<j; i++)" _
				& "if(re.test(els[i].className))a.push(els[i]);" _
			& "return a;" _
		& "}" _
		& "</script></head>"

	'Body Start String							
	LoadHTML = LoadHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
	
	'Access Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=accessButton style='height: 30px; width: 30px; text-align: center;background-color:orange;' disabled onclick='done.value=""access""'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 60px; height: 30px; width: 480px; text-align: left;' id='accessText'>Waiting for database connection&nbsp;</div>"
		
	'Scanner Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 60px; left: 25px;height: 30px; width: 30px;'>" _
		& "<button id=scannerButton style='height: 30px; width: 30px;background-color:orange;' disabled onclick='done.value=""scanner""'>&nbsp;</button></div>" _
		& "<div id=scannerText unselectable='on' class='unselectable' style='top: 60px; left: 60px;height: 30px; width: 480px;'>Waiting for scanner connection&nbsp;</div>" 
		
	'Scrap Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button style='height: 30px; width: 30px;' onclick='showScrap()' id='scrapButton'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 60px;height: 30px; width: 480px;' id='scrapText'>Click to show scrap</div>" 
		
	'Input String
	LoadHTML = LoadHTML _
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id='manualSerialNumber' style='height: 30px; width: 30px;' onclick='manualButton()'>&nbsp;</button></div>" _
		& "<div id='SerialNumberText' unselectable='on' class='unselectable' style='top: 130px; left: 60px;height: 30px; width: 480px;'>Click to enter data manually&nbsp;</div>" _
		& "<div id='inputFormDiv' style='top: 130px; left: 60px; height: 30px; width: 480px;visibility:hidden;'>" _
			& "<form id=inputForm onsubmit='inputComplete();' disabled>" _
				& "<input id=stringInput style='top: 0px; left: 0px; height: 30px; width: 480px;' value='' disabled /></form></div>"
	
	'Location String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable locationText' style='top: 300px; left: 25px;height: 60px; width: 200px;text-align: right;'>Location:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable locationText' style='top: 300px; left: 255px;height: 60px; width: 285px;' id=LocationDiv></div>" 
		
	'Location Count String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable locationText' style='top: 230px; left: 25px;height: 60px; width: 200px;text-align: right;'>Count:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable locationText' style='top: 230px; left: 255px;height: 60px; width: 285px;' id=locationCount></div>" 
		
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div id=errorDiv style='top: 475; left: 0px; height: 265px; width: 600px;'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 485px; left: 25px; height: 240px; width: 525px; text-align: center;' id=errorString></div>"
		
	'Modal MRB Div String
	LoadHTML = LoadHTML _
		& "<div id='commentModal' style='top: 1px; left: 576px; height: 778px; width: 575px;'>" _
		& "<div style='top: 50px; left: 50px; height: 550px; width: 430px;' id='table_wrapper'></div>" _		
		& "</div>"
		
	'All Op String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 505px;height: 30px; width: 30px;'><button class='opButton' style='height: 30px; width: 30px;' onclick='done.value=""allOps""'>&#10010;</button></div>"
		
	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 545px;height: 30px; width: 30px;'><button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'>X</button></div>" _
		& "<div><button type=hidden id=returnToHTA 		style='visibility:hidden;' value=false onclick='HTAReturn()'><center>&nbsp;</button></div>" _
		& "<div><input type=hidden id=done 				style='visibility:hidden;' value=false><center>&nbsp;</div>" _
		& "<div><input type=hidden id=submitText		style='visibility:hidden;' value=false><center>&nbsp;</div>" _
		& "<div><input type=hidden id=submitButton		style='visibility:hidden;' value=false><center>&nbsp;</div>" _
		& "<div><input type=hidden id=saveBlade			style='visibility:hidden;' value=false><center>&nbsp;</div>"
		
	'End Body String
	LoadHTML = LoadHTML _
		& "</body>"

	'Footer String
	LoadHTML = LoadHTML _
		& "<footer><script language='javascript'>" _
			& "document.getElementById('stringInput').focus();" _
		& "</script></footer>"

 End Function
 