Option Explicit

'********* VERSION HISTORY ************
' 1.0	8/2/2018	Initial Release for production
' 1.1	9/17/2018	Added email feature to notify production to create work orders and load slugs into AX
' 1.2	11/19/2018	Change email feature to write txt file (email does not work on tablet accounts)
' 1.3	5/30/2019	Changed email txt file to SQL update for 00_Invoice [Received] = true
' 1.4	6/20/2019	Added GfE slug information
'
'************** TO DO *****************
' Edit mode for existing data
' E-Tag mode
' Close input window when in scanner mode
'****** CHANGE THESE SETTINGS *********
Dim adminMode : adminMode = false
Dim debugMode : debugMode = false
'***************************************

Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
Dim objShell : Set objShell = CreateObject("WScript.Shell")
Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")

Dim allOPSsource : allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\All Operations.vbs"
Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """" & sArg
Dim dataSource : dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"
Dim adminPassword : adminPassword = "FLOW288"
Dim tabletPassword : tabletPassword = "Fl0wSh0p17"

Dim closeWindow : closeWindow = false
Dim errorWindow : errorWindow = false
Dim notFoundCount : notFoundCount = 0

Dim winsock0
Dim strData, windowBox, AccessArray, AccessResult
Dim SendData, RecieveData, wmi, cProcesses, oProcess
Dim machineBox, strSelection, RemoteHost, RemotePort, machineString


'*********************************************************
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
objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "30", "REG_DWORD"	'Changes TCP timeout settings if needing to restart program w/in 5 minutes
On Error Goto 0

'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
Set wmi = GetObject("winmgmts:root\cimv2") 
Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
For Each oProcess in cProcesses
	oProcess.Terminate()
Next

    

'// CREATE WINSOCK: 0 - QA Scanner
Set winsock0 = Wscript.CreateObject("OSWINSCK.Winsock", "winsock0_")
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
'Function to check for access connection and load info from database
AccessResult = Load_Access

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
set windowBox = HTABox("white", 780, 600, 300, 0) 

with windowBox
	.document.title = "Operation 00"
	
	checkAccess
	
	'Connects to the scanner
	connect2Scanner
	
	'.document.accessText.focus
	'.document.accessText.select
	do until closeWindow = true													'Run loop until conditions are met
		do until .done.value = "cancel" or .done.value = "access" or .done.value = "scanner" or .submitButton.value = "true" or .done.value = "allOps"
			wsh.sleep 50
			On Error Resume Next
			If .done.value = true Then
				wsh.quit
			End If
			On Error GoTo 0
		loop
		if .done.value = "cancel" then											'If the x button is clicked
			closeWindow = true													'Variable to end loop
		ElseIf .done.value = "access" then
			.done.value = false
			windowBox.accessText.innerText = "Retrying connection."
			windowBox.accessButton.style.backgroundcolor = "orange"
			If windowbox.SlugID.innerText <> "" and windowbox.operatorID.innerText <> "" Then
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
		windowBox.partMarkText.innerText = "Connect to " & machineString
		windowBox.partMarkButton.style.backgroundcolor = "orange"
		windowBox.partMarkButton.disabled = true
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
	ElseIf winsock0.state = sckConnected Then 
		windowBox.partMarkText.innerText = "Connected to " & machineString
		windowBox.partMarkButton.style.backgroundcolor = "limegreen"
		windowBox.partMarkButton.disabled = true
	Else
		windowBox.partMarkText.innerText = "Error: " & machineString
		windowBox.partMarkButton.style.backgroundcolor = "red"
		windowBox.partMarkButton.disabled = false
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
	Dim inputString
	
	inputString = TrimString(stringFromScanner)
	windowbox.errorDiv.style.background = ""
	windowBox.errorString.innerText = ""
	windowbox.submitText.value = ""
	If inputString = tabletPassword Then
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
	ElseIf Left(inputString, 4) = "QA_" Then
		machineString = inputString
		sArg = """" & inputString & """"
		RemoteHost = ""
		RemotePort = ""
		Load_IP
		connect2Scanner
		'D_TESTSN-0
	ElseIF Left(inputString, 1) = "D" and ((Len(inputString) = 10 and Mid(inputString, 9, 1) = "-") or (Len(inputString) = 13 and Mid(inputString, 9, 1) = "M")) Then
		If inputString = windowbox.SlugID.innerText Then
			windowBox.errorString.innerText = windowBox.errorString.innerText & "Serial number scanned: " & Chr(13) & inputString
			CleanUpScreen
			windowbox.errorDiv.style.background = "red"
			Exit Function
		Else
			windowbox.SlugID.innerText = inputString
		End If
	ElseIf Left(inputString, 4) = "AEFL" Then
		Exit Function
	Else
		windowbox.operatorID.innerText = inputString
	End If	
	If windowbox.SlugID.innerText <> "" and windowbox.operatorID.innerText <> "" Then
		LoadSNtoAccess
	End if
End Function

Function Logout()
	windowBox.operatorID.innerText = ""
	windowBox.errorString.innerText = "Logged Out"
	windowBox.accessButton.disabled = true
	CleanUpScreen
End Function

Function LoadSNtoAccess()
	Dim strQuery, CurrentTime, Inspector, SlugID, strQueryPre, sqlString, rs, Duplicate, SCFound, Comments, SCCnt
	Dim objCmd : set objCmd = GetNewConnection
		
	Inspector = windowBox.operatorID.innerText
	SlugID = windowBox.SlugID.innerText
	CurrentTime = Now
	Comments = windowBox.commentTextSave.Value
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
	
	sqlString = "SELECT [Slug S/N] FROM [00_Initial] WHERE [Slug S/N]='" & SlugID & "';"
	
	set rs = objCmd.Execute(sqlString)		
	DO WHILE NOT rs.EOF
		Duplicate = rs.Fields(0)
		rs.MoveNext
	Loop	
	Set rs = Nothing
	If Duplicate <> "" Then
		windowBox.errorString.innerText = "Serial number alread scanned: " & SlugID
		windowbox.errorDiv.style.background = "red"
		windowBox.SlugID.innerText = ""
	Else
		strQueryPre = "INSERT INTO [00_Initial] ([Slug S/N], [Slug Inspection Date], [Inc Insp Inspector Last Name], [Comments]) "
		strQuery = strQueryPre & "VALUES ('" & SlugID & "', '" & CurrentTime & "', '" & Inspector & "', '" & Comments & "'); "
		objCmd.Execute(strQuery)
		windowBox.errorString.innerText = "S.N. scan successful: " & SlugID
		windowbox.errorDiv.style.background = "limegreen"
	End If
	sqlString = "SELECT TOP 1 [Invoice Number] FROM [00_AE_SN_Control] WHERE [Slug Serial Number]='" & SlugID & "';"
	set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		SCFound = rs.Fields(0)
		rs.MoveNext
	Loop
	Set rs = Nothing
	If SCFound = "" Then
		windowBox.errorString.innerText = windowBox.errorString.innerText & vbCrLf & vbCrLf & "ERROR:" & vbCrLf & "Slug SN not in database"
		windowbox.errorDiv.style.background = "red"
		notFoundCount = notFoundCount + 1
		windowBox.notFoundCnt.InnerHTML = notFoundCount
		windowBox.notFoundText.style.visibility = "visible"
	Else
		windowBox.InvoiceID.innerHTML = SCFound
		sqlString = "SELECT COUNT([Invoice Number]) FROM [00_AE_SN_Control] WHERE [Invoice Number]='" & SCFound & "';"
		set rs = objCmd.Execute(sqlString)
		SCCnt = rs.Fields(0) / 2
		Set rs = Nothing
		windowbox.InvoiceCnt.innerHTML = SCCnt
		strQuery = "SELECT DISTINCT [00_AE_SN_Control].[Invoice Number], [00_AE_SN_Control].[Slug Serial Number]"
		strQuery = strQuery + "FROM [00_Initial] INNER JOIN [00_AE_SN_Control] ON [00_Initial].[Slug S/N] = [00_AE_SN_Control].[Slug Serial Number]"
		strQuery = strQuery + "GROUP BY [00_AE_SN_Control].[Invoice Number], [00_AE_SN_Control].[Slug Serial Number]"
		strQuery = strQuery + "HAVING ((([00_AE_SN_Control].[Invoice Number])='" & SCFound & "'));"
		set rs = objCmd.Execute(strQuery)
		windowBox.SlugCnt.innerText = rs.RecordCount
		If rs.RecordCount = SCCnt Then
			Set rs = Nothing
			sqlString = "UPDATE [00_Invoice] SET [Received]='1' WHERE  [Invoice Number]='" & SCFound & "';"
			set rs = objCmd.Execute(strQuery)
			Send_Email(SCFound)
			windowBox.errorString.innerText = windowBox.errorString.innerText & chr(10) & "Invoice completely received."
		End If
		Set rs = Nothing
	End If
	
	objCmd.Close
	Set objCmd = Nothing
	CleanUpScreen

End Function

Function CleanUpScreen()
	windowBox.SlugID.innerText = ""	
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

Sub Send_Email(Message)
	Dim objFSO : Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim fileName : fileName = "Operation00_" & Year(date) & Month(date) & Day(date) & "_" & Int(Timer()) & ".txt"
	
	' How to write file
	Dim outFile : outFile="G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\TXTFiles\" & fileName
	Dim objFile : Set objFile = objFSO.CreateTextFile(outFile,True)
	objFile.Write Message
	objFile.Close
End Sub

'// WINSOCK DATA ARRIVES
Sub winsock0_OnDataArrival(bytesTotal)
    winsock0.GetData strData, vbString
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

'// EXIT SCRIPT
Sub ServerClose()

	If debugMode = False Then On Error Resume Next

	WScript.Sleep 1000  '// REQUIRED OR ERRORS
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"

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
		& ".opButton {" _
			& "background-color: blue;" _
			& "height: 30px;" _
			& "width: 30px;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
			& "color: white;" _
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
				& "document.getElementById('footID').focus({preventScroll:false});" _
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
			& "document.getElementById('SlugID').innerText = '';" _
			& "document.getElementById('errorDiv').style.background = '';" _
			& "document.getElementById('accessButton').disabled = true;" _
			& "document.getElementById('errorString').innerText = 'Logged Out';" _
			& "document.getElementById('SlugCnt').innerText = 0;" _
			& "document.getElementById('logoutButton').disabled = true;" _
			& "document.getElementById('logoutButton').disabled = false;" _
		& "}" _
		& "function resetFunction() {" _
			& "document.getElementById('SlugID').innerText = '';" _
			& "document.getElementById('errorDiv').style.background = '';" _
			& "document.getElementById('accessButton').disabled = true;" _
			& "document.getElementById('errorString').innerText = 'Fields Reset';" _
			& "document.getElementById('SlugCnt').innerText = 0;" _
			& "document.getElementById('resetButton').disabled = true;" _
			& "document.getElementById('resetButton').disabled = false;" _
		& "}" _
		& "function commentFunction() {" _
			& "document.getElementById('commentModal').style.visibility = 'visible';" _
		& "};" _
		& "function okComment() {" _
			& "document.getElementById('commentModal').style.visibility = 'hidden';" _
			& "commentValue = document.getElementById('commentText').firstChild.value.replace('\'', '\'\'');" _
			& "document.getElementById('commentTextSave').value = commentValue;" _
			& "if (commentValue == '') {" _ 
				& "document.getElementById('commentButton').style.backgroundColor  = '';" _ 
				& "document.getElementById('commentButtonText').innerHTML  = 'Click to add a comment&nbsp;';" _ 
			& "} else {" _
				& "document.getElementById('commentButton').style.backgroundColor  = 'limegreen';" _ 
				& "document.getElementById('commentButtonText').innerHTML  = 'Click to edit comment&nbsp;';" _ 
			& "}" _
		& "};" _
		& "function cancelComment() {" _
			& "document.getElementById('commentModal').style.visibility = 'hidden';" _
			& "document.getElementById('commentText').firstChild.value = document.getElementById('commentTextSave').value;" _
		& "};" _
		& "function keepCommentFunction() {" _
			& "if (document.getElementById('keepCommentSave').value == 'False') {" _
				& "document.getElementById('keepCommentButton').style.backgroundColor  = 'limegreen';" _ 
				& "document.getElementById('keepCommentSave').value = 'True';" _
			& "} else {" _
				& "document.getElementById('keepCommentButton').style.backgroundColor  = '';" _ 
				& "document.getElementById('keepCommentSave').value = 'False';" _
			& "}" _
		& "};" _
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
		& "<button id=partMarkButton style='height: 30px; width: 30px;background-color:orange;' disabled onclick='done.value=""scanner""'>&nbsp;</button></div>" _
		& "<div id=partMarkText unselectable='on' class='unselectable' style='top: 60px; left: 60px;height: 30px; width: 480px;'>Waiting for scanner connection&nbsp;</div>" 
		
	'Reset Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id=resetButton style='height: 30px; width: 30px;' onclick='resetFunction()'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 95px; left: 60px;height: 30px; width: 480px;'>Click to reset fields&nbsp;</div>" 
		
	'Logout Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id=logoutButton style='height: 30px; width: 30px;' onclick='logoutFunction()'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 60px;height: 30px; width: 480px;'>Click to logout&nbsp;</div>" 
		
	'Comment Button String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id='commentButton' style='height: 30px; width: 30px;' onclick='commentFunction()'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 165px; left: 60px;height: 30px; width: 480px;' id='commentButtonText'>Click to add a comment&nbsp;"  _
		& "<input type=hidden id='commentTextSave' value=''>" _
		& "<input type=hidden id='keepCommentSave' value='False'></div>"
				
	'Input String
	LoadHTML = LoadHTML _
		& "<div unselectable='on' class='unselectable' style='top: 200px; left: 25px;height: 30px; width: 30px;'>" _
			& "<button id='manualSerialNumber' style='height: 30px; width: 30px;' onclick='manualButton()'>&nbsp;</button></div>" _
		& "<div id='SerialNumberText' unselectable='on' class='unselectable' style='top: 200px; left: 60px;height: 30px; width: 480px;'>Click to enter data manually&nbsp;</div>" _
		& "<div id='inputFormDiv' style='top: 200px; left: 60px; height: 30px; width: 480px;visibility:hidden;'>" _
			& "<form id=inputForm onsubmit='inputComplete();' disabled>" _
				& "<input id=stringInput style='top: 0px; left: 0px; height: 30px; width: 480px;' value='' disabled /></form></div>"
	
	'Inspector String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 265px; left: 25px; height: 30px; width: 175px; text-align: right;'>Inspector:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 265px; left: 200px; height: 30px; width: 375px; text-align: center;' id=operatorID></div>"
	
	'Slug S/N String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 295px; left: 25px; height: 30px; width: 175px; text-align: right;'>Slug S/N:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 295px; left: 200px; height: 30px; width: 375px; text-align: center;' id=SlugID></div>"
		
	'Invoice Number String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 325px; left: 25px; height: 30px; width: 175px; text-align: right;'>Invoice:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 325px; left: 200px; height: 30px; width: 375px; text-align: center;' id=InvoiceID></div>"
		
	'Slug Count String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 355px; left: 25px; height: 30px; width: 175px; text-align: right;'>Slug Count:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 355px; left: 200px; height: 30px; width: 140px; text-align: right;' id=SlugCnt>0</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 355px; left: 340px; height: 30px; width: 70px; text-align: center;'>&nbsp; of &nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 355px; left: 410px; height: 30px; width: 140px;' id=InvoiceCnt>0</div>" _
		
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div id=errorDiv style='top: 405px; left: 0px; height: 355px; width: 600px;'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 430px; left: 50px; height: 250px; width: 500px; text-align: center;' id=errorString></div>"
		
	'Bad Scan String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 690px; left: 25px; height: 30px; width: 275px; text-align: right;visibility:hidden;' id=notFoundText>S/N not found count:&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 690px; left: 300px; height: 30px; width: 275px;' id=notFoundCnt></div>" _
	
	'All Op String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 505px;height: 30px; width: 30px;'><button class='opButton' style='height: 30px; width: 30px;' onclick='done.value=""allOps""'>&#10010;</button></div>"
		
	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 545px;height: 30px; width: 30px;'><button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'>X</button></div>" _
		& "<div><button type=hidden id=returnToHTA 		style='visibility:hidden;' value=false onclick='HTAReturn()'><center>&nbsp;</button></div>" _
		& "<div><input type=hidden id=done 				style='visibility:hidden;' value=false><center>&nbsp;</div>" _
		& "<div><input type=hidden id=submitButton 		style='visibility:hidden;' value=false><center>&nbsp;</div>" _
		& "<div><input type=hidden id=submitText 		style='visibility:hidden;' value=false><center>&nbsp;</div>" 
		
	'Modal Comment Div String
	LoadHTML = LoadHTML _
		& "<div id='commentModal' style='top: 1px; left: 1px; height: 778px; width: 598px;'>" _
		& "<div unselectable='on' class='unselectable' style='top: 50px; left: 50px; height: 80px; width: 498px;'>Enter any comments<br>(can leave blank)</div>" _
		& "<div style='top: 150px; left: 50px; height: 400px; width: 498px;' id='commentText'><input type='text' style='height: 400px; width: 498px;'></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 600px; left: 50px;height: 30px; width: 30px;'>" _
			& "<button id='keepCommentButton' style='height: 30px; width: 30px;' onclick='keepCommentFunction()'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 600px; left: 85px;height: 30px; width: 463px; text-align: left; font:normal 18px Tahoma;'>Click to keep comment for additional slugs&nbsp;</div>" _
		& "<div style='top: 650px; left: 50px; height: 48px; width: 100px;'><input type=button value='Ok' 	   style='height: 48px; width: 100px;' onclick='okComment()'></div>" _
		& "<div style='top: 650px; left: 450px; height: 48px; width: 100px;'><input type=button value='Cancel' style='height: 48px; width: 100px;' onclick='cancelComment()'></div>" _
		& "</div>"
		
	'End Body String
	LoadHTML = LoadHTML _
		& "</body>"

	'Footer String
	LoadHTML = LoadHTML _
		& "<footer id=footID><script language='javascript'>" _
			& "document.getElementById('stringInput').focus();" _
		& "</script></footer>"

End Function