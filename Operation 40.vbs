Dim THIS_IS_A_PLACEHOLDER
 'OPTION EXPLICIT
 'Version 3	8/29/2019	Updating to CMM Revision E

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
 Dim searchTime : searchTime = 10 / 60 / 60 / 24
 Dim logoutTime : logoutTime = 45 / 60 / 24

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


 Dim adjHeight
 Dim closeWindow : closeWindow = false
 Dim CMMID
 'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
 Set wmi = GetObject("winmgmts:root\cimv2") 
 Dim cProcesses : Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
 Dim oProcesses : For Each oProcess in cProcesses
	oProcess.Terminate()
 Next

 Set Processes = GetObject("winmgmts:").InstancesOf("Win32_Process")

 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%PCDLRN.exe%'") 
 Dim PCDMIS_ID : For Each oProcesses In cProcesses
    PCDMIS_ID = oProcesses.ProcessId
	Exit For
 Next

 If Not WScript.Arguments.Count = 0 Then
	sArg = ""
	For Each Arg In Wscript.Arguments
		If InStr(1, Arg, "CMM") > 0 Then
			 CMMID = Arg
		Else
		  sArg = sArg & " " & """" & Arg & """"
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
	machineString = "No Scanner"
 End If
 
 Dim WshShell : Set WshShell = WScript.CreateObject("WScript.Shell")
 If Left(machineString, 3) = "COM" Then
	Dim objComport : Set objComport = CreateObject("AxSerial.ComPort")      ' Create new instance
	objComport.Clear()
	objComport.LicenseKey = "FD2C1-DC93A-6BFBF"
	objComport.Device = machineString
	objComport.BaudRate  = 112500
	objComport.ComTimeout = 1000  ' Timeout after 1000msecs 
 End If
 
 Dim startTime
 AccessResult = Load_Access
 set windowBox = HTABox("white") : with windowBox
	Call checkAccess
	Call connect2Scanner				'Connects to the scanner			
	do until closeWindow = true													'Run loop until conditions are met
		startTime = now
		logoutStart = now
		do while .done.value = false  or UCase(.done.value) = "FALSE"
			wsh.sleep 50
			On Error Resume Next
			If .done.value = true Then
				wsh.quit
			End If
			On Error GoTo 0
			If Left(machineString, 3) = "COM" Then ReadResponse(objComport)
			If startTime + searchTime < now Then
				Call Update_Final
				startTime = now
			End If
			If logoutStart + logoutTime < now Then
				windowBox.operator.innerText = ""
				windowbox.errorDiv.style.background = "red"
				windowBox.errorString.innerText = "Logged out"
				logoutStart = now
			End If
		loop
		If .done.value = "access" then
			.done.value = false
			.accessText.innerText = "Retrying."
			.accessButton.style.backgroundcolor = "orange"
			AccessResult = Load_Access
			Call checkAccess	
		ElseIf .done.value = "send" Then
			.done.value = false
			.sendButton.disabled = true
			.sendButton.disabled = false
			Call Check_String(InputBox("Scan"))
		ElseIf .done.value = "SaveAs" Then
			.done.value = false
			If .SaveAsButton.style.backgroundcolor = "limegreen" Then
				.SaveAsButton.innerText = "Manual Save"
				.SaveAsButton.style.backgroundcolor = "orange"
				manualSave = True
			Else
				.SaveAsButton.style.backgroundcolor = "limegreen"
				.SaveAsButton.innerText = "Auto Save"
				manualSave = False
			End If
		Else
			closeWindow = true													'Variable to end loop
		End If
	loop
	.close																		'Closes the window
 end with
 'ServerClose()																	'Function to close open connections and return settings back to original	
 Wscript.Quit
 
Function HTABox(sBgColor)
	Dim nRnd : randomize : nRnd = Int(1000000 * rnd) 
	Dim strComputer : strComputer = "."
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Dim colItems : Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor",,48)
	Const HTAheight = 50
	Const l = 0
	Dim w : w = 1366
	Dim t : t = 768
	Dim objItem : For Each objItem in colItems
		If objItem.ScreenHeight <> "" and objItem.ScreenWidth <> "" Then
			w = objItem.ScreenWidth
			t = objItem.ScreenHeight - 50
		End If
	Next

	Dim sCmd : sCmd = "mshta.exe ""javascript:{new " _ 
		& "ActiveXObject(""InternetExplorer.Application"")" _ 
		& ".PutProperty('" & nRnd & "',window);" _ 
		& "window.moveTo(" & l & ", " & t & ");    " _
		& "window.resizeTo(" & w & "," & HTAheight & ")}""" 
	with CreateObject("WScript.Shell")
		.Run sCmd, 1, False 
		do until .AppActivate("javascript:{new ") : WSH.sleep 10 : loop 
	end with
	Dim Scr
	Dim IE : For Each IE In CreateObject("Shell.Application").windows 
		If IsObject(IE.GetProperty(nRnd)) Then 
			Set HTABox = IE.GetProperty(nRnd) 
			IE.Quit 
			With HTABox.Document.parentWindow.screen
				adjHeight = .availheight - HTAheight
				HTABox.moveTo l, adjHeight
				HTABox.resizeTo .availwidth, HTAheight


				If .availwidth >= 1920 Then
					HTABox.document.write LoadHTML(sBgColor, .availwidth)
				ElseIf .availwidth >= 1350 Then
					HTABox.document.write LoadMidHTML(sBgColor, .availwidth)
				Else
					HTABox.document.write LoadSmHTML(sBgColor, .availwidth)
				End If
				HTABox.document.title = "CMM" 
			End With
			Exit Function 
		End If 
	Next 
	MsgBox "HTA window not found." 
	wsh.quit
 End Function

Sub checkAccess()
	If AccessResult = false Then
		windowBox.accessText.innerText = "SQL Fail"
		windowBox.accessButton.style.backgroundcolor = "red"
	Else
		windowBox.accessText.innerText = "SQL"
		windowBox.accessButton.style.backgroundcolor = "limegreen"
		windowBox.accessButton.disabled = true
	End If
 End Sub

Sub connect2Scanner()
	Dim secs : secs = 0
	If machineString <> "No Scanner" and machineString <> "" Then
		windowBox.scannerText.innerText = machineString
		windowBox.scannerButton.style.backgroundcolor = "orange"
		windowBox.scannerButton.disabled = true
		windowBox.errorString.innerText = ""
	Else
		windowBox.scannerText.innerText = "Error"
		windowBox.scannerButton.style.backgroundcolor = "red"
		windowBox.scannerButton.disabled = false
		windowBox.errorDiv.style.backgroundColor = "red"
		windowBox.errorString.innerText = "No Scanner ID"
		Exit Sub
	End If
	
	'Stores variable if connected to part marker
	If Left(machineString, 3) = "COM" Then
		objComport.Open
		If( objComport.LastError <> 0 ) Then
			windowBox.scannerText.innerText = "Error: " & machineString
			windowBox.errorDiv.style.backgroundColor = "red"
			windowBox.errorString.innerText = objComport.LastError & " (" & objComport.GetErrorDescription( objComport.LastError ) & ")"
			windowBox.scannerButton.style.backgroundcolor = "red"
			windowBox.scannerButton.disabled = false
		Else
			windowBox.scannerText.innerText = machineString
			windowBox.scannerButton.style.backgroundcolor = "limegreen"
			windowBox.scannerButton.disabled = true
		End If
	Else
		windowBox.scannerText.innerText = "Error: " & machineString
		windowBox.scannerButton.style.backgroundcolor = "red"
		windowBox.scannerButton.disabled = false
	End If
 End Sub

Sub Check_String(stringFromScanner)
	Dim objCmd, userName, searchString, resultQty, sqlString, rs, sqlQTYString, rsQTY, sqlCutString, duplicateCnt, CMM_String, ProdID, partString, slugString
	Dim duplicate : duplicate = false
	Dim changeMade : changeMade = false
	Dim inputString : inputString = TrimString(stringFromScanner)
	Dim fixture : fixture = false
	
	logoutStart = now
	windowBox.sendButton.style.backgroundColor = ""
	If windowBox.duplicate.value = true or UCase(windowBox.duplicate.value) = "TRUE" Then
		If windowBox.duplicateSave.value <> inputString Then
			windowBox.duplicate.value = false
		Else
			duplicate = true
		End If
		windowBox.duplicateSave.value = false
	End If
	windowBox.duplicate.value = false
	windowbox.errorDiv.style.background = ""
	windowBox.errorString.innerText = ""
	windowBox.FixtureList.value = 0
	windowBox.MachineList.value = 0
	
	If inputString = "" or inputString = tabletPassword or inputString = computerPassword Then
	ElseIF Left(inputString, 3) = "WON" Then
			windowbox.errorDiv.style.background = "red"
			windowBox.errorString.innerText = "Scanning work orders is disabled"
			Exit Sub		
	ElseIF Len(inputString) = 10 and Mid(inputString, 9, 1) = "-" and Left(inputString, 1) = "H" Then
		windowBox.errorString.innerText = "CMM started"
		Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkAccess : Exit Sub
		If windowBox.operator.innerText = "" or windowBox.operator.innerText = "Not Authorized" Then
			windowbox.errorDiv.style.background = "red"
			windowBox.errorString.innerText = "Missing operator name"
			Exit Sub
		End If
		sqlString = "SELECT TOP 1 [FIC Blade Part Number], [Dash1ProdID], [Dash2ProdID], [FIC Slug Part Number] FROM [00_AE_SN_Control] " & _
					"LEFT JOIN [00_Invoice] ON  [00_AE_SN_Control].[Invoice Number] = [00_Invoice].[Invoice Number] " & _
					"WHERE [Blade Serial Number] = '" & inputString & "';"
		Set rs = objCmd.Execute(sqlString)
		changeMade = False
		DO WHILE NOT rs.EOF
			dashNum = rs.Fields(0)
			partString = rs.Fields(0)
			slugString = rs.Fields(3)
			If Right(dashNum, 1) = 1 Then
				ProdID = rs.Fields(1)
				If windowBox.dash1WO.innerText <> ProdID Then
					windowBox.dash1WO.innerText = ProdID	
					changeMade = True
				End If
			ElseIf Right(dashNum, 1) = 2 Then
				ProdID = rs.Fields(2)
				If windowBox.dash2WO.innerText <> ProdID Then
					windowBox.dash2WO.innerText = ProdID	
					changeMade = True
				End If
			End If
			rs.MoveNext
		Loop
		If duplicate = false Then
			sqlString = "SELECT COUNT([Blade S/N]) FROM [50_Final] WHERE [Blade S/N] = '" & inputString & "';"
			Set rs = objCmd.Execute(sqlString)
			If rs(0).value <> 0 Then
				windowbox.errorDiv.style.background = "red"
				windowBox.errorString.innerText = "Blade already scanned" & chr(10) & "Scan blade again for new CMM file"
				windowBox.duplicate.value = true
				windowBox.duplicateSave.value = inputString
				Exit Sub
			End If
			Set rs = Nothing
			If IsNull(ProdID) Then
				If partString  = "060053-1" or partString  = "062085-1" Then
					prodID = windowBox.dash1WO.innerText
				ElseIf partString = "060053-2" or partString  = "062085-1" Then
					prodID = windowBox.dash2WO.innerText
				End If
				If ProdID = "" Then
					windowbox.errorDiv.style.background = "red"
					windowBox.errorString.innerText = "Missing work order"
					Exit Sub
				End If
				Set rs = Nothing
			End If
		Else
			sqlString = "SELECT COUNT([Serial Number]) FROM [40_CMM_LPT5] WHERE [Serial Number] = '" & inputString & "';"
			Set rs = objCmd.Execute(sqlString)
			duplicateCnt = CInt(rs(0).value) + 1
			Set rs = Nothing
			sqlString = "SELECT [Comments] FROM [50_Final] WHERE [Blade S/N] = '" & inputString & "';"
			Set rs = objCmd.Execute(sqlString)
			DO WHILE NOT rs.EOF
				finalComment = rs.Fields(0)
				rs.MoveNext
			Loop	
			Set rs = Nothing
		End If
		sqlCutString = "SELECT TOP 1 [Fixture Location] FROM [20_LPT5] WHERE [Blade SN Dash 1]='" & inputString & "' OR [Blade SN Dash 2]='" & inputString & "' ORDER BY [Cut Date] DESC;"	
		set rs = objCmd.Execute(sqlCutString)
		DO WHILE NOT rs.EOF
			fixture = rs.Fields(0)
			rs.MoveNext
		Loop	
		Set rs = Nothing
		
		If fixture = false Then
			windowbox.errorDiv.style.background = "red"
			windowBox.errorString.innerText = "No fixture data found" & chr(10) & "Please scan the blade to the WaterJet"
			Exit Sub
		End If
		sqlString = "SELECT TOP 1 [MachineName], [Location], [CMMID] FROM [30_Fixtures] WHERE [FixtureID] = '" & fixture & "';"
		set rs = objCmd.Execute(sqlString)
		DO WHILE NOT rs.EOF
			windowBox.MachineList.value = rs.Fields(0)
			if rs.Fields(1) mod 2 = 0 then
				windowBox.FixtureList.value = rs.Fields(1) - 1
			else
				windowBox.FixtureList.value = rs.Fields(1)
			end if
			windowBox.CMMID.value = rs.Fields(2)
			rs.MoveNext
		Loop
		Set rs = Nothing
		
		If duplicate = false Then
			sqlString = "INSERT INTO [50_Final] ([Blade S/N], [Blade Inspected Date], [Final Insp Inspector Last Name], [ProdID], [CMMID]) VALUES (" _
										& "'" & inputString & "', '" & now & "', '" & windowBox.operator.innerText & "', '" & prodID & "', '" & CMMID & "');"
		Else
			If finalComment <> "" Then
				finalComment = finalComment & ", CMM rerun " & now & " on " & CMMID
			Else
				finalComment = "CMM rerun " & now
			End If
			sqlString = "UPDATE [50_Final] Set [Comments] = '" & finalComment & "' WHERE [Blade S/N] = '"  & inputString & "';"
			Set rs = objCmd.Execute(sqlString)
		End If
		
		If duplicate = false Then
			CMM_String = inputString & ".txt"
		Else
			CMM_String = inputString & " rerun " & DuplicateCnt & ".txt"
		End If
						'serial number, won,   part num,   operator,                     cmm number,            cmm filename
		CMM_Array = array(inputString, prodID, partString, slugString, windowBox.operator.innerText, windowBox.CMMID.value, CMM_String)
		
		If CMM_Windows(CMM_Array) Then
			Set rs = objCmd.Execute(sqlString)
			windowBox.errorString.innerText = inputString
			windowbox.errorDiv.style.background = "limegreen"
			Call updateWO
		Else
			windowBox.errorString.innerText = "CMM sequencing failed. Scan part again."
			windowbox.errorDiv.style.background = "red"
		End If
	ElseIf IsNumeric(inputString) Then
		Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkAccess : Exit Sub
		sqlString = "SELECT TOP 1 [CMMID] FROM [00_Personnel] WHERE [USERID]='" & stringFromScanner & "';"
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
	ElseIf inputString = "Calibrate" Then
		windowBox.errorString.innerText = "Calibration started"
		partString = "Calibration"
		If windowBox.operator.innerText = "" or windowBox.operator.innerText = "Not Authorized" Then
			windowbox.errorDiv.style.background = "red"
			windowBox.errorString.innerText = "Missing operator name"
			Exit Sub
		End If
		CMM_String = "CALIBRATION " & Replace(FormatDateTime(now,2),"/","-") & "_" & Replace(FormatDateTime(now,4),":","") & "_" & CMMID & ".txt"
		CMM_Array = array(inputString, false, partString, slugString, windowBox.operator.innerText, false, CMM_String)
		
		If CMM_Windows(CMM_Array) Then
			windowBox.errorString.innerText = inputString
			windowBox.errorString.innerText = "Calibration complete"
			windowbox.errorDiv.style.background = "limegreen"
		Else
			windowBox.errorString.innerText = windowBox.errorString.innerText & chr(10) & "Calibration sequencing failed. Try again."
			windowbox.errorDiv.style.background = "red"
		End If
	
	Else
		inputArray = Split(stringFromScanner)
		If UBound(inputArray) > 0 Then
			sqlString = "SELECT TOP 1 [CMMID] FROM [00_Personnel] WHERE [FirstName]='" & inputArray(0) & "' and [LastName]='" & inputArray(1) & "';"
		Else
			sqlString = "SELECT TOP 1 [CMMID] FROM [00_Personnel] WHERE [LastName]='" & inputArray(0) & "';"
		End If
		Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkAccess : Exit Sub
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

Sub updateWO()
	Set objCmd = GetNewConnection									'Creates the connection object to the database
	If objCmd is Nothing Then AccessResult = false : checkAccess : Exit Sub						'If a connection is not found, returns false to the function and then exits
	
	For n = 0 to 1
		If n = 0 Then
			prodID = windowBox.dash1WO.innerText
		Else
			prodID = windowBox.dash2WO.innerText
		End If
		If prodID <> "" Then
			sqlQTYString = "Select Count([ProdID]) From [50_Final] WHERE [PRODID] = '" & prodID & "'"
			set rsQTY = objCmd.Execute(sqlQTYString)												'Sends the request to the database
			If n = 0 Then
				windowBox.dash1WOCnt.innerText = rsQTY(0).value
			Else
				windowBox.dash2WOCnt.innerText = rsQTY(0).value
			End If
			Set rsQTY = Nothing
		End If
	Next
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing
	If CInt(windowBox.dash1WOCnt.innerText) >= CInt(windowBox.dash1WOQTY.innerText) and CInt(windowBox.dash1WOQTY.innerText) > 0 Then
		windowBox.errorString.innerText = "Work order complete" & chr(10) & windowBox.errorString.innerText
		windowBox.dash1WOQTY.innerText = 0
		windowBox.dash1WOCnt.innerText = 0
		windowBox.dash1WO.innerText = ""
		windowbox.errorDiv.style.background = "blue"
	ElseIf CInt(windowBox.dash2WOCnt.innerText) >= CInt(windowBox.dash2WOQTY.innerText) and CInt(windowBox.dash2WOQTY.innerText) > 0 Then
		windowBox.errorString.innerText = "Work order complete" & chr(10) & windowBox.errorString.innerText
		windowbox.errorDiv.style.background = "blue"
		windowBox.dash2WOQTY.innerText = 0
		windowBox.dash2WOCnt.innerText = 0
		windowBox.dash2WO.innerText = ""
	End If
 End Sub
 
Function CMM_Windows(CMM_Array)
	Dim shl : set shl = createobject("wscript.shell")
	Dim SerialNumber : SerialNumber = CMM_Array(0)
	Dim RunNumber 	 : RunNumber 	= CMM_Array(1)
	Dim PartNumber 	 : PartNumber 	= CMM_Array(2)
	Dim SlugPN	 : SlugPN	= CMM_Array(3)
	Dim OperatorID 	 : OperatorID 	= CMM_Array(4)
	Dim MachineID 	 : MachineID 	= CMM_Array(5)
	Dim CMM_File	 : CMMFile	= CMM_Array(6)
	Const newLine = 13
	Const TimeOut = 5000
	Const SaveTimeOut = 120000
	Const WaitMS = 50

	 If shl.AppActivate(PCDMIS_ID) Then
		wsh.sleep WaitMS
		shl.SendKeys "^q"
	 End If
	
	TimeCnt = 0
	
	Do
		wsh.sleep WaitMS
		TimeCnt = TimeCnt + WaitMS
	Loop While shl.AppActivate("Input Comment") = False and shl.AppActivate("PC-DMIS Message") = False and TimeCnt < SaveTimeOut 

	If TimeCnt >= SaveTimeOut Then
		windowBox.errorDiv.style.backgroundColor = "red"
		windowBox.errorString.innerText = "Program not executed"
		CMM_Windows = False
		Exit Function
	End If
	CMM_Windows = True
	If shl.AppActivate("PC-DMIS Message") Then
		wsh.sleep WaitMS
		shl.SendKeys "{ENTER}"
	End If
		
	Dim n, TimeCnt : For n = 0 to UBound(CMM_Array)
		TimeCnt = 0
		Do
			wsh.sleep WaitMS
			TimeCnt = TimeCnt + WaitMS
		Loop While shl.AppActivate("Input Comment") = False and TimeCnt < TimeOut
		If TimeCnt < TimeOut Then
			wsh.sleep WaitMS
			shl.SendKeys "{TAB}"
			If IsNull(CMM_Array(n)) Then
				shl.SendKeys "{DELETE}"
			Else
				shl.SendKeys CMM_Array(n)
			End If
			shl.SendKeys "{ENTER}"
			wsh.sleep WaitMS
		Else
			windowBox.errorDiv.style.backgroundColor = "red"
			windowBox.errorString.innerText = "Error finding the correct input window"
		End If
	Next
	TimeCnt = 0
	Do
		wsh.sleep WaitMS
		TimeCnt = TimeCnt + WaitMS
	Loop While shl.AppActivate("PC-DMIS Message") = False and TimeCnt < TimeOut
	If TimeCnt < TimeOut Then
		wsh.sleep WaitMS
		shl.SendKeys "{TAB}"
		If manualSave = True Then
			shl.SendKeys "N"
		Else
			shl.SendKeys "Y"
		End If
		wsh.sleep WaitMS
	Else
		windowBox.errorDiv.style.backgroundColor = "red"
		windowBox.errorString.innerText = "Error finding the correct autosave window"
	End If
	TimeCnt = 0
	Do
		wsh.sleep WaitMS
		TimeCnt = TimeCnt + WaitMS
	Loop While shl.AppActivate("Input Comment") = False and TimeCnt < TimeOut
	If TimeCnt < TimeOut Then
		wsh.sleep WaitMS
		shl.SendKeys "{TAB}"
		If PartNumber = "060053-1" Then
			shl.SendKeys 3
		ElseIf PartNumber = "060053-2" Then
			shl.SendKeys 4
		ElseIf PartNumber = "062085-1" Then
			shl.SendKeys 3
		ElseIf PartNumber = "062085-2" Then
			shl.SendKeys 4
		Else
			shl.SendKeys 1
		End If
		shl.SendKeys "{ENTER}"
	Else
		windowBox.errorDiv.style.backgroundColor = "red"
		windowBox.errorString.innerText = "Error finding the correct program window"
	End If
 End Function
 
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
 
Sub Update_Final()
	Dim rs, sqlString 
	If CMMID = "" Then Exit Sub
	Dim objCmd : Set objCmd = GetNewConnection
	If objCmd is Nothing Then Exit Sub

	Dim objFSO : Set objFSO=CreateObject("Scripting.FileSystemObject")
	Const FolderLocations = "C:\PC-DMIS Programs\AEROEDGE\AcceptedList\"
 
	Dim objFolder : Set objFolder = objFSO.GetFolder(FolderLocations)
	Dim colFiles : Set colFiles = objFolder.Files
	Dim objFile : For Each objFile in colFiles
		If Right(UCase(objFile.Name),4) = ".TXT" Then
			On Error Resume Next
				Err.Clear
				Accepted = ""
				Dim fileContents : Set fileContents = objFSO.OpenTextFile(objFile)
				If Err.Number = 0 Then
					Dim SerialNumber : SerialNumber = Left(objFile.Name, len(objFile.Name) - 4)
					Do Until fileContents.AtEndOfStream
						Accepted = Accepted & fileContents.ReadLine
					Loop
					fileContents.Close
					sqlString = "SELECT COUNT([Blade S/N]) FROM [50_Final] WHERE [Blade S/N] = '" & SerialNumber & "';"
					Set rs = objCmd.Execute(sqlString)
					If rs(0).value > 0 Then
						If Accepted <> "Y" and Accepted <> "N" Then
							Accepted = "N"
						End If
						Set rs = Nothing
						sqlString = "UPDATE [50_Final] Set [Accepted Y/N] = '" & Accepted & "' WHERE [Blade S/N] = '"  & SerialNumber & "';"
						Set rs = objCmd.Execute(sqlString)
						Set rs = Nothing
						objFSO.DeleteFile(objFile)
						windowBox.sendButton.style.backgroundColor = "limegreen"
					End If
				End If
			On Error GoTo 0
		End if
	Next
	Set colFiles = Nothing
	Set objFolder = Nothing
	objCmd.Close
	Set objCmd = Nothing
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
 
'Function to create all of the JS and HTML code for the window
Function LoadHTML(sBgColor, HTAwidth)
	'HTA String
	LoadHTML = "<HTA:Application " _
				 & "border=none " _
				 & "caption=no " _
				 & "contextMenu=no " _
				 & "innerborder=no " _
				 & "maximizebutton=no " _
				 & "minimizebutton=no " _
				 & "scroll=no " _
				 & "showintaskbar=no " _
				 & "singleinstance=yes " _
				 & "sysmenu=no " _
			 & "/>"
	
	'CSS String
	LoadHTML = LoadHTML _	
		& "<head><style>" _
		& "body {" _
			& "background-color: " & sBgColor & ";" _
			& "font:normal 28px Tahoma;" _
			& "border-Style:outset" _
			& "border-Width:3px" _
			& "}" _
		& ".unselectable {" _
			& "-moz-user-select: -moz-none;" _
			& "-khtml-user-select: none;" _
			& "-webkit-user-select: none;" _
			& "-o-user-select: none;" _
			& "user-select: none;" _
			& "}" _
		& ".buttonText {" _
			& "font: bold 12px Tahoma;" _
			& "}" _
		& ".errorFont {" _
			& "font: normal 16px Tahoma;" _
			& "color: white;" _
			& "}" _
		& ".closeButton {" _
			& "background-color: red;" _
			& "height: 30px;" _
			& "width: 30px;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
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
		& "</script></head>"

	'Body Start String							
	LoadHTML = LoadHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
	
	'SQL Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 4px; left: 4px; height: 19px; width: 19px; text-align: left;'>" _
		& "<button class=HTAButton id=accessButton style='height: 19px; width: 19px; text-align: center;background-color:orange;' disabled onclick='done.value=""access""'></button></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 4px; left: 27px; height: 19px; width: 70px; text-align: left;' id='accessText'>SQL</div>"
	
	'Scanner Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 27px; left: 4px; height: 19px; width: 19px; text-align: left;'>" _
		& "<button class=HTAButton id=scannerButton style='height: 19px; width: 19px; text-align: center;background-color:orange;' disabled onclick='done.value=""scanner""'></button></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 27px; height: 19px; width: 70px; text-align: left;' id='scannerText'>Scanner</div>"
	
	'Save As String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 27px; left: 100px; height: 19px; width: 91px; text-align: left;'>" _
		& "<button class='HTAButton buttonText' id=saveAsButton style='height: 19px; width: 91px; text-align: center;background-color:limegreen;' onclick='done.value=""SaveAs""'>Auto Save</button></div>"
	
	'Work Order String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 200px; height: 19px; width: 200px; text-align: left;'>060053-1 Work Order:</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 400px; height: 19px; width: 100px; text-align: center; font: normal;' id='dash1WO'></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 500px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash1WOCnt'>0</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 540px; height: 19px; width: 20px; text-align: center; font: normal;'>of</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 560px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash1WOQTY'>0</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 200px; height: 19px; width: 200px; text-align: left;'>060053-2 Work Order:</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 400px; height: 19px; width: 100px; text-align: center; font: normal;' id='dash2WO'></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 500px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash2WOCnt'>0</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 540px; height: 19px; width: 20px; text-align: center; font: normal;'>of</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 560px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash2WOQTY'>0</div>"
	
	'Operator String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top:  10px; left: 625px; height: 30px; width: 150px; text-align: left; font: bold;'>Operator:</div>" _
		& "<div unselectable='on' class='unselectable' style='top:  10px; left: 775px; height: 30px; width: 250px; text-align: left; font: normal;' id='operator'></div>" _
	
	'Machine Drop Down String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 15px; left: 1050px; height: 30px; width: 150px;'>" _
		& "<select id='MachineList' class='firstHidden' style='height: 30px; width: 250px;' onchange='machineFunction(this.value)' disabled>" _
			& "<option value='0' selected disabled>Select Machine</option>" _
			& "<option value='WJM_AE1' id='location1'>AMP 1 (WJM_AE1)</option>" _
			& "<option value='WJM_AE2' id='location2'>AMP 2 (WJM_AE2)</option>" _
			& "<option value='WJM_M500_1' id='location0'>Machine X (WJM_M500_1)</option>" _
			& "<option value='WJM_M500_2' id='location3'>AMP 6 (WJM_M500_2)</option>" _
			& "<option value='WJM_M500_3' id='location4'>AMP 7 (WJM_M500_3)</option>" _
			& "<option value='WJM_M500_4' id='location5'>AMP 5 (WJM_M500_4)</option>" _
			& "<option value='WJM_M500_5' id='location6'>AMP 4 (WJM_M500_5)</option>" _
			& "<option value='WJM_M500_6' id='location7'>AMP 3 (WJM_M500_6)</option>" _
		& "</select></div>"
		
	'Fixture Drop Down String
	LoadHTML = LoadHTML _
		& "<div unselectable='on' class='unselectable' style='top: 15px; left: 1325px; height: 30px; width: 150px;'>" _
		& "<select id='FixtureList' class='firstHidden' style='height: 30px; width: 150px;' onchange='fixtureFunction(this.value)' disabled>" _
			& "<option value='0' selected disabled >Select Fixture</option>" _
			& "<option value='1' id='location1'>Location 1 / 2</option>" _
			& "<option value='3' id='location3'>Location 3 / 4</option>" _
			& "<option value='5' id='location5'>Location 5 / 6</option>" _
			& "<option value='7' id='location7'>Location 7 / 8</option>" _
		& "</select></div>"
	
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div id=errorDiv class='errorFont' style='top:  0px; left: 1500px; height: 50px; width: " & HTAwidth - 1600 & "px;'>" _
			& "<div id=errorString class='errorFont' style='top: 5px; left: 10px; height: 40px; width: " & HTAwidth - 1620 & "px; text-align: center'></div></div>"
				
	'Send String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 10px; left: " & HTAwidth - 80 & "px;height: 30px; width: 30px;'><button id=sendButton style='height: 30px; width: 30px;' onclick='done.value=""send""'></button></div>"

	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 10px; left: " & HTAwidth - 40 & "px;height: 30px; width: 30px;'><button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'>&#10006;</button></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=done 				style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=CMMID 				style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=duplicate			style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=duplicateSave			style='visibility:hidden;' value=false><center></div>"

	'End Body String
	LoadHTML = LoadHTML _
		& "</body>"

 End Function
 
'Function to create all of the JS and HTML code for the window
Function LoadMidHTML(sBgColor, HTAwidth)
	'HTA String
	LoadMidHTML = "<HTA:Application " _
				 & "border=none " _
				 & "caption=no " _
				 & "contextMenu=no " _
				 & "innerborder=no " _
				 & "maximizebutton=no " _
				 & "minimizebutton=no " _
				 & "scroll=no " _
				 & "showintaskbar=no " _
				 & "singleinstance=yes " _
				 & "sysmenu=no " _
			 & "/>"
	
	'CSS String
	LoadMidHTML = LoadMidHTML _	
		& "<head><style>" _
		& "body {" _
			& "background-color: " & sBgColor & ";" _
			& "font:normal 20px Tahoma;" _
			& "border-Style:outset" _
			& "border-Width:3px" _
			& "}" _
		& ".unselectable {" _
			& "-moz-user-select: -moz-none;" _
			& "-khtml-user-select: none;" _
			& "-webkit-user-select: none;" _
			& "-o-user-select: none;" _
			& "user-select: none;" _
			& "}" _
		& ".buttonText {" _
			& "font: bold 16px Tahoma;" _
			& "}" _
		& ".errorFont {" _
			& "font: normal 16px Tahoma;" _
			& "color: white;" _
			& "}" _
		& ".firstHidden {" _
			& "font: normal 10px Tahoma;" _
			& "}" _
		& ".closeButton {" _
			& "background-color: red;" _
			& "height: 30px;" _
			& "width: 30px;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
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
		LoadMidHTML = LoadMidHTML _
			& "border-style: solid;" _
			& "border-Width:1px;"
	End If
	LoadMidHTML = LoadMidHTML _
			& "}" _
		& "</style>"
			
	'JS String
	LoadMidHTML = LoadMidHTML _
		& "<script language='javascript'>" _
		& "</script></head>"

	'Body Start String							
	LoadMidHTML = LoadMidHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
	
	'SQL Connect String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 4px; left: 4px; height: 19px; width: 19px; text-align: left;'>" _
		& "<button class=HTAButton id=accessButton style='height: 19px; width: 19px; text-align: center;background-color:orange;' disabled onclick='done.value=""access""'></button></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 4px; left: 27px; height: 19px; width: 40px; text-align: left;' id='accessText'>SQL</div>"
	
	'Scanner Connect String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 27px; left: 4px; height: 19px; width: 19px; text-align: left;'>" _
		& "<button class=HTAButton id=scannerButton style='height: 19px; width: 19px; text-align: center;background-color:orange;' disabled onclick='done.value=""scanner""'></button></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 27px; height: 19px; width: 40px; text-align: left;font: normal 12px Tahoma;' id='scannerText'>Scanner</div>"
	
	'Save As String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 27px; left: 70px; height: 19px; width: 63px; text-align: left;'>" _
		& "<button class='HTAButton buttonText' id=saveAsButton style='height: 19px; width: 63px; text-align: center;background-color:limegreen;font: normal 12px Tahoma;' onclick='done.value=""SaveAs""'>Auto Save</button></div>"
	
	'Work Order String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 140px; height: 19px; width: 60px; text-align: left;'>-1 WO:</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 200px; height: 19px; width: 100px; text-align: center; font: normal;' id='dash1WO'></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 300px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash1WOCnt'>0</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 340px; height: 19px; width: 20px; text-align: center; font: normal;'>of</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top:  4px; left: 360px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash1WOQTY'>0</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 140px; height: 19px; width: 60px; text-align: left;'>-2 WO:</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 200px; height: 19px; width: 100px; text-align: center; font: normal;' id='dash2WO'></div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 300px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash2WOCnt'>0</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 340px; height: 19px; width: 20px; text-align: center; font: normal;'>of</div>" _
		& "<div unselectable='on' class='unselectable buttonText' style='top: 27px; left: 360px; height: 19px; width: 40px; text-align: center; font: normal;' id='dash2WOQTY'>0</div>"
	
	'Operator String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable' style='top:  10px; left: 425px; height: 30px; width: 60px; text-align: left; font: bold;'>Oper:</div>" _
		& "<div unselectable='on' class='unselectable' style='top:  10px; left: 485px; height: 30px; width: 200px; text-align: left; font: normal;' id='operator'></div>" _
	
	'Machine Drop Down String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 700px; height: 15px; width: 200px;'>" _
		& "<select id='MachineList' class='firstHidden' style='height: 15px; width: 200px;' onchange='machineFunction(this.value)' disabled>" _
			& "<option value='0' selected disabled>Select Machine</option>" _
			& "<option value='WJM_AE1' id='location1'>AMP 1 (WJM_AE1)</option>" _
			& "<option value='WJM_AE2' id='location2'>AMP 2 (WJM_AE2)</option>" _
			& "<option value='WJM_M500_1' id='location0'>Machine X (WJM_M500_1)</option>" _
			& "<option value='WJM_M500_2' id='location3'>AMP 7 (WJM_M500_2)</option>" _
			& "<option value='WJM_M500_3' id='location4'>AMP 6 (WJM_M500_3)</option>" _
			& "<option value='WJM_M500_4' id='location5'>AMP 5 (WJM_M500_4)</option>" _
			& "<option value='WJM_M500_5' id='location6'>AMP 3 (WJM_M500_5)</option>" _
			& "<option value='WJM_M500_6' id='location7'>AMP 4 (WJM_M500_6)</option>" _
		& "</select></div>"
		
	'Fixture Drop Down String
	LoadMidHTML = LoadMidHTML _
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 700px; height: 15px; width: 200px;'>" _
		& "<select id='FixtureList' class='firstHidden' style='height: 15px; width: 200px;' onchange='fixtureFunction(this.value)' disabled>" _
			& "<option value='0' selected disabled >Select Fixture</option>" _
			& "<option value='1' id='location1'>Location 1 / 2</option>" _
			& "<option value='3' id='location3'>Location 3 / 4</option>" _
			& "<option value='5' id='location5'>Location 5 / 6</option>" _
			& "<option value='7' id='location7'>Location 7 / 8</option>" _
		& "</select></div>"
	
	'Error Output String
	LoadMidHTML = LoadMidHTML _	
		& "<div id=errorDiv class='errorFont' style='top:  0px; left: 950px; height: 50px; width: " & HTAwidth - 900 - 150 & "px;'>" _
			& "<div id=errorString class='errorFont' style='top: 5px; left: 10px; height: 40px; width: " & HTAwidth - 900 - 170 & "px; text-align: center'></div></div>"
				
	'Send String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 10px; left: " & HTAwidth - 80 & "px;height: 30px; width: 30px;'><button id=sendButton style='height: 30px; width: 30px;' onclick='done.value=""send""'></button></div>"

	'Close Box String
	LoadMidHTML = LoadMidHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 10px; left: " & HTAwidth - 40 & "px;height: 30px; width: 30px;'><button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'>&#10006;</button></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=done 				style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=CMMID 				style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=duplicate			style='visibility:hidden;' value=false><center></div>" _
		& "<div style='top: 0px; left: " & HTAwidth + 50 & "px;'><input type=hidden id=duplicateSave			style='visibility:hidden;' value=false><center></div>"

	'End Body String
	LoadMidHTML = LoadMidHTML _
		& "</body>"

 End Function
 