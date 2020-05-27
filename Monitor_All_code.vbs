Option Explicit	
	Const allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\All Operations.vbs"
	Const dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"

	Dim closeWindow : closeWindow = false
	Dim errorWindow : errorWindow = false
	Dim ColCount : ColCount = 0
	Dim bladeSNString : bladeSNString = ""
	Dim animateString : animateString = "flashButtonLegend;"
	Dim shiftColor : shiftColor = Array("Red", "BLACK", "DIMGREY", "MIDNIGHTBLUE", "MEDIUMBLUE", "DODGERBLUE")
	Dim RowCount : RowCount = 0
	Dim CMMHistory : CMMHistory = 6
	Dim tolName : tolName = Array("Dim 1.1",	"Dim 1.2",	"Dim 2.1",	"Dim 2.2",	"Dim 3.1",	"Dim 3.2",	"Dim 4.1",	"Dim 4.2",	"Dim 5.1",	"Dim 5.2",	"Dim 9.1",	"Dim 9.2",	"Dim 10.1",	"Dim 10.2",	"Dim 11 Max",	"Dim 11 Min",	"Dim 12 Max",	"Dim 12 Min")
	Dim minTol : minTol =   Array(40.795, 		40.795, 	155.1, 		155.1, 		168, 		168,  		155.1, 		155.1, 		26.7, 		26.7, 		16.55,  	16.55,  	32,  		32,  		-0.7,  	 		-0.7,  			-0.7,  			-0.7)
	Dim maxTol : maxTol =   Array(41.805, 		41.805, 	156.5, 		156.5,  	169.4,  	169.4,  	156.5,  	156.5,  	28.1,  		28.1,  		17.95,  	17.95,  	99.99,  	99.99,   	0.7,   	 		0.7,   	 		0.7,  			 0.7)
	
	Dim HTAWidth : HTAWidth = 1920
	Dim HTAHeight : HTAHeight = 1080
	
	Const pastDueStartDate = "4/15/2019"
	Const weeklyShipAmount = 1086
	Const weeklyShipBarAdj = 1.5514
	Const weeklyProdAmount = 1128
	Const weeklyProdBarAdj = 1.6114
	
	Const adminMode = false
	Const debugMode = false
	Const AbrasiveLimit = 40
	Const MixLimit = 40
	Const OrificeLimit = 600
	Const ColFooter = 705
	Const footerTop = 730
	Const sBgColor = "white"
	Const machineEffic = .8
	Const yieldRate = .95
	
	Dim FixtureArray()
	Dim machineNameArray()
	Dim machineNomenclatureArray()
	Dim mixingTubeArray()
	Dim allCMMArray()
	Dim offsetArray()
	Dim failureArray()
	Dim blankArray()
	Dim HTAX : HTAX = 0
	Dim HTAY : HTAY = 0
	Dim machineCount : machineCount = 0
	
	Dim x, timeCnt, Result, NewDic
	Dim strData, windowBox, AccessArray, AccessResult, HostID
	Dim SendData, RecieveData, wmi, cProcesses, oProcess
	Dim machineBox, strSelection, RemoteHost, RemotePort, machineString, MachineColWidth
	Dim strCommandLine, strParams, strParam, arrParams
	Dim processDate, createDate, PID, machineSummaryArray, machineNomenArray

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

	Const adOpenDynamic			= 2	 '// Uses a dynamic cursor.
	Const adOpenForwardOnly		= 0	 '// Default.
	Const adOpenKeyset			= 1	 '// Uses a keyset cursor.
	Const adOpenStatic			= 3	 '// Uses a static cursor.
	Const adOpenUnspecified		= -1 '// Does not specify the type of cursor.

	Const adLockBatchOptimistic	= 4	 '// Indicates optimistic batch updates. Required for batch update mode.
	Const adLockOptimistic		= 3	 '// Indicates optimistic locking, record by record.
	Const adLockPessimistic		= 2	 '// Indicates pessimistic locking, record by record.
	Const adLockReadOnly		= 1	 '// Indicates read-only records. You cannot alter the data.
	Const adLockUnspecified		= -1 '// Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.

	Const adStateClosed			= 0  '// The object is closed
	Const adStateOpen			= 1  '// The object is open
	Const adStateConnecting		= 2  '// The object is connecting
	Const adStateExecuting		= 4  '// The object is executing a command
	Const adStateFetching		= 8  '// The rows of the object are being retrieved
	
	Class TestClass
		Public ID
		Public TestText
		Private Sub Class_Initialize
				TestText  = ""
		End Sub
	End Class
	'*********************************************************
	
	Dim WshShell	
	Set WshShell = CreateObject("WScript.Shell")
	On Error Resume Next
		WshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Styles\MaxScriptStatements", "1107296255", "REG_DWORD"	'Changes security settings to ignore extended script operations
		WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
		WshShell.RegWrite "HKLM\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
		Set WshShell = nothing
	On Error GoTo 0
	Set wmi = GetObject("winmgmts:root\cimv2") 
	Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
	For Each oProcess in cProcesses
		processDate = left(oProcess.CreationDate, len(oProcess.CreationDate) - 4)
		If CDbl(processDate) > CDbl(createDate)Then
			createDate = processDate
			PID = oProcess.ProcessId
		End If
	Next
	For Each oProcess in cProcesses
		If PID <> oProcess.ProcessId Then oProcess.Terminate()
	Next
	
	strCommandLine = CMMHta.commandline
	strParams = Right(strCommandLine, Len(strCommandLine) - InStr(2, strCommandLine, """", 0))
	arrParams = Split(strParams, " ")
	For Each strParam In arrParams
		If strParam ="" Then
			'empty string, don't do anything!
				HostID = "QA0"
		Else
			If strParam = "QA1" Then 'My computer
				'HTAWidth = 2560
				'HTAHeight = 1600
				HTAX = 0
				HTAY = 0
				HostID = strParam
			ElseIf strParam = "QA2" Then
				HTAX = HTAWidth
				HTAY = 0
				HostID = strParam
			ElseIf strParam = "QA3" Then
				HTAX = 0
				HTAY = 0
				HostID = strParam
			Else
				HostID = "QA0"
			End If
		End If
	Next		
	
	'Function to check for access connection and load info from database
	AccessResult = Load_Access

	InitialParameters
	
	Result = HTABox(sBgColor, HTAWidth, HTAHeight, HTAX, HTAY)
	document.title = "Show all CMM"
	WaitSeconds .05
	CMMLoop	
	If CMMHistory <= 28 Then
		chartButton.disabled = false
	Else
		chartButton.disabled = true
		waitForInput
	End If
	google.setOnLoadCallback(drawChart)
	
Function waitForInput()
		Dim graphTime
		waitForLoop.value = true
		HostID = ""
		do until closeWindow = true													'Run loop until conditions are met
			do until done.value = "cancel" or done.value = "complete" or done.value = "allOps" or done.value = "okCorrect" or done.value = "changeHistory"
				WaitSeconds .05
				On Error Resume Next
				If done.value = true Then
					Self.Close()
				End If
				On Error GoTo 0
			loop
			if done.value = "cancel" then											'If the x button is clicked
				closeWindow = true													'Variable to end loop
			ElseIf done.value = "complete" Then
				done.value = false
				For timeCnt = 600 To 1 Step -1
					counterString.innerHTML = "Restarting in:<BR> " & Int(timeCnt/60) & ":" & Right("00" & timeCnt - Int(timeCnt/60) * 60,2)
					animateButtons timeCnt
					WaitSeconds 1
					On Error Resume Next
					If done.value = true Then
						Self.Close()
					End If
					On Error GoTo 0
					If done.value = "cancel" or done.value = "okCorrect" or done.value = "changeHistory" Then
						Exit For
					ElseIf done.value = "allOps" Then
						allOps
					End If
				Next
				counterString.innerHTML = ""
				If done.value <> "cancel" and done.value <> "okCorrect" and done.value <> "changeHistory" Then
					CMMLoop
					drawChart
				End If
			ElseIf done.value = "okCorrect" Then
				done.value = false
				addReason
				CMMLoop
				drawChart
			ElseIf done.value = "changeHistory" Then
				done.value = false
				InitialParameters
				buttonDiv.innerHTML = ""
				WaitSeconds .05
				CMMLoop
				drawChart
			ElseIf done.value = "allOps" Then
				allOps
			End If 
		loop
		close
		ServerClose()
 End Function
	
Sub WaitSeconds (intNumSecs) 
	' Because WScript.Sleep () is not available in HTA 
	' scripts, invoke a VBScript file to do the waiting. 
 
	Dim strScriptFile, strCommand, intRetcode, objWS 
 
	If intNumSecs <= 0 Then Exit Sub 
 
	Set objWS = CreateObject ("WScript.Shell") 
 
	strScriptFile = "%temp%\wait" & intNumSecs & "seconds.vbs" 
 
	strCommand = "cmd /c ""echo WScript.Sleep " & intNumSecs * 1000 & " >" & strScriptFile & _ 
				"&start /wait """" wscript.exe " & strScriptFile & """" 
 
	intRetCode = objWS.Run (strCommand, 0, True) 
 
	If intRetCode = 0 Then Exit Sub 
 
	LogLine "ERROR " & CStr (intRetCode) & " DURING WAITSECONDS PROCEDURE" 
 End Sub 

Function HTABox(sBgColor, Width, Height, HTAX, HTAY)
	window.resizeTo Width, Height
	window.moveTo HTAX, HTAY
	document.title = "HTABox" 
	document.write LoadHTML(sBgColor)
	WaitSeconds .05
	document.write LoadModalHTML
	chartButton.disabled = false
	HTABox = true
 End Function

Function InitialParameters()
	Dim objCmd : Set objCmd = GetNewConnection
	Dim sqlQuery(3), rs, a, FixtureID, FixtureID2, LocationID, n, strAnswer, machineArrayString, machineNomenString
	
	If Left(HostID, 2) <> "QA" Then
		strAnswer = InputBox("Please enter the number of history to display:" & chr(10) & chr(10) & "Default is 7 days" & chr(10) & "Warning - more than 14 days is slow to load.")
		If IsNumeric(strAnswer) = True Then
			If Int(strAnswer) > 7 Then CMMHistory = Int(strAnswer)
		Else CMMHistory = 7
		End If
	End If
	sqlQuery(0) = "Select COUNT(*) From [30_Fixtures] WHERE ((([ActiveFixture]) = 1) and ([ProgramName] = 'Cut1' or [ProgramName] = 'AllCut'));"
	set rs = objCmd.Execute(sqlQuery(0))
	If rs(0).value <> 0 Then
		ColCount = rs(0).value - 1
	End If
	Set rs = Nothing
	ReDim FixtureArray(ColCount + 1)
	ReDim machineNameArray(ColCount + 1)
	ReDim machineNomenclatureArray(ColCount + 1)
	ReDim mixingTubeArray(CMMHistory + 1, ColCount + 1)
	ReDim offsetArray(CMMHistory + 1, ColCount + 1)
	ReDim failureArray(3, ColCount + 1)
	ReDim allCMMArray(UBound(tolName))
	MachineColWidth = Int((HTAWidth - 10) / (ColCount + 2))
	
	sqlQuery(1) = "SELECT [FixtureID], [MachineName], [ProgramName], [Nomenclature]" _
				& " FROM [30_Fixtures]" _
				& " WHERE ((([ActiveFixture]) = 1) and ([ProgramName] = 'Cut1' or [ProgramName] = 'AllCut'))" _
				& " ORDER BY [Nomenclature] ASC;"
	a = 0
	set rs = objCmd.Execute(sqlQuery(1))
	DO WHILE NOT rs.EOF
		FixtureArray(a) = rs.Fields(0)
		machineNameArray(a) = rs.Fields(1)
		machineNomenclatureArray(a) = rs.Fields(3)
		rs.MoveNext
		a = a + 1
	Loop	
	Set rs = Nothing
	
	
	For a = 0 to ColCount
		LocationID = Right(FixtureArray(a), 1)
		FixtureID = FixtureArray(a)
		FixtureID2 = Left(FixtureID, Len(FixtureID) - 1) & LocationID + 1
		sqlQuery(2) = "SELECT COUNT(*) FROM [20_LPT5] WHERE (([Fixture Location]='" & FixtureID & "') and ([Cut Date] >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - CMMHistory) & "'));"
		Set rs = objCmd.Execute(sqlQuery(2))
			If rs(0).value > RowCount Then RowCount = rs(0).value
		Set rs = Nothing
	Next
	RowCount = RowCount + 1
	ReDim blankArray(RowCount, 7)
	
	sqlQuery(3) = "SELECT DISTINCT [MachineName], [MachNomen] FROM [30_Fixtures] WHERE ActiveFixture = 1 ORDER BY [MachNomen];"
	set rs = objCmd.Execute(sqlQuery(3))
	DO WHILE NOT rs.EOF
		machineCount = 	machineCount + 1
		machineArrayString = machineArrayString & rs.Fields(0) & ";"
		machineNomenString = machineNomenString & rs.Fields(1) & ";"
		rs.MoveNext
	Loop	
	Set rs = Nothing
	machineSummaryArray = Split(machineArrayString, ";")
	machineNomenArray = Split(machineNomenString, ";")
 End Function

Function animateButtons(timeCnt) 
	Dim animateID, flashColor
	
	If timeCnt mod 2 = 0 Then
	   flashColor = "Red"
	Else
	   flashColor = "Yellow"
	End If
	
	for each animateID in Split(animateString, ";")
		If animateID <> "" Then
			Document.getElementByID(animateID).style.backgroundcolor = flashColor
		End If
	Next
 End Function
 
Function allOps()
	Dim ScriptHost : ScriptHost = "WScript.exe"
	Dim objShell : Set objShell = CreateObject("WScript.Shell")
	Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")
	Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """"
	objShell.Run sOPsCmd
	Self.Close()
 End Function
 
Function CMMLoop()
	Dim sqlQuery, rs
	Dim a, b, offsetDate
	bladeSNString = ""
	animateString = "flashButtonLegend;"
	
	buttonDiv.scrollTop = 0
	Document.getElementByID("closeDiv").style.visibility="hidden"
	For a = 0 to UBound(FixtureArray) - 1
		Document.getElementByID("machineCol" & a).innerHTML = machineNomenclatureArray(a)
		Document.getElementByID("machineCol" & a).title = FixtureArray(a)
	Next
	Dim objCmd : Set objCmd = GetNewConnection
	For b = 0 to UBound(machineNameArray) - 1
		a = 0
		sqlQuery = "SELECT [MaintDate] FROM [20_Maint_History] WHERE [MachineID] = '" & machineNameArray(b) & "' AND [MaintType] = 'Mixing Tube' AND [MaintDate] >= '" _
				 & (CDate(FormatDateTime(Now, vbShortDate)) - CMMHistory) & "' ORDER BY [MaintDate] DESC;"
		set rs = objCmd.Execute(sqlQuery)
		DO WHILE NOT rs.EOF
			If a > 0 Then
				If CDate(FormatDateTime(rs.Fields(0), vbShortDate)) <> CDate(FormatDateTime(mixingTubeArray(a - 1, b), vbShortDate)) Then
					mixingTubeArray(a, b) = rs.Fields(0)
					a = a + 1
				End If
			Else
				mixingTubeArray(a, b) = rs.Fields(0)
				a = a + 1
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
		a = 0
		
		sqlQuery = "SELECT [CreateDate] FROM [30_Offset] WHERE [MachineNumber] = '" & machineNameArray(b) & "' AND [FileName] = '" & "Fixture" & CInt(Right(FixtureArray(b),1)) + 1 & "' AND [CreateDate] >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - CMMHistory) & "'  ORDER BY [CreateDate] DESC;"
		set rs = objCmd.Execute(sqlQuery)
		DO WHILE NOT rs.EOF
			If a > 0 Then
				If CDate(FormatDateTime(rs.Fields(0), vbShortDate)) <> CDate(FormatDateTime(offsetArray(a - 1, b), vbShortDate)) Then
					offsetArray(a, b) = rs.Fields(0)
					a = a + 1
				End If
			Else
				offsetArray(a, b) = rs.Fields(0)
				a = a + 1
			End If
			rs.MoveNext
		Loop
		Set rs = Nothing
		sqlQuery = "SELECT TOP 1 [CreateDate] FROM [30_Offset] WHERE [MachineNumber] = '" & machineNameArray(b) & "' AND [FileName] = '" & "Fixture" & CInt(Right(FixtureArray(b),1)) + 1 & "' ORDER BY [CreateDate] DESC;"
		set rs = objCmd.Execute(sqlQuery)
		DO WHILE NOT rs.EOF
			offsetDate = CDate(FormatDateTime(rs.Fields(0), vbShortDate))
			Document.getElementByID("mon3Cnt" & b & "Text").innerHTML = Left(offsetDate, Len(offsetDate) - 4) & Right(offsetDate, 2)
			rs.MoveNext
		Loop
		Set rs = Nothing
	Next
	a = UBound(FixtureArray)
	machineNameArray(a) = ""
	FixtureArray(a) = "Not scanned"
	Document.getElementByID("machineCol" & a).innerHTML = machineNameArray(a)
	sqlQuery = "SELECT [MachineID], [AbrasiveCnt], [MixCnt], [OrificeCnt] FROM [20_Counters];"
	set rs = objCmd.Execute(sqlQuery)
	DO WHILE NOT rs.EOF
		For b = 0 To machineCount - 1
			If machineSummaryArray(b) = rs.Fields(0) Then
				Document.getElementByID("MachMix" & b).style.backgroundcolor = ""
				Document.getElementByID("MachOri" & b).style.backgroundcolor = ""
				Document.getElementByID("MachMix" & b).innerHTML = rs.Fields(2)
				Document.getElementByID("MachOri" & b).innerHTML = rs.Fields(3)
				If rs.Fields(2) >= MixLimit Then 		Document.getElementByID("MachMix" & b).style.backgroundcolor = "red"
				If rs.Fields(3) >= OrificeLimit Then 	Document.getElementByID("MachOri" & b).style.backgroundcolor = "red"
			End If 
		Next
		rs.MoveNext
	Loop
	Set rs = Nothing
	
	
	'************ MACHINE INFO *************
	
	sqlQuery = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Failures] = 0 AND [Date] >= '" & now - 7 & "';"
	set rs = objCmd.Execute(sqlQuery)
	Dim weekFail : weekFail = rs(0).value
	sqlQuery = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Date] >= '" & now - 7 & "';"
	set rs = objCmd.Execute(sqlQuery)
	Dim weekTotal : weekTotal = rs(0).value
	If weekTotal <> 0 Then
		Document.getElementByID("totalYield").innerHTML = Int(weekFail / weekTotal * 1000)/10 & " %"
	Else
		Document.getElementByID("totalYield").innerHTML = "Error"
	End If
	Set rs = Nothing
	'totalRunTime
	
	sqlQuery = "SELECT COUNT(*) FROM [20_LPT5] WHERE [Cut Date] >= '" & now - 7 & "';"
	set rs = objCmd.Execute(sqlQuery)
	Dim runCount : runCount = rs(0).value
	If runCount <> 0 Then
		Document.getElementByID("totalRunTime").innerHTML = Int(runCount / 140 * 54 / 60 / machineCount * 2000)/10 & " %"
	Else
		Document.getElementByID("totalRunTime").innerHTML = "Error"
	End If
	Set rs = Nothing
	
	
	'MachMix" & a & ">0</div>" _
	'MachOri" & a & ">0</div>" _
	'MachRun" & a & ">0</div>" _
	'MachYld" & a & ">0</div>"
	
	Dim weekEndDate : weekEndDate = DateAdd("d", -((Weekday(date()) + 7 - 1) Mod 7), date())
	sqlQuery = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Failures] = 0 AND [Date] > '" & weekEndDate & "';"
	set rs = objCmd.Execute(sqlQuery)
	Document.getElementByID("prodText").innerHTML = rs(0).value
	If rs(0).value > weeklyProdAmount Then
		Document.getElementByID("prodBar").style.width = Int(weeklyProdAmount / weeklyProdBarAdj)
	Else
		Document.getElementByID("prodBar").style.width = Int(rs(0).value / weeklyProdBarAdj)
	End If
	Set rs = Nothing
	sqlQuery = "SELECT COUNT(*) FROM [60_SHIPPING] WHERE [Date Shipped] >= '" & weekEndDate + 2 & "';"
	set rs = objCmd.Execute(sqlQuery)
	Document.getElementByID("shipText").innerHTML = rs(0).value
	If rs(0).value > weeklyShipAmount Then
		Document.getElementByID("shipBar").style.width = Int(weeklyShipAmount / weeklyShipBarAdj)
	Else
		Document.getElementByID("shipBar").style.width = Int(rs(0).value / weeklyShipBarAdj)
	End If
	sqlQuery = "SELECT COUNT(*) FROM [40_CMM_LPT5] WHERE [Failures] = 0 AND [Date] > '" & weekEndDate - 7 & "' AND [Date] <= '" & weekEndDate & "';"
	Set rs = objCmd.Execute(sqlQuery)
	Document.getElementByID("prodTextPrev").innerHTML = Int(rs(0).value)
	If rs(0).value > weeklyProdAmount Then
		Document.getElementByID("prodBarPrev").style.width = Int(weeklyProdAmount / weeklyProdBarAdj)
	Else
		Document.getElementByID("prodBarPrev").style.width = Int(rs(0).value / weeklyProdBarAdj)
	End If
	Set rs = Nothing
	sqlQuery = "SELECT COUNT(*) FROM [60_SHIPPING] WHERE [Date Shipped] > '" & weekEndDate - 7 + 2 & "' AND [Date Shipped] < '" & weekEndDate + 2 & "';"
	Set rs = objCmd.Execute(sqlQuery)
	Document.getElementByID("shipTextPrev").innerHTML = Int(rs(0).value)
	If rs(0).value > weeklyShipAmount Then
		Document.getElementByID("shipBarPrev").style.width = Int(weeklyShipAmount / weeklyShipBarAdj)
	Else
		Document.getElementByID("shipBarPrev").style.width = Int(rs(0).value / weeklyShipBarAdj)
	End If
	Set rs = Nothing
	Select Case Weekday(date(), 2)  
		Case 1, 2, 3, 4
			Document.getElementByID("currentLocation").style.width = Int(124 * (Weekday(date(), 2) - 1 + now() - date()))
		Case Else
			Document.getElementByID("currentLocation").style.width = Int(124 * 4 + 68 * (Weekday(date(), 2) - 5 + ((now() - date()) * 24 - 3)/ 13))
	End Select
		
	' sqlQuery = "SELECT COUNT(*) FROM [60_Shipping] WHERE [Date Shipped] >= '" & pastDueStartDate & "' AND [Date Shipped] <= '" & weekEndDate + 2 & "';"
	' Set rs = objCmd.Execute(sqlQuery)
	' Dim totalShipped : totalShipped = rs(0).value
	' Dim shouldHaveShipped : shouldHaveShipped = Int((weekEndDate + 2 - CDate(pastDueStartDate)) / 7) * weeklyShipAmount

	' If totalShipped >= shouldHaveShipped Then
		' Document.getElementByID("pastDue").innerText = totalShipped - shouldHaveShipped & " blades ahead"
	' Else
		' Document.getElementByID("pastDue").innerText = shouldHaveShipped - totalShipped  & " blades behind"
	' End If
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
'************ MACHINE INFO *************
		
	Set objCmd = GetNewConnection
	For a = 0 to ColCount + 1
		If done.value = "cancel" Then
			objCmd.Close
			Set objCmd = Nothing
			ServerClose()																	'Function to close open connections and return settings back to original	
			Self.Close()
			Exit Function
		ElseIf done.value = "changeHistory" Then
			objCmd.Close
			Set objCmd = Nothing
			errorString.innerHTML = ""
			Exit Function
		End If
		CMMSearch FixtureArray(a), Right(FixtureArray(a), 1), objCmd, a
		WaitSeconds .05
	Next
	objCmd.Close
	Set objCmd = Nothing
	paretoArray
	errorString.innerHTML = ""
	If done.value <> "okCorrect" or done.value = "changeHistory" Then done.value = "complete"
 End Function

Function paretoArray()
	Dim i, temp
	Dim gDic : Set gDic = CreateObject("Scripting.Dictionary")
	Set NewDic = nothing
	
	For i = 0 to UBound(tolName)
		Set temp = new TestClass
		temp.ID = tolName(i)
		temp.TestText = allCMMArray(i)

		gDic.Add i,temp
	Next
	Set NewDic = SortDict(gDic)
 End Function

Function SortDict(dict)
	Dim i, j, arrKeys, arrItems, tempObj
	arrKeys = dict.keys                                               'Array containing the keys
	arrItems = dict.Items                                             'Array containing the Items(which are nothing but objects of class TestClass)
	Set tempObj = New TestClass
	For i=0 To UBound(arrItems)-1                                     'From 1st element to the penultimate element
		For j=i+1 To UBound(arrItems)                                 'From i+1th element to last element
			If arrItems(i).TestText < arrItems(j).TestText Then                  'Sorting in DESCENDING ORDER by the Property "ID"
				tempObj.ID = arrItems(i).ID
				tempObj.TestText = arrItems(i).testText
				dict.item(arrKeys(i)).ID = arrItems(j).ID
				dict.item(arrKeys(i)).TestText = arrItems(j).TestText
				dict.item(arrKeys(j)).ID = tempObj.ID
				dict.item(arrKeys(j)).TestText = tempObj.TestText
			End If
		Next
	Next
	Set SortDict = dict
 End Function

Function CMMSearch(FixtureID, LocationID, objCmd, ColX)
	Dim cParts1 : Set cParts1 = CreateObject("Scripting.Dictionary")
	Dim cParts2 : Set cParts2 = CreateObject("Scripting.Dictionary")
	Dim cDates : Set cDates = CreateObject("System.Collections.ArrayList")
	Dim cBlade(6)
	Dim noCMMArray
	Dim sqlQuery, efficiency
	Dim toleranceArray(17)
	Dim FixtureID2, rs, n, dateValue, i, bladeArray, slugArray, DateString, DateSerial, duplicate, a, b, tolResult, ReasonString, ReasonArray(), CMMDay, YesterdayArray()
	ReDim bladeArray(RowCount + 1, 20) : ReDim slugArray(RowCount + 1, 20) 
	ReDim YesterdayArray(CMMHistory + 1,7)
	For a = 0 to CMMHistory
		For b = 0 to 7
			YesterdayArray(a, b) = 0
		Next
	Next
	If IsNumeric(LocationID) = False Then
		sqlQuery    = "SELECT Count([40_CMM_LPT5].[Serial Number])" _
					& " FROM ([40_CMM_LPT5]" _
					& " LEFT JOIN [20_LPT5] 					ON [40_CMM_LPT5].[Serial Number] = [20_LPT5].[Blade SN Dash 1])" _
					& " LEFT JOIN [20_LPT5] AS [20_LPT5_2] ON [40_CMM_LPT5].[Serial Number] = [20_LPT5_2].[Blade SN Dash 2]" _
					& " WHERE ((([20_LPT5].[Fixture Location] Is Null) AND ([20_LPT5_2].[Fixture Location] Is Null)) AND ([40_CMM_LPT5].Date >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - CMMHistory) & "'));"
		set rs = objCmd.Execute(sqlQuery)
		ReDim noCMMArray(rs(0).value - 1, 1)
		sqlQuery    = "SELECT [40_CMM_LPT5].[Serial Number], [40_CMM_LPT5].[Date]" _
					& " FROM ([40_CMM_LPT5]" _
					& " LEFT JOIN [20_LPT5] 					ON [40_CMM_LPT5].[Serial Number] = [20_LPT5].[Blade SN Dash 1])" _
					& " LEFT JOIN [20_LPT5] AS [20_LPT5_2] ON [40_CMM_LPT5].[Serial Number] = [20_LPT5_2].[Blade SN Dash 2]" _
					& " WHERE ((([20_LPT5].[Fixture Location] Is Null) AND ([20_LPT5_2].[Fixture Location] Is Null)) AND ([40_CMM_LPT5].Date >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - CMMHistory) & "'))" _
					& " ORDER BY [40_CMM_LPT5].[Date] DESC;"
		set rs = objCmd.Execute(sqlQuery)
		a = 0
		DO WHILE NOT rs.EOF
			noCMMArray(a, 0) = rs.Fields(0)
			noCMMArray(a, 1) = rs.Fields(1)
			rs.MoveNext
			a = a + 1
		Loop	
		SlugRow = 0
		For a = 0 to UBound(noCMMArray,1)
			sqlQuery = "SELECT [00_AE_SN_Control].[Blade Serial Number] as [Dash 1 SN], [00_2].[Blade Serial Number] as [Dash 2 SN] " _
					 & "FROM [00_AE_SN_Control] " _
					 & "LEFT JOIN [00_AE_SN_Control] as [00_2] ON [00_2].[Slug Serial Number]=[00_AE_SN_Control].[Slug Serial Number] " _
					 & "WHERE [00_AE_SN_Control].[Slug Serial Number] = (" _
						& "SELECT TOP 1 [00_AE_SN_Control].[Slug Serial Number] " _
						& "FROM [00_AE_SN_Control] " _
						& "WHERE [00_AE_SN_Control].[Blade Serial Number]='" & noCMMArray(a, 0) & "') " _
					 & "and [00_AE_SN_Control].[FIC Blade Part Number] = '060053-1' " _
					 & "and [00_2].[FIC Blade Part Number] = '060053-2';"
			set rs = objCmd.Execute(sqlQuery)
			DO WHILE NOT rs.EOF
				For b = 0 to RowCount + 1 
					If slugArray(b, 1) = rs.Fields(0) Then Exit For
				Next
				If b > RowCount + 1 Then
					slugArray(SlugRow, 0) = noCMMArray(a, 1)			'Scan Date
					slugArray(SlugRow, 1) = rs.Fields(0)		'Blade 1 SN
					tolResult = toleranceCheck(rs.Fields(0), objCmd)
					slugArray(SlugRow, 2) = tolResult(0)		'Pass or Fail
					slugArray(SlugRow, 3) = tolResult(1)		'failed String
					slugArray(SlugRow, 4) = tolResult(2)		'failed features
					
					slugArray(SlugRow, 11) = rs.Fields(1)		'Blade 2 SN
					tolResult = toleranceCheck(rs.Fields(1), objCmd)
					slugArray(SlugRow, 12) = tolResult(0)		'Pass or Fail
					slugArray(SlugRow, 13) = tolResult(1)		'failed String
					slugArray(SlugRow, 14) = tolResult(2)		'failed features
					SlugRow = SlugRow + 1
				End If
				rs.MoveNext
			Loop
		Next
		HTASlug slugArray, ColX, False
		Exit Function
	End If
	
	sqlQuery    = "SELECT [Blade SN Dash 1], [Blade SN Dash 2], [Cut Date]" _
				& " FROM [20_LPT5]" _
				& " WHERE (([Fixture Location]='" & FixtureID & "') and ([Cut Date] >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - CMMHistory) & "'))" _
				& " ORDER BY [Cut Date] DESC;"
	set rs = objCmd.Execute(sqlQuery)
	Dim SlugRow : SlugRow = 0
	DO WHILE NOT rs.EOF
		DateString = Split(rs.Fields(2), " ")
		If UBound(DateString) = 2 Then
			DateSerial = CDbl(CDate(DateString(0))) + CDbl(CDate(DateString(1) & " " & DateString(2)))
		Else				
			DateSerial = CDbl(CDate(DateString(0))) + CDbl(CDate(DateString(1)))
		End If
	
		slugArray(SlugRow, 0) = DateSerial			'Scan Date
		slugArray(SlugRow, 1) = rs.Fields(0)		'Blade 1 SN
		tolResult = toleranceCheck(rs.Fields(0), objCmd)
		slugArray(SlugRow, 2) = tolResult(0)		'Pass or Fail
		slugArray(SlugRow, 3) = tolResult(1)		'failed String
		slugArray(SlugRow, 4) = tolResult(2)		'failed features
		
		slugArray(SlugRow, 11) = rs.Fields(1)		'Blade 2 SN
		tolResult = toleranceCheck(rs.Fields(1), objCmd)
		slugArray(SlugRow, 12) = tolResult(0)		'Pass or Fail
		slugArray(SlugRow, 13) = tolResult(1)		'failed String
		slugArray(SlugRow, 14) = tolResult(2)		'failed features

		CMMDay = Int(CDate(FormatDateTime(Now, vbShortDate)) - CDate(FormatDateTime(rs.Fields(2) - 0.15, vbShortDate)))
		YesterdayArray(CMMDay,0) = YesterdayArray(CMMDay,0) + 2
		If slugArray(SlugRow, 2)  = "Pass" Then YesterdayArray(CMMDay,4) = YesterdayArray(CMMDay,4) + 1
		If slugArray(SlugRow, 12) = "Pass" Then YesterdayArray(CMMDay,4) = YesterdayArray(CMMDay,4) + 1
		If Weekday(rs.Fields(2) - 0.15, vbMonday) < 5 and CDate(TimeValue(rs.Fields(2))) > CDate(.15) and CDate(TimeValue(rs.Fields(2))) < CDate(.65) Then
			YesterdayArray(CMMDay,1) = YesterdayArray(CMMDay,1) + 2
		ElseIf Weekday(rs.Fields(2) - 0.15, vbMonday) < 5 Then
			YesterdayArray(CMMDay,2) = YesterdayArray(CMMDay,2) + 2
		Else
			YesterdayArray(CMMDay,3) = YesterdayArray(CMMDay,3) + 2
		End If
		rs.MoveNext
		SlugRow = SlugRow + 1
	Loop	
	Set rs = Nothing
	sqlQuery = "SELECT COUNT(*) FROM [30_Reason] WHERE [ReasonDate] >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - 10) & "';"
	set rs = objCmd.Execute(sqlQuery)
	ReDim ReasonArray(rs(0).value, 2)
	Set rs = Nothing
	sqlQuery = "SELECT [OperName], [Reason], [SNs] FROM [30_Reason] WHERE [ReasonDate] >= '" & (CDate(FormatDateTime(Now, vbShortDate)) - 10) & "';"
	set rs = objCmd.Execute(sqlQuery)
	a = 0
	DO WHILE NOT rs.EOF
		ReasonString = ReasonString & rs.Fields(2)
		ReasonArray(a, 0) = rs.Fields(0)
		ReasonArray(a, 1) = rs.Fields(1)
		ReasonArray(a, 2) = rs.Fields(2)
		a = a + 1
		rs.MoveNext
	Loop
	Set rs = Nothing
	For a = 0 To UBound(slugArray)
		If InStr(1, ReasonString, slugArray(a, 1)) <> 0 Then
			slugArray(a, 5) = 1
			For b = 0 to UBound(ReasonArray)
				If InStr(1, ReasonArray(b, 2), slugArray(a, 1)) Then
					slugArray(a, 6) = ReasonArray(b, 0)		'Operator
					slugArray(a, 7) = ReasonArray(b, 1)		'Reason
				End If
			Next
		End If
		If InStr(1, ReasonString, slugArray(a, 11)) <> 0 Then
			slugArray(a, 15) = 1
			For b = 0 to UBound(ReasonArray)
				If InStr(1, ReasonArray(b, 2), slugArray(a, 11)) Then
					slugArray(a, 16) = ReasonArray(b, 0)		'Operator
					slugArray(a, 17) = ReasonArray(b, 1)		'Reason
				End If
			Next
		End If
	Next
	
	HTASlug slugArray, ColX, False
 End Function

Function HTASlug(slugArray, ColX, isEmptyArray)
	Dim a, b, TextID, shiftID, arrayDate, rowDate, arrayDay, divRow, divHeight
	Dim dash1title, dash2title, dash1vis, dash2vis, dash1color, dash2color, dash1disable, dash2disable, dateText, dateTextColor, dateTextBackcolor
	Dim slugSN, slugDate, dash1tolResult, dash1tolValue, dash2tolResult, dash2tolValue
	Dim dash1cor, dash1oper, dash1corDesc, dash2cor, dash2oper, dash2corDesc, dash1failMode, dash2failMode
	Dim mixTubeCnt : mixTubeCnt = 0
	Dim mixTubeDate : mixTubeDate = mixingTubeArray(mixTubeCnt, ColX)
	Dim offsetCnt : offsetCnt = 0
	Dim offsetDate : offsetDate = offsetArray(offsetCnt, ColX)
	Dim RowCnt(6)
	
	For b = 1 to RowCount
		dash1title = ""
		dash2title = ""
		dash1vis = "hidden"
		dash2vis = "hidden"
		dash1color = ""
		dash2color = ""
		dash1disable = true
		dash2disable = true
		dateText = ""
		dateTextColor = ""
		dateTextBackcolor = ""
		shiftID = 0
		divRow = 0
		For a = 0 to CMMHistory + 1
			Document.getElementByID(a & "Part" & ColX & "_" & b & "Div").style.display = "none"
			Document.getElementByID(a & "Part" & ColX & "_" & b & "Text").innerHTML = dateText
			Document.getElementByID(a & "Part" & ColX & "_" & b & "Text").style.color = shiftColor(shiftID)
			Document.getElementByID(a & "Part" & ColX & "_" & b & "Text").style.backgroundcolor = dateTextBackcolor
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_1Button").SNValue = ""
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_2Button").SNValue = ""
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_1Button").title = dash1title
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_2Button").title = dash2title
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_1Button").style.visibility = dash1vis
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_2Button").style.visibility = dash2vis
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_1Button").disabled = dash1disable
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_2Button").disabled = dash2disable
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_1Button").style.backgroundcolor = dash1color
			Document.getElementByID(a & "Part" & ColX & "_" & b & "_2Button").style.backgroundcolor = dash2color
		Next
		slugDate = slugArray(b - 1, 0)
		
		dash1title = slugArray(b - 1, 1)
		dash1tolResult = slugArray(b - 1, 2)
		dash1tolValue = slugArray(b - 1, 3)
		dash1failMode = slugArray(b - 1, 4)
		dash1cor = slugArray(b - 1, 5)
		dash1oper = slugArray(b - 1, 6)
		dash1corDesc = slugArray(b - 1, 7)
		
		dash2title = slugArray(b - 1, 11)
		dash2tolResult = slugArray(b - 1, 12)
		dash2tolValue = slugArray(b - 1, 13)
		dash2failMode = slugArray(b - 1, 14)
		dash2cor = slugArray(b - 1, 15)
		dash2oper = slugArray(b - 1, 16)
		dash2corDesc = slugArray(b - 1, 17)
		
		
		If dash1title = "" Then dash1title = "No serial number found"		
		If dash2title = "" Then dash2title = "No serial number found"
		
		If slugDate <> 0 Then
			divRow = CInt(Date() - Int(slugDate))
			Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_1Button").SNValue = dash1title
			Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_2Button").SNValue = dash2title
			If dash1tolValue <> "" Then dash1title = dash1title & chr(10) & Replace(dash1tolValue, ", " , chr(10)) : dash1disable = false
			If dash2tolValue <> "" Then dash2title = dash2title & chr(10) & Replace(dash2tolValue, ", " , chr(10)) : dash2disable = false
		
			rowDate = FormatDateTime(CDate(slugDate), vbShortDate)
			If slugDate <= mixTubeDate Then
				If b > 1 Then
					Document.getElementByID(divRow & "Part" & ColX & "_" & b - 1 & "Text").style.backgroundcolor = "Orange"
				Else
					Document.getElementByID(divRow & "Part" & ColX & "_" & b & "Text").style.backgroundcolor = "Orange"
				End If
				mixTubeCnt = mixTubeCnt + 1
				mixTubeDate = mixingTubeArray(mixTubeCnt, ColX)
			End If
			dateText = Left(rowDate, len(rowDate) - 5) & " " & FormatDateTime(CDate(slugDate), vbShortTime)
			arrayDate = slugDate - 0.15
			arrayDay = Weekday(arrayDate, vbMonday)
			If arrayDay < 5 and CDate(arrayDate - Int(arrayDate)) < CDate(.5) Then
				shiftID = 1
			ElseIf arrayDay < 5 Then
				shiftID = 2
			ElseIF arrayDay = 5 Then
				shiftID = 3
			ElseIF arrayDay = 6 Then
				shiftID = 4
			Else
				shiftID = 5
			End If
			
			If slugDate <= offsetDate Then
				If b > 1 Then
					Document.getElementByID(divRow & "Part" & ColX & "_" & b - 1 & "Text").innerHTML = "&#10004; " & Document.getElementByID(divRow & "Part" & ColX & "_" & b - 1& "Text").innerHTML
				Else
					dateText = "&#10004; " & dateText
				End If
				offsetCnt = offsetCnt + 1
				offsetDate = offsetArray(offsetCnt, ColX)
			End If
			
			If dash1tolResult = "Pass" Then
				dash1color = "limegreen"
			ElseIF dash1tolResult = "Fail" Then
				If dash1cor = 1 Then
					dash1color = "blue"
					dash1title = dash1title & chr(10) & chr(10) & "Operator Name: " & dash1oper & chr(10) & "Correction Made: " & dash1corDesc
				Else
					dash1color = checkFailHistory(dash1failMode, b, slugArray, ColX, "_1Button", divRow)
				End If
			Else
				dash1color = ""
				dash1title = dash1title & chr(10) & chr(10) & "Missing CMM File"
			End If
			
			If dash2tolResult = "Pass" Then
				dash2color = "limegreen"
			ElseIF dash2tolResult = "Fail" Then
				If dash2cor = 1 Then
					dash2color = "blue"
					dash2title = dash2title & chr(10) & chr(10) & "Operator Name: " & dash2oper & chr(10) & "Correction Made: " & dash2corDesc
				Else
					dash2color = checkFailHistory(dash2failMode, b, slugArray, ColX, "_2Button", divRow)
				End If
			Else
				dash2color = ""
				dash2title = dash2title & chr(10) & chr(10) & "Missing CMM File"
			End If
			dash1vis = "visible"
			dash2vis = "visible"
		Else
			Document.getElementByID(divRow & "Part" & ColX & "_" & b & "Text").style.backgroundcolor = ""
		End If
		If dash1vis = "hidden" Then
			Document.getElementByID(divRow & "Part" & ColX & "_" & b & "Div").style.display = "none"
		Else
			Document.getElementByID(divRow & "Part" & ColX & "_" & b & "Div").style.display = "inline-block"
			Document.getElementByID(divRow & "Part" & ColX & "_" & b & "Div").style.top = RowCnt(divRow) * 15 & "px"
			RowCnt(divRow) = RowCnt(divRow) + 1
		End If
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "Text").style.color = shiftColor(shiftID)
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_1Button").title = dash1title
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_2Button").title = dash2title
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_1Button").disabled = dash1disable
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_2Button").disabled = dash2disable
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_1Button").style.backgroundcolor = dash1color
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_2Button").style.backgroundcolor = dash2color
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "Text").innerHTML = "&nbsp;&nbsp;"'dateText
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_1Button").style.visibility = dash1vis
		Document.getElementByID(divRow & "Part" & ColX & "_" & b & "_2Button").style.visibility = dash2vis
		errorString.innerHTML = ColX + 1 & " of " & ColCount + 2
	Next
	For a = 0 to UBound(RowCnt)
		divHeight = Document.getElementByID("divDay" & a).style.height
		If CInt(Left(divHeight, len(divHeight) - 2)) < RowCnt(a) * 15 Then
			Document.getElementByID("divDay" & a).style.height = RowCnt(a) * 15 & "px"
		End If
	Next
 End Function

Function checkFailHistory(failMode, b, slugArray, ColX, buttonType, divRow)
	Dim a, correctiveCol, failID, failModeCol
	checkFailHistory = "red"
	
	If buttonType = "_1Button" Then
		failModeCol = 4
		correctiveCol = 5
	Else
		failModeCol = 14
		correctiveCol = 15
	End If
	
	For Each failID in Split(failMode, ";")
		If failID = "" Then
		ElseIf failID = "Dim 1.2" 	 or failID = "Dim 3.2"	  or failID = "Dim 10.1" 	or failID = "Dim 10.2" 	 Then
		Else
			For a = -2 to 2
				If (b + a > 0) and (b + a <= RowCount) and (a <> 0) Then
					If (InStr(slugArray(b + a - 1, correctiveCol), failID) = 1) Then
					ElseIf (InStr(slugArray(b + a - 1, failModeCol), failID) <> 0) Then
						If (InStr(animateString, "Part" & ColX & "_" & b & buttonType & ";") = 0) Then
							animateString = animateString & divRow & "Part" & ColX & "_" & b & buttonType & ";" 
						End If
						If (InStr(animateString, "Part" & ColX & "_" & b + a & buttonType & ";") = 0) Then
							animateString = animateString & divRow & "Part" & ColX & "_" & b + a & buttonType & ";"
						End If
						checkFailHistory = "Yellow"
					End If
				End If
			Next
		End If
	Next
 End Function
 
Function toleranceCheck(serial_number, objCmd)
	Dim n, rs ,j
	Dim toleranceArray(17)
	Dim failString : failString = ""
	Dim failMode : failMode = ""
	Dim CMMExist : CMMExist = False
	
	Dim sqlQuery : sqlQuery = "SELECT TOP 1 [Dim 1_1], [Dim 1_2], [Dim 2_1], [Dim 2_2], [Dim 3_1], [Dim 3_2], [Dim 4_1], [Dim 4_2], [Dim 5_1], [Dim 5_2], " _
							& "[Dim 9_1], [Dim 9_2], [Dim 10_1], [Dim 10_2], [Dim 11 Max], [Dim 11 Min], [Dim 12 Max], [Dim 12 Min]" _
							& " FROM [40_CMM_LPT5] " _
							& " WHERE [Serial Number]='" & serial_number & "' ORDER BY [Date] DESC;"
	set rs = objCmd.Execute(sqlQuery)
	DO WHILE NOT rs.EOF
		CMMExist = True
		For j = 0 to 17
			If Not IsNull(rs.Fields(j)) Then toleranceArray(j) = rs.Fields(j)
		Next
		rs.MoveNext
	Loop
	If CMMExist = True Then
		toleranceCheck = "Pass"
		For n = lbound(toleranceArray) to ubound(toleranceArray)
			If IsNull(toleranceArray(n)) or IsEmpty(toleranceArray(n)) Then
			ElseIf toleranceArray(n) < minTol(n) or toleranceArray(n) > maxTol(n) Then
				toleranceCheck = "Fail"
				failString = failString & tolName(n) & ": " & toleranceArray(n) & " (" &  minTol(n) & " to " & maxTol(n) & "), "
				failMode = failMode & tolName(n) & ";"
				allCMMArray(n) = allCMMArray(n) + 1
			End If
		Next
		If toleranceCheck = "Fail" Then failString = Left(failString, len(failString) - 2)
	Else
		toleranceCheck = "Missing"
		failString = ""
		failMode = ""
	End If
	toleranceCheck = Array(toleranceCheck, failString, failMode)
 End Function

Function addReason()
	Dim objCmd : set objCmd = GetNewConnection
	Dim sqlQuery : sqlQuery = "INSERT INTO [30_Reason] ([OperName], [Reason], [SNs], [ReasonDate]) "
	Dim SN_IDs, SN_ID, rs
	If objCmd is Nothing Then : Exit Function
	SN_IDs = Replace(SNIDs.innerHTML, "<BR>", ";") & ";"
	sqlQuery = sqlQuery & "VALUES ('" & opNameInput.value & "', '" & reasonInput.value & "', '" & SN_IDs & "', '" & Now & "');"
	set rs = objCmd.Execute(sqlQuery)
	opNameInput.value = ""
	reasonInput.value = ""
	submitText.value = ""
	SNIDs.innerHTML = ""
	objCmd.Close
	Set objCmd = Nothing
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

'// EXIT SCRIPT
Sub ServerClose()
	If debugMode = False Then On Error Resume Next

	WScript.Sleep 1000  '// REQUIRED OR ERRORS
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"
	
	close
	
	On Error GoTo 0
	Self.Close()
 End Sub


'Function to create all of the JS and HTML code for the window
Function LoadHTML(sBgColor)
	Dim a : a = 0
	Dim b : b = 0
	Dim monitorText : monitorText = Array("Abrasive", "Mixing Tube", "Orifice", "Last Offset")
	
	'CSS String
	LoadHTML = LoadHTML _	
		& "<head><style>" _
		& "body {" _
			& "background-color: " & sBgColor & ";" _
			& "font:normal 20px Tahoma;" _
			& "border-Style:outset" _
			& "border-Width:3px" _
			& "}" _
		& ".CMM {" _
			& "font:normal " & int(MachineColWidth / 10) & "px Tahoma;" _
			& "}" _
		& ".divOutline {" _
			& "border-style: solid;" _
			& "border-Width:1px;" _
			& "}" _
		& ".machineHead {" _
			& "font:normal 20px Tahoma;" _
			& "}" _
		& ".machineSummary {" _
			& "font:normal 18px Tahoma;" _
			& "}" _
		& ".statusSummary {" _
			& "font:normal 16px Tahoma;" _
			& "border-style: solid;" _
			& "border-Width:1px;" _
			& "}" _
		& ".locBar {" _
			& "border-right: 5px solid red;" _
			& "}" _
		& ".machineText {" _
			& "font:normal " & int(MachineColWidth / 5) & "px Tahoma;" _
			& "}" _
		& ".fixtureText, .sumText {" _
			& "font:normal " & int(MachineColWidth / 10) & "px Tahoma;" _
			& "}" _
		& ".dailyValueText, .dailyText {" _
			& "font:normal " & int(MachineColWidth / 12) & "px Tahoma;" _
			& "}" _
		& ".legendText {" _
			& "font:normal 15px Tahoma;" _
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
		& ".opButton, .closeButton, .historyButton {" _
			& "height: 30px;" _
			& "width: 30px;" _
			& "font: 20px;" _
			& "}" _
		& ".opButton {" _
			& "background-color: blue;" _
			& "color: white;" _
			& "}" _
		& ".closeButton {" _
			& "background-color: red;" _
			& "color: white;" _
			& "}" _
		& ".historyButton {" _
			& "background-color: green;" _
			& "color: white;" _
			& "padding: -10px 0px 10px 0px;" _
			& "}" _
		& ".modal {" _
			& "background-color: red;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
			& "}" _
		& "#MRBModal {" _
			& "font:normal 20px Tahoma;" _
			& "background-color = 'grey';" _
			& "}" _
		& "#SNIDs, .fixtureDiv {" _
			& "overflow-y: scroll;" _
			& "}" _
		& "div, span{" _
			& "position:absolute;" _
			& "}" _
		& ".dayDiv{" _
			& "position:relative;" _
			& "display:block;" _
			& "}" _
		& ".dayColDiv{" _
			& "display:block;" _
			& "}"
	If adminMode = true Then
		LoadHTML = LoadHTML _
		& "div, span{" _
			& "border-style: solid;" _
			& "border-Width:1px;" _
			& "}"
	Else
		LoadHTML = LoadHTML _
		& "#MRBModal {" _
			& "visibility: hidden;" _
			& "}"
	End If
	LoadHTML = LoadHTML _
		& "</style>"

	'Body Start String							
	LoadHTML = LoadHTML & "<div unselectable='on' class='unselectable divOutline' style='top:0px; left:1670px; height: " & HTAHeight & "px; width: 250px; text-align: center; background-color:WhiteSmoke;'></div>" _
		
	For a = 0 to ColCount + 1
		If a Mod 2 = 0 Then
			LoadHTML = LoadHTML _	
				& "<div unselectable='on' class='unselectable divOutline' style='top:  0px; left: " & a * MachineColWidth -  10 & "px; height: " & ColFooter + 30 _
				& "px; width: " & MachineColWidth & "px; text-align: center; background-color:AliceBlue;'></div>"
		Else
			LoadHTML = LoadHTML _	
				& "<div unselectable='on' class='unselectable divOutline' style='top:  0px; left: " & a * MachineColWidth -  10 & "px; height: " & ColFooter + 30 _
				& "px; width: " & MachineColWidth & "px; text-align: center; background-color:LightCyan;'></div>"
		End If
		LoadHTML = LoadHTML _	
			& "<div unselectable='on' class='unselectable machineText' 	style='top:  5px; left: " & a * MachineColWidth & "px; height: 60px; width: " & MachineColWidth _
			& "px; text-align: center;' id=machineCol" & a & ">0-0</div>"
	Next
		
	For a = 0 to ColCount
		LoadHTML = LoadHTML _	
			& "<div unselectable='on' class='unselectable sumText' style='top: " & ColFooter & "px; left: " & a * MachineColWidth - 5 & "px;height: 25px; width: " _
			& Int(MachineColWidth * .9) & "px; text-align: center;' id=mon3Cnt" & a & "Text>0</div>"
		For b = 7 to 8
			LoadHTML = LoadHTML _	
				& "<div unselectable='on' class='unselectable sumText' style='top: " & b * 25 + ColFooter - 15 & "px; left: " & a * MachineColWidth  & "px;  height: 25px; width: " & Int(MachineColWidth * .60) _
				& "px; text-align: right;'>" & monitorText(b - 6) & "&nbsp;</div>" _
				& "<div unselectable='on' class='unselectable sumText' style='top: " & b * 25 + ColFooter - 15 & "px; left: " & a * MachineColWidth + Int(MachineColWidth * .60) & "px;height: 25px; width: " _
				& Int(MachineColWidth * .35) & "px; text-align: center;' id=mon" & b - 6 & "Cnt" & a & "Text>0</div>"
		Next
	Next
	
	'Machine Table String
	LoadHTML = LoadHTML _	
		& "<div id=machineDiv style='top: " & footerTop & "px; left: -1px; height: 350px; width: 802px; background-color:WhiteSmoke;' class='divOutline'>" _
			& "<div unselectable='on' class='unselectable machineHead' style='top: 15px; left: 0px;  height: 25px; width: 150px; text-align: center;'>Machine Name</div>"  _
			& "<div unselectable='on' class='unselectable machineHead' style='top: 5px; left: 150px;  height: 25px; width: 80px; text-align: center;'>Mixing Tube</div>"   _
			& "<div unselectable='on' class='unselectable machineHead' style='top: 15px; left: 230px;  height: 25px; width: 80px; text-align: center;'>Orifice</div>" _
			& "<div unselectable='on' class='unselectable machineHead' style='top: 5px; left: 310px;  height: 25px; width: 80px; text-align: center; visibility:hidden;'>Run Time</div>"
		
		For a = 0 to machineCount - 1
			LoadHTML = LoadHTML _	
				& "<div unselectable='on' class='unselectable machineSummary' style='top: " & 55 + a * 30 & "px; left: 0px;  height: 25px; width: 150px; text-align: center;'>" & machineNomenArray(a) & "</div>" _
				& "<div unselectable='on' class='unselectable machineSummary' style='top: " & 55 + a * 30 & "px; left: 150px; height: 25px; width: 80px; text-align: center;' id=MachMix" & a & ">0</div>" _
				& "<div unselectable='on' class='unselectable machineSummary' style='top: " & 55 + a * 30 & "px; left: 230px; height: 25px; width: 80px; text-align: center;' id=MachOri" & a & ">0</div>" _
				& "<div unselectable='on' class='unselectable machineSummary' style='top: " & 55 + a * 30 & "px; left: 310px; height: 25px; width: 80px; text-align: center;' id=MachRun" & a & "></div>"
		Next
		
	LoadHTML = LoadHTML _	
		& "</div>"
	
	'Status Div String
	LoadHTML = LoadHTML _	
		& "<div class='unselectable divOutline' id=statusDiv style='top: " & footerTop & "px; left: 500px; height: 350px; width: 895px; background-color:LightGrey;' class='divOutline'>" _
			& "<div class='unselectable machineHead' style='top: 10px; left: 150px; height: 50px; width: 230px; text-align: center;'>1st Pass Yield (&#8805; 95%)</div>" _
				& "<div class='unselectable machineHead' style='top: 60px; left: 150px; height: 50px; width: 230px; text-align: center;' id=totalYield></div>" _
			& "<div class='unselectable machineHead' style='top: 10px; left: 380px; height: 50px; width: 200px; text-align: center;'>Run Time (&#8805; 80%)</div>" _
				& "<div class='unselectable machineHead' style='top: 60px; left: 380px; height: 50px; width: 200px; text-align: center;' id=totalRunTime></div>" _
			& "<!--<div class='unselectable machineHead' style='top: 10px; left: 580px; height: 50px; width: 250px; text-align: center;'>Past Due</div>" _
				& "<div class='unselectable machineHead' style='top: 60px; left: 580px; height: 50px; width: 250px; text-align: center;' id=pastDue></div>-->" _
			& "<div class='unselectable machineHead' style='top: 160px; left: 10px; height: 50px; width: 100px; text-align: center;'>Shipping</div>"_	
				& "<div class='unselectable' style='top: 150px; left: 110px; height: 50px; width: 700px; background-color:WhiteSmoke;'></div>"_
				& "<div class='unselectable' style='top: 150px; left: 110px; height: 50px; width: 0px; background-color:Blue;' id=shipBar></div>"_
				& "<div class='unselectable' style='top: 160px; left: 810px; height: 50px; width: 68px; text-align: center;' id=shipText>0</div>"_
				& "<div class='unselectable' style='top: 195px; left: 110px; height:  5px; width: 0px; background-color:cyan; font:1px;' id=shipBarPrev></div>"_
				& "<div class='unselectable' style='top: 190px; left: 810px; height: 50px; width: 68px; text-align: center; font:12px;' id=shipTextPrev>0</div>"_
			& "<div class='unselectable machineHead' style='top: 220px; left: 10px; height: 50px; width: 100px; text-align: center;'>Production</div>"_
				& "<div class='unselectable' style='top: 210px; left: 110px; height: 50px; width: 700px; background-color:WhiteSmoke;'></div>"_
				& "<div class='unselectable' style='top: 210px; left: 110px; height: 50px; width: 0px; background-color:Blue;' id=prodBar></div>"_
				& "<div class='unselectable' style='top: 220px; left: 810px; height: 50px; width: 68px; text-align: center;' id=prodText>0</div>"_
				& "<div class='unselectable' style='top: 255px; left: 110px; height:  5px; width: 0px; background-color:cyan; font:1px;'' id=prodBarPrev></div>"_
				& "<div class='unselectable' style='top: 250px; left: 810px; height: 50px; width: 68px; text-align: center; font:12px;' id=prodTextPrev>0</div>"_
			& "<div class='unselectable statusSummary' style='top: 121px; left: 110px; height: 140px; width: 125px; text-align: center;'>Monday</div>"_
			& "<div class='unselectable statusSummary' style='top: 121px; left: 234px; height: 140px; width: 125px; text-align: center;'>Tuesday</div>"_
			& "<div class='unselectable statusSummary' style='top: 121px; left: 358px; height: 140px; width: 125px; text-align: center;'>Wednesday</div>"_
			& "<div class='unselectable statusSummary' style='top: 121px; left: 482px; height: 140px; width: 125px; text-align: center;'>Thursday</div>"_
			& "<div class='unselectable statusSummary' style='top: 121px; left: 606px; height: 140px; width: 69px; text-align: center;'>Friday</div>"_
			& "<div class='unselectable statusSummary' style='top: 121px; left: 674px; height: 140px; width: 69px; text-align: center;'>Saturday</div>"_
			& "<div class='unselectable statusSummary' style='top: 121px; left: 742px; height: 140px; width: 69px; text-align: center;'>Sunday</div>"_
			& "<div class='unselectable locBar' style='top: 150px; left: 110px; height:110px; width: 0px;' id=currentLocation></div>" _
			& "<div class='unselectable' style='top: 275px; left: 110px; height: 30px; width: 100px; background-color:Blue;'></div>"_
				& "<div class='unselectable' style='top: 278px; left: 215px; height: 25px; width: 100px;'>This Week</div>"_
			& "<div class='unselectable' style='top: 320px; left: 110px; height:  5px; width: 100px; background-color:cyan; font:1px;'></div>"_
				& "<div class='unselectable' style='top: 310px; left: 215px; height:  25px; width: 100px;'>Last Week</div>"_
			& "<div class='unselectable locBar' style='top: 275px; left: 375px; height:25px; width: 5px;'></div>" _
				& "<div class='unselectable' style='top: 275px; left: 385px; height:  25px; width: 100px;'>Now</div>"_
		& "</div>"
		
	'Legend Table String	
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable divOutline' style='top:" & footerTop & "px; left:1390px; height: 350px; width: 282px; background-color:WhiteSmoke;'>" _
		_
		& "<div unselectable='on' class='unselectable legendText' style='top: 10px; left: 35px; height: 20px; width: 210px;'>Passed CMM</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 10px; left: 5px; height: 20px; width: 20px;'>" _
			& "<button style='height: 20px; width: 20px; background-color: limegreen;' disabled></button></div>" _
		& "<div unselectable='on' class='unselectable legendText' style='top: 40px; left: 35px; height: 20px; width: 210px;'>Failed CMM</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 40px; left: 5px;height: 20px; width: 20px;'>" _
			& "<button style='height: 20px; width: 20px; background-color: red;' disabled></button></div>" _
		& "<div unselectable='on' class='unselectable legendText' style='top: 70px; left: 35px; height: 20px; width: 210px;'>Failed CMM, correction made</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 70px; left: 5px;height: 20px; width: 20px;'>" _
			& "<button style='height: 20px; width: 20px; background-color: blue;' disabled></button></div>" _
		& "<div unselectable='on' class='unselectable legendText' style='top: 100px; left: 35px; height: 20px; width: 210px;'>Failed CMM, needs correction</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 100px; left: 5px;height: 20px; width: 20px;'>" _
			& "<button id='flashButtonLegend' style='height: 20px; width: 20px; background-color: yellow;' disabled></button></div>" _
		& "<div unselectable='on' class='unselectable legendText' style='top: 130px; left: 35px; height: 20px; width: 210px;'>Missing CMM Data</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 130px; left: 5px;height: 20px; width: 20px;'>" _
			& "<button style='height: 20px; width: 20px;' disabled></button></div>" _
		_
		& "<div unselectable='on' class='unselectable' style='top: 155px; left: 5px; height: 20px; width: 20px;'>&#10004;</div>" _
		& "<div unselectable='on' class='unselectable legendText' style='top: 160px; left: 35px; height: 20px; width: 250px;'>1st part after an offset was made</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 190px; left: 5px; height: 20px; width: 20px; background-color: Orange;'>&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable legendText' style='top: 190px; left: 35px; height: 20px; width: 245px;'>1st part with a new mixing tube</div>" _
		_
		& "<div unselectable='on' class='unselectable' style='top: 310px; left: 5px;height: 20px; width: 20px;'>" _
			& "<button style='height: 20px; width: 20px; background-color=grey;' onclick='javascript:showChart();' disabled id=chartButton></button></div>" _
		& "<div unselectable='on' class='unselectable legendText' style='top: 310px; left: 35px; height: 20px; width: 200px;'>Show/Hide Pareto chart</div></div>"
		
	'Scrolling Div Strings
	LoadHTML = LoadHTML _
		& "<div unselectable='on' class='unselectable fixtureDiv' id=buttonDiv style='top:75px; left:0px; height: " & ColFooter - 82 & "px; width: " & HTAWidth - 3 & "px;'>"
		For a = 0 to CMMHistory + 1
			LoadHTML = LoadHTML & "<div unselectable='on' class='unselectable dayDiv' style='top:-1px; width: " & HTAWidth - 3 & "px; height: 2px; background-color:blue; font: 1px;'></div>"
			LoadHTML = LoadHTML & "<div unselectable='on' class='unselectable dayDiv' style='width: " & HTAWidth - 3 & "px; height: 0px;' id=divDay" & a & ">"
			LoadHTML = LoadHTML & AddScrollingHTML(a)
			LoadHTML = LoadHTML & "</div>"
		Next
	LoadHTML = LoadHTML & "</div>"
		
	'Error Output String
	LoadHTML = LoadHTML _	
		& "<div id=errorDiv style='top: " & footerTop & "px; left: 1670px; height: 252px; width: 250px; background-color:LightGrey;' class='divOutline'>" _
		& "<div unselectable='on' class='unselectable' style='top: 20px; left: 10px; height: 210px; width: 230px; text-align: center;' id=errorString>&nbsp;</div></div>"
	
	'Counter String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable divOutline' style='top:" & footerTop + 250 & "px; left: 1670px; height: 125px; width: 250px; background-color:WhiteSmoke;'>"_
		& "<div unselectable='on' class='unselectable' style='top: 20px; left: 10px; height: 85px; width: 230px; text-align: center;' id=counterString>&nbsp;</div></div>"
		
	'Close Box String
	LoadHTML = LoadHTML _		
		& "<div id=closeDiv unselectable='on' class='unselectable divOutline' style='top: 0px; left: 1795px; height: 42px; width: 125px; background-color:LightGrey;visibility:hidden;'>" _
		& "<div	unselectable='on' class='unselectable' style='top: 5px; left:  5px;height: 30px; width: 30px;'><button class='historyButton'  title='Change number of days' 	onclick='done.value=""changeHistory""'>&#9473;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 45px;height: 30px; width: 30px;'><button class='opButton' 		title='Open up all operations'	onclick='done.value=""allOps""'>&#10010;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 85px;height: 30px; width: 30px;'><button class='closeButton'   	title='Close script'	onclick='done.value=""cancel""'>&#10006;</button></div></div>" _
		& "<div style='top: 5px; left: 2000px;'><button type=hidden id=returnToHTA 		style='visibility:hidden;' value=false onclick='HTAReturn()'><center><span>&nbsp;</span></button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 0px; left: 1795px; height: 42px; width: 125px;' onmouseover='document.getElementById(""closeDiv"").style.visibility=""visible""'></div>"_
		& "<div style='top: 5px; left: 2000px;'><input type=hidden id=done 				style='visibility:hidden;' value=false><center><span>&nbsp;</span></div>" _
		& "<div style='top: 5px; left: 2000px;'><input type=hidden id=submitButton 		style='visibility:hidden;' value=false><center><span>&nbsp;</span></div>" _
		& "<div style='top: 5px; left: 2000px;'><input type=hidden id=submitText 		style='visibility:hidden;' value=''><center><span>&nbsp;</span></div>"  _
		& "<div style='top: 5px; left: 2000px;'><input type=hidden id=waitForLoop 		style='visibility:hidden;' value=false><center><span>&nbsp;</span></div>" 
		
 End Function

Function AddScrollingHTML(divRow)
	Dim a : a = 0
	Dim b : b = 0
	Dim dayText : dayText = Array("Today", "Yesterday")
	Dim ScrollHTML
	
	'Part Buttons String
	For a = 0 to ColCount + 1
		For b = 1 to RowCount
			ScrollHTML = ScrollHTML _	
				& "<div unselectable='on' class='unselectable dayColDiv' style='left: " & a * MachineColWidth & "px; height: 15px; width: " & .9 * MachineColWidth & "px; display:none;' id='" & divRow & "Part" & a & "_" & b & "Div'>" _
					& "<div unselectable='on' class='unselectable CMM' style='left: 0px; height: 15px; width: " & Int(MachineColWidth * .8) _
						& "px; text-align: left;' id=" & divRow & "Part" & a & "_" & b & "Text>&nbsp;</div>" _
					& "<div unselectable='on' class='unselectable' style='left: " & Int(MachineColWidth * .5) - 15 & "px;height: 15px; width: 15px;'>" _
						& "<button class= 'blank' style='height: 15px; width: 15px;' title='No serial number found' onclick=""javascript:CMMFunction(this.id, this.SNValue);"" id=" & divRow & "Part" & a & "_" & b & "_1Button SNValue='' disabled></button></div>" _
					& "<div unselectable='on' class='unselectable' style='left: " & Int(MachineColWidth * .5) & "px;height: 15px; width: 15px;'>" _
						& "<button class= 'blank' style='height: 15px; width: 15px;' title='No serial number found' onclick=""javascript:CMMFunction(this.id, this.SNValue);"" id=" & divRow & "Part" & a & "_" & b & "_2Button SNValue='' disabled></button></div>" _
					& "</div>"
		Next
		AddScrollingHTML = AddScrollingHTML + ScrollHTML
		ScrollHTML = ""
	Next
 End Function
 
 Function LoadModalHTML()
 
	'Modal MRB Div String
	LoadModalHTML = "<div id='MRBModal' style='top: " & footerTop  & "px; left: 1px; height: 365px; width: 875px;'>" _
					& "<div style='top: 200px; left: 625px; height: 48px; width: 200px;'><input type=button value='Save Changes' style='height: 48px; width: 200px;' onclick='okCorrection()'></div>" _
					& "<div style='top: 275px; left: 625px; height: 48px; width: 200px;'><input type=button value='Cancel' style='height: 48px; width: 200px;' onclick='cancelCorrection()'></div>" 
		
		'Operator Name String
		LoadModalHTML = LoadModalHTML _
			& "<div unselectable='on' class='unselectable' style='top: 20px; left: 20px;height: 30px; width: 160px;'>Operator Name:</div>" _
			& "<div id='opNameFormDiv' style='top: 20px; left: 180px; height: 30px; width: 645px;'>" _
				& "<input id=opNameInput style='top: 0px; left: 0px; height: 30px; width: 645px;' value='' /></div>"
		
		'Reason String
		LoadModalHTML = LoadModalHTML _
			& "<div unselectable='on' class='unselectable' style='top: 75px; left: 20px;height: 30px; width: 160px;'>Correction Made:</div>" _
			& "<div id='reasonFormDiv' style='top: 75px; left: 180px; height: 100px; width: 645px;'>" _
				& "<input id=reasonInput style='top: 0px; left: 0px; height: 100px; width: 645px;' value='' /></div>"
		
		'SN String
		LoadModalHTML = LoadModalHTML _	
			& "<div unselectable='on' class='unselectable' style='top: 210px; left: 25px; height: 120px; width: 150px; text-align: right;'>SN List:&nbsp;</div>" _
			& "<div unselectable='on' class='unselectable' style='top: 210px; left: 175px; height: 120px; width: 225px;' id=SNIDs></div>"
				
		LoadModalHTML = LoadModalHTML _	
		& "</div>" _
		& "<div id='columnchart_values' style='top: 0px; left: 0px; height: " & footerTop - 1 & "px; width: " & HTAWidth & "px; visibility: hidden;'></div>" _
		& "<div id='chart_div' style='top: 0px; left: 0px; height: " & footerTop - 1 & "px; width: " & HTAWidth & "px; visibility: hidden;'></div>"
 
 
 End Function