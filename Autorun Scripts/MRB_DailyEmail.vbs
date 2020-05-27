Dim THIS_IS_A_PLACEHOLDER 'OPTION EXPLICIT

 '******  CHANGE THESE SETTINGS *********
 Dim adminMode : adminMode = false
 Dim debugMode : debugMode = false
 '***************** Database Settings *******************
 Dim dataSource : dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"
 Dim initialCatalog : initialCatalog = "CMM_Repository"								'Initial database
 Dim AXdataSource : AXdataSource = "PRODSQLAX01\PRODSQLAX01"										'Server location for AX (Backup database used for testing)
 Dim initialAXCatalog : initialAXCatalog = "FLO_US_AX_Prod"							'Initial database
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
 Const searchTime = 10000

 If Not WScript.Arguments.Count = 0 Then
	Dim AXbool
	Dim Arg : For Each Arg In Wscript.Arguments
		  If Arg = "AX" Then AXbool = true
	Next
 End If
 
 If Load_Access and Load_AX Then
	Call CheckAX
	If AXbool = true Then Wscript.Quit
	Call Check_New_Invoice
 End If
 Wscript.Quit
 
Sub Check_New_Invoice()
	Set objCmd = GetNewConnection									'Creates the connection object to the database
	If objCmd is Nothing Then Exit Sub
	Dim sqlString : sqlString = "SELECT [Invoice Number], [Serial Control Date], [Blade Count], [SlugProdID], [Dash1ProdID], [Dash2ProdID] FROM [00_Invoice] WHERE ([SlugProdID] IS NULL OR [Dash1ProdID] IS NULL OR [Dash2ProdID] IS NULL) AND [Received] = 1;"
	Dim rs, rsPN, sqlPNString : set rs = objCmd.Execute(sqlString)
	Dim msgString : msgString = ""
	DO WHILE NOT rs.EOF
		sqlPNString = "SELECT DISTINCT [FIC Slug Part Number], [FIC Blade Part Number] FROM [00_AE_SN_Control] WHERE [Invoice Number] = '" & rs.Fields(0) & "';"
		Set rsPN = objCmd.Execute(sqlPNString)
		DO WHILE NOT rsPN.EOF
			SlugPN = rsPN.Fields(0)
			If Right(rsPN.Fields(1), 1) = "1" Then
				Dash1PN = rsPN.Fields(1)
			Else
				Dash2PN = rsPN.Fields(1)
			End IF
			rsPN.MoveNext
		Loop
		msgString = msgString & "Invoice " & rs.Fields(0) & " has been received and is missing Work Order information."
		If IsNull(rs.Fields(3)) Then
			msgString = msgString & "<br>Please receive " & SlugPN & " slugs, quantity of " & rs.Fields(2) / 2 & "."
		End If
		If IsNull(rs.Fields(4)) or IsNull(rs.Fields(5)) Then
			msgString = msgString & "<br>Please create work orders for " & Dash1PN & " & " & Dash2PN & ", quantity " & rs.Fields(2) / 2 & " each."
		End If
		msgString = msgString & "<br><br>"
		rs.MoveNext
	Loop
	If msgString <> "" Then
		Const Subject = "Contract cutting work orders"
		Const EmailList = "koliver@flowcorp.com;czarlengo@flowcorp.com"
		Dim messageBody : messageBody = "<body><p><span style='font-size:12pt; color:red'>This is an automatically generated email.</span></p><br>" _
			& "<p><span>" & msgString & "</span></p>"
		
		Call Send_Email(messageBody, subject, EmailList, "")
	End If
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing	
 End Sub

Sub CheckAX
	Dim AXrs, rs, updateRS, WONCnt, updateSqlString
	Dim objCmd : Set objCmd = GetNewConnection
	Dim UpdateobjCmd : Set UpdateobjCmd = GetNewConnection
	Dim AXobjCmd : Set AXobjCmd = GetNewAXConnection
 	Dim sqlString : sqlString = "SELECT [Blade Count], [SlugProdID], [Dash1ProdID], [Dash2ProdID], [Invoice Number] FROM [00_Invoice] WHERE [SlugProdID] IS NULL OR [Dash1ProdID] IS NULL OR [Dash2ProdID] IS NULL ORDER BY [Invoice Number];"
	Set rs = objCmd.Execute(sqlString)
	DO WHILE NOT rs.EOF
		WONCnt = Int(rs.Fields(0) / 2)
		sqlString = "SELECT [PRODID] , [ITEMID], [QTYSTUP] FROM [PRODTABLE] WHERE ([ITEMID] = '060052-1' OR [ITEMID] = '060053-1' OR [ITEMID] = '060053-2') AND [CREATEDDATETIME] >= GetDate()-14 AND [QTYSTUP] = " & WONCnt & ";"
		Set AXrs = AXobjCmd.Execute(sqlString)
		DO WHILE NOT AXrs.EOF
			Select Case AXrs.Fields(1)
				Case "060052-1"
					updateSqlString = "SELECT COUNT(*) FROM [00_Invoice] WHERE [SlugProdID] = '" & AXrs.Fields(0) & "';"
					Set updateRS = objCmd.Execute(updateSqlString)
					If updateRS(0).value = 0 Then
						Set updateRS = Nothing
						updateSqlString = "UPDATE [00_Invoice] SET [SlugProdID]='" & AXrs.Fields(0) & "' WHERE [Invoice Number]='" & rs.Fields(4) & "';"
						Set updateRS = objCmd.Execute(updateSqlString)
						Set updateRS = Nothing
					End If
				Case "060053-1"
					updateSqlString = "SELECT COUNT(*) FROM [00_Invoice] WHERE [Dash1ProdID] = '" & AXrs.Fields(0) & "';"
					Set updateRS = objCmd.Execute(updateSqlString)
					If updateRS(0).value = 0 Then
						Set updateRS = Nothing
						updateSqlString = "UPDATE [00_Invoice] SET [Dash1ProdID]='" & AXrs.Fields(0) & "' WHERE [Invoice Number]='" & rs.Fields(4) & "';"
						Set updateRS = objCmd.Execute(updateSqlString)
						Set updateRS = Nothing
					End If
				Case "060053-2"
					updateSqlString = "SELECT COUNT(*) FROM [00_Invoice] WHERE [Dash2ProdID] = '" & AXrs.Fields(0) & "';"
					Set updateRS = objCmd.Execute(updateSqlString)
					If updateRS(0).value = 0 Then
						Set updateRS = Nothing
						updateSqlString = "UPDATE [00_Invoice] SET [Dash2ProdID]='" & AXrs.Fields(0) & "' WHERE [Invoice Number]='" & rs.Fields(4) & "';"
						Set updateRS = objCmd.Execute(updateSqlString)
						Set updateRS = Nothing
					End If
			End Select
			AXrs.MoveNext
		Loop
		rs.MoveNext
	Loop
	objCmd.Close
	Set objCmd = Nothing
	AXobjCmd.Close
	Set AXobjCmd = Nothing
  End Sub
 
Sub Check_String(stringFromScanner)
	Dim objCmd, objAXCmd, userName, searchString, resultQty, sqlString, rs, sqlQTYString, rsQTY, sqlCutString, duplicateCnt, CMM_String
	Dim duplicate : duplicate = false
	Dim changeMade : changeMade = false
	Dim inputString : inputString = TrimString(stringFromScanner)
	Dim fixture : fixture = false
	
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
		If AXResult = false Then
			Dim result : result = InputBox("Choose dash number" & chr(10) & " (1) 060053-1" & chr(10) & " (2) 060053-2", "Work order")
			If result = "" Then
			ElseIf result = 1 Then
				resultQty = InputBox("Enter WO quantity", "Work order quantity")
				windowBox.dash1WO.innerText = inputString
				windowBox.dash1WOQTY.innerText = CInt(resultQty)
			ElseIf result = 2 Then
				resultQty = InputBox("Enter WO quantity", "Work order quantity")
				windowBox.dash2WO.innerText = inputString
				windowBox.dash2WOQTY.innerText = CInt(resultQty)
			End If
		Else
			Set objAXCmd = GetNewAXConnection									'Creates the connection object to the database
			If objAXCmd is Nothing Then AXResult = false : checkAccess : Exit Sub						'If a connection is not found, returns false to the function and then exits
			sqlString = "Select [ITEMID], [QTYCALC] From [PRODTABLE] WHERE [PRODID] = '" & inputString & "';" ' and [PRODSTATUS] <= 4;"
			Set rs = objAXCmd.Execute(sqlString)												'Sends the request to the database
			DO WHILE NOT rs.EOF
				changeMade = true
				If rs.Fields(0) = "060053-1" Then
					windowBox.dash1WO.innerText = inputString
					windowBox.dash1WOQTY.innerText = rs.Fields(1)
				ElseIf rs.Fields(0) = "060053-2" Then
					windowBox.dash2WO.innerText = inputString
					windowBox.dash2WOQTY.innerText = rs.Fields(1)
				End If
				rs.MoveNext
			Loop
			objAXCmd.Close																	'Closes the connection object
			Set objAXCmd = Nothing	
		End If
		If changeMade = true Then
			Call updateWO
		Else
			windowbox.errorDiv.style.background = "red"
			windowBox.errorString.innerText = "Invalid work order"
		End If
	ElseIF Len(inputString) = 10 and Mid(inputString, 9, 1) = "-" and Left(inputString, 1) = "H" Then
		windowBox.errorString.innerText = "CMM started"
		Set objCmd = GetNewConnection
		If objCmd is Nothing Then AccessResult = false : checkAccess : Exit Sub
		If windowBox.operator.innerText = "" or windowBox.operator.innerText = "Not Authorized" Then
			windowbox.errorDiv.style.background = "red"
			windowBox.errorString.innerText = "Missing operator name"
			Exit Sub
		End If
		sqlString = "SELECT TOP 1 [FIC Blade Part Number] FROM [00_AE_SN_Control] WHERE [Blade Serial Number] = '" & inputString & "';"
		Set rs = objCmd.Execute(sqlString)
		DO WHILE NOT rs.EOF
			dashNum = rs.Fields(0)
			If dashNum = "060053-1" Then
				partString = "060053-1"
			ElseIf dashNum = "060053-2" Then
				partString = "060053-2"
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
			ProdID = ""
			If partString  = "060053-1" Then
				prodID = windowBox.dash1WO.innerText
			ElseIf partString = "060053-2" Then
				prodID = windowBox.dash2WO.innerText
			End If
			If ProdID = "" Then
				windowbox.errorDiv.style.background = "red"
				windowBox.errorString.innerText = "Missing work order"
				Exit Sub
			End If
			Set rs = Nothing
		Else
			prodID = "Rerun"
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
		CMM_Array = array(inputString, prodID, partString, windowBox.operator.innerText, windowBox.CMMID.value, CMM_String)
		
		If CMM_Windows(CMM_Array) Then
			Set rs = objCmd.Execute(sqlString)
			windowBox.errorString.innerText = inputString
			windowbox.errorDiv.style.background = "limegreen"
			Call updateWO
		Else
			windowBox.errorString.innerText = "Calibration sequencing failed. Try again."
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
		CMM_Array = array(inputString, false, partString, windowBox.operator.innerText, false, CMM_String)
		
		If CMM_Windows(CMM_Array) Then
			windowBox.errorString.innerText = inputString
			windowBox.errorString.innerText = "Calibration complete"
			windowbox.errorDiv.style.background = "limegreen"
		Else
			windowBox.errorString.innerText = "CMM sequencing failed. Scan part again."
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
	Dim OperatorID 	 : OperatorID 	= CMM_Array(3)
	Dim MachineID 	 : MachineID 	= CMM_Array(4)
	Dim CMM_File	 : CMMFile	= CMM_Array(5)
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
			shl.SendKeys CMM_Array(n)
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

Function GetNewAXConnection()														'Function to connect to AX
	Dim objCmd : Set objCmd = CreateObject("ADODB.Connection")						'Object to connect to the SQL source
	Dim sConnection : sConnection = "Data Source=" & AXdataSource & ";Initial Catalog=" & initialAXCatalog & ";Integrated Security=SSPI;"
	Dim sProvider : sProvider = "SQLOLEDB.1;"										'Connection type 

	objCmd.ConnectionString	= sConnection											'Contains the information used to establish a connection to a data store.
	objCmd.Provider = sProvider														'Indicates the name of the provider used by the connection.
	objCmd.CursorLocation = adOpenStatic											'Sets or returns a value determining who provides cursor functionality.
	On Error Resume Next															'Bypasses any error messages that occur during connection
		objCmd.Open																	'Opens connection to the SQL database
	On Error GoTo 0 																'Resets the error reporting
	If objCmd.State = adStateOpen Then  											'Checks if the connection is open
        Set GetNewAXConnection = objCmd  											'Returns the object to the function
	Else																			'If the connection is not open
        Set GetNewAXConnection = Nothing											'Returns false to the function
    End If  
 End Function 
 
Function Load_AX()																	'Function to check for AX connection and load initial variables
	Dim objCmd : set objCmd = GetNewAXConnection									'Creates the connection object to the database
	If objCmd is Nothing Then Load_AX = false : Exit Function						'If a connection is not found, returns false to the function and then exits
	objCmd.Close																	'Closes the connection object
	Set objCmd = Nothing															'Erases the connection information
	Load_AX = true	
 End Function
 
 
Function Send_Email(Message, subject, EmailTo, EmailBCC)
' exit function
	Dim MyEmail : Set MyEmail=CreateObject("CDO.Message")
	
	Dim Signature : Signature = "<footer><div>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span>&nbsp;</span><br>" _
		& "<span style='font-size:7.5pt;font-family:""Franklin Gothic Medium"",sans-serif; color:teal'>Chris Zarlengo</span><span style='color:#1F497D'></span><br>" _
		& "</div></footer>"
	
	MyEmail.Subject = subject
	MyEmail.From="czarlengo@flowcorp.com"
	MyEmail.To = EmailTo
	MyEmail.BCC = EmailBCC
	MyEmail.HTMLBody = Message & Signature

	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

	'SMTP Server
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="SKENEXC60.flowcorp.com"

	'SMTP Port
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 

	'SMTP Auth (For Windows Auth set this to 2)
	MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate")=2

	MyEmail.Configuration.Fields.Update
	MyEmail.Send

	set MyEmail = nothing


 End Function