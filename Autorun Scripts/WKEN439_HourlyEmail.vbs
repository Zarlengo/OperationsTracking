Option Explicit
 '************** NOTES *****************
 ' v2.0 has updated the script to work with net-inspect v5. Excel files are exported into the CC_Script folder and then uploaded to the SQL with an excel application portion
 '********* VERSION HISTORY ************
 ' 1.0	10/8/2018	Initial Release for production
 ' 1.1	10/20/2018	Removed functions transfered to PowerShell
 ' 1.2	11/08/2018	Added supplier issue identification
 ' 1.3	11/14/2018	Added searching for missing e-tags
 ' 1.4	1/10/2019	Added termination steps for running iexplorer and wscript
 '					Added an auto-email when the termination happens (script must have failed previously)
 '					Changed the net-inspect columns from fixed to variable (they added a new one and it broke the script)
 ' 1.5	2/11/2019	Updated to V5 only script (v4 no longer available)
 '					Added AppActivate() prior to IE.focus(), script would fail when PowerShell script ran and momentarily took focus
 '
 ' 2.0	2/19/2019	Changed to excel export. v5 was too slow and would not update
 ' 2.1	6/28/2019	Added full reset of database to update E-Tags if they were changed, to be done once daily
 ' 2.2	8/1/2019	Added e-tag last edit date to database
 '************** TO DO *****************
 ' 
 '
 '****** CHANGE THESE SETTINGS *********
 Const adminMode = false
 Const debugMode = false
 Dim resetMode : resetMode = false
 
 Const startTagPrefix = "C12"
 Const tabletPassword = "Fl0wSh0p17"
 Const computerPassword = "Snowball18!"
 
 Dim DispositionArray : DispositionArray = Array("E-Tag Open", "Scrap", "Return to Customer", "Supplier Rework/Remake", "Advanced Rejection", "Rework", "in MRB", "Use As Is", "Void")
 Dim StatusArray : StatusArray = Array("Open", "Closed")

 '***************** Database Settings *******************
 Const dataSource = "PRODSQLAPP01\PRODSQLAPP01"
 Const initialCatalog = "CMM_Repository"								'Initial database
 Const MSNL = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Master Serial Number Listing-AeroEdge.xlsx"
 Const MRBL = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\MRB Review\MRB Inventory.xlsx"
 
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
 Dim notFoundCount : notFoundCount = 0
 Dim errorSTring : errorString = ""
 Dim fileCnt : fileCnt = 0
 
 Dim strData, AccessArray, fieldArray(4), fieldsBad, sArg, Arg
 Dim SendData, RecieveData, wmi, cProcesses, oProcess
 Dim machineBox, strSelection, RemoteHost, RemotePort, machineString
 Dim objFolder, colFiles, objFile, objWorkbook,objExcel

 '**************** EXCEL CONSTANTS *******************
 
 Const xlDown				= -4121
 Const xlToLeft				= -4159
 Const xlToRight			= -4161
 Const xlUp					= -4162
 
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
 If Not WScript.Arguments.Count = 0 Then
	sArg = ""
	For Each Arg In Wscript.Arguments
		If Arg = "reset" Then
			resetMode = true
		Else
			sArg = Arg
		End If
	Next
 End If
 If resetMode = true Then resetSQL
 
 If debugMode = False Then On Error Resume Next
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "30", "REG_DWORD"	'Changes TCP timeout settings if needing to restart program w/in 5 minutes
 On Error Goto 0

 'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
 Set wmi = GetObject("winmgmts:root\cimv2") 
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
 For Each oProcess in cProcesses
	errorString = errorString & "MSHTA.exe PID: " & oProcess.ProcessID & "<br>"
	oProcess.Terminate()
 Next
 
 'Checks for existing vbs scripts and terminates them, ignores this script
 Dim sCmd : sCmd = "powershell -command exit (gwmi Win32_Process -Filter \""processid='$PID'\"").parentprocessid"
 Set objShell = CreateObject("WScript.Shell")
 Dim wsPID : wsPID = objShell.Run(sCmd, 0, True)
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%wscript.exe%'") 
 For Each oProcess in cProcesses
	If oProcess.ProcessId <> wsPID Then
		errorString = errorString & "WScript.exe PID: " & oProcess.ProcessID & "<br>"
		oProcess.Terminate()
	End If
 Next
 
 'Checks for existing internet explorer instances and terminates any open ones
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%excel.exe%'") 
 For Each oProcess in cProcesses
	errorString = errorString & "Excel.exe PID: " & oProcess.ProcessID & "<br>"
	oProcess.Terminate()
 Next
 
 'Checks for existing internet explorer instances and terminates any open ones
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%iexplore.exe%'") 
 For Each oProcess in cProcesses
	errorString = errorString & "IExplorer.exe PID: " & oProcess.ProcessID & "<br>"
	oProcess.Terminate()
 Next
 
 If errorString <> "" Then Error_File(errorString)
'Calls function to create ie window
	
	'Function to check for access connection and load info from database
	If Load_Access = false Then WScript.Quit
	
	'Deletes all XLSX files in ie temp folder and CC_Script folder
	Const IEFiles = "C:\Users\CZarlengo\AppData\Local\Microsoft\Windows\Temporary Internet Files\Content.IE5"
	Const CCFiles = "C:\CC_Script"
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objSuperFolder : Set objSuperFolder = objFSO.GetFolder(IEFiles)
	Dim objCCFolder : Set objCCFolder = objFSO.GetFolder(CCFiles)
	If adminMode = false Then
		Call deleteXLSFiles(objSuperFolder)
		Call deleteXLSFiles(objCCFolder)
	End If
	
	Dim IE : Set IE = CreateObject("InternetExplorer.Application")
	IE.Visible = True

	'Gets the PID of the internet explorer window, fixes issue with IE.document.focus() moving into the background when the powershell script runs
	Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%iexplore.exe%'") 
	Dim iePID : iePID = 0 : For Each oProcess in cProcesses
		iePID = oProcess.ProcessID
	Next
	If iePID = 0 Then
		Error_File("Error finding PID of iexplore.exe")
		Wscript.Quit
	End If
	IE.Navigate ("https://www.net-inspect.com/default.aspx")
	Do While (IE.Busy)
		wscript.sleep 200
	Loop
	objShell.AppActivate(iePID)
	IE.document.focus()	
	IE.document.getElementByID("UserID").innerText = "chris zarlengo"
	IE.document.getElementByID("Password").innerText = "UyE2fDENol2I"
	IE.document.getElementByID("CompanyName").innerText = "Flow"
	IE.document.getElementByID("Submit").click
	wscript.sleep 1000
	Do While (IE.Busy)
	   wscript.sleep 200
	Loop
	wscript.sleep 5000
	Call ExportETags
	wscript.sleep 2000
	Call searchForXLSFiles(objSuperFolder)
	IE.Quit
	
	Set objExcel = CreateObject("Excel.Application")
	Set objFolder = objFSO.GetFolder(objCCFolder.Path)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		If UCase(objFSO.GetExtensionName(objFile.name)) = "XLSX" Then
			Set objWorkbook = objExcel.Workbooks.Open(objFile.path, , True)   'true here means readonly=yes.
			If adminMode = false Then
				objExcel.Application.Visible = False
			Else
				objExcel.Application.Visible = True
			End If
			Call UpdateETags			'Goes through all open e-tags
			objWorkbook.Close False 
		End If
	Next
	objExcel.Application.Quit 'close excel
	If adminMode = false Then Call deleteXLSFiles(objCCFolder)
	Call UpdateRejections
	Call UpdateSupplier
 ServerClose()																	'Function to close open connections and return settings back to original	
 Wscript.Quit
 
Sub resetSQL()
	Dim objCmd : set objCmd = GetNewConnection
	Dim sqlString : sqlString = "UPDATE [40_Rejections] SET [Summary Status] = 'Open';"
	Dim rs : set rs = objCmd.Execute(sqlString)	
	Set rs = Nothing
	sqlString = "UPDATE [40_E-Tags] SET [Status]= 'Open';"
	set rs = objCmd.Execute(sqlString)
	Set rs = Nothing
 End Sub
 
Sub ExportETags()	
	Dim shl : set shl = createobject("wscript.shell")
	IE.Navigate ("https://v5.net-inspect.com/QualityManagement/eTags/")
	Call waitV5LoadPage
	objShell.AppActivate(iePID)
	IE.document.focus()
	IE.document.getElementByID("closedDate").click
	Call waitV5LoadPage
	Dim objCmd : set objCmd = GetNewConnection
	Dim sqlString : sqlString = "Select TOP 1 [Tag Number] FROM [40_E-Tags] WHERE STATUS = 'Open' ORDER BY [Tag Number] ASC;"
	Dim rs : set rs = objCmd.Execute(sqlString)	
	Dim TagFound : TagFound = ""
	DO WHILE NOT rs.EOF
		TagFound = rs.Fields(0)
		rs.MoveNext
	Loop
	Set rs = Nothing
	Dim tagStart : tagStart = mid(TagFound, 2, 2)
	
	sqlString = "Select TOP 1 [Tag Number] FROM [40_E-Tags] ORDER BY [Tag Number] DESC;"
	set rs = objCmd.Execute(sqlString)	
	TagFound = ""
	DO WHILE NOT rs.EOF
		TagFound = rs.Fields(0)
		rs.MoveNext
	Loop
	Set rs = Nothing
	Dim tagEnd : tagEnd = Mid(TagFound,2,2) + 1
	
	Call waitV5LoadPage
	Dim NewTag : NewTag = get1stTag
	Dim tableChildren : Set tableChildren = IE.document.getElementsByClassName("etag-list")(0).children(0).children(0).children(0).children(1).children(0).Children
	Dim tagID : tagID = tableChildren(0).id
	Const timeStart = 5000
	Const sleepMS = 200
	
	Call waitV5LoadPage
	IE.document.getElementsByClassName("btn-warning")(0).click
	Call waitV5LoadPage
	
	Dim tagNum, tagPrefix : For tagNum = tagStart to tagEnd
		objShell.AppActivate(iePID)
		tagPrefix = "C" & tagNum
		IE.document.getElementByID(tagID).Children(0).click
		Call waitV5LoadPage
		IE.document.getElementsByClassName("k-textbox")(0).value = ""
		IE.document.getElementsByClassName("k-textbox")(0).focus()
		shl.SendKeys tagPrefix
		shl.SendKeys "~"
		wscript.sleep 200
		Call waitV5LoadPage
		wscript.sleep 200
		IE.document.getElementsByClassName("action-area-list")(0).children(0).children(1).children(1).click
		wscript.sleep 1000
		Call waitV5LoadPage
	Next
 End Sub


Sub deleteXLSFiles(fFolder)
	Set objFolder = objFSO.GetFolder(fFolder.Path)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		If UCase(objFSO.GetExtensionName(objFile.name)) = "XLSX" Then
			objFSO.DeleteFile objFile.Path, true
		End If
	Next

	Dim Subfolder : For Each Subfolder in fFolder.SubFolders
		Call deleteXLSFiles(Subfolder)
	Next
 End Sub

Sub searchForXLSFiles(fFolder)
	Set objFolder = objFSO.GetFolder(fFolder.Path)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		If UCase(objFSO.GetExtensionName(objFile.name)) = "XLSX" Then
			objFSO.MoveFile objFile.Path, "C:\CC_Script\NetInspect_" & fileCnt &  ".XLSX"
			fileCnt = fileCnt + 1
		End If
	Next

	Dim Subfolder : For Each Subfolder in fFolder.SubFolders
		Call searchForXLSFiles(Subfolder)
	Next
 End Sub
 
Function get1stTag()
	get1stTag = ""
	On Error Resume Next
	Do While get1stTag = ""
		wscript.sleep 200
		get1stTag = IE.document.getElementsByClassName("k-selectable")(0).children(1).children(0).children(0).InnerText
	Loop
	On Error GoTo 0

	wscript.sleep 200
 End Function
 
Sub waitV5LoadPage()
	Do While (IE.Busy)
	   wscript.sleep 200
	Loop
	wscript.sleep 500
	Do While (IE.document.getElementByID("atlwdg-blanket").style.visibility = "visible")
	   wscript.sleep 200
	Loop
	Do While (IE.Busy)
	   wscript.sleep 200
	Loop
 End Sub
 
Function UpdateETags()
	'Goes through each row in the file and updates tag dispositions and status
	Dim objSheet : Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	Dim objCmd : set objCmd = GetNewConnection
	Const tagCol = 1
	Const serialCol = 6
	Const defectCol = 8
	Const causeCol = 9
	Const correctCol = 10
	Const dispCol = 11
	Const statCol = 16
	Const operCol = 17
	Const openCol = 21
	Const closeCol = 22
	Dim sqlString, rs, rsUpdate, TagID, Disposition, Status, rejString, Defect, Serial, OpenDate, CloseDate, Operator
	Dim tag, tagString, SNFound, TagsFound, DispFound, StatFound, Change, Cause, Corrective
	Dim SummaryDisp, SummaryStatus, TagNumArr, DispArr, StatArr, n, a
	
	Const startRow = 2
	Dim endRow : endRow = 83112
	endRow = objSheet.cells(endRow,1).End(xlUp).Row
	Dim rowX : For rowX = startRow to endRow
		TagID = objSheet.cells(rowX,tagCol).Value
		Disposition = objSheet.cells(rowX,dispCol).Value
		Status = objSheet.cells(rowX,statCol).Value
		Defect = objSheet.cells(rowX,defectCol).Value
		Serial = objSheet.cells(rowX,serialCol).Value
		OpenDate = objSheet.cells(rowX,openCol).Value
		CloseDate = objSheet.cells(rowX,closeCol).Value
		Operator = objSheet.cells(rowX,operCol).Value
		Cause = objSheet.cells(rowX, causeCol).Value
		Corrective = objSheet.cells(rowX, correctCol).Value
		sqlString = "Select Count(*) FROM [40_E-Tags] WHERE [Tag Number] = '" & TagID & "';"
		Set rs = objCmd.Execute(sqlString)	
		If rs(0).value = 0 Then
			Set rs = Nothing
			sqlString = "INSERT INTO [40_E-Tags] ([Tag Number], [Defect Type], [Disposition], [Status], [Created By], [Open Date], [Close Date], [Cause], [Corrective]) VALUES"
			rejString = "INSERT INTO [40_Rejections] ([Serial Number], [Tag Numbers], [Dispositions], [Status], [Summary Disposition], [Summary Status]) VALUES"
			sqlString = sqlString & " ('" & TagID & "', '" & Defect & "', '" & Disposition & "', '" & Status & "', '" & Operator & "', '" & OpenDate & "', '" & CloseDate & "', '" & Cause & "', '" & Corrective & "'),"
				
			tagString = "Select TOP 1 [Serial Number], [Tag Numbers], [Dispositions], [Status] FROM [40_Rejections] WHERE [Serial Number] = '" & Serial & "';"
			Set rsUpdate = objCmd.Execute(tagString)
			SNFound = ""
			DO WHILE NOT rsUpdate.EOF
				SNFound = rsUpdate.Fields(0)
				TagsFound = rsUpdate.Fields(1)
				DispFound = rsUpdate.Fields(2)
				StatFound = rsUpdate.Fields(3)
				rsUpdate.MoveNext
			Loop
			Set rsUpdate = Nothing
			If SNFound <> "" Then
				'Update
				TagNumArr = Split(TagsFound, ";")
				DispArr = Split(DispFound, ";")
				StatArr = Split(StatFound, ";")
				SummaryDisp = UBound(DispositionArray)
				SummaryStatus = UBound(StatusArray)
				Change = False
				For n = 0 to UBound(TagNumArr)
					If TagNumArr(n) <> "" and TagNumArr(n) <> TagID Then
						Change = True
						For a = 0 to UBound(DispositionArray)
							If InStr(1, DispArr(n), DispositionArray(a)) <> 0 Then Exit For
						Next
						If a > UBound(DispositionArray) Then a = 0
						If SummaryDisp > a Then SummaryDisp = a
						
						For a = 0 to UBound(StatusArray)
							If InStr(1, StatArr(n), StatusArray(a)) <> 0 Then Exit For
						Next
						If a > UBound(StatusArray) Then a = 0
						If SummaryStatus > a Then SummaryStatus = a
					End If
				Next
				If Change = True Then
					Set rsUpdate = objCmd.Execute("UPDATE [40_Rejections] SET [Tag Numbers]='" & TagsFound & TagID & ";', [Dispositions]='" & DispFound & Disposition & ";', [Status]='" & StatFound & Status & ";', [Summary Disposition]='" & DispositionArray(SummaryDisp) & "', [Summary Status]='" & StatusArray(SummaryStatus) & "' WHERE [Serial Number]='" & Serial & "';")
					Set rsUpdate = Nothing
				End If
			Else
				Set rsUpdate = objCmd.Execute(rejString & " ('" & Serial & "', '" & TagID & ";', '" & Disposition & ";', '" & Status & ";', '" & Disposition & "', '" & Status & "');")
				Set rsUpdate = Nothing
			End If
			sqlString = Left(sqlString, len(sqlString) - 1) & ";"
			set rsUpdate = objCmd.Execute(sqlString)	
			Set rsUpdate = Nothing
		Else
			Set rs = Nothing
			sqlString = "Select [Status] FROM [40_E-Tags] WHERE [Tag Number] = '" & TagID & "';"
			Set rs = objCmd.Execute(sqlString)	
			DO WHILE NOT rs.EOF
				If rs.Fields(0) = "Open" or Status <> rs.Fields(0) Then
					sqlString = "UPDATE [40_E-Tags] SET [Disposition]='" & Disposition & "', [Status]='" & Status & "', [Close Date]='" & CloseDate & "' WHERE  [Tag Number]='" & TagID & "';"
					set rsUpdate = objCmd.Execute(sqlString)	
					Set rsUpdate = Nothing
				End If
				rs.MoveNext
			Loop
			Set rs = Nothing
			sqlString = "Select [Tag Number], [Defect Type], [Cause], [Corrective], [Disposition] FROM [40_E-Tags] " & _
								"WHERE ([Cause] is null or [Corrective] is null or " & _
								"[Cause] = '' or [Corrective] = '') and [Status] <> 'Open' and [Tag Number]='" & TagID & "';"
			Set rs = objCmd.Execute(sqlString)	
			DO WHILE NOT rs.EOF
				If InStr(1, UCase(Disposition), "VOID") <> 0 Then
					If Defect = "" Then Defect = "Void"
					If Cause = "" Then Cause = "Void"
					If Corrective = "" Then Corrective = "Void"
				End If
				If Defect <> "" or Cause <> "" or Corrective <> "" Then
					Set rsUpdate = objCmd.Execute("UPDATE [40_E-Tags] SET [Defect Type]='" & Defect & "', [Cause]='" & Cause & "', [Corrective]='" & Corrective & "', [Close Date]='" & CloseDate & "' WHERE [Tag Number]='" & TagID & "';")
					Set rsUpdate = Nothing
				End If
				rs.MoveNext
			Loop
			Set rs = Nothing
		End If
	Next
	objCmd.Close
	Set objCmd = Nothing
 End Function
 
Function UpdateRejections()
	Dim strQuery, CurrentTime, Operator, strQueryPre, sqlString, Duplicate, SCFound, POID, ShipDate, PalletID, BoxID
	Dim ErrorFound : ErrorFound = False
	Dim errorNote : errorNote = ""
	
	On Error GoTo 0
	Dim objCmd : set objCmd = GetNewConnection : If objCmd is Nothing Then WScript.Quit
	
	Dim objDictionary : Set objDictionary = CreateObject("Scripting.Dictionary")
	sqlString = "Select [Tag Number], [Disposition], [Status] From [40_E-Tags];"
	Dim rs : Set rs = objCmd.Execute(sqlString)
	Dim tagArray(1), tagNum, tagString
	DO WHILE NOT rs.EOF
		tagNum = rs.Fields(0)
		tagArray(0) = rs.Fields(1)
		tagArray(1) = rs.Fields(2)
		objDictionary.Add tagNum, tagArray
		rs.MoveNext
	Loop	
	Set rs = Nothing
	sqlString = "Select [Serial Number], [Tag Numbers], [Dispositions], [Status] From [40_Rejections] WHERE [Summary Status] = 'Open';"
	Set rs = objCmd.Execute(sqlString)	
	Dim rsUpdate
	Dim SN, TagNumber, Disposition, Status, SummaryDisp, SummaryStatus, n, Change, TagNumArr, DispArr, StatArr, a, lineString, IfOpen
	Dim lineStart : lineStart = "UPDATE [40_Rejections] SET "
	DO WHILE NOT rs.EOF
		SN = rs.Fields(0)
		TagNumArr = Split(rs.Fields(1), ";")
		DispArr = Split(rs.Fields(2), ";")
		StatArr = Split(rs.Fields(3), ";")
		Change = False
		TagNumber = ""
		Disposition = ""
		Status = ""
		IfOpen = True
		lineString = ""
		SummaryDisp = UBound(DispositionArray)
		SummaryStatus = UBound(StatusArray)
		For n = 0 to UBound(TagNumArr)
			tagNum = TrimString(TagNumArr(n))
			If objDictionary.Exists(tagNum) Then
				tagArray(0) = objDictionary(tagNum)(0)
				tagArray(1) = objDictionary(tagNum)(1)
				If DispArr(n) <> tagArray(0) or StatArr(n) <> tagArray(1) or resetMode = true Then
					Change = True
				End If
				TagNumber = TagNumber & tagNum & ";"
				Disposition = Disposition & tagArray(0) & ";"
				Status = Status & tagArray(1) & ";"
				If tagArray(1) <> "Closed" Then
					IfOpen = False
					SummaryDisp = 0
					SummaryStatus = 0
				End If
				For a = 0 to UBound(DispositionArray)
					If InStr(1, Disposition, DispositionArray(a)) <> 0 Then Exit For
				Next
				If a > UBound(DispositionArray) Then a = 0
				If SummaryDisp > a Then SummaryDisp = a
				
				For a = 0 to UBound(StatusArray)
					If InStr(1, Status, StatusArray(a)) <> 0 Then Exit For
				Next
				If a > UBound(StatusArray) Then a = 0
				If SummaryStatus > a Then SummaryStatus = a
			End If
		Next
		If Change = True Then
			sqlString = lineStart & "[Tag Numbers]='" & TagNumber & "', [Dispositions]='" & Disposition & "', [Status]='" & Status & _
				"', [Summary Disposition]='" & DispositionArray(SummaryDisp) & "', [Summary Status]='" & StatusArray(SummaryStatus) & "' WHERE [Serial Number]='" & SN & "';"
			Set rsUpdate = objCmd.Execute(sqlString)
			Set rsUpdate = Nothing
		ElseIf IfOpen = False Then
			sqlString = lineStart & "[Summary Disposition]='E-Tag Open' WHERE [Serial Number]='" & SN & "';"
			Set rsUpdate = objCmd.Execute(sqlString)
			Set rsUpdate = Nothing
		End If
		rs.MoveNext
	Loop
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
 End Function
 
Function UpdateSupplier()
	Dim strQuery, CurrentTime, Operator, strQueryPre, sqlString, Duplicate, SCFound, POID, ShipDate, PalletID, BoxID, supplier
	Dim ErrorFound : ErrorFound = False
	Dim errorNote : errorNote = ""
	
	On Error GoTo 0
	Dim objCmd : set objCmd = GetNewConnection : If objCmd is Nothing Then WScript.Quit
	
	Dim objDictionary : Set objDictionary = CreateObject("Scripting.Dictionary")
	sqlString = "Select [Tag Number], [Defect Type] From [40_E-Tags];"
	Dim rs : Set rs = objCmd.Execute(sqlString)
	Dim tagArray(2), tagNum, tagString
	DO WHILE NOT rs.EOF
		tagNum = rs.Fields(0)
		tagArray(0) = rs.Fields(1)
		objDictionary.Add tagNum, tagArray
		rs.MoveNext
	Loop	
	Set rs = Nothing
	
	sqlString = "Select [Serial Number], [Tag Numbers] From [40_Rejections] WHERE [Supplier Issue] is null and [Summary Status] = 'Closed';"
	Set rs = objCmd.Execute(sqlString)	
	Dim rsUpdate
	Dim SN, TagNumber, Disposition, Status, SummaryDisp, SummaryStatus, n, Change, TagNumArr, DispArr, StatArr, a, lineString, IfOpen
	DO WHILE NOT rs.EOF
		SN = rs.Fields(0)
		TagNumArr = Split(rs.Fields(1), ";")
		Change = False
		TagNumber = ""
		supplier = 0
		For n = 0 to UBound(TagNumArr)
			tagNum = TrimString(TagNumArr(n))
			If objDictionary.Exists(tagNum) Then
				tagArray(0) = objDictionary(tagNum)(0)
				If Instr(1, tagArray(0), "Supplier") <> 0 or supplier = 1 Then
					supplier = 1
				Else
					supplier = 0
				End If
			End If
		Next
		sqlString = "UPDATE [40_Rejections] SET [Supplier Issue]=" & supplier & " WHERE [Serial Number]='" & SN & "';"
		Set rsUpdate = objCmd.Execute(sqlString)
		Set rsUpdate = Nothing
		rs.MoveNext
	Loop
	Set rs = Nothing
	objCmd.Close
	Set objCmd = Nothing
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


'// EXIT SCRIPT
Sub ServerClose()
	If debugMode = False Then On Error Resume Next

	WScript.Sleep 1000  '// REQUIRED OR ERRORS
	objShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 1, "REG_DWORD"
	objShell.RegWrite "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\TcpTimedWaitDelay", "240", "REG_DWORD"

	 'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
	Set wmi = GetObject("winmgmts:root\cimv2") 
	Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
	For Each oProcess in cProcesses
		oProcess.Terminate()
	Next

	'Checks for existing internet explorer instances and terminates any open ones
	Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%excel.exe%'") 
	For Each oProcess in cProcesses
		oProcess.Terminate()
	Next

	'Checks for existing internet explorer instances and terminates any open ones
	Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%iexplore.exe%'") 
	For Each oProcess in cProcesses
		oProcess.Terminate()
	Next
	On Error GoTo 0
    Wscript.Quit
 End Sub

Sub Error_File(Message)
	Dim objFSO : Set objFSO=CreateObject("Scripting.FileSystemObject")
	Dim fileName : fileName = "Rejections_" & Year(date) & Month(date) & Day(date) & "_" & Int(Timer()) & ".txt"
	
	' How to write file
	Dim outFile : outFile="G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\TXTFiles\" & fileName
	Dim objFile : Set objFile = objFSO.CreateTextFile(outFile,True)
	objFile.Write Message
	objFile.Close
 End Sub
 