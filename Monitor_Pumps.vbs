Option Explicit	
	'*******************************************************************
	' The www.iot.shapetechnologies.com needs to be added to the trusted site in ie
	' Internet Options -> Security -> Trusted Sites    : Low
	' Internet Options -> Security -> Internet         : Medium + unchecked Enable Protected Mode
	' Internet Options -> Security -> Restricted Sites : unchecked Enable Protected Mode
	'*******************************************************************
	Const adminMode = true
	Const debugMode = false
	Const dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"

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
	'*********************************************************
	
	Dim WshShell : Set WshShell = CreateObject("WScript.Shell")
	On Error Resume Next
		WshShell.RegWrite "HKCU\Software\Microsoft\Internet Explorer\Styles\MaxScriptStatements", "1107296255", "REG_DWORD"	'Changes security settings to ignore extended script operations
		WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
		WshShell.RegWrite "HKLM\Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\1\1406", 0, "REG_DWORD"	'Changes security settings on ie to allow HTA
		Set WshShell = nothing
	On Error GoTo 0
	
	Call Load_Access
	
	On Error Resume Next
	 Dim wmi : Set wmi = GetObject("winmgmts:root\cimv2") 
	 Dim cProcesses, oProcess : Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%iexplore.exe%'") 
	 For Each oProcess in cProcesses
		oProcess.Terminate()
	 Next
	On Error GoTo 0
	 
	 

	' 'Gets the PID of the internet explorer window, fixes issue with IE.document.focus() moving into the background when the powershell script runs
	' Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%iexplore.exe%'") 
	' Dim iePID : iePID = 0 : For Each oProcess in cProcesses
		' iePID = oProcess.ProcessID
	' Next
	' If iePID = 0 Then
		' Error_File("Error finding PID of iexplore.exe")
		' Wscript.Quit
	' End If
	
	
	
	Dim IE1, IE2, IE3, IE4, IE5, IE6, IE7
	Const winTop = -1080
	Const winLeft = 50
	Const winWidth = 900
	Const winHeight = 350
	Const heightSpace = -5
	Const widthSpace = -12
	
	' Dim nameArray : nameArray = array("AMP 1", "AMP 2", "AMP 5", "AE 3", "AMP 7")
	' Dim ieArray : ieArray = array("https://www.iot.shapetechnologies.com/asset_dashboard/54774590", _
								  ' "https://www.iot.shapetechnologies.com/asset_dashboard/28414785", _
								  ' "https://www.iot.shapetechnologies.com/asset_dashboard/42720835", _
								  ' "https://www.iot.shapetechnologies.com/asset_dashboard/2827073", _
								  ' "https://www.iot.shapetechnologies.com/asset_dashboard/53638573")
						
	
	Dim nameArray : nameArray = array("AMP 1", "AMP 2", "AMP 3", "AMP 4", "AMP 5", "AMP 7")
	Dim ieArray : ieArray = array("https://www.iot.shapetechnologies.com/asset_dashboard/70071741", _
								  "https://www.iot.shapetechnologies.com/asset_dashboard/65645590", _
								  "https://www.iot.shapetechnologies.com/asset_dashboard/46734137", _
								  "https://www.iot.shapetechnologies.com/asset_dashboard/30996124", _
								  "https://www.iot.shapetechnologies.com/asset_dashboard/21802717", _ 
								  "https://www.iot.shapetechnologies.com/asset_dashboard/16741869")
								  
	Dim LeftArray : LeftArray = array(0, winWidth + widthSpace, 					  0,   winWidth + widthSpace,							  0, winWidth + widthSpace)
	Dim TopArray : TopArray =   array(0, 					 0, winHeight + heightSpace, winHeight + heightSpace, (winHeight + heightSpace) * 2, (winHeight + heightSpace) * 2)

	'AMP1
	Set IE1 = CreateObject("InternetExplorer.Application")
	IE1.Visible = True
	Set IE2 = CreateObject("InternetExplorer.Application")
	Set IE3 = CreateObject("InternetExplorer.Application")
	Set IE4 = CreateObject("InternetExplorer.Application")
	Set IE5 = CreateObject("InternetExplorer.Application")
	Set IE6 = CreateObject("InternetExplorer.Application")
	Dim CurrentObj : Set CurrentObj = IE1
	
	CurrentObj.Navigate ("https://www.iot.shapetechnologies.com/users/login")
	Do While (CurrentObj.Busy)
		wsh.sleep 1000
	Loop	
	
	Dim loginHTML : Set loginHTML = IE1.document.getElementByID("loginForm").Children(1).Children(1).Children(0)
	loginHTML.Children(0).Children(1).Children(0).innerText = "czarlengo@flowcorp.com"
	loginHTML.Children(1).Children(1).Children(0).innerText = "Snowball16"
	loginHTML.Children(2).Children(0).Children(1).click
	Do While (CurrentObj.Busy)
		wsh.sleep 1000
	Loop
	Dim reportTime
	Dim n : For n = 1 to 6
		Select Case n
			Case 1
				Set CurrentObj = IE1
			Case 2
				Set CurrentObj = IE2
			Case 3
				Set CurrentObj = IE3
			Case 4
				Set CurrentObj = IE4
			Case 5
				Set CurrentObj = IE5
			Case 6
				Set CurrentObj = IE6
		End Select
	
		CurrentObj.Visible = True
		CurrentObj.Width = winWidth
		CurrentObj.Height = winHeight
		CurrentObj.menubar = 0 
		CurrentObj.toolbar = 0 
		CurrentObj.statusbar = 0 
		CurrentObj.AddressBar = 0 
		CurrentObj.Resizable = 0
		CurrentObj.Left = LeftArray(n - 1) + winLeft
		CurrentObj.Top = TopArray(n - 1) + winTop
		CurrentObj.Navigate ieArray(n - 1)
		Do While (CurrentObj.Busy)
			wsh.sleep 1000
		Loop
		
		CurrentObj.document.getElementByID("navHeader").style.visibility = "hidden"
		CurrentObj.document.getElementByID("pageFooter").style.visibility = "hidden"
		CurrentObj.document.getElementByID("pageHeader").style.top = "0px"
		CurrentObj.document.getElementByID("pageHeader").Children(0).Children(0).innerText = nameArray(n - 1)
		CurrentObj.document.getElementByID("dashboardWrapper").style.top = "-70px"
		CurrentObj.document.getElementByID("dashboardWrapper").style.left = "0px"
		CurrentObj.document.getElementByID("dashboardWrapper").style.position = "absolute"
		CurrentObj.document.getElementsByClassName("mainBody")(0).style.backgroundcolor = "rgb(0, 146, 159)"
		CurrentObj.document.getElementsByClassName("mainBody")(0).style.overflow = "hidden"
		CurrentObj.document.getElementsByClassName("mainBody")(0).style.zoom = "85%"
		
		wsh.sleep 1000
		reportTime = now()
		on Error Resume Next
		reportTime = CDate(CurrentObj.document.getElementByID("reportTime").innerText)
		If reportTime < Date - 1 Then
			CurrentObj.document.getElementsByClassName("mainBody")(0).style.backgroundcolor = "red"
			CurrentObj.document.getElementByID("pageContent").style.backgroundcolor = "red"
			CurrentObj.document.getElementByID("pageHeader").Children(0).Children(0).innerHTML = nameArray(n - 1) & "         <span style='color:red;font-weight: bold;'>PUMP OFFLINE</span>"
		End If
		On Error Goto 0
	Next
	WScript.Quit	
	
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
