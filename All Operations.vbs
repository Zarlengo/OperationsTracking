Option Explicit
 '// Script to provide a single source access to every file used in production

 '*** KNOWN ISSUE: If an error happens on line of:    .Run    **********
 '*** Windows Defender has a security block for all of these scripts ***
 '*** Contact IT and provide the computer name so they can add it to ***
 '*** the exclusion list in it's security profile **********************

 '********* VERSION HISTORY ************
 ' 1.0	8/10/2018	Initial Release for production
 ' 1.1	9/5/2018	Retrieve script information from the database
 '					Added ME mode (allows using TCPIP Builder to simulate scanner)
 ' 1.2	9/18/2018	Added comments to the code, cleaned up code
 '					Uploaded to repository
 '					Removed extra code
 '					Created reference document All Operations.md to contain code summary
 
 '********** Repository information ************
 ' Files are located at https://gitlab.com/FlowCorp_CC/contract_cutting
 ' Account username is czarlengo@flowcorp.com
 ' Account password is Snowball18!

 '************* CHANGE THESE SETTINGS *********************************************************************************************************
 Dim adminMode : adminMode = false													'Variable to show the borders on elements in the window
 Dim debugMode : debugMode = false													'Variable for not bypassing database operations, used for debugging
 Dim meMode : meMode = false														'Variable to load Manufacturing Engineering computers as a scanner option
 Dim documentTitle : documentTitle = "All Operations"								'IE window title
 
  '***************** Database Settings *******************
 Dim dataSource : dataSource = "PRODSQLAPP01.shapetechnologies.com\PRODSQLAPP01"							'SQL location
 Dim initialCatalog : initialCatalog = "CMM_Repository"								'Initial database
 
 '*********************************************************************************************************************************************
 '**************** INITIAL PARAMETERS *******************
 Dim ScriptHost : ScriptHost = Mid(WScript.FullName, InStrRev(WScript.FullName, "\") + 1, Len(WScript.FullName))
 Dim objShell : Set objShell = CreateObject("WScript.Shell")
 Dim oProcEnv : Set oProcEnv = objShell.Environment("Process")
 Dim LocDict : Set LocDict = CreateObject("Scripting.Dictionary") : LocDict.CompareMode = vbTextCompare
 Dim ArgDict : Set ArgDict = CreateObject("Scripting.Dictionary") :  ArgDict.CompareMode = vbTextCompare

 Dim allOPSsource : allOPSsource = "G:\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\All Operations.vbs"
 Dim sOPsCmd : sOPsCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & allOPSsource & """" & sArg

 Dim closeWindow : closeWindow = false
 Dim errorWindow : errorWindow = false

 Dim strData, windowBox, AccessArray, AccessResult, sArg, Arg
 Dim SendData, RecieveData, wmi, cProcesses, oProcess
 Dim machineBox, strSelection, RemoteHost, RemotePort, machineString
 Dim OPNameArray, OPIDArray, ADNameArray, ADIDArray, windowHeight


'**************** DATABASE CONSTANTS *******************

 Const adOpenDynamic			= 2	 '// Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
 Const adOpenForwardOnly		= 0	 '// Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
 Const adOpenKeyset				= 1	 '// Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
 Const adOpenStatic				= 3	 '// Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.
 Const adOpenUnspecified		= -1 '// Does not specify the type of cursor.

 Const adLockBatchOptimistic	= 4	 '// Indicates optimistic batch updates. Required for batch update mode.
 Const adLockOptimistic			= 3	 '// Indicates optimistic locking, record by record. The provider uses optimistic locking, locking records only when you call the Update method.
 Const adLockPessimistic		= 2	 '// Indicates pessimistic locking, record by record. The provider does what is necessary to ensure successful editing of the records, usually by locking records at the data source immediately after editing.
 Const adLockReadOnly			= 1	 '// Indicates read-only records. You cannot alter the data.
 Const adLockUnspecified		=-1 '// Does not specify a type of lock. For clones, the clone is created with the same lock type as the original.

 Const adStateClosed			= 0  '// The object is closed
 Const adStateOpen				= 1  '// The object is open
 Const adStateConnecting		= 2  '// The object is connecting
 Const adStateExecuting			= 4  '// The object is executing a command
 Const adStateFetching			= 8  '// The rows of the object are being retrieved

 '*********************************************************


 'Checks for existing vbs scripts that are running and terminates them, avoids locking up ports
 Set wmi = GetObject("winmgmts:root\cimv2") 
 Set cProcesses = wmi.ExecQuery("select * from win32_process where Name like '%mshta.exe%'") 
 For Each oProcess in cProcesses
	oProcess.Terminate()
 Next

 ' Checks for any arguments being passed into the script
 If Not WScript.Arguments.Count = 0 Then
	sArg = ""
	For Each Arg In Wscript.Arguments
		  sArg = sArg & " " & """" & Arg & """"
	Next
 End If
 If InStr(sArg, "ME") <> 0 Then meMode = true										'Checks if the passed argument is for ME mode
 AccessResult = Load_Access			 												'Function to check for access connection and load info from database
 set windowBox = HTABox("white", 180, 600, 300, 0) :  with windowBox	 	 		'Calls function to create ie window and sets as the active object
	checkAccess																		'Updates buttons on the form if the connection is good
	do until closeWindow = true														'Run loop until conditions are met
		do until .done.value = "cancel" or .done.value = "access" or _
				 .done.value = "done" or .done.value = "allOps"						'Conditions to stop the wait loop
			wsh.sleep 50															'Pause for 50 ms
			On Error Resume Next													'Ignores error if window is closed manually
			If .done.value = true Then												'Checks for a non-existent case
				ServerClose()														'If the window is closed, runs the close script
			End If
			On Error GoTo 0															'Resets the error prompting
		Loop
		If .done.value = "cancel" then												'If the x button is clicked
			closeWindow = true	 													'Variable to end loop	
		ElseIf .done.value = "done" then											'If the button is clicked
			LoadOPScript .OPText.value, .argText.value								'Function to load the new script
			closeWindow = true														'Variable to end loop	
		ElseIf .done.value = "access" then											'If the database button is clicked
			.done.value = false														'Resets the variable to return to the loop when finished
			windowBox.accessText.innerText = "Retrying connection."					'Changes AX button text
			windowBox.accessButton.style.backgroundcolor = "orange"					'Changes the AX button color
			AccessResult = Load_Access												'Function to check for SQL connection
			checkAccess																'Updates buttons on the form if the connection is good
		ElseIf .done.value = "allOps" Then											'If the all operations button is pressed
			objShell.Run sOPsCmd													'Runs the command string for the all ops script
			ServerClose()															'If the window is closed, runs the close script
		End If 
	loop
	.close																			'Once the loop is complete closes the window
 end with
 ServerClose()																		'Function to close open connections and return settings back to original	
 Wscript.Quit																		'Ends the script

Function HTABox(sBgColor, h, w, l, t)  												'Function to create the ie window
	Dim IE, nRnd, sCmd																'Initializes variables
	randomize : nRnd = Int(1000000 * rnd)  											'Creates a random number to ID the ie window
	
	sCmd = "mshta.exe ""javascript:{new " _ 
		& "ActiveXObject(""InternetExplorer.Application"")" _ 
		& ".PutProperty('" & nRnd & "',window);" _ 
		& "window.moveTo(" & l & ", " & t & ");    " _
		& "window.resizeTo(" & w & "," & h & ")}""" 								'Parameters used to define the initial window
	with objShell																	'Object to run commands
		.Run sCmd, 1, False  														'Creates the user ie window
		do until .AppActivate("javascript:{new ") : WSH.sleep 10 : loop 			'Loop to wait until the window has been loaded
	end with
	For Each IE In CreateObject("Shell.Application").windows 						'Loops through each ie window
		If IsObject(IE.GetProperty(nRnd)) Then 										'If the window is the correct one
			set HTABox = IE.GetProperty(nRnd) 										'Defines the window
			IE.Quit 																'Quits any ie instances open
			HTABox.document.write LoadHTML(sBgColor)								'Loads the HTML to the ie object
			HTABox.document.title = documentTitle									'Changes the window's title
			HTABox.resizeTo w, windowHeight											'Changes the height based on the number of scripts loaded in LoadHTML
			Exit Function 															'Exits the function
		End If 
	Next 
	MsgBox "HTA window not found." 													'Messages the user if the window is unable to be created (or closed to quickly)
	wsh.quit																		'Ends the script	
 End Function

Function checkAccess()																'Function to change the buttons based on the connection results
	If AccessResult = false Then													'If the connection was not successful
		windowBox.accessText.innerText = "Database not loaded"						'Updates the button text
		windowBox.accessButton.style.backgroundcolor = "red"						'Colors the button
	Else																			'If the connection was successful
		windowBox.accessText.innerText = "Database connection successful"			'Updates the button text
		windowBox.accessButton.style.backgroundcolor = "limegreen"					'Colors the button
		windowBox.accessButton.disabled = true										'Makes the button unable to be clicked upon
	End If
 End Function

Function GetNewConnection()															'Function to connect to the database
	Dim objCmd : Set objCmd = CreateObject("ADODB.Connection")						'Object to connect to the SQL source
	'Connection string to the SQL database
	Dim sConnection : sConnection = "Data Source=" & dataSource & ";Initial Catalog=" & initialCatalog & ";Integrated Security=SSPI;"
	Dim sProvider : sProvider = "SQLOLEDB.1;"										'Connection type
	
	objCmd.ConnectionString	= sConnection											'Contains the information used to establish a connection to a data store.
	objCmd.Provider = sProvider														'Indicates the name of the provider used by the connection.
	objCmd.CursorLocation = adOpenStatic											'Sets or returns a value determining who provides cursor functionality.
	If debugMode = False Then On Error Resume Next									'Bypasses any error messages that occur during connection
    objCmd.Open																		'Opens connection to the SQL database
	On Error GoTo 0  																'Resets the error reporting
	If objCmd.State = adStateOpen Then    											'Checks if the connection is open
        Set GetNewConnection = objCmd    											'Returns the object to the function
	Else																			'If the connection is not open
        Set GetNewConnection = Nothing												'Returns false to the function
    End If  
 End Function 

Function Load_Access()																'Function to check for database connection and load script information
	Dim objCmd : set objCmd = GetNewConnection										'Creates the connection object to the database
	Dim adminCnt: adminCnt = 0
	Dim allCnt : allCnt = 0
	
	Dim sqlString, rsScript, rsCount, OPName, OPLoc, ArgVal, OPstring, OPID			'Initial variables
	Dim machineString, rsMachine, machineID, cnt, Admin, ScanArray, Scan			'Initial variables
	
	On Error GoTo 0																	'Resets the error warnings
	If objCmd is Nothing Then Load_Access = false : Exit Function					'If the database connection fails returns false to function
	sqlString = "Select COUNT(*) From [00_Script] WHERE [Admin] is null;"			'SQL string to count the number of operations scripts in the database
	set rsCount = objCmd.Execute(sqlString)											'Sends the SQL query to the database
	If rsCount(0).value <> 0 Then													'If scripts are found in the table
		Redim OPNameArray(rsCount(0).value - 1)										'Changes dimension on the operation name array
		Redim OPIDArray(rsCount(0).value - 1)										'Changes dimension on the operation id array
	End If
	Set rsCount = Nothing															'Erases the sql query
	
	sqlString = "Select COUNT(*) From [00_Script] WHERE [Admin] = 1;"				'SQL string to count the number of admin scripts
	set rsCount = objCmd.Execute(sqlString)											'Sends the SQL query to the database
	If rsCount(0).value <> 0 Then													'If scripts are found in the table
		Redim ADNameArray(rsCount(0).value - 1)										'Changes dimension on the admin name array
		Redim ADIDArray(rsCount(0).value - 1)										'Changes dimension on the admin id array
	End If
	Set rsCount = Nothing															'Erases the sql query
	
	'SQL string to get all of the scripts from the database
	sqlString = "Select [OPID], [OPName], [ScriptLocation], [ScriptArg], [ScanPrefix], [Admin], [OPDescription] From [00_Script];"
	Set rsScript = objCmd.Execute(sqlString)										'Sends the SQL query to the database
	DO WHILE NOT rsScript.EOF														'Loops through all of the resulting rows
		OPID = rsScript.Fields(0)													'Stores the operation ID
		OPName = rsScript.Fields(1)													'Stores the operation name
		OPLoc = rsScript.Fields(2)													'Stores the folder location
		ArgVal = rsScript.Fields(3)													'Stores the allowed scanners
		If IsNull(rsScript.Fields(4)) Then											'Checks if the scanner array string is empty
			ScanArray = Array()														'Creates a blank array for the scanner
		Else																		'If the scanner array is not empty
			ScanArray = Split(rsScript.Fields(4), ";")								'Splits the data into separate rows
		End If
		Admin = rsScript.Fields(5)													'Stores the admin script identifier
		If Admin = True Then														'If the script is an admin script
			ADNameArray(adminCnt) = rsScript.Fields(6)								'Stores the admin name in a new row
			ADIDArray(adminCnt) = rsScript.Fields(0)								'Stores the admin id in the new row
			adminCnt = adminCnt + 1													'Adjusts the admin counter
		Else																		'If the script is not an admin script
			OPNameArray(allCnt) = rsScript.Fields(6)								'Stores the script name
			OPIDArray(allCnt) = rsScript.Fields(0)									'Stores the operations ID
			allCnt = allCnt + 1
		End If
		LocDict.Add OPID, OPLoc														'Adds the information to the script array, used in the HTML creation 
		If ArgVal = false Then														'If there are no scanners supported in the script
			'Creates a button
			OPstring = "<button id='" & OPID & "' style='height: 30px; width: 150;' onclick='operationFunction(&#39;" & OPID & "&#39;)'>" & OPName & "&nbsp;</button>"
		Else																		'If there are a scanner option for the scripts
			'Creates a dropdown menu for selecting which scanner to connect to
			OPstring = "<select size='1' id='" & OPID & "' style='height: 30px; width: 150;' onChange='argumentFunction(&#39;" & OPID & "&#39;)'>" _
					 & "<option value='0'>" & OPName & "&nbsp;</option>" _
					 & "<option value='1'>Manual</option>"
			cnt = 2																	'Starts the dropdown row to 2. 1 is used for the default (manual) scanner
			For each Scan in ScanArray												'Cycles through each scanner option
				'Creates the SQL query string for the scanner type
				machineString = "Select [MachineName] From [00_Machine_IP] Where [MachineName] Like '" & Scan & "%' and [DeviceID] Like '%Scanner' and [Inactive] = 'false' ORDER BY [MachineName] ASC;"
				Set rsMachine = objCmd.Execute(machineString)						'Sends the SQL command to the database
				DO WHILE NOT rsMachine.EOF											'Loops through each resulting scanner
					OPstring = OPstring & "<option value='" & cnt & "'>" & rsMachine.Fields(0) & "</option>"	'Creates the dropdown row with the scanner ID
					cnt = cnt + 1													'Increments the counter
					rsMachine.MoveNext												'Moves to the next SQL result
				Loop
				Set rsMachine = Nothing												'Erases the SQL query result
			Next
			If meMode = true Then													'If the M.E. option is enabled
				'Creates the SQL query for the M.E. computers
				machineString = "Select [MachineName] From [00_Machine_IP] Where [MachineName] Like 'ME%' and [Inactive] = 'false' ORDER BY [MachineName] ASC;"
				Set rsMachine = objCmd.Execute(machineString)						'Sends the SQL command to the database
				DO WHILE NOT rsMachine.EOF											'Loops through each resulting scanner
					OPstring = OPstring & "<option value='" & cnt & "'>" & rsMachine.Fields(0) & "</option>"	'Creates the dropdown row with the scanner ID
					cnt = cnt + 1													'Increments the counter
					rsMachine.MoveNext												'Moves to the next SQL result
				Loop
				Set rsMachine = Nothing												'Erases the SQL query result
			End If
			OPstring = OPstring & "</select></div>"									'HTML items to close the dropdown menu
		End If
		ArgDict.Add OPID, OPstring													'Adds the script information to the array
		rsScript.MoveNext															'Moves to the next script SQL result
	Loop
	
	Set rsScript = Nothing															'Erases the script query
	objCmd.Close																	'Closes the database connection
	Set objCmd = Nothing															'Erases the database connection
	Load_Access = true																'Sends true to the function
 End Function

Function LoadOPScript(OPNum, ScriptArg)												'Function to load the scripts
	Dim ScriptLocation, sCmd														'Initializes variables
	
	'ShowWindow() Commands:
	Const SW_HIDE = 0
	Const SW_SHOWNORMAL = 1
	Const SW_NORMAL = 1
	Const SW_SHOWMINIMIZED = 2
	Const SW_SHOWMAXIMIZED = 3
	Const SW_MAXIMIZE = 3
	Const SW_SHOWNOACTIVATE = 4
	Const SW_SHOW = 5
	Const SW_MINIMIZE = 6
	Const SW_SHOWMINNOACTIVE = 7
	Const SW_SHOWNA = 8
	Const SW_RESTORE = 9
	Const SW_SHOWDEFAULT = 10
	Const SW_FORCEMINIMIZE = 11
	Const SW_MAX = 11

	ScriptLocation = LocDict(OPNum)																							'Retrieves the script location
	If UCase(Right(ScriptLocation, 3)) = "HTA" Then																			'Checks for HTA scripts
		CreateObject("Shell.Application").ShellExecute ScriptLocation, , , "open", SW_SHOWNORMAL							'Runs command to open HTA script
	ElseIf UCase(Right(ScriptLocation, 3)) = "PS1" Then																		'For PowerShell scripts
		Dim scriptFolder : scriptFolder = Left(ScriptLocation, InStrRev(ScriptLocation,"\"))								'Variable of the folder location for the script
		Dim scriptName : scriptName = Right(ScriptLocation, Len(ScriptLocation) - InStrRev(ScriptLocation,"\"))				'Variable of the script file name
		sCmd = """" &  "explorer.exe" & """  /e, """ & scriptFolder & """"													'Creates the command for opening windows explorer
		msgbox "Open file: '" & scriptName & "' by right clicking and selecting 'Run with PowerShell'"						'Sends a message to the user with steps to be taken
		objShell.Run sCmd																									'Opens up the windows explorer
	Else																													'All vbs scripts
		sCmd = """" &  oProcEnv("windir") & "\SysWOW64\" & ScriptHost & """" & " """ & ScriptLocation & """" & ScriptArg	'Creates the command for running the script
		objShell.Run sCmd																									'Sends the command to run the new script
	End If
    ServerClose()																											'Function to close script
 End Function

Sub ServerClose()																	'// EXIT SCRIPT
	If debugMode = False Then On Error Resume Next							 		'Used for debugging process
	WScript.Sleep 1000  													 		'// REQUIRED OR ERRORS
	windowBox.close																	'Closes the ie window
	On Error GoTo 0																	'Resets the error capturing
    Wscript.Quit																	'Ends the script
 End Sub

Function LoadHTML(sBgColor)															'Function to create all of the JS and HTML code for the window
	Dim opTop : opTop = 75
	Dim a, adTop
	
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
		& ".opButton {" _
			& "background-color: blue;" _
			& "height: 30px;" _
			& "width: 30px;" _
			& "font-weight: bold;" _
			& "font: 20px;" _
			& "color: white;" _
			& "}" _
		& ".closeButton {" _
			& "background-color: red;" _
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
		& "function operationFunction(OpNum) {" _
			& "document.getElementById('done').value = 'done';" _
			& "document.getElementById('OPText').value = OpNum;" _
		& "}" _
		& "function argumentFunction(OpNum) {" _
			& "document.getElementById('done').value = 'done';" _
			& "document.getElementById('OPText').value = OpNum;" _
			& "document.getElementById('argText').value = ' ' + document.getElementById(OpNum).options[document.getElementById(OpNum).selectedIndex].innerHTML;" _
		& "}" _
		& "</script></head>"

	'Body Start String							
	LoadHTML = LoadHTML & "<body scroll=no unselectable='on' class='unselectable'>"	
		
	'Access Connect String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 25px; height: 30px; width: 30px; text-align: left;'>" _
		& "<button class=HTAButton id=accessButton style='height: 30px; width: 30px; text-align: center;background-color:orange;' disabled onclick='done.value=""access""'>&nbsp;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 25px; left: 60px; height: 30px; width: 480px; text-align: left;' id='accessText'>Waiting for database connection&nbsp;</div>"
				
	'Operations String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & opTop + 0 & "px; left: 25px;height: 30px; width: 400px; text-align: left'>Operations</div>" _
	
	For a = 0 to UBound(OPNameArray)
		LoadHTML = LoadHTML _	
			& "<div unselectable='on' class='unselectable' style='top: " & a * 45 + opTop + 45 & "px; left: 200px;height: 30px; width: 350px;'>" & OPNameArray(a) & "</div>" _
			& "<div unselectable='on' class='unselectable' style='top: " & a * 45 + opTop + 48 & "px; left: 25px;height: 30px; width: 150px;'>" _
				& ArgDict(OPIDArray(a)) & "</div>" 
	Next
	
	
	 adTop = opTop + a * 45 + 70
	
	'ADMIN String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: " & adTop + 0 & "px; left: 25px;height: 30px; width: 350px;' id=adminText>Administration</div>"
			
	For a = 0 to UBound(ADNameArray)
		LoadHTML = LoadHTML _	
			& "<div unselectable='on' class='unselectable' style='top: " & a * 45 + adTop + 45 & "px; left: 200px;height: 30px; width: 350px;'>" & ADNameArray(a) & "</div>" _
			& "<div unselectable='on' class='unselectable' style='top: " & a * 45 + adTop + 48 & "px; left: 25px;height: 30px; width: 150px;'>" & ArgDict(ADIDArray(a)) & "</div>" 
	Next
	windowHeight = adTop + a * 45 + 90
			
	'All Op String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 505px;height: 30px; width: 30px;'><button class='opButton' style='height: 30px; width: 30px;' onclick='done.value=""allOps""'>&#10010;</button></div>"
		
	'Close Box String
	LoadHTML = LoadHTML _	
		& "<div unselectable='on' class='unselectable' style='top: 5px; left: 545px;height: 30px; width: 30px;'><button class='closeButton' style='height: 30px; width: 30px;' onclick='done.value=""cancel""'>&#10006;</button></div>" _
		& "<div unselectable='on' class='unselectable' style='top: 0px; left: 700px;'><input type=hidden id=done 			style='visibility:hidden;' value=false><center>&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 0px; left: 700px;'><input type=hidden id=argText 		style='visibility:hidden;' value=''><center>&nbsp;</div>" _
		& "<div unselectable='on' class='unselectable' style='top: 0px; left: 700px;'><input type=hidden id=OPText 			style='visibility:hidden;' value=false><center>&nbsp;</div>"
		
	'End Body String
	LoadHTML = LoadHTML _
		& "</body>"
		
		
 End Function
 