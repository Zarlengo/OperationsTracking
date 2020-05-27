Option Explicit
' HTTP strings to Login, search, load ECO approvals, and load ceviation details
Dim testMode
Dim loginURL, ApprovalURLPre, ApprovalURLSuf, SearchUrlPre, SearchUrlSuf, ECODetailURL
Dim loginENOVIA, oConn, StartTime, ObjIE, FirstECO
Dim Data_Source, Initial_Catalog, Database_Name, SQL_Table

    testMode = False
    UpdateECO False
    WScript.Quit

Sub UpdateECO(ByVal VarIn)
Dim RunDuration, StartTime, intTimeout, strMessage, strTitle, WshShell, intResult
    InitialVariables False
    StartTime = Now
    RunUpdate False
    If testMode = True Then
            RunDuration = Now - StartTime                                                                                                   ' Calculates macro duration
            strMessage = "Completed in: " & Right("00" & hour(RunDuration), 2) & ":" & Right("00" & minute(RunDuration), 2) & ":" & Right("00" & second(RunDuration), 2)
            intTimeout = 5      'Number of seconds to wait
            strTitle = "Update complete"
            Set WshShell = CreateObject("WScript.Shell")
            intResult = WshShell.Popup(strMessage, intTimeout, strTitle)
    End If
End Sub

Sub InitialVariables(ByVal VarIn)
    loginURL = "http://plm.flowcorp.com:8080/enovia/emxLogin.jsp"
    ApprovalURLPre = "http://plm.flowcorp.com:8080/enovia/common/emxTable.jsp?parentOID="
    ApprovalURLSuf = "&table=AEFLifecycleApprovalsSummary&program=emxLifecycle%3AgetAllTaskSignaturesOnObject&pagination=0&objectId="
    SearchUrlPre = "http://plm.flowcorp.com:8080/enovia/common/emxTable.jsp?txtWhere=revision%253D%253Dlast%2B%2B%2526%2526%2528Name%2B%257E%257E%2B%2522"
    SearchUrlSuf = "*%2522%2529%2B%2526%2526%2528Current%2B%253D%253D%2Bconst%2522Release%2522%2529%2B%257C%257C%2528Current%2B%253D%253D%2Bconst%2522Review%2522%2529%2B%257C%257C%2528Current%2B%253D%253D%2Bconst%2522Design Work%2522%2529&txtSearch=&txtFormat=*&ckChangeQueryLimit=&queryLimit=1000&pagination=0&selType=ECO&table=ENCGeneralSearchResult&program=emxPart%3AgetPartSearchResult&vaultAwarenessString=true&sortColumnName=Name&sortDirection=ascending"
    ECODetailURL = "http://plm.flowcorp.com:8080/enovia/common/emxForm.jsp?form=type_Flo_Deviation&objectId="

    Data_Source = "AXPRODSQL02\AXPRODSQL02" & Chr(59) ' Name of server
    Initial_Catalog = "WJH_DB_Warehouse" & Chr(59) ' Name of database
	Database_Name = "ECO.ECOApprovers" ' Name of database
    SQL_Table = "ECO_Approvers"
End Sub

Sub RunUpdate(ByVal VarIn)
    If ENOVIAStartUp(False) And SQLStartUp(False) Then                                                                                                          ' Runs function to load all initial variable
        LoadSystemTable (True)
        OpenECOData False                                                                                                             ' Function to load ie window with ECO search
        LoadSystemTable False
        ieQuit False                                                                                                                  ' Function to close the open ie window
    End If
End Sub

Function ENOVIAStartUp(ByVal VarIn)
' Initializes routine variables
Dim ENOVIACheck, ENOVIATable, subTable, elem
Dim login_name, login_password, TableString

    loginENOVIA = False                                                                                                             ' Variable to store if logged into other account
    ENOVIACheck = False                                                                                                             ' Sets ENOVIA check to false
    Set ObjIE = GetObject("new:{D5E8041D-920F-45e9-B8FB-B1DEB82C6E5E}")  ' this creates a medium IL (PM off) tab by default
        With ObjIE                                                                                                                      ' Sets ie object
        If testMode = True Then .Visible = True Else .Visible = False
        .navigate loginURL                                                                                                          ' Opens ENOVIA url
                If WaitURLLoad(False) Then Exit Function
        ieLogOut (False)
        Do While ENOVIACheck = False                                                                                                ' Do while login screen is present
            If WaitURLLoad(False) Then Exit Function
            ENOVIACheck = True                                                                                                      ' Sets ENOVIA check variable for loop
            Set ENOVIATable = .document.getElementsByTagName("Table")                                                                       ' Searches for HTML Table objects
            For Each subTable In ENOVIATable                                                                                                ' Cycles through each object
                TableString = subTable.innerText                                                                                    ' Stores text for the object
                If InStr(1, UCase(TableString), "USERNAME") > 0 Then                                                                ' If string username is found
                    loginENOVIA = True                                                                                              ' Change variable to show logged into other account in ENOVIA
                    ENOVIACheck = False                                                                                             ' Changes the ENOVIA check variable
                    login_name = "wkg_cleanroom"                                                                                    ' Login username for ENOVIA
                    login_password = "clean1"                                                                                       ' Login password for ENOVIA
                    For Each elem In subTable.document.getElementsByTagName("input")                                                ' Cycles through each input object on the page
                        If elem.Name = "login_name" Then                                                                            ' Checks if the input is the username field
                            elem.Value = login_name                                                                                 ' Puts the user name into the username field
                        ElseIf elem.Name = "login_password" Then                                                                    ' Checks if the input is the login field
                            elem.Value = login_password                                                                             ' Puts the password into the login field
                        ElseIf elem.Name = "enter" Then                                                                             ' Checks if the input is the enter button
                            elem.Click                                                                                              ' Clicks the button
                        End If
                    Next
                End If
            Next
            If WaitURLLoad(False) Then Exit Function
        Loop
    End With
ENOVIAStartUp = True
End Function

Function WaitURLLoad(ByVal VarIn)
' Initializes routine variables
Dim childHWND, hWND, loadStart

    loadStart = Now()                                                                                                               ' Stores the current time
    With ObjIE                                                                                                                      ' Sets ie object
       Do While .Busy Or .readyState <> 4                                                                                          ' Pause counter to wait for page to load
'            DoEvents                                                                                                                ' Pauses macro to run pending windows events
            If Now() >= loadStart + #12:01:00 AM# Then                                                                              ' Checks if the loading is taking too long
                ieQuit (False)                                                                                                      ' Function to close the open ie window
                WaitURLLoad = True
                Exit Function                                                                                                                ' Ends the macro
            End If
        Loop
    End With
End Function

Sub LoadSystemTable(ByVal VarIn)
Dim eleFound, ifra_Frame, Cnt, n, tagElement, TableName, tableElement, ifra_Inner
Dim HTMLString, tableSearchString, tagSearch, tagSearchString

    If VarIn = True Then
        TableName = "System Table..."
        tableSearchString = "id=ENCGeneralSearchResult"
    Else
        TableName = "ECO_Notes..."
        tableSearchString = "id=ECO_Notes~ENCGeneralSearchResult"
    End If
   
    LoadWebpage (SearchUrlPre & 1 & SearchUrlSuf)                                                     ' Function to load ENOVIA webpage
   
    HTMLString = ObjIE.document.getElementById("divPageBody").FirstChild.contentWindow.document.getElementsByTagName("Body")(0).innerHTML     ' Gets html string for the whole enovia table
    If InStr(1, HTMLString, tableSearchString) Then Exit Sub
    tagSearch = Array("td", "span")
    tagSearchString = Array("title=View", TableName)
    Cnt = 0
    For n = 0 To UBound(tagSearch)
        eleFound = False
        Do
            For Each tagElement In ObjIE.document.getElementsByTagName(tagSearch(n))
                If InStr(1, tagElement.outerHTML, tagSearchString(n)) <> 0 Then eleFound = True: Exit For
            Next
            If eleFound = True And n = 0 Then
                tagElement.Click
            ElseIf eleFound = True And n <> 0 Then
                tagElement.FireEvent ("onmousedown")
                WScript.Sleep 1000
                If WaitURLLoad(False) Then Exit Sub
            Else
                WScript.Sleep 1000
            End If
            If Cnt > 10 Then Exit Sub Else Cnt = Cnt + 1
        Loop While eleFound = False
    Next
End Sub

Function LoadWebpage(ByVal VarIn)                                                                                  ' Function to load ENOVIA webpage
' Initializes routine variables
Dim hWND, childHWND, loadStart

    loadStart = Now()                                                                                                               ' Stores the current time
    LoadWebpage = False                                                                                                             ' Sets the initial value for the function                                                                                                             '
    With ObjIE                                                                                                                      ' Sets ie object
        .navigate VarIn                                                                                                             ' Navigates to the new webpage
        If WaitURLLoad(False) Then Exit Function
    End With
End Function

 
Function SQLStartUp(ByVal VarIn)
' Initializes routine variables
Dim strSQL, rs

	Set rs = CreateObject("ADODB.Recordset")
    FirstECO = "D0000000"
    SQLOpen False
    strSQL = "SELECT ECO_Number FROM " & Database_Name & Chr(59)
    rs.Open strSQL, oConn, 3, 1
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst 'Unnecessary in this case, but still a good habit
        Do Until rs.EOF = True
            If rs("ECO_Number").Value < FirstECO Then FirstECO = rs("ECO_Number").Value
            'Move to the next record. Don't ever forget to do this.
            rs.MoveNext
        Loop
    Else
        FirstECO = "C4050000"
    End If
    
    strSQL = "DELETE FROM " & Database_Name & Chr(59)
    oConn.Execute strSQL
    SQLClose False
    SQLStartUp = True
End Function

Function SQLOpen(ByVal VarIn)
    'connect to MySQL server using Connector/ODBC
    Set oConn = CreateObject("ADODB.Connection")
        oConn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;" _
        & "Data Source=" & Data_Source _
        & "Initial Catalog=" & Initial_Catalog & Chr(59)
    oConn.Open
End Function

Sub SQLClose(ByVal VarIn)
    oConn.Close
    Set oConn = Nothing
End Sub

Sub OpenECOData(ByVal VarIn)                                                                                                   ' Function to load ie window with ECO search
' Initializes routine variables
Dim ECOString, ECONum, MaxECO
    ECONum = CInt(Mid(FirstECO, 2, 3))
    MaxECO = 410                                                                               ' Trims the top ECO number
    Do While ECONum < MaxECO                                                                                                   ' While the current ECO number is below the maximum
        ECOString = "C" & ECONum & "*"                                                                                          ' Creates the ENOVIA search string
        ECONum = ECONum + 1                                                                                                     ' Increments the ECO counter
        LoadWebpage (SearchUrlPre & ECOString & SearchUrlSuf)                                                                   ' Function to load ENOVIA webpage
        If LoadECOData(False) Then Exit Do                                                                                                      ' Function to load the ECO data from the webpage
    Loop
End Sub

Function LoadECOData(ByVal VarIn)                                                                                                   ' Function to load status update data from CSV
' Initializes routine variables
Dim ECOIDs
Dim HTMLString
Dim ECOArray, HeaderArray
Dim n, ECO_Name_Loc, j
    
    Set ECOIDs = CreateObject("Scripting.Dictionary")
    ECOArray = Array()
    HeaderArray = Array("ECOID")
        HTMLString = ObjIE.document.getElementById("divPageBody").FirstChild.contentWindow.document.getElementById("ENCGeneralSearchResult").innerHTML & "<TABLE>ENDOFSTRING"     ' Gets html string for the whole enovia table
    Do While InStr(1, HTMLString, "</TH>") <> 0
        ReDim Preserve HeaderArray(UBound(HeaderArray) + 1)
        HeaderArray(UBound(HeaderArray)) = TrimString(Left(HTMLString, InStr(1, HTMLString, "</TH>") - 1))                                                                 ' Stores the ECO ID value
        If HeaderArray(UBound(HeaderArray)) = "Name" Then ECO_Name_Loc = UBound(HeaderArray)
        HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TH>") - 4)                                ' Trims the string to the ECO number
    Loop
    ECOIDs.Add "Header", HeaderArray
    If TrimString(Left(HTMLString, InStr(1, HTMLString, "</TD>") - 1)) = "No Objects Found" Then
        LoadECOData = True
    Else
        Do While InStr(1, HTMLString, "rmbID=") <> 0
            ReDim ECOArray(UBound(HeaderArray))
            HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "rmbID=") - 6)                                     ' Trim the string to the first ECO ID
            ECOArray(0) = Left(HTMLString, InStr(1, HTMLString, Chr(34)) - 1)                                                                 ' Stores the ECO ID value
            For n = 1 To UBound(ECOArray)
                HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TD>") - 4)                                         ' ECO Name: Trims the string to the next column
                Do While Left(HTMLString, 1) = Chr(60) Or Left(HTMLString, 1) = Chr(59) & Chr(32) Or Left(HTMLString, 1) = Chr(59) & Chr(32)
                    HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, Chr(62)))
                Loop
                ECOArray(n) = TrimString(Left(HTMLString, InStr(1, HTMLString, "</TD>") - 1))
            Next
            HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TR>") - 4)                                       ' Trims the string for the next ECO
            If Left(ECOArray(ECO_Name_Loc), 6) <> "C40500" Then
                ECOIDs.Add ECOArray(ECO_Name_Loc), ECOArray
            End If
        Loop
        SQLAdd ECOIDs
    End If
End Function

Function TrimString(ByVal VarIn)                                                                                                            ' Function to trim approval string to just the text
    If InStr(1, VarIn, "</A>") <> 0 Then VarIn = Left(VarIn, InStr(1, VarIn, "</A>") - 1)                                           ' Checks if there is a </a> code at the end of the string
    If InStr(1, VarIn, "</B>") <> 0 Then VarIn = Left(VarIn, InStr(1, VarIn, "</B>") - 1)                                           ' Checks if there is a </a> code at the end of the string
    Do While InStr(1, VarIn, ">") <> 0                                                                                              ' Checks for preceeding HTML code blocks
        VarIn = Right(VarIn, Len(VarIn) - InStr(1, VarIn, ">"))                                                                     ' Removes preceeding HTML code blocks
    Loop
    If InStr(1, VarIn, "&nbsp;") <> 0 Then VarIn = Left(VarIn, InStr(1, VarIn, "&nbsp;") - 1)                                       ' Checks for HTML code for space at the end of the string
    TrimString = Trim(VarIn)                                                                                                              ' Sets output string for the function
End Function

Sub SQLAdd(ByVal ECOIDs)
Dim ECOIDCol, ECONameCol, i
Dim strSQL, SQL_Table, ECOID, ECOName, ECOStatus, ECOAging, CurApp, AppAging, DueDate, RemApp, ECOStatusCol
Dim ECO_Data, ApproverArray
    
    strSQL = "INSERT INTO " & Database_Name & " (ECO_Number, ECO_Status, ECO_Aging, Current_Approver, Approver_Aging, Due_Date, Next_Approvers) VALUES "
	ECOIDCol = 0
    For Each ECO_Data In ECOIDs.Items
        If ECO_Data(ECOIDCol) = "ECOID" Then
            For i = LBound(ECO_Data) To UBound(ECO_Data)
                If ECO_Data(i) = "Name" Then ECONameCol = i
                If ECO_Data(i) = "State" Then ECOStatusCol = i
            Next
        Else
            ECOID = ECO_Data(ECOIDCol)
            ECOName = Chr(39) & ECO_Data(ECONameCol) & Chr(39) & Chr(44)
            ECOStatus = Chr(39) & ECO_Data(ECOStatusCol) & Chr(39) & Chr(44)
            LoadWebpage (ApprovalURLPre & ECOID & ApprovalURLSuf & ECOID)                                                    ' Function to load ENOVIA webpage
            ApproverArray = LoadApprovers(ECO_Data(ECOStatusCol))
            ECOAging = Chr(39) & ApproverArray(0) & Chr(39) & Chr(44)
            CurApp = Chr(39) & ApproverArray(1) & Chr(39) & Chr(44)
            AppAging = Chr(39) & ApproverArray(2) & Chr(39) & Chr(44)
            DueDate = Chr(39) & ApproverArray(3) & Chr(39) & Chr(44)
            RemApp = Chr(39) & ApproverArray(4) & Chr(39)
            strSQL = strSQL & "(" & ECOName & ECOStatus & ECOAging & CurApp & AppAging & DueDate & RemApp & "), "
        End If
    Next
    strSQL = Left(strSQL, Len(strSQL) - 2) & Chr(59)
    If ECOIDs.Count > 1 Then
        SQLOpen (False)
        oConn.Execute strSQL
        SQLClose (False)
    End If
End Sub

Function LoadApprovers(ByVal ECOStatus)                                                                                                 ' Function to load approvers from CSV
' Initializes routine variables
Dim AppCollection
Dim LastApproved
Dim PendLevel, State_Loc, Assignee_Loc, Title_Loc, Action_Loc, Due_Loc, Complete_Loc, ECOAging
Dim NextApp, State, Title, RejectString, Assignee, FinalApp, HTMLString, NextString, Action, CompleteDate, DueDate, Cnt, i, n, Wait_Due, DesignComplete, ReviewComplete, DefineComplete
Dim HeaderArray, AppArray, Table_Line
    
    Set AppCollection = CreateObject("Scripting.Dictionary")
    AppArray = Array()
    HTMLString = ObjIE.document.getElementById("divPageBody").FirstChild.contentWindow.document.getElementById("AEFLifecycleApprovalsSummary").innerHTML & "<TABLE>ENDOFSTRING"     ' Gets html string for the whole enovia table
    PendLevel = 99                                                                                                                  ' Sets the pending approval max value
    HeaderArray = Array("CheckBox")
    HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TH>") - 9)                               ' Trims the string to the ECO number
    Do While InStr(1, HTMLString, "</TH>") <> 0
        ReDim Preserve HeaderArray(UBound(HeaderArray) + 1)
        HeaderArray(UBound(HeaderArray)) = TrimString(Left(HTMLString, InStr(1, HTMLString, "</TH>") - 1))                                                                 ' Stores the ECO ID value
        HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TH>") - 4)                                ' Trims the string to the ECO number
    Loop
    HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TR>") - 4)                                       ' Trims the string for the next ECO
    For i = LBound(HeaderArray) To UBound(HeaderArray)
        If HeaderArray(i) = "State" Then State_Loc = i
        If HeaderArray(i) = "Assignee" Then Assignee_Loc = i
        If HeaderArray(i) = "Task Title" Then Title_Loc = i
        If HeaderArray(i) = "Action" Then Action_Loc = i
        If HeaderArray(i) = "Due Date" Then Due_Loc = i
        If HeaderArray(i) = "Completed Date" Then Complete_Loc = i
    Next
    Cnt = 0
    Do While InStr(1, HTMLString, "</TR>") <> 0
        ReDim AppArray(UBound(HeaderArray))
        HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TD>") - 4)
        For n = 1 To UBound(AppArray)
            Do While Left(HTMLString, 1) = Chr(60) Or Left(HTMLString, 1) = Chr(59) & Chr(32) Or Left(HTMLString, 1) = Chr(59) & Chr(32)
                HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, Chr(62)))
            Loop
            AppArray(n) = TrimString(Left(HTMLString, InStr(1, HTMLString, "</TD>") - 1))
            HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TD>") - 4)                                         ' ECO Name: Trims the string to the next column
        Next
        HTMLString = Right(HTMLString, Len(HTMLString) - InStr(1, HTMLString, "</TR>") - 4)                                       ' Trims the string for the next ECO
        AppCollection.Add Cnt, AppArray
        Cnt = Cnt + 1
    Loop
    For Each Table_Line In AppCollection.Items
        State = Table_Line(State_Loc)
        Assignee = Table_Line(Assignee_Loc)
        Title = Table_Line(Title_Loc)
        Action = Table_Line(Action_Loc)
        DueDate = Table_Line(Due_Loc)
        CompleteDate = Table_Line(Complete_Loc)
        If State <> "Notify Only" Then                                                                                             ' If the state is not notify only
            If CompleteDate <> "" And CompleteDate <> HeaderArray(Complete_Loc) Then If CDate(CompleteDate) > LastApproved Then LastApproved = CDate(CompleteDate)
            If Action = "Awaiting Approval" Then                                                                                    ' If waiting on assignee
                If (Assignee = "Durand, Mike" Or Assignee = "Hu, David") And Title = "ECO Implement Tasks" Then                     ' Checks for change analysist final approval
                    FinalApp = FinalApp & Chr(59) & Chr(32) & "Final: " & Assignee                                                             ' Stores final task
                Else                                                                                                                ' All other approvers
                    NextString = NextString & Chr(59) & Chr(32) & Assignee                                                    ' Stores waiting on string
                    If DueDate <> "" Then Wait_Due = CDate(DueDate)
                End If
            ElseIf Left(Action, 7) = "Pending" Then                                                                                 ' If status is pending
                If IsNumeric(Right(Action, 1)) Then                                                                                 ' Checks if there are steps to approvals
                    If PendLevel > CInt(Right(Action, Len(Action) - InStrRev(Action, Chr(32)))) Then                                                                    ' Checks if this is the next approval step
                        PendLevel = CInt(Right(Action, Len(Action) - InStrRev(Action, Chr(32))))                                                                                ' Stores the pending level
                        NextApp = Chr(59) & Chr(32) & Assignee                                                                      ' Adds assignee to the next string
                    ElseIf PendLevel = CInt(Right(Action, Len(Action) - InStrRev(Action, Chr(32)))) Then                                                                  ' Checks if there are same tier approvers
                        NextApp = NextApp & Chr(59) & Chr(32) & Assignee                                                            ' Adds assignee to the next string
                    End If
                Else                                                                                                                ' If there is only 1 pending level
                    NextApp = NextApp & Chr(59) & Chr(32) & Assignee                                                                ' Stores next approver string
                End If
            ElseIf Action = "Approved" Then
                If State = "Design Work" And CompleteDate <> "" And (ECOStatus = "Review" Or ECOStatus = "Release") Then
                    If CDate(CompleteDate) > DesignComplete Then DesignComplete = CDate(CompleteDate)
                ElseIf State = "Review" And CompleteDate <> "" And ECOStatus = "Release" Then
                    If CDate(CompleteDate) > ReviewComplete Then ReviewComplete = CDate(CompleteDate)
                End If
            ElseIf Action = "Rejected" Then RejectString = "ECO Rejected by: " & Assignee                                           ' Checks if ECO has been rejected
            ElseIf State = "Define Components" Then
                DefineComplete = CDate(CompleteDate)
            End If
        End If
    Next
    If RejectString <> "" Then NextString = RejectString              ' Replaces the next approver if the ECO is rejected
    If ECOStatus = "Design Work" Then
        ECOAging = Date - DefineComplete
    ElseIf ECOStatus = "Review" Then
        ECOAging = Date - DesignComplete
    Else
        ECOAging = Date - ReviewComplete
    End If
    
    If NextString <> "" Then
        NextString = Right(NextString, Len(NextString) - 1)
    ElseIf FinalApp <> "" Then
        NextString = Right(FinalApp, Len(FinalApp) - 1)
		FinalApp = ""
    End If
    If NextApp <> "" Then
		NextApp = Right(NextApp, Len(NextApp) - 1)
	ElseIf FinalApp <> "" Then
		NextApp = Right(FinalApp, Len(FinalApp) - 1)
	End If
	LoadApprovers = Array(ECOAging, NextString, Date - LastApproved, Wait_Due, NextApp)
End Function

Sub ieLogOut(ByVal VarIn)
' Initializes routine variables
Dim ENOVIAHeader, HeaderButtons

    If InStr(1, ObjIE.document.body.innerHTML, "pageHeadDiv") Then
        Set ENOVIAHeader = ObjIE.document.getElementById("pageHeadDiv").getElementsByTagName("td")                              ' Locates the header and stores as an object
        For Each HeaderButtons In ENOVIAHeader                                                                                  ' Cycles through each button on the header
            If InStr(1, HeaderButtons.outerHTML, "title=Logout") <> 0 Then                                                      ' If the table column is for the logout button
                HeaderButtons.Click                                                                                             ' Clicks the logout button
                WScript.Sleep 1000
                Exit For                                                                                                        ' Exits the loop
            End If
        Next
    End If
End Sub

Sub ieQuit(ByVal VarIn)
' Initializes routine variables
Dim loadStart, hWND, childHWND, ENOVIAHeader, HeaderButtons

    loadStart = Now()                                                                                                               ' Stores the current time
        ObjIE.navigate loginURL                                                                                                      ' Opens ENOVIA url
        If WaitURLLoad(False) = 0 Then
                ieLogOut (False)
        End If
    ObjIE.Quit                                                                                                                      ' Closes ie window
    Set ObjIE = Nothing                                                                                                             ' Clears memory
End Sub
