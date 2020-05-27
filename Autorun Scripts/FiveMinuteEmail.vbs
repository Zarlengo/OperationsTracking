
Set objFSO=CreateObject("Scripting.FileSystemObject")
 Const FolderLocations = "\\shapetechnologies.com\root\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\TXTFiles\"
 
		Set objFolder = objFSO.GetFolder(FolderLocations)
		Set colFiles = objFolder.Files
		For Each objFile in colFiles
			If Right(UCase(objFile.Name),4) = ".TXT" and Left(UCase(objFile.Name),12) = "OPERATION00_" Then
				ReadOP00(objFile)
			ElseIf Right(UCase(objFile.Name),4) = ".TXT" and Left(UCase(objFile.Name),11) = "REJECTIONS_" Then
				ReadReject(objFile)
			End if
		Next
		Set colFiles = Nothing
		Set objFolder = Nothing



Sub ReadOP00(fileObject)
	Set objFile = objFSO.OpenTextFile(fileObject)
	Do Until objFile.AtEndOfStream
		strLine = strLine & objFile.ReadLine
	Loop
	objFile.Close
	objFSO.DeleteFile(fileObject)
	
	Dim objShell
	Set objShell = Wscript.CreateObject("WScript.Shell")
	Const scriptLocation = "\\shapetechnologies.com\root\Flow\Operations\Seattle\Quality\Contract Cutting\Operation Documents\Scripts\Autorun Scripts\Invoice_Email_v1_1.vbs"
	
	objShell.Run Chr(34) & scriptLocation & Chr(34)

	' Using Set is mandatory
	Set objShell = Nothing
	
	' Const Subject = "Contract cutting material received"
	' Const EmailList = "koliver@flowcorp.com;czarlengo@flowcorp.com"
	' Dim messageBody : messageBody = "<body><p><span style='font-size:12pt; color:red'>This is an automatically generated email.</span></p><br>" _
		' & "<p><span>" & strLine & " slugs have completed initial inspection. Please create a work order for 060052-1 (quantity " & strLine & ") and receive into AX.</span></p>" _
		' & "<p><span>Also please create work orders for 060053-1 (quantity " & strLine & " in batches of 20) and 060053-2 (quantity " & strLine & " in batches of 20).</span></p><br>" _
		' & "<p><span>Thank you,</span></p>"
	
	
	'Call Send_Email(messageBody, subject, EmailList, "")
 End Sub
	
Sub ReadReject(fileObject)
	Set objFile = objFSO.OpenTextFile(fileObject)
	Do Until objFile.AtEndOfStream
		strLine = strLine & objFile.ReadLine
	Loop
	objFile.Close
	objFSO.DeleteFile(fileObject)
	
	
	Const Subject = "Contract cutting failed script"
	Const EmailList = "czarlengo@flowcorp.com"
	Dim messageBody : messageBody = "<body><p><span style='font-size:12pt; color:red'>This is an automatically generated email.</span></p><br>" _
		& "<p><span>There was a failure in the script running on WKEN439.</span></p><br>" _
		& "<p><span>" & strLine & "</span></p>"
	
	Call Send_Email(messageBody, subject, EmailList, "")
 End Sub
 
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