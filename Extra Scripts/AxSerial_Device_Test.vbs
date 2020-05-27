' ********************************************************************
' Sample: Query a device / send commands and receive responses
' (c) Copyright 1999-2007 by ActiveXperts Software
'   http://www.activexperts.com
' ********************************************************************

Option Explicit

Dim objComport, str

Set objComport       = CreateObject( "AxSerial.ComPort" )

' Clear (good practise)
objComport.Clear()

' A license key is required to unlock this component after the trial period has expired.
' Assign the LicenseKey property every time a new instance of this component is created (see code below). 
' Alternatively, the LicenseKey property can be set automatically. This requires the license key to be stored in the registry.
' For details, see manual, chapter "Product Activation".
objComport.LicenseKey = "FD2C1-DC93A-6BFBF"

' Component info
dim echoString : echoString = "ActiveXperts Serial Port Component " & objComport.Version & vbCrLf & vbCrLf
echoString =  echoString & "Build: " & objComport.Build & vbCrLf & "Module: " & objComport.Module & vbCrLf & vbCrLf
echoString =  echoString & "License Status: " & objComport.LicenseStatus & vbCrLf & "License Key: " & objComport.LicenseKey & vbCrLf

' Set Logfile
'Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
'objComport.LogFile = fso.GetSpecialFolder(2) & "\SerialPort.log"
'WScript.Echo "Log file: " & objComport.LogFile


objComport.Device = AskDevice( objComport )

' Optionally override defaults for direct COM ports
If( Left( objComport.Device, 3 ) = "COM" ) Then
  objComport.BaudRate  = 9600 'Ask( "Enter baud rate (no input means: default baud rate):", "9600", True )
' objComport.HardwareFlowControl  = True
' objComport.SoftwareFlowControl  = False
End If

' Set Logging - for troubleshooting purposes
' objComport.LogFile = "C:\SerialPort.log"

' Open the port
objComport.Open
Wscript.Echo "Open, result:" & objComport.LastError & " (" & objComport.GetErrorDescription( objComport.LastError ) & ")"

If( objComport.LastError <> 0 ) Then
  WScript.Quit
End If



' Set all Read functions (e.g. ReadString) to timeout after a specified number of millisconds
objComport.ComTimeout = 1000  ' Timeout after 1000msecs 

ReadResponse(objComport)
While (WriteCommand(objComport))
  ReadResponse(objComport)
WEnd

objComport.Close()
WScript.Echo "Close, result: " & objComport.LastError & " (" & objComport.GetErrorDescription(objComport.LastError) & ")"

'WScript.Echo "Ready."



' ********************************************************************
' Sub Routines
' ********************************************************************

Function AskDevice( ByVal objComport )
    Dim strInput, strTitle, strDevice
    Dim i, j

    strTitle = echoString & vbCrLf & vbCrLf & "Select a device" & vbCrLf

    For i = 0 To objComport.GetDeviceCount - 1
        strTitle = strTitle & "  " & i & ": " & objComport.GetDeviceName( i ) & vbCrLf
    Next

    For j = 0 To objComport.GetPortCount - 1
        strTitle = strTitle & "  " & ( i + j ) & ": " & objComport.GetPortName( j ) & vbCrLf
    Next

    While (strDevice = "")
        strInput = InputBox( strTitle, "Select device:", "0" )
        If( strInput = "" ) Then
            strDevice = ""
        ElseIf( CInt( strInput) < i) Then
            strDevice = objComport.GetDeviceName( CInt(strInput) )
        ElseIf ( CInt( strInput ) < i + j ) Then
             strDevice = objComport.GetPortName( CInt(strInput) - i)
        End If
    WEnd
    WScript.Echo "Selected device: " & strDevice & vbCrLf

    AskDevice = strDevice
End Function


' ********************************************************************

Function Ask( ByVal strTitle, ByVal strDefault, ByVal bAllowEmpty )

  Dim strInput, strReturn

  Do
     strInput = inputbox( strTitle, strTitle, strDefault )
     If ( strInput <> "" ) Then
          strReturn = strInput
     End If
  Loop until strReturn <> "" Or bAllowEmpty

  Ask = strReturn
End Function


' ********************************************************************

Sub ReadResponse(ByVal objComport)
  Dim str

  str = "notempty"
  objComport.Sleep(200)
  While (str <> "")
    str = objComport.ReadString()
    If (str <> "") Then
      wScript.Echo "  <- " & str
    End If

  WEnd
End Sub


' ********************************************************************

Function WriteCommand(ByVal objComport)

  Dim str

  str = InputBox( "Enter command (enter QUIT to stop the program):", "Enter value" ) 
  objComport.WriteString(str)
  If( objComport.LastError = 0 ) Then
    wScript.Echo "  -> " & str
  Else
    WScript.Echo "Write failed, result: " & objComport.LastError & " (" & objComport.GetErrorDescription(objComport.LastError) & ")"
  End If

  If (LCase( str ) = "quit") Then
    WriteCommand = False
  Else
    WriteCommand = True
  End If
End Function
