Private Sub Form_Load()
    ' Set up the Winsock control
    With sock
        .LocalPort = 0 ' Use any available local port
        .RemoteHost = "192.168.1.100" ' Replace with the IP address of your server
        .RemotePort = 1234 ' Replace with the port number of your server
    End With
    
    ' Attempt to connect to the remote server
    On Error GoTo connect_error
    sock.Connect ' Connect to the remote server
    On Error GoTo 0

    ' Start sending temperature values
    Dim temp As Double
    Do While True
        ' Replace with your code to get the temperature value
        temp = GetTankTemperature()
        
        ' Check if the socket is connected
        If sock.State = sckConnected Then
            ' Send the temperature value as a string message
            On Error Resume Next
            sock.SendData CStr(temp)
            On Error GoTo 0
        Else
            ' Attempt to reconnect to the remote server
            On Error GoTo connect_error
            sock.Connect
            On Error GoTo 0
            GoTo continue_loop
        End If
        
continue_loop:
        ' Wait for 1 second before sending the next message
        Sleep 1000
    Loop

connect_error:
    MsgBox "Error connecting to remote server: " & Err.Description, vbCritical
    Exit Sub
End Sub
