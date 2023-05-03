Private Sub Form_Load()
    ' Set up the Winsock control
    With sock
        .LocalPort = 0 ' Use any available local port
        .RemoteHost = "192.168.1.100" ' Replace with the IP address of your server
        .RemotePort = 1234 ' Replace with the port number of your server
        .Connect ' Connect to the remote server
    End With

    ' Start sending temperature values
    Dim temp As Double
    Do While True
        temp = GetTankTemperature() ' Replace with your code to get the temperature value
        If sock.State = sckConnected Then ' Check if the socket is connected
            sock.SendData CStr(temp) ' Send the temperature value as a string message
        End If
        Sleep 1000 ' Wait for 1 second before sending the next message
    Loop
End Sub
