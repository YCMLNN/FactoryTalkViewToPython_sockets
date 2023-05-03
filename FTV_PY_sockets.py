import socket

HOST = '127.0.0.1'  # Replace with the IP address of your VBA application
PORT = 1234  # Replace with the port number used by your VBA application

# Create a socket object
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

# Connect to the VBA application's socket
sock.connect((HOST, PORT))

# Receive data from the VBA application's socket and write it to the console
while True:
    data = sock.recv(1024)
    if not data:
        break
    print(data.decode('utf-8'))

# Close the socket
sock.close()
