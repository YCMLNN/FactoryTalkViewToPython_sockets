import socket
import logging

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s'
)

# Replace with the IP address and port number used by your VBA application
HOST = '127.0.0.1'
PORT = 1234

# Create a socket object
sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)

# Attempt to connect to the VBA application's socket
try:
    sock.connect((HOST, PORT))
    logging.info('Connected to VBA application')
except ConnectionRefusedError:
    logging.error('Unable to connect to VBA application')
    exit()

# Receive data from the VBA application's socket and write it to the console
while True:
    try:
        data = sock.recv(1024)
        if not data:
            break
        message = data.decode('utf-8')
        # Process the message as needed
        logging.info('Received message: %s', message)
    except ConnectionResetError:
        logging.error('Connection to VBA application reset')
        break
    except UnicodeDecodeError:
        logging.error('Unable to decode message')
    except Exception as e:
        logging.error('Unhandled exception: %s', e)

# Close the socket
sock.close()
logging.info('Socket closed')
