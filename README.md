# VBA-Python-Socket-Communication

This repository contains sample code for communicating data between FactoryTalk View and a Python application using sockets.

## Overview

The project includes two components:

- A VBA application that runs within FactoryTalk View and sends temperature data to a client application over a socket
- A Python client application that receives temperature data from the VBA application over a socket and processes it for further analysis or visualization

## Requirements

- FactoryTalk View 7.0 or higher
- Microsoft Excel 2010 or higher
- Python 3.6 or higher

### Why Excel?
Microsoft Excel is required for the VBA portion of the "FactoryTalkViewToPython_sockets" project because it provides an environment for creating VBA macros and projects.

In the project, the VBA application is created using Microsoft Excel's Visual Basic Editor (VBE), which is a built-in tool that provides a graphical interface for creating, editing, and debugging VBA code.

The VBA application uses the Winsock control to send temperature data over a socket to a Python client application. The Winsock control is available in Microsoft Office applications, including Excel, and provides an easy way to establish and manage network connections.

Overall, while Microsoft Excel is not strictly required for the Python portion of the project, it is needed for the VBA portion in order to create the VBA application that communicates with the Python client application over a socket.

## Installation

To use the VBA application, follow these steps:

1. Open FactoryTalk View and create a new display
2. Insert a new ActiveX control and choose the "Microsoft Winsock Control" option
3. Name the control "Winsock1" and set its properties as follows:
   - RemoteHost: set to the IP address of the machine running the Python client application
   - RemotePort: set to the port number used by the Python client application
4. Add the VBA code provided in the "FTV_VBA_sockets.bas" file to the VBA project associated with the display
5. Save and run the display

To use the Python client application, follow these steps:

1. Clone this repository to your local machine
2. Install the required Python packages using pip:
```
pip install -r requirements.txt
```

3. Modify the `HOST` and `PORT` variables in the "python FTV_PY_sockets.py" file to match the IP address and port number used by the VBA application
4. Run the Python script:
```
python FTV_PY_sockets.py
```


The Python script will connect to the VBA application's socket and continuously read temperature data that is sent over the socket, writing it to the console for further analysis or visualization.

## License

This code is licensed under the MIT License. See the "LICENSE" file for more information.

## Credits

This project was developed by YCML@novonordisk.com (Marco Molinari) for NNUSBPI (Novo Nordisk US Bio Production, Inc.).
