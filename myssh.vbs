' <---------- Made By The Genius ---------->

Dim username
Dim port
Dim defaultIP
Dim defaultPort
Dim configFile
Dim ipAddress
Dim portInput

' Set username
username = "u0_a123" ' Replace with your username

' Default values for IP and Port
defaultIP = "192.168.4.1" ' Default IP address
defaultPort = "8022" ' Default port

' Configuration file to store last used IP address and port
configFile = "ssh_config.txt" ' This file will store the last used IP address and port

' Initialize previousIP and previousPort
Dim previousIP, previousPort
previousIP = defaultIP
previousPort = defaultPort

' Create FileSystemObject to read from the config file (if it exists)
Dim fso, file
Set fso = CreateObject("Scripting.FileSystemObject")

' Check if the configuration file exists, if yes, read the last IP and port
If fso.FileExists(configFile) Then
    Set file = fso.OpenTextFile(configFile, 1) ' Open file for reading
    previousIP = file.ReadLine() ' Read last used IP address
    previousPort = file.ReadLine() ' Read last used port
    file.Close
End If

' Prompt user for IP address with last used IP address as the default value
ipAddress = InputBox("Welcome Genius. Please enter the IP address to connect:", "SSH Connection by Genius", previousIP)

' Check if the IP address is entered
If ipAddress = "" Then
    MsgBox "No IP address entered. Exiting.", vbExclamation, "Error"
    WScript.Quit
End If

' Prompt user for Port number with last used port as the default value
portInput = InputBox("Please enter the port number to connect:", "Enter Port", previousPort)

' Check if the port is entered
If portInput = "" Then
    MsgBox "No port number entered. Exiting.", vbExclamation, "Error"
    WScript.Quit
End If

' Show the connection details in a message box
Dim message
message = "Connecting to IP: " & ipAddress & vbCrLf & _
          "Username: " & username & vbCrLf & _
          "Port: " & portInput
MsgBox message, vbInformation, "Connection Details"

' Create the shell object to run the command
Dim shell
Set shell = CreateObject("WScript.Shell")

' Specify the full path to SSH (you may need to change this path if OpenSSH is installed elsewhere)
Dim command
command = "C:\Windows\System32\OpenSSH\ssh.exe " & username & "@" & ipAddress & " -p " & portInput ' Example SSH command

' Execute the command, and wait for it to finish (use 1 for normal window)
shell.Run command, 1, True

' Save the entered IP and port for next time
Set file = fso.CreateTextFile(configFile, True) ' Create the file if it doesn't exist
file.WriteLine(ipAddress) ' Save the new IP address
file.WriteLine(portInput) ' Save the new port number
file.Close

' Show completion message
MsgBox "Connection attempt completed.", vbInformation, "Done"
