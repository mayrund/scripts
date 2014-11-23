' CopyFilesFromNetworkShare
' Mayrun Digmi

' Variables
' ---------
ServerShare = "\\IP.ADDRESS.HERE\c$"
ServerPath = "\windows\test.log"
ShareUserName = "domain\username"
SharePassword = "password"
DestinationPath = "c:\test.log"

' Start
Set NetworkObject = CreateObject("WScript.Network")
Set FSO = CreateObject("Scripting.FileSystemObject")

NetworkObject.MapNetworkDrive "", ServerShare, False, ShareUserName, SharePassword
FSO.CopyFile ServerShare & ServerPath, DestinationPath
NetworkObject.RemoveNetworkDrive ServerShare, True, False

Set FSO = Nothing
Set ShellObject = Nothing
Set NetworkObject = Nothing

MsgBox("Done.")
