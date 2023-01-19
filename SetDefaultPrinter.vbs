''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
On Error Resume Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' SET NETBIOS DOMAINNAME AND GROUPNAMES HERE !
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

DomainString = ""

Group001 = ""
Group002 = ""
Group003 = ""
Group004 = ""
Group005 = ""
Group006 = ""
Group007 = ""
Group008 = ""
Group009 = ""
Group010 = ""
Group011 = ""
Group012 = ""
Group013 = ""
Group014 = ""
Group015 = ""

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Dim DomainString
Dim UserObj
Dim GroupObj
Dim strUser
Dim WSHShell


Set WSHShell = WScript.CreateObject("WScript.Shell")
Set WshNetwork = WScript.CreateObject("WScript.Network")
Set fso = CreateObject("Scripting.FileSystemObject")
strDesktop = WshShell.SpecialFolders("Desktop")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Retrieve currently logged-on user
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

strUser = WshNetwork.UserName

'WScript.Echo strUser

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Determine Groupmemberships for current user
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Set UserObj = GetObject("WinNT://" & DomainString & "/" & strUser)

For Each GroupObj in UserObj.Groups
	Group = GroupObj.Name
	WshCreateShortcut Group
Next

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' SUBROUTINE'S
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Sub routine for creating Printer connections
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub WshCreateShortcut(Group)

    Select Case Group
	Case Group001
	   WshNetwork.SetDefaultPrinter ""

	Case Group002
	   WshNetwork.SetDefaultPrinter ""

	Case Group003
	   WshNetwork.SetDefaultPrinter ""

	Case Group004
	   WshNetwork.SetDefaultPrinter ""

	Case Group005
	   WshNetwork.SetDefaultPrinter ""

	Case Group006
	   WshNetwork.SetDefaultPrinter ""

	Case Group007
	   WshNetwork.SetDefaultPrinter ""

	Case Group008
	   WshNetwork.SetDefaultPrinter ""

	Case Group009
	   WshNetwork.SetDefaultPrinter ""

	Case Group010
	   WshNetwork.SetDefaultPrinter ""

	Case Group011
	   WshNetwork.SetDefaultPrinter ""

	Case Group012
	   WshNetwork.SetDefaultPrinter ""

	Case Group013
	   WshNetwork.SetDefaultPrinter ""

	Case Group014
	   WshNetwork.SetDefaultPrinter ""

	Case Group015
	   WshNetwork.SetDefaultPrinter ""

	End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Sub routine for writing log information in the eventlog
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub WshLogEvent(strDescription)
	If Err.Number = 0 Then
		If InfoMessages Then
			strErrorMsg = "User: " & strUser & vbCrLf & _ 
                        "Script: " & ScriptPath & "\" & WScript.ScriptName & ", Version: " & version & _
                        vbCrLf & "Message : " & vbCrLf & strDescription
                WshShell.LogEvent 0, strErrorMsg
		End IF
	Else
		         strErrorMsg = "User: " & strUser & vbCrLf & _ 
                         "Script: " & ScriptPath & "\" & WScript.ScriptName & ", version: " & Version & vbCrLf & _
                         "Message : " & vbCrLf & strDescription & vbCrLf & "Error Number Is : " & _ 
                         Err.Number & vbCrLf & "Error Is : " & Err.Description
                WshShell.LogEvent 2, strErrorMsg
	End If
	Err.Clear
End Sub

Sub WshCopyEvent(strCrt)
    strNotifMsg = "Script: " & WScript.ScriptName & VbCrLf & "Version: " & version & _
    VbCrLf & "Copyright: " & strCrt
    WshShell.LogEvent 0, strNotifMsg
End Sub