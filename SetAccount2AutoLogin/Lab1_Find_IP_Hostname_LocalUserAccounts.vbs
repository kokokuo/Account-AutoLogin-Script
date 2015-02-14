On Error Resume Next 
 
Dim computer_ip,computer_hostname
Dim user_accounts_name()
accounts_size = 0
strComputer = "." 
Dim output_message
Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

'	==== Get Ip Address ====
Set IPConfigSet = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True")
 
For Each IPConfig in IPConfigSet 
    If Not IsNull(IPConfig.IPAddress) Then
    '	IPAddress(0) is ip   
        computer_ip = IPConfig.IPAddress(0)  
    End If 
Next

'	==== Get Computer Name ====
Set objWshNetwork = WScript.CreateObject("WScript.Network")
computer_hostname = objWshNetwork.Computername

' Add previous message 
output_message = "Ip Address =" & computer_ip & vbNewLine _
				 & "Computer Name = " & computer_hostname & vbNewLine _ 
				 & "User Accounts" & vbNewLine _
				 & "======" & vbNewLine


'	==== Get User Accounts ====
Set  objLocalUsersInfo = objWMIService.ExecQuery _ 
    ("Select * from Win32_UserAccount Where LocalAccount = True") 
For Each objLocalUser in objLocalUsersInfo 
	ReDim Preserve user_accounts_name(accounts_size)
    user_accounts_name(accounts_size) = objLocalUser.Name
    'Another way to get user accounts by Caption...
    'user_accounts_name(accounts_size) =  user_accountsobjLocalUser.Caption
    accounts_size = accounts_size + 1
Next

size = 0
For Each userName in user_accounts_name
	 output_message = output_message & size + 1 & ". " & userName & vbNewLine
	 size = size + 1
Next

MsgBox(output_message)