


'   ==== Method ===='
Function GetIPAddress(ByRef objWMIService)
    Dim local_computer_ip
    '   ==== Get Ip Address ====
    Set IPConfigSet = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True")
    For Each IPConfig in IPConfigSet 
        If Not IsNull(IPConfig.IPAddress) Then
        '   IPAddress(0) is ip   
            local_computer_ip = CStr(IPConfig.IPAddress(0))  
        End If 
    Next
    MsgBox("Hello = " & local_computer_ip)
    
    local_computer_ip
End Function


Function GetComputerName()
    CompName =""
    Set WshNetwork = WScript.CreateObject("WScript.Network")
    CompName = WshNetwork.Computername
    MsgBox("Hello = " & CompName)
    CompName
End Function

Function GetUserAccounts(objWMIService, computer_user_accounts_name)
    accounts_size = 0
    '   ==== Get User Accounts ====
    Set  objLocalUsersInfo = objWMIService.ExecQuery _ 
        ("Select * from Win32_UserAccount Where LocalAccount = True") 
    For Each objLocalUser in objLocalUsersInfo 
        ReDim Preserve computer_user_accounts_name(accounts_size)
        computer_user_accounts_name(accounts_size) = objLocalUser.Name
        'Another way to get user accounts by Caption...
        'computer_user_accounts_name(accounts_size) =  user_accountsobjLocalUser.Caption
        accounts_size = accounts_size + 1
    Next
End Function

'   ==== Main : Script code Starter ===='


Sub Main
'    ==== Init ===='
    Dim computer_ip,computer_hostname
    Dim computer_user_accounts_name()
    strComputer = "." 
    Dim output_message

    Const IP_INDEX = 0
    Const HOSTNAME_INDEX = 1
    Const AUTO_LOGIN_ACCOUNT = 2

    mapping_choice = 0 'mapping by ip or computer name'
    Dim password
    has_passord = True

    Dim found_user_account
    '   ====  Get Command Line input ====='
    Set cmd_args = WScript.Arguments
    '   Script 
    '   If there is no args , Default mapping by ip and means don't need password
    '   If there is 1 arg , Default mapping by ip and means need password
    '   If there are 2 args , means need choose mapping by ip or hostname(input 0 is ip and 1 is computer name) and need password
    If cmd_args.Count = 2 Then
        mapping_arg = CInt(cmd_args(0))
        If mapping_arg = 0 Or mapping_arg = 1 Then
            mapping_choice = mapping_arg
        Else
            MsgBox("Mapping Choice Error")
            Exit Sub
        End If
        password = cmd_args(1)
    ElseIf cmd_args.Count = 1 Then
        password = cmd_args(0)
        WScript.Echo "Your Input Password is :" & password
    ElseIf cmd_args.Count = 0 Then
        has_passord = False
    Else
        MsgBox("Arguments Number Error")
        Exit Sub
    End If 

    '   ==== Get Computer Information ===='
    Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
    '   ==== Get IP Address ===='
    Set IPConfigSet = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True")
    For Each IPConfig in IPConfigSet 
        If Not IsNull(IPConfig.IPAddress) Then
        '   IPAddress(0) is ip   
            computer_ip = IPConfig.IPAddress(0)
        End If 
    Next
    '  ==== Get Computer Name ====

    Set WshNetwork = WScript.CreateObject("WScript.Network")

    computer_hostname = WshNetwork.Computername
    
    ' Add previous message 
    output_message = "Ip Address =" & computer_ip & vbNewLine _
                 & "Computer Name = " & computer_hostname & vbNewLine _ 
                 & "User Accounts" & vbNewLine _
                 & "======" & vbNewLine

    '   ==== Get User Accounts ====
    Call GetUserAccounts(objWMIService,computer_user_accounts_name)

    size = 0
    For Each userName in computer_user_accounts_name
         output_message = output_message & size + 1 & ". " & userName & vbNewLine
         size = size + 1
    Next

    MsgBox(output_message)

'   ==================================================================='
    MsgBox("Start off Reading File and Mapping...")
    '   ===== Read Mapping Files ===='
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFileMappingDictionary = CreateObject("Scripting.Dictionary")

    Const ForReading = 1

    ' ==== Read File ===='
    Set objFile = objFSO.OpenTextFile ("D:\SetAccount2AutoLogin\Example_Mapping_Files.txt", ForReading)


    Do Until objFile.AtEndOfStream
        strNextLine = objFile.Readline
        If strNextLine <> "" Then
            '   ==== Split String : return type is array ==== '
            split_array  = Split(strNextLine,",")
            '   ==== Add to dicitonary ===='
            If mapping_choice = IP_INDEX Then
                objFileMappingDictionary.Add split_array(IP_INDEX), split_array(AUTO_LOGIN_ACCOUNT)
            Else
                objFileMappingDictionary.Add split_array(HOSTNAME_INDEX), split_array(AUTO_LOGIN_ACCOUNT)
            End If
        End If
    Loop
    objFile.Close

    '   ==== Get dictionary Data ===='
    '   ==== Dictionary => Ip,User Account ===='
    has_found_account = False
    has_found_ip_hostname = False
    mapping_keys = objFileMappingDictionary.Keys
    For Each key in mapping_keys
        '   ==== Maping Ip ===='
        'MsgBox CInt(computer_ip) & "," & CInt(ip_key)
        If has_found_account = True Then
            Exit For
        Else
            ' Check the mapping choice ,mapping choice is ip'
            If mapping_choice = IP_INDEX Then    
                If  computer_ip = key  Then
                    has_found_ip_hostname = True
                    For Each userName in computer_user_accounts_name
                        If userName = objFileMappingDictionary.Item(key) Then
                            MsgBox "Find ip = " & key & vbNewLine & "User Account = " & objFileMappingDictionary.Item(key)
                            found_user_account =  objFileMappingDictionary.Item(key)
                            has_found_account = True
                            Exit For
                        End If 
                    Next 
                End If
            ' mapping choice is hostname'
            Else            
                ' The hostname is uppercase when get the data from WScript.CreateObject("WScript.Network")
                ' So need to uppercase mapping data
                If  computer_hostname = ucase(key)  Then
                    has_found_ip_hostname = True
                    For Each userName in computer_user_accounts_name
                        If userName = objFileMappingDictionary.Item(key) Then
                            MsgBox "Find hostname = " & key & vbNewLine & "User Account = " & objFileMappingDictionary.Item(key)
                            found_user_account =  objFileMappingDictionary.Item(key)
                            has_found_account = True
                            Exit For
                        End If 
                    Next 
                End If
            End If
        End If
    Next
    
    If has_found_ip_hostname = False Then
        MsgBox("No mapping machine")
    End If
    If has_found_account = False Then
        MsgBox("No mapping Account")
    End If



    '   ==== OK! Run Auto Login .. ===='
    Const HKEY_LOCAL_MACHINE = &H80000002
    Set WshNetwork = WScript.CreateObject("WScript.Network")
    strUserName = found_user_account
    Dim strPassword
    strComputer = WshNetwork.ComputerName
    If has_passord Then
        If  has_found_account And has_found_ip_hostname Then
    '   Validate'
'       strDomain = computer_hostname
'        strUsername = found_user_account
'        strPassword = password
'        Set objDS = GetObject("LDAP:")
'        On Error Resume Next
'        Set objDomain = objDS.OpenDSObject("LDAP://" & strDomain, strUsername, strPassword, ADS_SECURE_AUTHENTICATION)
'        If Err.Number Then
'            WScript.Echo "Validiate not success"
'        Else
'            WScript.Echo "Validiate success"
'        End If''

            strPassword = password


            Wscript.echo "click ok to continue please wait for the completed message before logging off or shutting down!"

             
            '''  Setting forceAutoLogon to true
            Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &  strComputer &  "\root\default:StdRegProv")
            strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
            strValueName = "ForceAutoLogon"
            strValue = "1"
            Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
            WScript.Echo  Return
            If (Return = 0) And (Err.Number = 0) Then   
        '        Wscript.Echo strValueName & "Added Value = " & strValue
            End If
            ''' Setting the default username to be the same as strUserName
            strValueName = "DefaultUserName"
            strValue = strUsername
            Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
            If (Return = 0) And (Err.Number = 0) Then   
        '        Wscript.Echo strValueName & "Added Value = " & strValue
            End If
            '''  Setting the default password to be same as strPassword
            strValueName = "DefaultPassword"
            strValue = strPassword
            Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
            If (Return = 0) And (Err.Number = 0) Then   
        '        Wscript.Echo strValueName & "Added Value = " & strValue
            End If
            '''  Setting AutoAdminLogon to True
            strValueName = "AutoAdminLogon"
            strValue = "1"
            Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
            If (Return = 0) And (Err.Number = 0) Then   
        '        Wscript.Echo strValueName & "Added Value = " & strValue
            End If
            '''  Setting the default login domain to be the local machine
            strValueName = "DefaultDomainName"
            strValue = strComputer
            Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
            If (Return = 0) And (Err.Number = 0) Then   
        '        Wscript.Echo strValueName & "Added Value = " & strValue
            End If
            
            Wscript.echo "Completed: please reboot to save changes"
        Else
            WScript.echo "Can't Complete autologin ,please check your password or ip or hostname correct!"  
        End iF

    Else
        WScript.echo "Can't Complete autologin ,please check your password or ip or hostname correct!"
    End If


End Sub




Main









