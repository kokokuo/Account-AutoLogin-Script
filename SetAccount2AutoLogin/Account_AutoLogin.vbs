   ==== Golbal Var ===================================='

Const IP_INDEX = 0
Const HOSTNAME_INDEX = 1
Const AUTO_LOGIN_ACCOUNT = 2
Const STR_COMPUTER = "." 
Const MAPPING_FILE_SOURCE_PATH = "D:\SetAccount2AutoLogin\Example_Mapping_Files.txt"

'   ==== Method ===================================='

'   ====  Auto Login ==== '
Sub ExecuteAddAccountToAutoLogin(found_user_account_name,password,computer_hostname)
    Wscript.echo "click ok to continue please wait for the completed message before logging off or shutting down!"
    Const HKEY_LOCAL_MACHINE = &H80000002
    '''  Setting forceAutoLogon to true
    Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &  STR_COMPUTER &  "\root\default:StdRegProv")
    strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
    strValueName = "ForceAutoLogon"
    strValue = "1"
    Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
    If (Return <> 0) Or (Err.Number <> 0) Then   
        Wscript.Echo strValueName & "Added Value = " & strValue & " Cause Fault!"
        Wscript.Quit
    End If
    ''' Setting the default username to be the same as strUserName
    strValueName = "DefaultUserName"
    strValue = found_user_account_name
    Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
    If (Return <> 0) Or (Err.Number <> 0) Then   
        Wscript.Echo strValueName & "Added Value = " & strValue & " Cause Fault!"
        Wscript.Quit
    End If
    '''  Setting the default password to be same as strPassword
    strValueName = "DefaultPassword"
    strValue = password
    Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
    If (Return <> 0) Or (Err.Number <> 0) Then   
        Wscript.Echo strValueName & "Added Value = " & strValue & " Cause Fault!"
        Wscript.Quit
    End If
    '''  Setting AutoAdminLogon to True
    strValueName = "AutoAdminLogon"
    strValue = "1"
    Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
    If (Return <> 0) Or (Err.Number <> 0) Then   
        Wscript.Echo strValueName & "Added Value = " & strValue & " Cause Fault!"
        Wscript.Quit
    End If
    '''  Setting the default login domain to be the local machine
    strValueName = "DefaultDomainName"
    strValue = computer_hostname
    Return = oReg.SetStringValue(HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue)
    If (Return <> 0) Or (Err.Number <> 0) Then   
        Wscript.Echo strValueName & "Added Value = " & strValue & " Cause Fault!"
        Wscript.Quit
    End If
    
    Wscript.echo "Completed: please reboot to save changes"
End Sub

'   ==== Read Mapping File ====='
Sub ReadAutoLoginMappingFile(file_source_path, objFileMappingDictionary)
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Const ForReading = 1

    '4.1   ==== Read File ===='
    Set objFile = objFSO.OpenTextFile (file_source_path, ForReading)

    '4.2   ===== Put Data into Dictionary (Split String) ====='
    Do Until objFile.AtEndOfStream
        strNextLine = objFile.Readline
        If strNextLine <> "" Then
            ' 4.2.1  ==== Split String : return type is array ==== '
            split_array  = Split(strNextLine,",")
            '4.2.2   ==== Add to dicitonary (According the mappint option)===='
            If mapping_option = IP_INDEX Then
                objFileMappingDictionary.Add split_array(IP_INDEX), split_array(AUTO_LOGIN_ACCOUNT)
            Else
                objFileMappingDictionary.Add split_array(HOSTNAME_INDEX), split_array(AUTO_LOGIN_ACCOUNT)
            End If

        End If
    Loop
    objFile.Close
End Sub

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


Function GetIPAddress(objWMIService)
    '   ==== Get Ip Address ====
    Set IPConfigSet = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled=True")
    For Each IPConfig in IPConfigSet 
        If Not IsNull(IPConfig.IPAddress) Then
        '   IPAddress(0) is ip   
            GetIPAddress = IPConfig.IPAddress(0)
        End If 
    Next
End Function

Function GetComputerName()
    Set WshNetwork = WScript.CreateObject("WScript.Network")
    GetComputerName = WshNetwork.Computername
End Function

'   ====  Get Command Line input ====='
Sub GetCommnadInputArgs(mapping_option,password)
    Set cmd_args = WScript.Arguments
    '   Script 
    '   If there is no args , Default mapping by ip and means don't need password
    '   If there is 1 arg , Default mapping by ip and means need password
    '   If there are 2 args , means need choose mapping by ip or hostname(input 0 is ip and 1 is computer name) and need password
    If cmd_args.Count = 2 Then
        mapping_arg = CInt(cmd_args(0))
        If mapping_arg = 0 Or mapping_arg = 1 Then
            mapping_option = mapping_arg
        Else
            MsgBox("Mapping Option Error,not 0 or 1")
            Wscript.Quit
        End If
        password = cmd_args(1)
    ElseIf cmd_args.Count = 1 Then
        password = cmd_args(0)
        WScript.Echo "Your Input Password is :" & password
    ElseIf cmd_args.Count = 0 Then
        has_passord = False
    Else
        MsgBox("The number of Arguments Error (only accept 2 args)")
        WScript.Quit
    End If 
End Sub



'   ==== Main : Script code Starter ======================================================='
Sub Main
'   1. ==== Init ===='
    Dim computer_ip,computer_hostname,user_account_name 
    Dim computer_user_accounts_name() 'all computer(hostname or ip) and user account from mapping files'
   
    Dim output_message

    mapping_option = 0 'mapping by ip or computer name'
    Dim password
    has_passord = True

    '2.   ====  Get Command Line input ====='
    Call GetCommnadInputArgs(mapping_option,password)

    '3.   ==== Get Computer Information ===='
    Set objWMIService = GetObject("winmgmts:" _ 
    & "{impersonationLevel=impersonate}!\\" & STR_COMPUTER & "\root\cimv2") 

    '3.1   ==== Get IP Address ===='
    computer_ip = GetIPAddress(objWMIService)
    '3.2  ==== Get Computer Name ====
    computer_hostname = GetComputerName()
    
    ' Add previous message 
    output_message = "Ip Address =" & computer_ip & vbNewLine _
                 & "Computer Name = " & computer_hostname & vbNewLine _ 
                 & "User Accounts" & vbNewLine _
                 & "======" & vbNewLine

    ' 3.3  ==== Get User Accounts ====
    Call GetUserAccounts(objWMIService,computer_user_accounts_name)

    size = 0
    For Each userName in computer_user_accounts_name
         output_message = output_message & size + 1 & ". " & userName & vbNewLine
         size = size + 1
    Next

    '   ==== Print computer info ===='
    WScript.Echo output_message

'   ==================================================================='
    WScript.Echo "Start off Reading File and Mapping..."
    '4.   ===== Read Mapping Files ====' 
    Set objFileMappingDictionary = CreateObject("Scripting.Dictionary")
    Call ReadAutoLoginMappingFile(MAPPING_FILE_SOURCE_PATH,objFileMappingDictionary)

    '5.   ==== File Data & Computer Info Mapping ===='
    has_found_account = False
    has_found_ip_hostname = False
    mapping_keys = objFileMappingDictionary.Keys
    For Each key in mapping_keys
       
        '   ====  If Found Account then break loop ====='
        If has_found_account = True Then
            Exit For
        Else
            '5.1    ===== Check the mapping choice ,mapping choice is ip ====='
            If mapping_option = IP_INDEX Then    
                If  computer_ip = key  Then
                    has_found_ip_hostname = True
                    For Each userName in computer_user_accounts_name
                        If userName = objFileMappingDictionary.Item(key) Then
                            WScript.Echo "Find ip = " & key & vbNewLine & "User Account = " & objFileMappingDictionary.Item(key)
                            user_account_name =  objFileMappingDictionary.Item(key)
                            has_found_account = True
                            Exit For
                        End If 
                    Next 
                End If
            '5.2    ===== mapping choice is hostname ====='
            Else            
                ' The hostname is uppercase when get the data from WScript.CreateObject("WScript.Network")
                ' So need to uppercase mapping data
                If  computer_hostname = ucase(key)  Then
                    has_found_ip_hostname = True
                    For Each userName in computer_user_accounts_name
                        If userName = objFileMappingDictionary.Item(key) Then
                            WScript.Echo "Find hostname = " & key & vbNewLine & "User Account = " & objFileMappingDictionary.Item(key)
                            user_account_name =  objFileMappingDictionary.Item(key)
                            has_found_account = True
                            Exit For
                        End If 
                    Next 
                End If
            End If
        End If
    Next
    
    If has_found_ip_hostname = False Then
        WScript.Echo "No mapping machine"
    End If
    If has_found_account = False Then
        WScript.Echo "No mapping Account"
    End If
 
    '   ====6.  OK! Run Auto Login .. ===='
    If has_passord Then
        If  has_found_account And has_found_ip_hostname Then
    '   Validate but these code can't work'
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
            Call ExecuteAddAccountToAutoLogin(user_account_name,password,computer_hostname)
        Else
            WScript.echo "Can't Complete autologin ,please check your password or ip or hostname correct!"  
        End iF

    Else
        WScript.echo "Can't Complete autologin ,please check your password or ip or hostname correct!"
    End If


End Sub



' ==== Rnu Main ===='
Main









