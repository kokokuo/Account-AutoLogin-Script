# Account-AutoLogin-Script
This is a VB Script which make your windows pc can autologin by given a certain account

# How to use

1.The main script is **"Account_AutoLogin.vbs"**
Please execute this script by command Line,and need thse Administrator privilege, or won't succese when writing the setting data to Registry key

2.The script will read "Example_Mapping_Files.txt" file and find the certain account you want to autologin through
ip and computer name

3.You can type 2 arguements when executing the script

#### If you input only 1 arguement
Please input the password that certain account you want to autologin
`BOT_Account_AutoLogin.vbs pass$$word`

#### If you input 2  arguements 
The first arguement is Mapping account method.
if you type 0 means you want to map account by ip ,and you type 1 means you want to use computer name to map account.

and the Second arguement is password

Sample :

`Account_AutoLogin.vbs 0 pass$$word`

`Account_AutoLogin.vbs 1 pass$$word`


4.If you want to modify the read file name and path, Please open the **"Account_AutoLogin.vbs"** and change the `MAPPING_FILE_SOURCE_PATH` string value.



