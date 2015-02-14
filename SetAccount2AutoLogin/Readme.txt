自動登入的主要程式碼是Account_AutoLogin.vbs
執行此腳本請透過CMD,並起要賦予最高執行權限啟動CMD來操作腳本
否則再寫入自動登入時會因權限不足失敗

CMD執行時可以帶兩個參數

如果是一個參數
請直接輸入帳號的密碼(範例如下)
>BOT_Account_AutoLogin.vbs pass$$word

如果是兩個參數
第一個參數是選擇Mapping的方法
如果填 0 表示要用ip對應
如果要填 1 表示用Hostname 主機名稱對應

第二個參數則是密碼

>Account_AutoLogin.vbs 0 pass$$word

>Account_AutoLogin.vbs 1 pass$$word

檔案的路徑請開起腳本修改,修改的變數名稱為
MAPPING_FILE_SOURCE_PATH

此腳本讀取範例檔案是Example_Mapping_Files.txt
透過,去區分
ip,主機名稱,要用來作自動登入的使用者帳號

By Eason Kuo (Yi-cheng Kuo) 20150214
======================================================
English:

1. The main script is "Account_AutoLogin.vbs"
Please execute this script by command Line,and need thse Administrator privilege, or won't succese when writing the setting data to Registry key

2. The script will read "Example_Mapping_Files.txt" file and find the certain account you want to autologin through
ip and computer name

3. You can type 2 arguements when executing the script

If you input only 1 arguement
Please input the password that certain account you want to autologin
>BOT_Account_AutoLogin.vbs pass$$word

If you input 2  arguements 
The first arguement is Mapping account method - if you type 0 means you want to map account by ip ,and you type 1 means you want to use computer name to map account.

and the Second arguement is password

Sample :
>Account_AutoLogin.vbs 0 pass$$word

>Account_AutoLogin.vbs 1 pass$$word


4. If you want to modify the read file name and path, Please open the "Account_AutoLogin.vbs" and change the MAPPING_FILE_SOURCE_PATH string value.


By Eason Kuo (Yi-cheng Kuo) 20150214