�۰ʵn�J���D�n�{���X�OAccount_AutoLogin.vbs
���榹�}���гz�LCMD,�ð_�n�ᤩ�̰������v���Ұ�CMD�Ӿާ@�}��
�_�h�A�g�J�۰ʵn�J�ɷ|�]�v����������

CMD����ɥi�H�a��ӰѼ�

�p�G�O�@�ӰѼ�
�Ъ�����J�b�����K�X(�d�Ҧp�U)
>BOT_Account_AutoLogin.vbs pass$$word

�p�G�O��ӰѼ�
�Ĥ@�ӰѼƬO���Mapping����k
�p�G�� 0 ��ܭn��ip����
�p�G�n�� 1 ��ܥ�Hostname �D���W�ٹ���

�ĤG�ӰѼƫh�O�K�X

>Account_AutoLogin.vbs 0 pass$$word

>Account_AutoLogin.vbs 1 pass$$word

�ɮת����|�ж}�_�}���ק�,�ק諸�ܼƦW�٬�
MAPPING_FILE_SOURCE_PATH

���}��Ū���d���ɮ׬OExample_Mapping_Files.txt
�z�L,�h�Ϥ�
ip,�D���W��,�n�Ψӧ@�۰ʵn�J���ϥΪ̱b��

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