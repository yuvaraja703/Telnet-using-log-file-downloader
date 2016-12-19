Dim Pno,Sno,ssno
msgbox ("Connecting to the server")
killprocess()


sub telnetconnect()
dim oShell
set oShell = CreateObject("WScript.Shell")
oShell.Run "cmd"
WScript.Sleep 100 

WScript.Sleep 100 
oShell.SendKeys "ftp xxx.xx.xxx.xxx"
WScript.Sleep 100
oShell.SendKeys "{enter}"
WScript.Sleep 100
oShell.SendKeys "xxxxxxxxxxxx"
WScript.Sleep 100
oShell.SendKeys "{enter}"
WScript.Sleep 100
oShell.SendKeys "xxxxxxxxxx"
WScript.Sleep 10
oShell.SendKeys "{enter}"
oShell.SendKeys "{enter}"
WScript.Sleep 50
Change_Directories="cd /mnt/xxxx/xxxxx/"
text = "Enter the Part Number" 
text = text + (Chr(13) & Chr(13) & Chr(10)) 
text = text + "Example:xxx-xxxxxxxxx-x" 
Pno=InputBox(text,"Opening Part Number Directory","xxx-xxxxxxx-x")
Pno=ucase(Pno)
Add1=Change_Directories+Pno
oShell.SendKeys Add1
oShell.SendKeys "{enter}"
Call download()


WScript.Sleep 1000
WScript.Quit
End Sub


Sub killprocess()
Set Whell = CreateObject("WScript.Shell")
WScript.Sleep 50
Whell.Run "cmd"
WScript.Sleep 50
WScript.Sleep 50
Whell.SendKeys "taskkill /im wscript.exe"
Whell.SendKeys "{enter}"
WScript.Sleep 50
Whell.SendKeys "exit{enter}"
Call telnetconnect()
end Sub


Sub download()
get_file="get "
save_location=" C:\Log\"
file_mode=".txt"
do while True
set kshell = CreateObject("WScript.Shell")

	text1 = "Enter the Serial Number" 
	text1 = text1 + (Chr(13) & Chr(13) & Chr(10)) 
	text1 = text1 + "Example:  xx-xxxx-12-1234" 
    Sno=InputBox(text1,"Serial Number",sno)
	Sno=ucase(Sno)
    Add2=get_file+Sno+save_location+Sno+file_mode
	kShell.SendKeys Add2
	kShell.SendKeys "{enter}"
	text2 = "Click YES --- Continue Process" 
	text2 = text2 + (Chr(13) & Chr(13) & Chr(10)) 
	text2 = text2 + "Click NO --- Change to New Part Number" 
	text2 = text2 + (Chr(13) & Chr(13) & Chr(10)) 
	text2 = text2 + "Click CANCEL --- Exit the Software " 	
	text2 = text2 + (Chr(13) & Chr(13) & Chr(10)) 
	text2 = text2 + "Downloaded LOGFile S.NO: " 
ssno=sno
    Num=MsgBox(text2 & sno,3,"Logfile downloaded")
       	if Num = 2 then
        MsgBox"Bye"
        kShell.SendKeys "bye"
        kShell.SendKeys "{enter}"
        kShell.SendKeys "exit"
        kShell.SendKeys "{enter}"
        Call close()
        Exit do
    elseif Num = 6 Then
    	MsgBox"continuing the process"
    ElseIf Num = 7 Then
        text3="FTP Timeout"  
    	text3 = text3 + (Chr(13) & Chr(13) & Chr(10)) 
		text3=text3 + "Connecting to the server"
        MsgBox (text3)
        kShell.SendKeys "bye"
        kShell.SendKeys "{enter}"
        kShell.SendKeys "exit"
        kShell.SendKeys "{enter}"
        Call telnetconnect()
    else
        MsgBox "Selected is wrong"
    end If
 Loop
 End Sub


Sub close()
set chell = CreateObject("WScript.Shell")
WScript.Sleep 50
chell.Run "cmd"
WScript.Sleep 50
WScript.Sleep 50
chell.SendKeys "taskkill /im wscript.exe"
chell.SendKeys "{enter}"
wScript.Sleep 50
chell.SendKeys "exit{enter}"
End Sub
