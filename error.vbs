sendUpLoad="0"
sendDownLoad="0"
sendDownloadSize="0"
sendUploadSize="0"
UpLoadTime="0"
DownloadTime="0"
serverIP="iperf4.speedtest.com.hk"
Set IE = wscript.CreateObject("internetexplorer.application")
IE.Navigate "https://reg.hkbn.net/IperfTestLogin/savelog.do?downLoad="&sendDownLoad&"&upLoad="&sendUpLoad&"&id=E6C69360BF&errorCode=1005&serverIP=" & serverIP
MsgBox "�դU���e�W���ɾ����䴩�����աA�бN�q���s�����𨭼e�W���y��A�դ@�� (1005)" & Chr(10) , vbOKOnly, "HKBN - Iperf�t�״���"

