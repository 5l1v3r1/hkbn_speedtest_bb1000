sendUpLoad="0"
sendDownLoad="0"
sendDownloadSize="0"
sendUploadSize="0"
UpLoadTime="0"
DownloadTime="0"
serverIP="iperf4.speedtest.com.hk"
Set IE = wscript.CreateObject("internetexplorer.application")
IE.Navigate "https://reg.hkbn.net/IperfTestLogin/savelog.do?downLoad="&sendDownLoad&"&upLoad="&sendUpLoad&"&id=E6C69360BF&errorCode=1005&serverIP=" & serverIP
MsgBox "閣下之寬頻分享器不支援本測試，請將電腦連接至牆身寬頻插座後再試一次 (1005)" & Chr(10) , vbOKOnly, "HKBN - Iperf速度測試"

