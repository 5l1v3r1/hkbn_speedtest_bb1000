ForReading = 1
ForWriting = 2
ForAppending = 8
TristateUseDefault = -2
TristateTrue = -1
TristateFalse = 0
sendUpLoad="0"
sendDownLoad="0"
sendDownloadSize="0"
sendUploadSize="0"

UpLoadTime="0"
DownloadTime="0"
serverIP="iperf3.speedtest.com.hk"



Set fso = CreateObject("Scripting.FileSystemObject")

Function deleteFiles()
	If (fso.FileExists("iperf.exe")) Then
		Set IperfFile = fso.GetFile("iperf.exe")
		IperfFile.Delete
	End If

	If (fso.FileExists("IperfSpeedTest.bat")) Then
		Set BatchFile = fso.GetFile("IperfSpeedTest.bat")
		BatchFile.Delete
	End If

	If (fso.FileExists("SpeedTest.vbs")) Then
		Set VbScriptFile = fso.GetFile("SpeedTest.vbs")
		VbScriptFile.Delete
	End If

	If (fso.FileExists("IperfResult.txt")) Then
		Set IperfResultFile = fso.GetFile("IperfResult.txt")
		IperfResultFile.Delete
	End If
End Function

Function getSpeedTestResult()
	Rem Read result

	If (fso.FileExists("error.txt")) Then
		Set ErrorFile = fso.GetFile("error.txt")
		If (ErrorFile.size > 0) Then
		rem	Set IperfResultFile = fso.GetFile("IperfResult.txt")
		rem	Set ts = IperfResultFile.OpenAsTextStream(ForReading, TristateUseDefault)
		rem	DataArray = split(ts.ReadLine, " ", -1, 1)
		rem	serverIP=DataArray(2)
			Set IE = wscript.CreateObject("internetexplorer.application")
			IE.Navigate "https://reg.hkbn.net/IperfTestLogin/savelog.do?downLoad="&sendDownLoad&"&upLoad="&sendUpLoad&"&id=E6C69360BF&errorCode=1004&serverIP=" & serverIP
			MsgBox "由於測試人數眾多，請稍後再試(1004)" & Chr(10) , vbOKOnly, "HKBN - Iperf速度測試"
			Exit Function
		End If
	End If

	If (fso.FileExists("IperfResult.txt")) Then
		Set IperfResultFile = fso.GetFile("IperfResult.txt")
		Set ts = IperfResultFile.OpenAsTextStream(ForReading, TristateUseDefault)

		i = 0
		On Error Resume Next
		If (ts.AtEndOfStream <> True) Then
			DataArray = split(ts.ReadLine, " ", -1, 1)
			UploadTime = DataArray(0)
			sendUploadSize= DataArray(1)
			UploadSize = DataArray(1) & " " & DataArray(2)
			Rem UploadSpeed = DataArray(2) & " " & DataArray(3)
			sendUpLoad=DataArray(3)
			UploadSpeed = DataArray(3) & " Mbps"
			i = i + 1
		End If

		If (ts.AtEndOfStream <> True) Then
			DataArray = split(ts.ReadLine, " ", -1, 1)
			DownloadTime = DataArray(0)
			DownloadSize = DataArray(1) & " " & DataArray(2)
			sendDownloadSize=DataArray(1)
			Rem DownloadSpeed = DataArray(3) & " " & DataArray(4)
			sendDownLoad=DataArray(3)
			DownloadSpeed = DataArray(3) & " Mbps"
			i = i + 1
		End If

		If Err.Number <> 0 Then 
			i= 0
		end if	
		
rem		If (ts.AtEndOfStream <> True) Then
rem			DataArray = split(ts.ReadLine, " ", -1, 1)
rem			DataArray = split(ts.ReadLine, " ", -1, 1)
rem			serverIP=DataArray(2)
rem		End If

		ts.Close

		Rem Show result
		If (i = 0) Then
			Set IE = wscript.CreateObject("internetexplorer.application")
			IE.Navigate "https://reg.hkbn.net/IperfTestLogin/savelog.do?downLoad="&sendDownLoad&"&upLoad="&sendUpLoad&"&id=E6C69360BF&errorCode=1002&serverIP=" & serverIP
			MsgBox "測試有誤，請再試一次 (1002)", vbOKOnly, "HKBN - Iperf速度測試"
		ElseIf (i = 1) Then
			Set IE = wscript.CreateObject("internetexplorer.application")
			IE.Navigate "https://reg.hkbn.net/IperfTestLogin/savelog.do?downLoad="&sendDownLoad&"&upLoad="&sendUpLoad&"&id=E6C69360BF&errorCode=1003&serverIP=" & serverIP
			MsgBox "測試有誤，請再試一次 (1003)", vbOKOnly, "HKBN - Iperf速度測試"
		ElseIf (i = 2) Then
		Set IE = wscript.CreateObject("internetexplorer.application")
			IE.Navigate "https://reg.hkbn.net/IperfTestLogin/savelog.do?downLoad="&sendDownLoad&"&upLoad="&sendUpLoad&"&id=E6C69360BF&errorCode=null&serverIP=" & serverIP
			rem MsgBox "你的速度測試結果" & Chr(10) & Chr(10) & "下載速度: " & DownloadSpeed & "(" & DownloadTime & "sec)" & Chr(10) & "上載速度: " & UploadSpeed & "(" & UploadTime & "sec)" & Chr(10) & "Download Size:" & DownloadSize & Chr(10) & "Upload Size:" & UploadSize, vbOKOnly, "HKBN - Iperf速度測試"			

			rem a=cdbl(DownloadSpeed)
			
			

			sendUpLoad=CDbl(sendUpLoad)*1024/8
			sendDownLoad=CDbl(sendDownLoad)*1024/8
			sendDownloadSize=CDbl(sendDownloadSize)*1024/8
			sendUploadSize=CDbl(sendUploadSize)*1024/8
			
			do
			Loop Until IE.ReadyState=4

			If (((sendDownLoad*8/1024) <= 100) And ((sendUpLoad*8/1024) <= 100)) Then
    				IE.Navigate "http://reg.hkbn.net/DosPatch_PingTest/output/bb100_speedtest.html?DOWNsize=" & sendDownloadSize & "&DOWNduration="& DownloadTime &"&DOWNspeed=" & sendDownLoad & "&UPsize="&sendUploadSize&"&UPduration=" & UploadTime & "&UPspeed=" & sendUpLoad
   			Else 
    				IE.Navigate "http://reg.hkbn.net/DosPatch_PingTest/output/bb1000_speedtest.html?DOWNsize=" & sendDownloadSize & "&DOWNduration="& DownloadTime &"&DOWNspeed=" & sendDownLoad & "&UPsize="&sendUploadSize&"&UPduration=" & UploadTime & "&UPspeed=" & sendUpLoad
  			End If
		    IE.Visible = True
			
		End If

		Rem Set IE = wscript.CreateObject("internetexplorer.application")
		Rem IE.Visible = True
		Rem IE.Navigate "http://www.hkbn.net"
	Else
		MsgBox "測試有誤，請再試一次 (1001)", vbOKOnly, "HKBN - Iperf速度測試"
	End If
End Function

Call getSpeedTestResult()
Rem Call deleteFiles()