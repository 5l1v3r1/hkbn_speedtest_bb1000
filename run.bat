@echo off
cls

echo.
echo ---------------------------------------------------------------
echo Iperf Speed Test by Hong Kong Broadband Network Ltd.
echo (for Windows only)
echo.
echo For an accurate result,
echo please DON'T perform any tasks when the Iperf test is running.
echo ---------------------------------------------------------------
echo.
echo Test started, please wait...
NSLOOKUP iperf3.speedtest.com.hk 1>DNS.txt 2>dns_error.txt 
FOR /F %%i IN ('type DNS.txt ^| gawk -f "DNS.awk"' ) DO (SET SERVER=%%i)
ipconfig > ip.txt
ver > ver.txt
wget -q -O wanip.txt http://www.speedtest.com.hk/testIP.jsp
for /F %%i in ('type wanip.txt ^| findstr "[0-9]"' ) DO (SET EIP=%%i)
findstr %EIP% ip.txt > nul
if %errorlevel%==0 goto Normal
for /F %%t in ('type ver.txt ^| find /I "Windows" ^| gawk -F"[" "{print $2}" ^| gawk "{print $2}" ^| gawk -F"." "{print($1*100+$2*1)}"' ) DO (SET VER=%%t)
if %VER% LSS 501 goto error
upnpc-shared -r 5001 TCP 2>>upnp_error.txt 1>>upnp.txt
findstr /m "No IGD UPnP Device found" upnp_error.txt > nul
if %errorlevel%==0 goto error
start /B iperf -s -P 40 -w 1M 2> download_error.txt 1> download.txt
iperf -c %SERVER% -t 10 -P 35 -f m -w 8k 2> error.txt 1> upload.txt
@ping 127.0.0.1 -n 21 -w 1000 > nul
taskkill /F /IM "iperf.exe" > nul
upnpc-shared -d 5001 TCP 2>>upnp_error.txt 1>>upnp.txt
goto report
:Normal
start /B iperf -s -P 40 -w 1M 2> download_error.txt 1> download.txt
iperf -c %SERVER% -t 10 -P 35 -f m -w 8k 2> error.txt 1> upload.txt
@ping 127.0.0.1 -n 21 -w 1000 > nul
taskkill /F /IM "iperf.exe" > nul
goto report
:report
type download_error.txt >> error.txt
type upload.txt | find /I "[SUM]" | gawk "{ print $2, $4, $5, $6, $7 }" | gawk -F"-" "{print $2}" > IperfResult.txt
type download.txt | find /I "[SUM]" | gawk "{ print $2, $4, $5, $6, $7 }" | gawk -F"-" "{print $2}" >> IperfResult.txt
type upload.txt | find "connecting to" | gawk "{ print $2, $3, $4}" | gawk -F"," "{print $1}" >> IperfResult.txt
echo Test finished
cscript //Nologo SpeedTest.vbs
goto END
:error
cscript //Nologo error.vbs
:END
