@ECHO OFF
taskkill /f /im "DATAMART.exe"
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /out:DATAMART.exe DATAMART.cs
pause
