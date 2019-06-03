C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe /r:"lib\System.Data.SQLite.dll"  /t:library  /out:lib\synchronize_resources.dll src\ConnectionCipher.cs src\ConnectionProperty.cs
C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe /r:"lib\System.Data.SQLite.dll","lib\synchronize_resources.dll" /out:TableSynchronizer.exe src\TableSynchronizer.cs
pause