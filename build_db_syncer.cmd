C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe /r:"lib\System.Data.SQLite.dll"  /t:library  /out:lib\database_synchronize_resources.dll src\ConnectionCipher.cs src\ConnectionProperty.cs src\TableSynchronizer.cs
C:\Windows\Microsoft.NET\Framework64\v3.5\csc.exe /r:"lib\System.Data.SQLite.dll","lib\database_synchronize_resources.dll" /out:DatabaseSynchronizer.exe src\DatabaseSynchronizer.cs
pause