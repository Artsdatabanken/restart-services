Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set colObjects = objWMI.ExecQuery("Select * From Win32_Process where name='raven.server.exe'")

Dim PName,objShell

For Each Item in colObjects
    Set PName = Item
Next

ProcessName =  PName.Name
' WorkingSetSize/1048576, konverterer fra bytes til megabytes
ProcessMemory = PName.WorkingSetSize/1048576

'wscript.echo processmemory
'restarter raven og tilhørende services om "working set size" overskrider 10 gigabyte
If ProcessMemory >= 10000 Then
         Set objShell = CreateObject("WScript.Shell")
 
		 objShell.Run "net stop ""Artportalen.Services.AuditService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net stop ""Artportalen.Services.ConvertService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net stop ""Artportalen.Services.ConvertViewService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net stop ""Artportalen.Services.ImportService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net stop ""Artportalen.Services.ImportViewService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net stop ""Artportalen.Services.ReportService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net stop ""Artportalen.Services.ReportTopListService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net stop ""Artportalen.Services.ReportViewService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
         objShell.Run "net stop RavenDB"
         WScript.Sleep 30000
         objShell.Run "net start RavenDB"
		 WScript.Sleep 30000
		 objShell.Run "net start ""Artportalen.Services.AuditService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net start ""Artportalen.Services.ConvertService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net start ""Artportalen.Services.ConvertViewService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net start ""Artportalen.Services.ImportService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net start ""Artportalen.Services.ImportViewService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net start ""Artportalen.Services.ReportService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net start ""Artportalen.Services.ReportTopListService-1.0.0 (Instance: nbn)"""
		 WScript.Sleep 5000
		 objShell.Run "net start ""Artportalen.Services.ReportViewService-1.0.0 (Instance: nbn)"""
		 
				'Skriver timestamp til logfil, når restart er foretatt
		 Const ForAppending = 8
		 Set objFSO = CreateObject("Scripting.FileSystemObject")
		 Set objFile = objFSO.OpenTextFile("D:\logs\restart_raven\restart_raven.txt", ForAppending)
		 objFile.Write Now
		 objFile.WriteLine
		 objFile.Close
		 
         
End If
