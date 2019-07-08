Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set colObjects = objWMI.ExecQuery("Select * From Win32_Process where name='raven.server.exe'")

Dim PName,objShell

Dim dagenIDag, klokkeTime, skipping
Dim oShell
Set oShell = CreateObject("WScript.Shell")
skipping = "0"

'Finner hele klokketimer
klokkeTime = Hour(Now())
'Finner dag og setter mandag som ukestart
dagenIDag = int(Weekday(Now(), 2))

'Sjekker om det er onsdag og om klokken er mellom 9.00 og 12.59, eller alle andre dager mellom 11.00 og 11.59. Skipper periodene hvor det eksporteres med smuggler fra Raven
If ((klokkeTime >= 9 And klokkeTime <= 12) And (dagenIDag = 3)) Then
	skipping = "1"
End If
If klokkeTime = 11 Then
	skipping = "1"
End If

ProcessName = PName.Name
' WorkingSetSize/1048576, konverterer fra bytes til megabytes
ProcessMemory = PName.WorkingSetSize/1048576

If skipping = "0" Then
'wscript.echo processmemory
'restarter raven og tilhørende services om "working set size" overskrider 10 gigabyte
	If ProcessMemory >= 13000 Then
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
End If
