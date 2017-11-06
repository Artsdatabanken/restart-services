Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set colObjects = objWMI.ExecQuery("Select * From Win32_Process where name='ravendb.exe'")

Dim PName,objShell

For Each Item in colObjects
    Set PName = Item
Next

ProcessName =  PName.Name
ProcessMemory = PName.WorkingSetSize/1048576

wscript.echo processmemory