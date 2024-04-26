Option Explicit
Dim objFSO, stopper

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set stopper = objFSO.OpenTextFile("Composants/stop.txt", 2)
stopper.Write ("0")
stopper.close