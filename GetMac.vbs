'''''''''''''''''''''''''''''''''''''''''''''
'                                           '
'       Script Created by : basox70         '
'        First Release : 2016/09/29         '
'        Last Release : 2016/11/14          '
'   Script Name : arp.vbs   Version : 1.9   '
'                                           '
'''''''''''''''''''''''''''''''''''''''''''''

Rem INFOS :
Rem chr(34) = "
Rem chr(38) = &
Rem wscript.exe //H:cscript
Rem wscript.exe //H:wscript

''''''''''''''''''''TODO'''''''''''''''''''''
'                                           '
' - ajouter le nom des peripheriques        '
' - SI il y a d'autres idées, les ajouter   '
'                                           '
'''''''''''''''''''''''''''''''''''''''''''''

Rem INITIALIZATION / INITIALISATION

Set wShell = WScript.CreateObject("WSCript.shell")

' Get script directory
' Recupere le chemin d'execution du script
Set fso = CreateObject("Scripting.FileSystemObject")
curDir = fso.GetAbsolutePathName(".")
strPath = WScript.ScriptFullName
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(strPath)
strFolder = objFSO.GetParentFolderName(objFile) 

Dim Ip2Mac
Set Ip2Mac = CreateObject("Scripting.Dictionary")

' define address (format arr1[x].arr2[x].0.x)
arr1=Array(59,99)
arr2=Array(0,1,7,8,10,12,15)
peripheralNb = 0
max = 255
MAC = false

FileContentStr = ""
FileContentArr = ""

' File1 = strFolder & "\vbstmp.txt" 'Temp File read to complete dictionary
' File2 = strFolder & "\arp.txt" 'Original File to read
' File3 = strFolder & "\arpTrie.txt" 'Final File

File1 = "vbstmp.txt" 'Temp File read to complete dictionary
File2 = "arp.txt" 'Original File to read
File3 = "arpTrie.txt" 'Final File

Set oWSH = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"

Call ForceConsole()

Function ForceConsole()
	If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
		oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
		WScript.Quit
	End If
End Function

'printw File1
'printw File2
'printw File3

Function wait(n)
	WScript.Sleep Int(n * 1000)
End Function

Function printl(txt)
	WScript.StdOut.Write txt
End Function

Function printw(txt)
	WScript.StdOut.WriteLine txt
End Function

If WScript.Arguments.Count = 0 then
    WScript.Echo "Script defaut"
Else
	For I=0 To WScript.Arguments.Count-1
		Select Case WScript.Arguments(I)
		Case "1"
			MAC = true
		Case "/h"
			help()
			Wscript.Quit
		Case Else
			printw "Parametre non reconnu"
			Wscript.Quit
		end Select
	Next
End If

Function help()
	printw "parametres:"
	printw "    - Renouveler adresses MAC : 1 (defaut = 0)"
	printw "        exemple 'nom du script' 1"
	printw "    - afficher cette aide : /h"
End Function

' Afficher les parametres du script
' Display script parameters
printw "renouvellement adresses MAC : " & MAC

' Add arp request line into dictionary, without "static" / "dynamic" or "new" , key = ip, item = mac address
' Ajoute la requete arp dans le dictionnaire, sans "statique"/"dynamique"/"new", cle = ip, objet = Mac
Function arp2dict( ByRef line, newIp) 
	line = Replace(line,"dynamique","")
	line = Replace(line,"statique","")
	line = Replace(line,"dynamic","")
	line = Replace(line,"static","")
	line = Replace(line,"new","")
	line = Replace(line," ","")
	tmpIp = Left(line, (Len(line)-17))
	tmpIp = Replace(tmpIp," ","")
	tmpMac = Right(line, 17)
	tmpMac = Replace(tmpMac," ","")
	If (newIp and Not MAC)then 
		tmpMac = tmpMac&"  new"
	End If
	line = tmpIp&"|"&tmpMac
	tmpArr = Split(line , "|")
	If Not Ip2Mac.Exists(tmpIp) Then
		Ip2Mac.add tmpArr(0),tmpArr(1)
	End If	
End Function

' File in parameter is read & put in array "FileContentArr"
' Le fichier en param est lu et stocké dans un tableau "FileContentArr"
Function FileReader(ByRef file) 
	Dim filesys, readfile, contents
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set readfile = filesys.OpenTextFile(file, 1, False)
	FileContentStr = ""
	Do While readfile.AtEndOfStream=False 
		contents = readfile.ReadLine
		If Len(contents)>15 Then
			FileContentStr = FileContentStr&contents&"||" 
		End If
	Loop 
	readfile.Close
	If Len(FileContentStr) > 2 Then 
		FilecontentStr = Left(FilecontentStr, (Len(FilecontentStr)-2))
	End If
	FileContentArr = Split(FileContentStr , "||")
	FileReader = FileContentArr
End Function

' Write in a file without new line
' Ecrire dans un fichier, sans retour à la ligne
Function FileWriter(file, data) 
	Dim filesys, writefile
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set writefile = filesys.OpenTextFile(file, 8,True)
	writefile.Write(data)
End Function

' Format string var (e.g : Format("tmp", 5, "?") => "tmp??")
' Formatter une var de type string ( exemple : Format("tmp", 5, "?") => "tmp??" )
Function Format(Str, lgh, char)
	Y = Len(Str)
	Format = Str
	For i=Y To lgh-1
		Format = Format&char
	Next
End Function

' Sort dictionnary function
' Fonction de tri d'un dictionnaire
Function SortDictionary(objDict)
	Dim strDict()
	Dim X, Y, Z
	Z = objDict.Count
	If Z > 1 Then
		ReDim strDict(Z,2)
		Y=0
		For Each X In objDict.Keys()
			' printw X&" : ["&objDict.Item(X)&"]"
			strDict(Y,0) = X
			strDict(Y,1) = objDict.Item(X)
			e=3
			s=0
			For Each n In Split(X,".")
				s=s+n*256^e
				e=e-1
			Next
			strDict(Y,2) = s
			Y = Y+1
		Next
		For X=0 To Z-2
			For Y=X To Z-1
				If strDict(X,2) > strDict(Y,2) Then
					strKey  = strDict(X,0)
					strItem = strDict(X,1)
					strValue = strDict(X,2)
					strDict(X,0)  = strDict(Y,0)
					strDict(X,1) = strDict(Y,1)
					strDict(X,2) = strDict(Y,2)
					strDict(Y,0)  = strKey
					strDict(Y,1) = strItem
					strDict(Y,2) = strValue
					' printw "permut: "&X&"|"&strDict(X,0)&"|"&strDict(X,1)&"|"&strDict(X,2)&" with "&Y&"|"&strDict(Y,0)&"|"&strDict(Y,1)&"|"&strDict(Y,2)
				End If
			Next
		Next
	End If
	
	objDict.RemoveAll
	
	For X=0 To UBound(strDict)-1
		objDict.add strDict(X,0),strDict(X,1)
	Next
	
End Function

printw "Le script dure environ 10 min."
' printw "Fin estimee vers "&DateAdd("n",15,FormatDateTime(Now))& "."

Rem BEGINNING / DEBUT TRAITEMENT

Rem EXTRACT INFOS FROM FILE / PARCOURS DU FICHIER

' If file File3 exists, take ip & Mac from it, else take from File2 (if it exists)
' Si le fichier File3 existe, prend l'ip et la Mac de ce fichier, sinon prend l'ip et la Mac à partir du fichier File2 (s'il existe)
taken = True
usedFile = File2
If fso.FileExists(File3) Then
	FileReader(File3)
	arr3 = UBound(FileContentArr)
	If arr3>100 Then
		For Each content In FileContentArr
			If Len(Content)>30 Then
				arp2dict content, False
			End If
		Next
		usedFile = File3
		taken = False
	End If
End If
If (fso.FileExists(File2) And taken) Then
	FileReader(File2)
	For Each content In FileContentArr
		arp2dict content, False
	Next
End If

REM BEGIN PING/ARP REQUEST / DEBUT REQUETES PING/ARP

nb = ((UBound(arr1)+1)/2)*(UBound(arr2)+1)*25+((UBound(arr1)+1)/2)*(UBound(arr2)+1)*max+((UBound(arr1)+1)/2)*255 'nombre de boucle au total
nbTotal = 0

For Each i In arr1
	If i = 59 Then
		' printw i&".0.0.0"
		For Each j In arr2
			For k=0 To 1
				' printw i&"."&j&"."&k&".0"
				If k = 1 Then
					max = 25
				Else
					max = 255
				End If
				For l=1 To max
					nbTotal = nbTotal + 1
					If (nbTotal Mod nb\100) = 0 Then
						printw FormatPercent(nbTotal/nb,0)
					End If
					ip = i&"."&j&"."&k&"."&l
					If (Not Ip2Mac.Exists(ip) Or MAC) Then
						result = wShell.run("cmd /K (ping -n 1 -w 50 "&ip&" || exit /B 0 ) "&Chr(38)&Chr(38)&" arp -a "&ip&" > " & File1 & " "&Chr(38)&" exit",7,True) '(EN) https://msdn.microsoft.com/en-us/library/d5fk67ky(v=vs.84).aspx || (FR) http://jc.bellamy.free.fr/fr/vbsobj/wsmthrun.html
						' printw "(ping -n 1 -w 100 "&ip&" || exit /B 0 ) "&Chr(38)&Chr(38)&" arp -a "&ip&" > " & File1 & " "&Chr(38)&" exit"
						If fso.FileExists(File1) Then
							FileArr = FileReader(File1)
							For Each fileStr In FileArr
								If InStr(fileStr,"  "&ip)>0 And InStr(fileStr,"Interface")<1  Then
									printw "ip: "&i&"."&j&"."&k&"."&l '&vbCrLf&fileStr
									arp2dict fileStr, True
									peripheralNb = peripheralNb+1
								End If
							Next
						End If
					End If
				Next
			Next
		Next
	Else
		j=0
		k=0
		' printw i&"."&j&"."&k&".0"
		For l=1 To 255
			ip = i&"."&j&"."&k&"."&l
			nbTotal = nbTotal + 1
			If (nbTotal Mod nb\100) = 0 Then
				printw FormatPercent(nbTotal/nb,0)
			End If
			If (Not Ip2Mac.Exists(ip) Or MAC) Then
				result = wShell.run("cmd /K (ping -n 1 -w 50 "&ip&" || exit /B 0 ) "&Chr(38)&Chr(38)&" arp -a "&ip&" > " & File1 & " "&Chr(38)&" exit",7,True) '(EN) https://msdn.microsoft.com/en-us/library/d5fk67ky(v=vs.84).aspx || (FR) http://jc.bellamy.free.fr/fr/vbsobj/wsmthrun.html
				FileReader(File1)
				If fso.FileExists(File1) Then
					FileArr = FileReader(File1)
					For Each fileStr In FileArr
						If InStr(fileStr,"  "&ip)>0 And InStr(fileStr,"Interface")<1  Then
							printw "ip: "&i&"."&j&"."&k&"."&l
							arp2dict fileStr, True
							peripheralNb = peripheralNb+1
						End If
					Next
				End If
			End If
		Next
	End If
Next

printw FormatPercent(nbTotal/nb,0)&" de "&nb

printl peripheralNb&" nouveaux peripheriques trouves base sur :"&vbCrLf
printw usedFile

If fso.FileExists(File1) Then 'supprime le fichier vbstmp
	fso.deleteFile File1
End If

If fso.FileExists(File3) Then 'supprime le fichier arpTrie.txt pour éviter d'avoir un fichier de 600 lignes au bout de 3 lancements
	fso.deleteFile File3
End If

' TRI DU DICTIONNAIRE ET ECRITURE DANS LE FICHIER
' SORT DICTIONNARY & WRITE INTO FILE
SortDictionary Ip2Mac

If Not fso.FileExists(File3) Then
	fso.CreateTextFile(File3)
End If

result = wShell.run("cmd /K Date /t > "&File3&" "&Chr(38)&" exit",7,True)

For Each elem In Ip2Mac 'formatage de la ligne puis ecriture dans le fichier
	' printw "!"&Format(elem,11," ")&"!"
	str = "  "&Format(elem,11," ")&"        "&Ip2Mac(elem)&vbCrLf
	' printl "!"&str&"!"
	FileWriter File3,str
Next

result = wShell.run("cmd /K "&File3&" "&Chr(38)&" exit",7,True) 'affiche le fichier final
