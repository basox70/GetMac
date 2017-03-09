'''''''''''''''''''''''''''''''''''''''''''''
'                                           '
'       Script Created by : basox70         '
'        First Release : 2016/09/29         '
'        Last Release : 2017/03/08          '
'        Script Name : GetMac.vbs           '
'             Version : 1.24                '
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

debugLoop = false
debugHelp = false

Rem INITIALIZATION / INITIALISATION

Set wShell = WScript.CreateObject("WSCript.shell")

' Get script directory
' Recupere le chemin d'execution du script
Set fso = CreateObject("Scripting.FileSystemObject")

Dim Ip2Mac
Set Ip2Mac = CreateObject("Scripting.Dictionary")

If debugHelp Then
    If debugLoop Then
        printw "Mode debug : ALL"
    Else
        printw "Mode debug : Help"
    End If
Else
    If debugLoop Then
        printw "Mode debug : Loop"
    End If
End If

redim arr2(1), arr3(1), arr4(1)
' define address (format arr1[x].arr2[x].arr3[x].arr4[x])

arr1 = Array(59, 99)
arr2(0) = Array(1, 1)
arr2(1) = Array(0, 1, 7, 8, 10, 12, 15, -1, 0)
arr3(0) = Array(1, 7)
arr3(1) = Array(0, 1, -1, 0)
arr4(0) = Array(9)
arr4(1) = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178, 179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193, 194, 195, 196, 197, 198, 199, 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225, 226, 227, 228, 229, 230, 231, 232, 233, 234, 235, 236, 237, 238, 239, 240, 241, 242, 243, 244, 245, 246, 247, 248, 249, 250, 251, 252, 253, 254, 255)

peripheralNb = 0
MAC = false

FileContentStr = ""
FileContentArr = ""

File1 = "vbstmp.txt" 'Temp File read to complete dictionary
File2 = "arp.txt" 'Original File to read
File3 = "arpTrie.txt" 'Final File

Function wait(n)
    WScript.Sleep Int(n * 1000)
End Function

Function printl(txt)
    WScript.StdOut.Write txt
End Function

Function printw(txt)
    WScript.StdOut.WriteLine txt
End Function

If WScript.Arguments.Count = 0 Then
    WScript.Echo "Script defaut"
Else
    For I = 0 To WScript.Arguments.Count - 1
        Select Case WScript.Arguments(I)
        Case "/m"
            MAC = true
        Case "/h", "/?", "/help"
            help()
        Case Else
            printw "Parametre non reconnu" & vbCrLf
            help()
        End Select
    Next
End If

Function help()
    printw "Parametres:"
    printw "    - Renouveler adresses MAC : /m (defaut = 0)"
    printw "        exemple 'nom du script' /m"
    printw "    - afficher cette aide : /help, /h, /?"
    Wscript.Quit
End Function

' Afficher les parametres du script
' Display script parameters
printw "renouvellement adresses MAC : " & MAC

' Remove items from array
' retire des objets d'un array
Function Ritems( arr )
    Ritems = join(arr, "|")
    Do While LEFT(Ritems, 2) <> "-1"
        Ritems = Right(Ritems, len(Ritems)-1)
    Loop
    Ritems = Right(Ritems, len(Ritems)-3)
    If len(Ritems) < 4 Then
        Ritems = Ritems & "|-1"
    End If
    Ritems = split(Ritems, "|")
End Function

' Add arp request line into dictionary, without "static" / "dynamic" Or "new" , key = ip, item = mac address
' Ajoute la requete arp dans le dictionnaire, sans "statique"/"dynamique"/"new", cle = ip, objet = Mac
Function arp2dict( ByRef line, newIp)
    line = Replace(line, "dynamique", "")
    line = Replace(line, "statique", "")
    line = Replace(line, "dynamic", "")
    line = Replace(line, "static", "")
    line = Replace(line, "new", "")
    line = Replace(line, " ", "")
    tmpIp = Left(line, (Len(line) - 17))
    tmpMac = Right(line, 17)
    If (newIp And Not MAC) Then
        tmpMac = tmpMac & "  new"
    End If
    line = tmpIp & "|" & tmpMac
    tmpArr = Split(line , "|")
    If Not Ip2Mac.Exists(tmpIp) Then
        Ip2Mac.add tmpIp, tmpMac
    End If
End Function

' File in parameter is read & put in array "FileContentArr"
' Le fichier en param est lu et stocké dans un tableau "FileContentArr"
Function FileReader(ByRef file)
    Dim filesys, readfile, contents
    Set filesys = CreateObject("Scripting.FileSystemObject")
    Set readfile = filesys.OpenTextFile(file, 1, False)
    FileContentStr = ""
    Do While readfile.AtEndOfStream = False
        contents = readfile.ReadLine
        If Len(contents) > 15 Then
            FileContentStr = FileContentStr & contents & "||"
        End If
    Loop
    readfile.Close
    If Len(FileContentStr) > 2 Then
        FilecontentStr = Left(FilecontentStr, (Len(FilecontentStr) - 2))
    End If
    FileContentArr = Split(FileContentStr , "||")
    FileReader = FileContentArr
End Function

' Write in a file without new line
' Ecrire dans un fichier, sans retour à la ligne
Function FileWriter(file, data)
    Dim filesys, writefile
    Set filesys = CreateObject("Scripting.FileSystemObject")
    Set writefile = filesys.OpenTextFile(file, 8, True)
    writefile.Write(data)
End Function

' Format string var (e.g : Format("tmp", 5, "?") => "tmp??")
' Formatter une var de type string ( exemple : Format("tmp", 5, "?") => "tmp??" )
Function Format(Str, lgh, char)
    Y = Len(Str)
    Format = Str
    For i = Y To lgh - 1
        Format = Format & char
    Next
End Function

' Sort dictionnary function
' Fonction de tri d'un dictionnaire
Function SortDictionary(objDict)
    Dim strDict()
    Dim X, Y, Z
    Z = objDict.Count
    If Z > 1 Then
        ReDim strDict(Z, 2)
        Y = 0
        For Each X In objDict.Keys()
            strDict(Y, 0) = X
            strDict(Y, 1) = objDict.Item(X)
            e = 3
            s = 0
            For Each n In Split(X, ".")
                s = s + n * 256^e
                e = e - 1
            Next
            strDict(Y, 2) = s
            Y = Y + 1
        Next
        For X = 0 To Z - 2
            For Y = X To Z - 1
                If strDict(X, 2) > strDict(Y, 2) Then
                    strKey  = strDict(X, 0)
                    strItem = strDict(X, 1)
                    strValue = strDict(X, 2)
                    strDict(X, 0) = strDict(Y, 0)
                    strDict(X, 1) = strDict(Y, 1)
                    strDict(X, 2) = strDict(Y, 2)
                    strDict(Y, 0) = strKey
                    strDict(Y, 1) = strItem
                    strDict(Y, 2) = strValue
                End If
            Next
        Next
    End If

    objDict.RemoveAll

    For X = 0 To UBound(strDict) - 1
        objDict.add strDict(X, 0), strDict(X, 1)
    Next

End Function

printw "Le script dure environ 15 min."

Rem BEGINNING / DEBUT TRAITEMENT

Rem EXTRACT INFOS FROM FILE / PARCOURS DU FICHIER

' If file File3 exists, take ip & Mac from it, else take from File2 (if it exists)
' Si le fichier File3 existe, prend l'ip et la Mac de ce fichier, sinon prend l'ip et la Mac à partir du fichier File2 (s'il existe)
taken = True
usedFile = File2
If fso.FileExists(File3) Then
    FileReader(File3)
    arr6 = UBound(FileContentArr)
    If arr6 > 100 Then
        For Each content In FileContentArr
            If Len(Content) > 30 Then
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

If Not debugLoop Then
    wShell.run "cmd /C Date /t > " & File3 ,7,True
    wShell.run "cmd /C time /t >> " & File3 ,7,True
End If

nb = 0

If debugHelp Then
    printw "arr1 : " & UBound(arr1) + 1
    printw "arr2(0) : " & UBound(arr2(0)) + 1
    printw "arr2(1) : " & UBound(arr2(1)) + 1
    printw "arr3(0) : " & UBound(arr3(0)) + 1
    printw "arr3(1) : " & UBound(arr3(1)) + 1
    printw "arr4(0) : " & UBound(arr4(0)) + 1
    printw "arr4(1) : " & UBound(arr4(1)) + 1
End If

nbTotal = arr4(0)(0) * (Ubound(arr4(1)) + 1) - 100 Or 100 'nombre de boucle au total
printw nbTotal

i = 0
j = 0
k = 0
l = 0
If debugLoop Then
    printw "-----1-----"
    printw "i:" & i
    printw "j:" & j &"|jj:" & jj & "|jjj:" & jjj
    printw "k:" & k &"|kk:" & kk & "|kkk:" & kkk
    printw "l:" & l &"|ll:" & ll & "|lll:" & lll
    printw "-----------"
End If

For i=Lbound(arr1) To Ubound(arr1)
    If Ubound(arr2(0)) > 0 and arr2(0)(0) < 1 Then
        arr2(1) = Ritems(arr2(1))
        tmp = join(arr2(0), "|")
        Do While LEFT(tmp, 1) <> "|"
            tmp = Right(tmp, len(tmp)-1)
        Loop
        tmp = Right(tmp, len(tmp)-1)
        arr2(0) = split(tmp, "|")
        tmp = null
    End If

    If arr2(0)(0) > 0 Then
        arr2(0)(0) = arr2(0)(0) - 1
        For j = Lbound(arr2(1)) To Ubound(arr2(1))
            If CInt(arr2(1)(j)) = -1 Then
                Exit For
            End If

            If Ubound(arr3(0)) > 0 and arr3(0)(0) < 1 Then
                arr3(1) = Ritems(arr3(1))
                tmp = join(arr3(0), "|")
                Do While LEFT(tmp, 1) <> "|"
                    tmp = Right(tmp, len(tmp)-1)
                Loop
                tmp = Right(tmp, len(tmp)-1)
                arr3(0) = split(tmp, "|")
                tmp = null
            End If
            If arr3(0)(0) > 0 Then
                arr3(0)(0) = arr3(0)(0) - 1
                For k = Lbound(arr3(1)) To Ubound(arr3(1))
                    If CInt(arr3(1)(k)) <> -1 Then
                        If arr4(0)(0) > 0 Then
                            If Ubound(arr3(0)) <> arr30 Then
                                arr4(0)(0) = arr4(0)(0) - 1
                                arr30 = Ubound(arr3(0))
                            End If
                            For l = Lbound(arr4(1)) To Ubound(arr4(1))
                                If CInt(arr4(1)(l)) <> -1 Then
                                    nb = nb + 1
                                    If debugLoop Then
                                        printw " 1.1: " & i & "/" & UBound(arr1) & " | " & j & "/" & UBound(arr2(1)) & " | " & k & "/" & UBound(arr3(1)) & " | " & l & "/" & UBound(arr4(1)) & chr(9) & " || " & arr1(i) & "." & arr2(1)(j) & "." & arr3(1)(k) & "." & arr4(1)(l) & chr(9) & " ||"
                                    Else
                                        ip = arr1(i) & "." & arr2(1)(j) & "." & arr3(1)(k) & "." & arr4(1)(l)
                                        If (nb Mod nbTotal\100) = 0 Then
                                           printw FormatPercent(nb/nbTotal, 0)
                                        End If
                                        If Not Ip2Mac.Exists(ip) Or MAC Then
                                           wShell.run "cmd /C (ping -n 1 -w 50 " & ip & " || exit /B 0 ) " & Chr(38) & Chr(38) & " arp -a " & ip & " > " & File1 , 7, True  '(EN) https://msdn.microsoft.com/en - us/library/d5fk67ky(v = vs.84).aspx || (FR) http://jc.bellamy.free.fr/fr/vbsobj/wsmthrun.html
                                           ' printw "(ping -n 1 -w 100 " & ip & " || exit /B 0 ) " & Chr(38) & Chr(38) & " arp -a " & ip & " > " & File1 & " " & Chr(38) & " exit"
                                           If fso.FileExists(File1) Then
                                               FileArr = FileReader(File1)
                                               For Each fileStr In FileArr
                                                   If InStr(fileStr, "  " & ip) > 0 And InStr(fileStr, "Interface") < 1  Then
                                                       printw "ip: " & ip '&vbCrLf & fileStr
                                                       arp2dict fileStr, True
                                                       peripheralNb = peripheralNb + 1
                                                   End If
                                               Next
                                           End If
                                        End If
                                    End If
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                        If Ubound(arr4(0)) > 0 and arr4(0)(0) < 1 Then
                            arr4(1) = Ritems(arr4(1))
                            tmp = join(arr4(0), "|")
                            Do While LEFT(tmp, 1) <> "|"
                                tmp = Right(tmp, len(tmp)-1)
                            Loop
                            tmp = Right(tmp, len(tmp)-1)
                            arr4(0) = split(tmp, "|")
                            tmp = null
                        End If
                    Else
                        Exit For
                    End If
                Next
            End If
        Next
    End If
Next

If debugHelp Then printw nb & " / " & nbTotal : a = 0

If debugLoop Then
    WScript.Quit
End If

printw FormatPercent(nb/nbTotal, 0) & " de " & nbTotal
printl peripheralNb & " nouveaux peripheriques trouves base sur : "
printw usedFile

If fso.FileExists(File1) Then 'supprime le fichier vbstmp
    fso.deleteFile File1
End If

' TRI DU DICTIONNAIRE ET ECRITURE DANS LE FICHIER
' SORT DICTIONNARY & WRITE INTO FILE
SortDictionary Ip2Mac

If Not fso.FileExists(File3) Then
    fso.CreateTextFile(File3)
End If

For Each elem In Ip2Mac 'formatage de la ligne puis ecriture dans le fichier
    ' printw "!" & Format(elem, 11, " ") & "!"
    str = "  " & Format(elem, 11, " ") & "        " & Ip2Mac(elem) & vbCrLf
    ' printl "!" & str & "!"
    FileWriter File3, str
Next
wShell.run "cmd /C time /t >> "&File3, 7, True

wShell.run "cmd /C " & File3 , 7, True 'affiche le fichier final
