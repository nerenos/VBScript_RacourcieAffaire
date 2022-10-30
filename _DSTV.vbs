'////////////////////////////////////////////////////////////
stRepProjets = "X:"
'////////////////////////////////////////////////////////////

Set Fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("Wscript.Shell")

repFic = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
TabRep = Split(repFic, "\")

CodeChantier = Left(TabRep(UBound(TabRep)-2),5)

    bChangeChemin = True
    bCheminExiste = False
    SelectFolder = ""
    CheminFind = SearchFolder(stRepProjets, CodeChantier)
	If Fso.FolderExists(CheminFind) Then
		WshShell.run("Explorer" &" " & CheminFind & "\")
	Else
		Fso.CreateFolder(stRepProjets & "\" & TabRep(UBound(TabRep)-2))
		WshShell.run("Explorer" &" " & stRepProjets & "\" & TabRep(UBound(TabRep)-2) & "\")
	End If

Set Fso = Nothing
Set WshShell = Nothing

		'-----------------------------------------------Fonction Cherche Chemin Affaire Auto-------------------------------------------------------
Function SearchFolder(Dir, CA)
	SearchFolder = ""
	If Fso.FolderExists(Dir) Then
		For each SoFld in  Fso.GetFolder(Dir).SubFolders
			If Left(SoFld.Name,5) = CA Then
				SearchFolder = SoFld
				Exit Function
			End If
		Next
	End If
End Function