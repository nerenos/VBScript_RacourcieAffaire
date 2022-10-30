'////////////////////////////////////////////////////////////
stRepProjets = "Y:\AFFAIRES\CHANTIERS\"
'////////////////////////////////////////////////////////////

Set Fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("Wscript.Shell")

repFic = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
TabRep = Split(repFic, "\")

CodeChantier = Left(TabRep(UBound(TabRep)-1),5)
stRepProjets = stRepProjets & "20"& Left(CodeChantier,2)

    bChangeChemin = True
    bCheminExiste = False
    SelectFolder = ""
    CheminFind = SearchFolder(stRepProjets, CodeChantier)
	If Fso.FolderExists(CheminFind) Then
		WshShell.run("Explorer" &" " & CheminFind & "\")
	Else
		MsgBox("Aucune Affaire Tekla Trouvé")
	End If

Set Fso = Nothing
Set WshShell = Nothing

		'-----------------------------------------------Fonction Cherche Chemin Affaire Auto-------------------------------------------------------
Function SearchFolder(Dir, CA)
	SearchFolder = ""
	If Fso.FolderExists(Dir) Then
		For Each SoFld In  Fso.GetFolder(Dir).SubFolders
			If Left(Right(SoFld.Name, 12),5) = CA Then
				SearchFolder = SoFld
				Exit Function
			End If
		Next
	End If
End Function