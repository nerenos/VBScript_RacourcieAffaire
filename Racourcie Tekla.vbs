'////////////////////////////////////////////////////////////
stRepProjets = "Z:\TeklaStructuresModels"
stRepChantiers = "Z:\TeklaStructures_serrurerie"
'////////////////////////////////////////////////////////////

Set Fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("Wscript.Shell")

repFic = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
CodeChantier = Left(Right(repFic, 13),5)

    bChangeChemin = True
    bCheminExiste = False
    SelectFolder = ""
    CheminFind = SearchFolder(stRepProjets, CodeChantier)
    If CheminFind = "" Then CheminFind = SearchFolder(stRepChantiers, CodeChantier) End If
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
		For each SoFld in  Fso.GetFolder(Dir).SubFolders
			If Left(SoFld.Name, 5) = CA Then
				SearchFolder = SoFld
				Exit Function
			End If
		Next
	End If
End Function