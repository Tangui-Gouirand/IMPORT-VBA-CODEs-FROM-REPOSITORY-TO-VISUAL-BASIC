Sub SynchroniserCodeVBA()

    Dim FSO As Object
    Dim ProjetVBA As Object
    Dim Fichier As Object
    Dim RepertoireSource As String
    Dim FichierSource As Object
    Dim DossierSource As FileDialog
    
    ' Créer un objet FileDialog pour demander à l'utilisateur de sélectionner un répertoire
    Set DossierSource = Application.FileDialog(msoFileDialogFolderPicker)
    DossierSource.Title = "Sélectionnez le répertoire contenant les fichiers VBA"
    
    ' Afficher la boîte de dialogue et récupérer le répertoire sélectionné
    If DossierSource.Show = -1 Then
        RepertoireSource = DossierSource.SelectedItems(1) & "\"
    Else
        MsgBox "Aucun répertoire sélectionné, l'opération est annulée.", vbExclamation
        Exit Sub
    End If

    ' Créer un objet FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' Récupérer le projet VBA actif
    Set ProjetVBA = Application.VBE.ActiveVBProject

    ' Supprimer tous les composants du projet VBA (modules, classes, formulaires, etc.)
    For Each Fichier In ProjetVBA.VBComponents
        ProjetVBA.VBComponents.Remove Fichier
    Next Fichier

    ' Appeler la fonction pour parcourir le répertoire et ses sous-dossiers
    ImporterFichiersVBA FSO, RepertoireSource, ProjetVBA

    MsgBox "Les modules ont été synchronisés avec succès.", vbInformation

    ' Libérer les objets
    Set FichierSource = Nothing
    Set FSO = Nothing
    Set ProjetVBA = Nothing
    Set DossierSource = Nothing

End Sub

' Fonction pour parcourir le répertoire et ses sous-dossiers
Sub ImporterFichiersVBA(FSO As Object, Repertoire As String, ProjetVBA As Object)
    Dim Dossier As Object
    Dim Fichier As Object
    Dim SousDossier As Object

    ' Parcourir tous les fichiers dans le répertoire
    For Each Fichier In FSO.GetFolder(Repertoire).Files
        If LCase(FSO.GetExtensionName(Fichier.Name)) = "cls" Or LCase(FSO.GetExtensionName(Fichier.Name)) = "frm" Or LCase(FSO.GetExtensionName(Fichier.Name)) = "bas" Then
            ProjetVBA.VBComponents.Import Fichier.Path
        End If
    Next Fichier

    ' Parcourir tous les sous-dossiers du répertoire
    For Each SousDossier In FSO.GetFolder(Repertoire).Subfolders
        ImporterFichiersVBA FSO, SousDossier.Path, ProjetVBA
    Next SousDossier

End Sub
