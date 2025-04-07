Attribute VB_Name = "Module1"
Sub SynchroniserCodeVBA()

    Dim FSO As Object
    Dim ProjetVBA As Object
    Dim Fichier As Object
    Dim RepertoireSource As String
    Dim FichierSource As Object
    Dim DossierSource As FileDialog
    
    ' Cr�er un objet FileDialog pour demander � l'utilisateur de s�lectionner un r�pertoire
    Set DossierSource = Application.FileDialog(msoFileDialogFolderPicker)
    DossierSource.Title = "S�lectionnez le r�pertoire contenant les fichiers VBA"
    
    ' Afficher la bo�te de dialogue et r�cup�rer le r�pertoire s�lectionn�
    If DossierSource.Show = -1 Then
        RepertoireSource = DossierSource.SelectedItems(1) & "\"
    Else
        MsgBox "Aucun r�pertoire s�lectionn�, l'op�ration est annul�e.", vbExclamation
        Exit Sub
    End If

    ' Cr�er un objet FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' R�cup�rer le projet VBA actif
    Set ProjetVBA = Application.VBE.ActiveVBProject

    ' Supprimer tous les composants du projet VBA (modules, classes, formulaires, etc.)
    For Each Fichier In ProjetVBA.VBComponents
        ProjetVBA.VBComponents.Remove Fichier
    Next Fichier

    ' Appeler la fonction pour parcourir le r�pertoire et ses sous-dossiers
    ImporterFichiersVBA FSO, RepertoireSource, ProjetVBA

    MsgBox "Les modules ont �t� synchronis�s avec succ�s.", vbInformation

    ' Lib�rer les objets
    Set FichierSource = Nothing
    Set FSO = Nothing
    Set ProjetVBA = Nothing
    Set DossierSource = Nothing

End Sub

' Fonction pour parcourir le r�pertoire et ses sous-dossiers
Sub ImporterFichiersVBA(FSO As Object, Repertoire As String, ProjetVBA As Object)
    Dim Dossier As Object
    Dim Fichier As Object
    Dim SousDossier As Object

    ' Parcourir tous les fichiers dans le r�pertoire
    For Each Fichier In FSO.GetFolder(Repertoire).Files
        If LCase(FSO.GetExtensionName(Fichier.Name)) = "cls" Or LCase(FSO.GetExtensionName(Fichier.Name)) = "frm" Or LCase(FSO.GetExtensionName(Fichier.Name)) = "bas" Then
            ProjetVBA.VBComponents.Import Fichier.Path
        End If
    Next Fichier

    ' Parcourir tous les sous-dossiers du r�pertoire
    For Each SousDossier In FSO.GetFolder(Repertoire).Subfolders
        ImporterFichiersVBA FSO, SousDossier.Path, ProjetVBA
    Next SousDossier

End Sub

