# Synchroniser-VBA-POWERPOINT

Ce projet fournit un code permettant d'importer rapidement du code VBA (`.cls`, `.bas`, `.frm`) depuis un dossier et ses sous-dossiers vers l'environnement Visual Basic de PowerPoint.

## Tutoriel : Ajouter le code "SYNCHRO" au ruban PowerPoint

### Méthode 1 : Pas à pas (avec personnalisation du ruban)

1.  **Ouvrez PowerPoint.**
2.  **Ouvrez l'éditeur VBA** en appuyant sur `Alt + F11` ou en allant dans le ruban, onglet `Développeur` puis `Visual Basic`.
3.  **Importez le code BAS** fourni dans ce dépôt, via `Fichier` > `Importer un fichier`. [Télécharger SYNCHRO.bas](https://github.com/Tangui-Gouirand/Synchroniser-VBA-POWERPOINT/blob/main/SYNCHRO.bas)
4.  **Enregistrez le fichier au format `.ppam` (Complément PowerPoint)** via `Fichier` > `Enregistrer la présentation sous...` et en choisissant le type de fichier "Complément PowerPoint (*.ppam)".
5.  **Installez et ouvrez l'éditeur RibbonX :** [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).
    * 5.1 Ouvrez le fichier `.ppam` avec l'éditeur RibbonX.
    * 5.2 Faites un clic droit sur le fichier dans l'éditeur RibbonX, puis cliquez sur `Insérer Office 2010+ Custom UI Part`.
    * 5.3 Copiez le contenu du fichier `SYNCHRO.xml` dans la zone de texte qui apparaît.
    * 5.4 Cliquez sur `Enregistrer` dans l'éditeur RibbonX.
6. fermer et relancer Powerpoint.
7. Glisser le fichier SYNCHRO.ppam dans le ruban

### Méthode 2 : Glisser-déposer (Drop and Play)

1.  **Téléchargez le fichier `SYNCHRO.ppam`** fourni dans ce dépôt.
2.  **Ouvrez PowerPoint.**
3.  **Glissez-déposez le fichier `SYNCHRO.ppam`** directement dans la fenêtre de PowerPoint.
4.  **Le ruban sera automatiquement mis à jour** avec le bouton "SYNCHRO".

## Utilisation

1.  **Organisez votre code VBA** dans un dossier (et ses sous-dossiers si nécessaire), en utilisant les extensions `.cls`, `.bas` et `.frm`.
2.  **Cliquez sur le bouton "SYNCHRO"** dans le ruban PowerPoint.
3.  **Sélectionnez le dossier racine** contenant votre code VBA.
4.  **Le code sera automatiquement importé** dans votre projet PowerPoint.

## Prérequis

* Microsoft PowerPoint avec prise en charge de VBA.
* (Pour la méthode 1) [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).

## Contribution

Les contributions sont les bienvenues ! N'hésitez pas à proposer des améliorations ou à signaler des problèmes.
