# IMPORT-VBA-CODEs-FROM-REPOSITORY-TO-VISUAL-BASIC

# Table des matières (Français)

1.  [Description](#import-vba-from-repository-to-visual-basic)
2.  [Tutoriel : Ajouter le code "SYNCHRO" au ruban PowerPoint](#tutoriel--ajouter-le-code-synchro-au-ruban-powerpoint)
    * [Méthode 1 : Pas à pas (avec personnalisation du ruban)](#méthode-1--pas-à-pas-avec-personnalisation-du-ruban)
    * [Méthode 2 : Glisser-déposer (Drop and Play)](#méthode-2--glisser-déposer-drop-and-play)
3.  [Utilisation](#utilisation)
4.  [Prérequis](#prérequis)
5.  [Contribution](#contribution)

# Table of Contents (English)

1.  [Description (English)](#import-vba-from-repository-to-visual-basic-1)
2.  [Tutorial: Add the "SYNCHRO" code to the PowerPoint ribbon](#tutorial--add-the-synchro-code-to-the-powerpoint-ribbon)
    * [Method 1: Step by step (with ribbon customization)](#method-1--step-by-step-with-ribbon-customization)
    * [Method 2: Drag and drop (Drop and Play)](#method-2--drag-and-drop-drop-and-play)
3.  [Usage](#usage-1)
4.  [Prerequisites](#prerequisites)
5.  [Contribution](#contribution-1)

Ce projet fournit un code permettant d'importer rapidement du code VBA (`.cls`, `.bas`, `.frm`) depuis un dossier et ses sous-dossiers vers l'environnement Visual Basic de PowerPoint.

Je l'utilise pour pouvoir travailler sur Visual Studio Code et importer mon code vers l'environnement de travail Visual Basic rapidement.

## Tutoriel : Ajouter le code "SYNCHRO" au ruban PowerPoint

### Méthode 1 : Pas à pas

1.  **Ouvrez PowerPoint.**
2.  **Ouvrez l'éditeur VBA** en appuyant sur `Alt + F11` ou en allant dans le ruban, onglet `Développeur` puis `Visual Basic`.
3.  **Importez le code BAS** fourni dans ce dépôt, via `Fichier` > `Importer un fichier`. [Télécharger SYNCHRO.bas](https://github.com/Tangui-Gouirand/Synchroniser-VBA-POWERPOINT/blob/main/SYNCHRO.bas)
4.  **Enregistrez le fichier au format `.ppam` (Complément PowerPoint)** via `Fichier` > `Enregistrer la présentation sous...` et en choisissant le type de fichier "Complément PowerPoint (*.ppam)".
5.  **Installez et ouvrez l'éditeur RibbonX :** [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).
    * 5.1 Ouvrez le fichier `.ppam` avec l'éditeur RibbonX.
    * 5.2 Faites un clic droit sur le fichier dans l'éditeur RibbonX, puis cliquez sur `Insérer Office 2010+ Custom UI Part`.
    * 5.3 Copiez le contenu du fichier `SYNCHRO.xml` dans la zone de texte qui apparaît. [Télécharger SYNCHRO.xml](https://github.com/Tangui-Gouirand/Synchroniser-VBA-POWERPOINT/blob/main/SYNCHRO.xml)
    * 5.4 Cliquez sur `Enregistrer` dans l'éditeur RibbonX.
6.  **Fermez et relancez PowerPoint.**
7.  **Glissez-déposez le fichier `SYNCHRO.ppam` directement dans le ruban de PowerPoint.**

### Méthode 2 : Glisser-déposer (Drop and Play)

1.  **Téléchargez le fichier `SYNCHRO.ppam`** fourni dans ce dépôt. [Télécharger SYNCHRO.ppam](https://github.com/Tangui-Gouirand/Synchroniser-VBA-POWERPOINT/blob/main/Synchro.ppam)
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

---

# IMPORT-VBA-FROM-REPOSITORY-TO-VISUAL-BASIC

This project provides code to quickly import VBA code (`.cls`, `.bas`, `.frm`) from a folder and its subfolders into the Visual Basic environment of PowerPoint.

I use it to work on Visual Studio Code and import my code into the Visual Basic work environment quickly.

## Tutorial: Add the "SYNCHRO" code to the PowerPoint ribbon

### Method 1: Step by step

1.  **Open PowerPoint.**
2.  **Open the VBA editor** by pressing `Alt + F11` or going to the ribbon, `Developer` tab, then `Visual Basic`.
3.  **Import the BAS code** provided in this repository, via `File` > `Import File`. [Download SYNCHRO.bas](https://github.com/Tangui-Gouirand/Synchroniser-VBA-POWERPOINT/blob/main/SYNCHRO.bas)
4.  **Save the file in `.ppam` format (PowerPoint Add-in)** via `File` > `Save Presentation As...` and choosing the "PowerPoint Add-in (*.ppam)" file type.
5.  **Install and open the RibbonX editor:** [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).
    * 5.1 Open the `.ppam` file with the RibbonX editor.
    * 5.2 Right-click on the file in the RibbonX editor, then click `Insert Office 2010+ Custom UI Part`.
    * 5.3 Copy the content of the `SYNCHRO.xml` file into the text box that appears. [Download SYNCHRO.xml](https://github.com/Tangui-Gouirand/Synchroniser-VBA-POWERPOINT/blob/main/SYNCHRO.xml)
    * 5.4 Click `Save` in the RibbonX editor.
6.  **Close and reopen PowerPoint.**
7.  **Drag and drop the `SYNCHRO.ppam` file directly into the PowerPoint ribbon.**

### Method 2: Drag and drop (Drop and Play)

1.  **Download the `SYNCHRO.ppam` file** provided in this repository. [Download SYNCHRO.ppam](https://github.com/Tangui-Gouirand/Synchroniser-VBA-POWERPOINT/blob/main/Synchro.ppam)
2.  **Open PowerPoint.**
3.  **Drag and drop the `SYNCHRO.ppam` file** directly into the PowerPoint window.
4.  **The ribbon will automatically update** with the "SYNCHRO" button.

## Usage

1.  **Organize your VBA code** into a folder (and its subfolders if necessary), using the extensions `.cls`, `.bas`, and `.frm`.
2.  **Click the "SYNCHRO" button** in the PowerPoint ribbon.
3.  **Select the root folder** containing your VBA code.
4.  **The code will automatically import** into your PowerPoint project.

## Prerequisites

* Microsoft PowerPoint with VBA support.
* (For Method 1) [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor).

## Contribution

Contributions are welcome! Feel free to propose improvements or report issues.
