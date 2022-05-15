# BacCollector

## Téléchargement

[Télécharger BacCollector](https://github.com/romoez/BacCollector/releases)

## Introduction
BacCollector est un outil de récupération des travaux des candidats à l'épreuve pratique d'informatique des examens du baccalauréat.

- Il détermine automatiquement le numéro d'inscription du candidat.
- Il cherche les dossiers et fichiers de travail du candidat, puis :
  - Il crée une sauvegarde dans un dossier verrouillé, sur le PC lui-même.
  - et crée une copie des travaux du candidat sur la clé USB.
- à la fin, il efface les travaux, et prépare le dossier de travail Bac202x pour le prochain candidat.

![Interface BacCollector 1](https://github.com/romoez/BacCollector/blob/main/captures_ecran/BacCollector-InAction.gif)

## Ce que fait BacCollector dès son exécution

- Il essaye de déterminer le numéro d'inscription du candidat à partir des dossiers et fichiers de travail du candidat.
- Il cherche et affiche les logiciels ouverts qui peuvent causer des problèmes lors de la récupération des travaux du candidat.
- Il cherche et affiche les dossiers et les fichiers de travail du candidat.
- Pour l'épreuve d'informatique ou d'algorithmique &amp; programmation :
    - Les dossiers C:\Bac\*2\*
    - Tous les dossiers récemment créés (< 80 min.) sur le bureau, dans Mes documents, dans le dossier du profil de l'utilisateur et sous les racines des lecteurs de disques locaux
    - Tous les fichiers récemment modifiés (< 80 min.) sur le bureau, dans Mes documents, dans le dossier du profil de l'utilisateur et sous les racines des lecteurs de disques locaux
- Pour l'épreuve de STI / TIC : (EasyPhp, Wamp, Xampp, apachefriends)
    - Les sites web créés dans les dossiers d'hébergement d'Apache.
    - Les bases de données MySQL.

![Interface BacCollector 2](https://github.com/romoez/BacCollector/blob/main/captures_ecran/BacCollector-Interface-2.png)

## Les étapes de récupération des travaux d'un candidat (bouton Récupérer)

- Création d'un dossier récupération sur la Clé USB portant comme nom le numéro d'inscription du candidat.
- **[STI]** Copie les sites web situés dans les dossiers d'hébergement des serveurs Apache (EasyPHP / Wamp / Xampp).
- **[STI]** Copie les bases de données à partir du dossier de stockage de MySQL.
- Copie des dossiers Bac\*2\* vers le dossier de Récupération.
- Copie des **dossiers récemment créés** (< 80 min.) sur le bureau, dans Mes documents, dans le dossier du profil de l'utilisateur et sous les racines des lecteurs de disques locaux.
- Copie des **fichiers récemment modifiés** (< 80 min.) sur le bureau, dans Mes documents, dans le dossier du profil de l'utilisateur et sous les racines des lecteurs de disques locaux.
- Création d'une copie des dossiers et fichiers récupérés vers un dossier verrouillé dans le disque dur local. (X:\Sauvegardes\0-BacCollector\123456\)
- **[Info/Algo]** Supprime es fichiers/dossiers CORRECTEMENT COPIÉS.
- Création du dossier C:\Bac202x

## Les contrôles réalisés avant la récupération

- Vérification de la validité du format du numéro d'inscription du candidat (un nombre de six chiffres)
- Vérification si un dossier avec le même nom (numéro du candidat) existe déjà sur la Clé USB.
- Vérification qu'aucun logiciel n'est ouvert. (À partir d'une liste qui comprend : les suites Office, VSCode...)
- [Inf] Vérification de l'existence d'au moins un dossier Bac\*2\* (Ex: Bac2024, Bac-2023...)
- [Inf] Vérification si le dossier Bac202x n'est pas vide.
- [Tic] Vérification de l'existence d'au moins un site web dans l'un des dossiers d'hébergement d'Apache.
- [Tic] Vérification de l'existence d'au moins une Base de données dans l'un des dossiers de stockage MySQL.

## L'opération de Sauvegarde (bouton Créer Sauvegarde sur PC)

Cette opération permet de :

- Créer une sauvegarde sur le disque local de tous les dossiers récupérés,
- Génèrer un rapport de l'opération.
- Génèrer la grille d'évaluation correspondante (Excel).

![Grille d'évaluation Excel](https://github.com/romoez/BacCollector/blob/main/captures_ecran/BacCollector-Grille_Excel.png)


#### Exemple de rapport généré:
![Rapport générer par BacCollector](https://github.com/romoez/BacCollector/blob/main/captures_ecran/BacCollector-Rapport.png)

## Autres informations/fonctionnalités

- BacCollector est un outil portable et il ne demande aucune configuration ou paramétrage préalable.
- On peut utiliser plus d'une Clé USB pour la récupération des dossiers, dans le même &quot;Labo&quot; et en même temps.
- En cas de besoin (Pc infecté, BacCollector bloqué par un antivirus...), on peut faire la récupération manuelle, et le travail récupéré sera affiché la prochaine fois dans la liste des travaux récupérés.
- Pour ajouter un candidat absent, il suffit de créer un dossier vide, à côté de BacCollector, qui porte le numéro d'inscription du candidat, et y ajouter un sous-dossier vide intitulé &quot;Absent&quot;. (Par analogie avec ce qu'on fait aux examens théoriques, avec une feuille d'examen blanche)
- Affichage en temps réel, des principales actions des opérations de récupération et sauvegarde.
- Description en détail, dans un fichier log, de toutes les étapes de récupération et de sauvegarde.
- Après la fin de l'opération de récupération, si BacBackup est installé, BacCollector lui demande de créer un nouveau dossier de captures d'écran.
- Un double-clic sur un fichier/dossier dans la listview, ouvre son emplacement dans l'explorateur.
- Un double-clic sur le log, l'ouvre dans Wordpad
- Un clic sur le nom de la matière, actualise toutes les informations affichées par BacCollector (Logiciels ouverts, liste de dossiers récupérés, contenu du dossier Bac20xx...)
- La modification du &quot;labo&quot; entraine la modification du nom de la Clé USB dans l'explorateur de Windows (Lab-1, Labo-2...).
- BacCollector récupère tous les fichiers nécessaires pour ouvrir,  dans un autre PC, les bases de données InnoDB.
- Le dossier verrouillé &quot;X:\Sauvegardes&quot; n'est accessible qu'à partir du bouton &quot;Ouvrir Dossier de Sauve&quot;.

## Logiciels vérifiés par BacCollector avant la récupération

- Algospear
- Brackets
- Dreamweaver
- Flash
- Free Pascal Ide
- Geany
- Gimp
- IDLE (Python)
- Jupyter notebook & Jupyter Lab
- Libreoffice Base
- Libreoffice Calc
- Libreoffice Draw
- Libreoffice Impress
- Libreoffice Math
- Libreoffice Writer
- Libreoffice Writer
- Microsoft Access
- Microsoft Excel
- Microsoft Frontpage
- Microsoft PowerPoint
- Microsoft Publisher
- Microsoft Word
- Ms Expression Web 4
- Mu editor
- MyPascal
- Notepad
- Notepad++
- Paint
- Paint .Net
- Pascal Editor
- Pascal XE
- PyCharm
- PyScipter
- Qt Designer
- Rapid Php
- Sublime Text
- Spyder IDE
- Thonny
- TPW 1.5
- UltraEdit
- Visual Studio Code
- Webuilder
- Windows Movie Maker
- Wing Python IDE
