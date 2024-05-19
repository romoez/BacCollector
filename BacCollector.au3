#NoTrayIcon
#Region ;**** AutoIt3Wrapper ****
#AutoIt3Wrapper_Compression=2
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#AutoIt3Wrapper_Res_Field=URL|https://github.com/romoez/BacCollector
#AutoIt3Wrapper_Res_Field=Email|moez.romdhane@tarbia.tn
#AutoIt3Wrapper_Run_After=copy "%out%" ".\Tests\Labo00"
#AutoIt3Wrapper_Run_After=copy "%out%" ".\Tests\Labo01"
#AutoIt3Wrapper_Run_After=copy "%out%" ".\Tests\Labo01"
#EndRegion ;**** AutoIt3Wrapper ****

#Region ;**** pragma ****
#pragma compile(Icon, BacCollector.ico)
#pragma compile(Out, BacCollector.exe)
#pragma compile(UPX, False) ;
#pragma compile(Compression, 9)
#pragma compile(FileDescription, Récupération des travaux des candidats à l'épreuve pratique d'informatique des examens du baccalauréat)
#pragma compile(ProductName, BacCollector)
#pragma compile(ProductVersion, 0.8.15.519)
#pragma compile(FileVersion, 0.8.15.519)
#pragma compile(LegalCopyright, 2018-2024 © La Communauté Tunisienne des Enseignants d'Informatique)
#pragma compile(Comments, BacCollector: Récupération des travaux des candidats à l'épreuve pratique d'informatique des examens du baccalauréat)
#pragma compile(CompanyName, La Communauté Tunisienne des Enseignants d'Informatique)
#pragma compile(AutoItExecuteAllowed, False)
#EndRegion ;**** pragma ****

#Region ;**** include ****
#include <Array.au3>
#include <Date.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GUIComboBox.au3>
#include <GuiConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <GuiMenu.au3>
#include <GuiRichEdit.au3>
#include <ListViewConstants.au3>
#include <Process.au3>
#include <StaticConstants.au3>
#include <WinAPIFiles.au3>
#include <WinAPIProc.au3>
#include <WindowsConstants.au3>
#include <GUIToolTip.au3>
#include "Utils.au3"
;~ #include "include\7Zip.au3"  ; https://www.autoitscript.fr/forum/viewtopic.php?f=21&t=1943
#include "include\ExtMsgBox.au3"  ; https://www.autoitscript.com/forum/topic/109096-extended-message-box-new-version-19-nov-21/
#include "include\GUIExtender.au3"  ; https://www.autoitscript.com/forum/topic/117909-guiextender-original-version/
#include "include\MPDF_UDF.au3" ; https://www.autoitscript.com/forum/topic/118827-create-pdf-from-your-application/
#include "include\StringSize.au3"  ; https://www.autoitscript.com/forum/topic/114034-stringsize-m23-new-version-16-aug-11/
#include "include\Zip.au3"  ; https://www.autoitscript.com/forum/topic/73425-zipau3-udf-in-pure-autoit/
;~ #include "CheckSumVerify2.a3x"  ; https://www.autoitscript.com/forum/topic/164148-checksumverify-verify-integrity-of-the-compiled-exe/
#EndRegion ;**** include ****


_KillOtherScript()

#Region ;**** Global Const ****
;#Program.Info.Const
Global Const $PROG_TITLE = "BacCollector "
Global Const $PROG_VERSION = FileGetVersion(@ScriptFullPath) ;"0.8"

Global Const $ANNEES_BAC[5] = ["2024", "2025", "2026", "2027", "2028"]

;#GUI.Taille.Const
Global Const $GUI_LARGEUR = 800
Global Const $GUI_HAUTEUR = 600
Global Const $GUI_MARGE = 10
Global Const $GUI_HEADER_HAUTEUR = 20
Global Const $GUI_LARGEUR_PARTIE = $GUI_LARGEUR / 4
;~ GLOBAL Const $GUI_HEADER_LARGEUR = 20

;#GUI.Colors
Global Const $GUI_COLOR_CENTER = 0x073b4c
;~ Global Const $GUI_COLOR_CENTER = 0x073b4c ; 0x2E363F
Global Const $GUI_COLOR_SIDES = 0x118ab2  ; 0x01A3FA
Global Const $GUI_COLOR_CENTER_HEADERS_TEXT = 0xEEEEEE ;0x5EE505

Global Const $AGE_DU_FICHIER_EN_MINUTES = 80
Global Const $FREE_SPACE_DRIVE_BACKUP = 1024 ;en Mb, >>1Gb
Global Const $TAILLE_MAX_DU_DOSSIER_SOUS_LECTEUR = 10 ;en Mb
#EndRegion ;**** Global Const ****


If $CMDLINE[0] Then
	Select
		Case StringInStr($CMDLINE[1], "Unlock") Or StringInStr($CMDLINE[1], "Deblo") _
				Or StringInStr($CMDLINE[1], "Clea") Or StringInStr($CMDLINE[1], "Nett") _
				Or StringInStr($CMDLINE[1], "Deverr") Or StringInStr($CMDLINE[1], "Déverr") _
				Or StringInStr($CMDLINE[1], "Déblo") ;Debloquer Deblocage Unlock Clean CleanUp Clear Unlock Deverrouillage Deverrouiller Nettoyage Nettoyer...
			_UnLockAll()
			Exit
		Case Else
			MsgBox(262144, $PROG_TITLE & $PROG_VERSION, "Syntaxe de ligne de commande: " & @CRLF & @CRLF _
					 & "» Déverrouillage des dossiers ""X:\Sauvegardes\"" : " & @CRLF _
					 & @TAB & "- BacCollector /Déverrouiller, ou:" & @CRLF _
					 & @TAB & "- BacCollector /Unlock, ou encore:" & @CRLF _
					 & @TAB & "- BacCollector /Débloquer" & @CRLF & @CRLF _
					 & "» Récupération automatique des dossiers de travail du candidat: " & @CRLF _
					 & @TAB & "- [Cette fonctionnalité n'est pas encore implémentée.]") ; & @CRLF _
			Exit
	EndSelect
EndIf


_CreateGui()
_InitLogging()
_InitialParams()
_Initialisation()
_CheckRunningFromUsbDrive()
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
WinActivate($hMainGUI)

While 1

	$aMsg = GUIGetMsg(1)
	Switch $aMsg[0]
		Case $GUI_EVENT_CLOSE
			_CheckBacCollectorExists()
			_FinLogging()
			_GUICtrlRichEdit_Destroy($GUI_Log)
			Exit
		Case $bTogglePart3
			_CheckBacCollectorExists()
			_TogglePart3()
		Case $bRecuperer
			_CheckBacCollectorExists()
			GUISetState(@SW_DISABLE, $hMainGUI) ;
			_UnLockFoldersBC()
			$Ret = Recuperer()
			GUISetState(@SW_ENABLE, $hMainGUI) ;
			_LockFoldersBC()
			If $Ret <> 0 Then
				_Initialisation(0) ; ->0:Ne pas initialiser le numero de candidat ->1:init
			EndIf
			WinActivate($hMainGUI)
		Case $bCreerSauvegarde
			_CheckBacCollectorExists()
			GUISetState(@SW_DISABLE, $hMainGUI) ;
			_UnLockFoldersBC()
			Sauvegarder()
			_LockFoldersBC()
			GUISetState(@SW_ENABLE, $hMainGUI) ;
			WinActivate($hMainGUI)
;~ 			WinActivate($hMainGUI) déplacé vers la fonction, si affich Rapport, on n'acive pas la fenêtre
		Case $bOpenBackupFldr
			_CheckBacCollectorExists()
			_InitialParams()
			WinActivate($hMainGUI)
			_UnLockFoldersBC()
			If Not FileExists($Lecteur & $DossierSauve & "\" & $DossierBacCollector) Then _InitialParams()
			Run("explorer.exe /e, " & '"' & $Lecteur & $DossierSauve & "\" & $DossierBacCollector & '"')
			_Logging("Ouverture du dossier de sauvegarde : " & @CRLF & "                            " _
					 & """" & $Lecteur & $DossierSauve & "\" & $DossierBacCollector & """", 2, 0)

		Case $lblBacBackup
			_CheckBacCollectorExists()
			_OpenBacBackupInterface()
;~ 			MsgBox('0',"","Hello")
		Case $bAide
			_CheckBacCollectorExists()
			_ShowHelp()
		Case $rInfoProg, $tInfoProg
			_CheckBacCollectorExists()
			If $Matiere <> 'InfoProg' Then
				_Logging("Changement de la matière : ""STI"" --> ""Info/Prog""", 2, 0)
				$Matiere = 'InfoProg'
				_InitialisationInfo()
			EndIf

		Case $rTic, $tTic
			_CheckBacCollectorExists()
			If $Matiere <> 'STI' Then
				_Logging("Changement de la matière : ""Info/Prog"" --> ""STI""", 2, 0)
				$Matiere = 'STI'
				_InitialisationTic()
			EndIf

		Case $lblMatiere
			_CheckBacCollectorExists()
			_Logging("Demande d'actualisation des données suite au clic sur le label de la matière : " & $Matiere, 2, 0)
			_Initialisation()

		Case $cBac
			_CheckBacCollectorExists()
			_SaveData_Bacxx()

		Case $cSeance
			_CheckBacCollectorExists()
			_SaveData_Seancexx()

		Case $cLabo
			_CheckBacCollectorExists()
			_SaveData_Laboxx()

		Case $lblComputerID
			If ClipPut(_GetUUID()) Then
				ToolTip ( "Identifiant copié!", Default, Default, Default, 1, 1)
				Sleep(1000)
				ToolTip("")
			EndIf
	EndSwitch
WEnd

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _Initialisation($initNumeroCandidat = 1)
	If $Matiere = "STI" Then
		_InitialisationTic($initNumeroCandidat)
	Else
		_InitialisationInfo($initNumeroCandidat)
	EndIf
EndFunc   ;==>_Initialisation

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

#Region Function "_InitialisationInfo" ------------------------------------------------------------------------
Func _InitialisationInfo($initNumeroCandidat = 1)
	SplashTextOn("Sans Titre", "Initialisation de l'opération." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	GUICtrlSetData($HeadContenuDossiersBac, "Contenu des dossiers ?Bac*2*?:")
	GUICtrlSetData($HeadContenuAutresDossiers, "Autres Fichiers dans ""Mes Documents"", ""Bureau"", racine des lecteurs du hdd :")

	GUICtrlSetTip($HeadContenuDossiersBac, " ", "Contenu des dossiers ?Bac*2*?", 1,1)
	GUICtrlSetTip($HeadContenuAutresDossiers, "- Bureau" & @CRLF & "- Mes documents" & @CRLF & "- Dossier du profil de l'utilisateur" & @CRLF & "- Racines des lecteurs C:, D: ..." & @CRLF & @CRLF &  "(Double-clic sur un élément pour l'afficher dans l'Explorateur.)", "Éléments récemment créés/modifiés dans :", 1,1)

	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesDossiersBac, 0, "Dossiers & Fichiers")
	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesDossiersBac, 1, "Taille")
	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesAutresDossiers, 0, "Dossiers & Fichiers")
	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesAutresDossiers, 1, "Taille")

	GUISetState(@SW_DISABLE, $hMainGUI) ;
	If $initNumeroCandidat = 1 Then
		GUICtrlSetData($GUI_NumeroCandidat, _NumeroCandidat())
	EndIf

	;;;Section - Début ===================================
	Local $sFiltreFichiersAChercher
	IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Matiere", $Matiere)
	Switch $Matiere
		Case 'InfoProg'
			GUICtrlSetState($rInfoProg, $GUI_CHECKED)
			GUICtrlSetState($rTic, $GUI_UNCHECKED)
			GUICtrlSetData($lblMatiere, "Info/Prog")
			$sFiltreFichiersAChercher = "*"
		Case 'STI'
			GUICtrlSetState($rInfoProg, $GUI_UNCHECKED)
			GUICtrlSetState($rTic, $GUI_CHECKED)
			GUICtrlSetData($lblMatiere, "STI")
			$sFiltreFichiersAChercher = "*"
	EndSwitch
	;;;Section - Fin ===================================
	Local $TmpSpaces = ""
	$sApps = _ListeDApplicationsOuvertes()
	If $sApps <> "" Then
		$TmpSpaces = "                                           "
		_Logging("Applications ouvertes: " & StringReplace($sApps, @CRLF, @CRLF & $TmpSpaces), 2, 0)
		GUICtrlSetState($TextApps_Header, BitOR($GUI_SHOW, $GUI_DISABLE))
		GUICtrlSetState($TextApps_Text, $GUI_SHOW)
		GUICtrlSetState($TextApps, $GUI_SHOW)
		GUICtrlSetData($TextApps, $sApps)
	Else
		GUICtrlSetState($TextApps_Header, BitOR($GUI_HIDE, $GUI_DISABLE))
		GUICtrlSetState($TextApps_Text, $GUI_HIDE)
		GUICtrlSetState($TextApps, $GUI_HIDE)
		GUICtrlSetData($TextApps, "")
	EndIf

	Local $sDossiersRecuperes = _ListeDeDossiersRecuperes()
	If $sDossiersRecuperes <> "" Then
		GUICtrlSetData($TextDossiersRecuperes, $sDossiersRecuperes)
		$TmpSpaces = "                         "
		_Logging("Liste de dossiers de travail déjà récupérés : " & StringReplace($sDossiersRecuperes, @CRLF, @CRLF & $TmpSpaces), 2, 0)
;~ 		_Logging("Liste : " & _ArrayToString($Liste, "", 1, -1, @CRLF & $TmpSpaces, 0, 0), 2, 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	Else
		GUICtrlSetData($TextDossiersRecuperes, "-")
	EndIf

	_AfficherLeContenuDesDossiersBac()
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesDossiersBac, 3) ;Hide FullPath column
	_AfficherLeContenuDesAutresDossiers($sFiltreFichiersAChercher)
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesAutresDossiers, 3) ;Hide FullPath column

	GUISetState(@SW_ENABLE, $hMainGUI) ;
	WinActivate ($hMainGUI);
	If $DossiersNonConformesPourCandidatsAbsents <> "" Then
		If _NbOccurrences("»", $DossiersNonConformesPourCandidatsAbsents) = 1 Then
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "Pour le canditat absent suivant :" & @CRLF & @CRLF _
					 & $DossiersNonConformesPourCandidatsAbsents & @CRLF _
					 & "Veuillez créer dans son dossier, un sous-dossier vide intitulé :" & @CRLF & @CRLF _
					 & @TAB & """Absent""", 0)
		Else
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "Pour les canditats absents suivants :" & @CRLF & @CRLF _
					 & $DossiersNonConformesPourCandidatsAbsents & @CRLF _
					 & "Veuillez créer dans le dossier de chacun, un sous-dossier vide intitulé :" & @CRLF & @CRLF _
					 & @TAB & """Absent""", 0)
		EndIf
	EndIf

EndFunc   ;==>_InitialisationInfo
#EndRegion Function "_InitialisationInfo" ------------------------------------------------------------------------

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

#Region Function "_InitialisationTic" ------------------------------------------------------------------------
Func _InitialisationTic($initNumeroCandidat = 1)
	SplashTextOn("Sans Titre", "Initialisation de l'opération." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	GUICtrlSetData($HeadContenuDossiersBac, "Contenu des Dossiers d'hébergement des serveurs Web :")
	GUICtrlSetData($HeadContenuAutresDossiers, "Bases de Données et tables MySql:")

	GUICtrlSetTip($HeadContenuDossiersBac, " ", "Contenu des Dossiers d'hébergement des serveurs Web", 1,1)
	GUICtrlSetTip($HeadContenuAutresDossiers, " ", "Bases de Données et tables MySql", 1,1)

	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesDossiersBac, 0, "Sites Web & Fichiers")
	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesDossiersBac, 1, "Type")
	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesAutresDossiers, 0, "Bases de Données & Tables")
	_GUICtrlListView_SetColumn($GUI_AfficherLeContenuDesAutresDossiers, 1, "Type")

	GUISetState(@SW_DISABLE, $hMainGUI) ;
	If $initNumeroCandidat = 1 Then
		GUICtrlSetData($GUI_NumeroCandidat, _NumeroCandidatTic())
	EndIf

	IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Matiere", $Matiere)
	Switch $Matiere
		Case 'InfoProg'
			GUICtrlSetState($rInfoProg, $GUI_CHECKED)
			GUICtrlSetState($rTic, $GUI_UNCHECKED)
			GUICtrlSetData($lblMatiere, "Info/Prog")
			$sFiltreFichiersAChercher = "*"
		Case 'STI'
			GUICtrlSetState($rInfoProg, $GUI_UNCHECKED)
			GUICtrlSetState($rTic, $GUI_CHECKED)
			GUICtrlSetData($lblMatiere, "STI")
			$sFiltreFichiersAChercher = "*"
	EndSwitch

	$sApps = _ListeDApplicationsOuvertes()
	If $sApps <> "" Then
		GUICtrlSetState($TextApps_Header, BitOR($GUI_SHOW, $GUI_DISABLE))
		GUICtrlSetState($TextApps_Text, $GUI_SHOW)
		GUICtrlSetState($TextApps, $GUI_SHOW)
		GUICtrlSetData($TextApps, $sApps)
	Else
		GUICtrlSetState($TextApps_Header, BitOR($GUI_HIDE, $GUI_DISABLE))
		GUICtrlSetState($TextApps_Text, $GUI_HIDE)
		GUICtrlSetState($TextApps, $GUI_HIDE)
		GUICtrlSetData($TextApps, "")
	EndIf

	Local $sDossiersRecuperes = _ListeDeDossiersRecuperes() ;_ListeDeDossiersRecuperes($SearchMask = "??????", $RegEx = "([0-9]{6})")
	If $sDossiersRecuperes <> "" Then
		GUICtrlSetData($TextDossiersRecuperes, $sDossiersRecuperes)
	Else
		GUICtrlSetData($TextDossiersRecuperes, "-")
	EndIf

	_AfficherLeContenuDesDossiersTic()
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesDossiersBac, 3) ;Hide FullPath column
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesAutresDossiers, 3) ;Hide FullPath column
	GUISetState(@SW_ENABLE, $hMainGUI) ;
	WinActivate ($hMainGUI);
	If $DossiersNonConformesPourCandidatsAbsents <> "" Then
		If _NbOccurrences("»", $DossiersNonConformesPourCandidatsAbsents) = 1 Then
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "Pour le canditat absent suivant :" & @CRLF & @CRLF _
					 & $DossiersNonConformesPourCandidatsAbsents & @CRLF _
					 & "Veuillez créer dans son dossier, un sous-dossier vide intitulé :" & @CRLF & @CRLF _
					 & @TAB & """Absent""", 0)
		Else
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "Pour les canditats absents suivants :" & @CRLF & @CRLF _
					 & $DossiersNonConformesPourCandidatsAbsents & @CRLF _
					 & "Veuillez créer dans le dossier de chacun, un sous-dossier vide intitulé :" & @CRLF & @CRLF _
					 & @TAB & """Absent""", 0)
		EndIf
	EndIf

EndFunc   ;==>_InitialisationTic
#EndRegion Function "_InitialisationTic" ------------------------------------------------------------------------

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

#Region Function "_InitialParams" --------------------------------------------------------------------------
Func _InitialParams()
	If not FileExists(StringTrimRight(@ScriptFullPath, 4) & ".ini") Then
		_GUIExtender_Section_Action($hMainGUI, 1, 1) ;2ème paramètre (1) Numéro de la Section  -  3ème paramètre (0) to retract the Section
	EndIf

;~ 	SplashTextOn("Sans Titre", "Lecture des paramètres de ""BacCollector""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Mise à jour des paramètres...", "[0%] Veuillez patienter un moment, initialisation...", Default, Default, 1)

	Global $DossierBacCollector = "0-BacCollector"
	Global $DossierSauve = "Sauvegardes"
	Global $Lecteur = ""

	;Contrôle ajouté après la Demo au Crefoc de Tunis 1
	Local $aDrive = DriveGetDrive('FIXED')

;~ 		_ArraySort($aDrive, 0, 1)
	$Lecteur = @HomeDrive ; "C:" ; $aDrive[1] ; $aDrive[1] peut être A: !!
	For $i = 1 To $aDrive[0]
		If $aDrive[$i] = @HomeDrive Then ContinueLoop
		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; to Exclude external Hdd(s)
				And _WinAPI_IsWritable($aDrive[$i]) _ ;writable
				And DriveSpaceFree($aDrive[$i] & "\") > $FREE_SPACE_DRIVE_BACKUP _
				Then
			$Lecteur = $aDrive[$i]
			ExitLoop
		EndIf
	Next
	$Lecteur = StringUpper($Lecteur) & "\"
	_Logging("Mise à jour des paramètres ____Début___", 2, 0)
	ProgressSet(20, "[" & 20 & "%] " & "")
	Local $Ok
	If Not FileExists($Lecteur & $DossierSauve) Then
		$Ok = DirCreate($Lecteur & $DossierSauve)
		If $Ok Then
			_Logging("Création du dossier de sauvegarde : " & $Lecteur & $DossierSauve, 1, 0)
		Else
			_Logging("Création du dossier de sauvegarde : " & $Lecteur & $DossierSauve, 0, 0)
		EndIf
	Else
		_Logging("Dossier de sauvegarde : """ & $Lecteur & $DossierSauve & """", 2, 0)
	EndIf

	Local $Repeated = 0, $Success = 0
	If Not FileExists($Lecteur & $DossierSauve & "\" & $DossierBacCollector) Then
		While ($Repeated < 5)
			_UnLockFolder($Lecteur & $DossierSauve)
			$Success = DirCreate($Lecteur & $DossierSauve & "\" & $DossierBacCollector)
			$Repeated += 1
			If $Success Then
				_Logging("Création du dossier  BacCollector : " & $Lecteur & $DossierSauve & "\" & $DossierBacCollector, 1, 0)
				ExitLoop
			Else
				Sleep(20)
			EndIf
		WEnd
		If Not $Success Then
			ProgressOff()
			_Logging("Création du dossier  BacCollector : " & $Lecteur & $DossierSauve & "\" & $DossierBacCollector, 0, 0)
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Échec lors de la création du dossier de sauvegarde local, " & @CRLF _
					 & "     """ & $Lecteur & $DossierSauve & "\" & $DossierBacCollector & """" & @CRLF _
					 & "ça peut provoquer des erreurs lors de la copie de récupération!", 0)
			ProgressOn($PROG_TITLE & $PROG_VERSION, "Mise à jour des paramètres...", "", Default, Default, 1)
		EndIf
	Else
		_Logging("Dossier BacCollector  : """ & $Lecteur & $DossierSauve & "\" & $DossierBacCollector & """", 2, 0)
	EndIf
	FileSetAttrib($Lecteur & $DossierSauve, "+SH")
	_LockFolder($Lecteur & $DossierSauve)
	ProgressSet(40, "[" & 40 & "%] " & "")

	Global $Matiere = IniRead(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Matiere", "-1")
;~ 	If ($Matiere = -1) Or ($Matiere <> 'InfoProg' And $Matiere <> 'STI') Then
	If ($Matiere = -1) Or (($Matiere <> 'InfoProg') And ($Matiere <> 'STI')) Then
		$Matiere = 'InfoProg' ;Valeur Par défaut
		IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Matiere", $Matiere)
	EndIf
	ProgressSet(55, "[" & 55 & "%] " & "")

	;;;Bac20xx?
	Local $Bac20xx = IniRead(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Bac", "-1")
	If StringIsInt($Bac20xx) = 0 Or $Bac20xx < $ANNEES_BAC[0] Or $Bac20xx > $ANNEES_BAC[4] Then
		$Bac20xx = $ANNEES_BAC[0]
		IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Bac", $Bac20xx)
	EndIf
	_GUICtrlComboBox_SelectString($cBac, $Bac20xx)
	GUICtrlSetData($lblBac, "Bac" & $Bac20xx)
	_Logging("Baccalauréat          : " & "Bac" & $Bac20xx, 2, 0)
	ProgressSet(70, "[" & 70 & "%] " & "")

	;;;Labo-x?
	Local $Laboxx = IniRead(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Labo", "-1")
	If StringIsInt($Laboxx) = 0 Or $Laboxx < 1 Or $Laboxx > 6 Then
		$Laboxx = '1'
		IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Labo", $Laboxx)
	EndIf
	_GUICtrlComboBox_SelectString($cLabo, "Labo-" & $Laboxx)
	GUICtrlSetData($lblLabo, "Labo-" & $Laboxx)
	_Logging("Laboratoire           : " & "Labo-" & $Laboxx, 2, 0)
	ProgressSet(85, "[" & 85 & "%] " & "")

	;;;Séance?
	Local $Seancexx = IniRead(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Seance", "-1")
	If StringIsInt($Seancexx) = 0 Or $Seancexx < 1 Or $Seancexx > 6 Then
		$Seancexx = '1'
		IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Seance", $Seancexx)
	EndIf
	_GUICtrlComboBox_SelectString($cSeance, "Séance-" & $Seancexx)
	GUICtrlSetData($lblSeance, "Séance-" & $Seancexx)
	_Logging("Séance                : " & "Séance-" & $Seancexx, 2, 0)
	ProgressSet(95, "[" & 95 & "%] " & "")

	;;;BacBackup détecté
	If _IsRegistryExist("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{498AA8A4-2CBE-4368-BFA0-E0CF3F338536}_is1", "DisplayName") And ProcessExists('BacBackup.exe') Then
		GUICtrlSetBkColor($lblBacBackup, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetFont(-1, 11, 100, 0, "Segoe UI Light")
	;~ 	GUICtrlSetState($lblBacBackup, $GUI_HIDE)
		GUICtrlSetTip($lblBacBackup, "Cliquez ici pour ouvrir BacBackup", "BacBackup surveille le PC.", 0,1)
		GUICtrlSetData($lblBacBackup, "BacBackup est Installé.")
;~ 		_Logging("BacBackup est installé, et surveille les dossiers de travail des candidats.", 2, 0)
	Else
		GUICtrlSetBkColor($lblBacBackup, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetFont(-1, 11, 800, 0, "Segoe UI Light")
;~ 		GUICtrlSetState($lblBacBackup, $GUI_HIDE)
		GUICtrlSetTip($lblBacBackup, "Cliquez ici pour aller à la page de téléchargement de BacBackup", "BacBackup", 0,1)
		GUICtrlSetData($lblBacBackup, "BacBackup non Installé ;(")
;~ 		_Logging("BacBackup n'est pas installé" & $Bac20xx, 2, 0)
	EndIf

	ProgressSet(99, "[" & 99 & "%] " & "")

	_Logging("Mise à jour des paramètres ____Fin___", 2, 0)
	ProgressOff()
;~ 	SplashOff()
;~ 	WinActivate ($hMainGUI)

EndFunc   ;==>_InitialParams
#EndRegion Function "_InitialParams" --------------------------------------------------------------------------

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

#Region Functions _SaveParams & _SaveData_xxxx
Func _SaveParams()
	;;; En cas où le fichier ini est effacé après l'ouvertur du BacCollector,
	;;; Exple: Après l'ouvertur de BacCollector, l'utilisateur a volu effacer le contenu du Flash USB
	SplashTextOn("Sans Titre", "Enregistrement des paramètres de ""BacCollector""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	;;;Save $Matiere From Interface to ini File
	If GUICtrlGetState($rInfoProg) = $GUI_CHECKED Then
		$Matiere = 'InfoProg'
	ElseIf GUICtrlGetState($rTic) = $GUI_CHECKED Then
		$Matiere = 'STI'
	EndIf
	IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Matiere", $Matiere)

	_SaveData_Bacxx()
	_SaveData_Seancexx()
	_SaveData_Laboxx()
	SplashOff()
;~ 	WinActivate ($hMainGUI)
EndFunc   ;==>_SaveParams

Func _SaveData_Bacxx()
	Local $Bac20xx_lbl = StringRight(GUICtrlRead($lblBac), 4)
	Local $Bac20xx_combo = GUICtrlRead($cBac)
	If $Bac20xx_lbl <> $Bac20xx_combo Then
;~ 		_ExtMsgBox ($vIcon, $vButton, $sTitle, $sText, [$iTimeout, [$hWin, [$iVPos, [$bMain = True]]]])
		_Logging("MsgBox: Modifier l'année du Bac : " & $Bac20xx_lbl & " --> " & $Bac20xx_combo & ". (Oui/Non)?", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & $PROG_VERSION, "Êtes-vous sûr(e) de vouloir modifier : " & @CRLF _
				 & "    " & "<<Bac" & $Bac20xx_lbl & ">>" & "   en   " & "<<Bac" & $Bac20xx_combo & ">>", 0)
		If $Rep = 2 Then
			GUICtrlSetData($lblBac, "Bac" & $Bac20xx_combo)
			_Logging("Oui. Nouveau dossier bac : """ & GUICtrlRead($lblBac) & """", 2, 0)
		Else
			_GUICtrlComboBox_SelectString($cBac, $Bac20xx_lbl)
			_Logging("Non.", 2, 0)
		EndIf
	EndIf
	IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Bac", GUICtrlRead($cBac))
EndFunc   ;==>_SaveData_Bacxx

Func _SaveData_Seancexx()
	Local $Seancexx_lbl = StringRight(GUICtrlRead($lblSeance), 1)
	Local $Seancexx_combo = StringRight(GUICtrlRead($cSeance), 1)
	If $Seancexx_lbl <> $Seancexx_combo Then
;~ 		_ExtMsgBox ($vIcon, $vButton, $sTitle, $sText, [$iTimeout, [$hWin, [$iVPos, [$bMain = True]]]])
		_Logging("MsgBox: Modifier la séance : " & $Seancexx_lbl & " --> " & $Seancexx_combo & ". (Oui/Non)?", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & $PROG_VERSION, "Êtes-vous sûr(e) de vouloir modifier : " & @CRLF _
				 & "    " & "<<Séance-" & $Seancexx_lbl & ">>" & "   en   " & "<<Séance-" & $Seancexx_combo & ">>", 0)
		If $Rep = 2 Then
			GUICtrlSetData($lblSeance, "Séance-" & $Seancexx_combo)
			_Logging("Oui. Nouvelle séance  : """ & GUICtrlRead($lblSeance) & """", 2, 0)
		Else
			_GUICtrlComboBox_SelectString($cSeance, "Séance-" & $Seancexx_lbl)
			_Logging("Non.", 2, 0)
		EndIf
	EndIf
	IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Seance", StringRight(GUICtrlRead($cSeance), 1))
EndFunc   ;==>_SaveData_Seancexx

Func _SaveData_Laboxx()
	Local $Laboxx_lbl = StringRight(GUICtrlRead($lblLabo), 1)
	Local $Laboxx_combo = StringRight(GUICtrlRead($cLabo), 1)
	If $Laboxx_lbl <> $Laboxx_combo Then
;~ 		_ExtMsgBox ($vIcon, $vButton, $sTitle, $sText, [$iTimeout, [$hWin, [$iVPos, [$bMain = True]]]])
		_Logging("MsgBox: Modifier du laboratoire : " & $Laboxx_lbl & " --> " & $Laboxx_combo & ". (Oui/Non)?", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & $PROG_VERSION, "Êtes-vous sûr(e) de vouloir modifier : " & @CRLF _
				 & "    " & "<<Labo-" & $Laboxx_lbl & ">>" & "   en   " & "<<Labo-" & $Laboxx_combo & ">>", 0)
		If $Rep = 2 Then
			GUICtrlSetData($lblLabo, "Labo-" & $Laboxx_combo)
			;Flash Disk Label
			Local $sScriptDrive = StringLeft(@ScriptFullPath, StringLen(@ScriptFullPath) - StringLen(@ScriptName))
			_Logging("Oui. Nouveau labo  : """ & GUICtrlRead($lblLabo) & """", 2, 0)
			If StringLen($sScriptDrive) = 3 And StringRegExp($sScriptDrive, "^([a-zA-Z]{1,1}:)?(\\|\/)$", 0) = 1 And DriveGetType($sScriptDrive) = "Removable" Then
				Local $Ok = DriveSetLabel($sScriptDrive, "Labo-" & $Laboxx_combo)
				If $Ok Then
					_Logging("Changemant du nom de la clé USB """ & $sScriptDrive & """ vers """ & GUICtrlRead($lblLabo) & """", 1, 0)
				Else
					_Logging("Changemant du nom de la clé USB """ & $sScriptDrive & """ vers """ & GUICtrlRead($lblLabo) & """", 0, 0)
				EndIf
			EndIf
		Else
			_GUICtrlComboBox_SelectString($cLabo, "Labo-" & $Laboxx_lbl)
			_Logging("Non.", 2, 0)
		EndIf
	EndIf
	IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Labo", StringRight(GUICtrlRead($cLabo), 1))
EndFunc   ;==>_SaveData_Laboxx
#EndRegion Functions _SaveParams & _SaveData_xxxx

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _ShowHelp()
	DirCreate(@TempDir & "\BacCollector")
	FileInstall(".\AideBacCollector\AideBacCollector.chm", @TempDir & "\BacCollector\")
	If FileExists(@TempDir & "\BacCollector" & "\AideBacCollector.chm") Then
		ShellExecute(@TempDir & "\BacCollector" & "\AideBacCollector.chm")
	EndIf
EndFunc   ;==>_ShowHelp

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

#Region Functions BB
Func _NouvelleSessionBacBackup()
	Local $iID = ProcessExists('BacBackup.exe')

	If $iID = 0 Then
		Return 0
	EndIf
	Run(_WinAPI_GetProcessFileName($iID))  ; Le lancement de BacBackup arrête l'ancienne instance, et crée un nouveau dossier de capture
	Return 1
EndFunc   ;==>_NouvelleSessionBacBackup

Func _OpenBacBackupInterface()
	Local $iID = ProcessExists('BacBackup.exe')

	If $iID = 0 Then
		_Logging("Ouverture de la page de télépchargement de BacBackup", 2, 0)
		ShellExecute("https://github.com/romoez/BacBackup", "", "", "open")
;~ 		GUICtrlSetState($lblBacBackup, $GUI_HIDE)
		Return 0
	EndIf
	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = "", $sBBInterface = ""
	Local $aPathSplit = _PathSplit(_WinAPI_GetProcessFileName($iID), $sDrive, $sDir, $sFileName, $sExtension)

	$sBBInterface = $sDrive & $sDir & "BacBackup_Interface.exe"
;~ 	MsgBox(0,"",$sBBInterface)
	If FileExists($sBBInterface) Then
		_Logging("Ouverture de la fenêtre de BacBackup", 2, 0)
		Return Run("""" & $sBBInterface & """")
	EndIf
	Return -1
EndFunc   ;==>_OpenBacBackupInterface
#EndRegion Functions BB

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _TogglePart3()
	Switch _GUIExtender_Section_State($hMainGUI, 1)
		Case 0
			; Extend to the right (default)
			_GUIExtender_Section_Action($hMainGUI, 1, 1)
			_Logging("Ouverture de la partie droite de l'interface", 2, 0)
		Case 1
			; Retract from the right (default
			_GUIExtender_Section_Action($hMainGUI, 1, 0)
			_Logging("Fermeture de la partie droite de l'interface", 2, 0)
	EndSwitch
EndFunc   ;==>_TogglePart3

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

#Region Function "_CreateGui" -----------------------------------------------------------------------------
Func _CreateGui()
	;;; Fenêtre principale  - Début
	Global $hMainGUI = GUICreate($PROG_TITLE & $PROG_VERSION, $GUI_LARGEUR, $GUI_HAUTEUR, -1, -1, -1, 0)
	GUISetFont(8, 400, 0, "Tahoma")
	GUISetBkColor($GUI_COLOR_CENTER)
	;;; Fenêtre principale  - Fin

	;;; Partie 3 à droite + bouton Slider - Début===================================================================================
	_GUIExtender_Init($hMainGUI, 1)
	Global $bTogglePart3 = GUICtrlCreateButton(" ", $GUI_LARGEUR - $GUI_LARGEUR_PARTIE - 5, $GUI_HAUTEUR / 4, 5, $GUI_HAUTEUR / 2) ;"text", left, top [, width [, height
	GUICtrlSetColor(-1, $GUI_COLOR_SIDES)
	GUICtrlSetBkColor(-1, $GUI_COLOR_SIDES)
	_GUIExtender_Section_Create($hMainGUI, $GUI_LARGEUR - $GUI_LARGEUR_PARTIE, $GUI_LARGEUR_PARTIE)
	_GUIExtender_Section_Activate($hMainGUI, 1)
	GUICtrlSetTip($bTogglePart3, " ", "Cliquez ici pour afficher/masquer la partie droite de la fenêtre.", 0,1)


	; Déplacé vers _InitialParams(), si le fichier de config existe...
	_GUIExtender_Section_Action($hMainGUI, 1, 1) ;2ème paramètre (1) Numéro de la Section  -  3ème paramètre (0) to retract the Section

	;;; Partie 3 à droite + bouton Slider - Fin=====================================================================================

	$Gui_Partie_Gauche = GUICtrlCreateGraphic(0, 0, $GUI_LARGEUR_PARTIE, $GUI_HAUTEUR)
	GUICtrlSetBkColor($Gui_Partie_Gauche, $GUI_COLOR_SIDES)
	GUICtrlSetState($Gui_Partie_Gauche, $GUI_DISABLE)

	$Gui_Partie_Droite = GUICtrlCreateGraphic($GUI_LARGEUR - $GUI_LARGEUR_PARTIE, 0, $GUI_LARGEUR_PARTIE, $GUI_HAUTEUR)
	GUICtrlSetBkColor($Gui_Partie_Droite, $GUI_COLOR_SIDES)
	GUICtrlSetState($Gui_Partie_Droite, $GUI_DISABLE)

	;;;Grand Cadre - Début==================================================================================
	Local $Left_Header = $GUI_LARGEUR_PARTIE + $GUI_MARGE
	Local $Top_Header = $GUI_MARGE
	Local $WidthHeader = 2 * ($GUI_LARGEUR_PARTIE - $GUI_MARGE)
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $GUI_HAUTEUR - 2 * $GUI_MARGE)
	GUICtrlSetColor($Header, 0x696A65)
	GUICtrlSetState($Header, $GUI_DISABLE)
	;;;Grand Cadre - Fin==================================================================================

	;;;Numéro du Candidat - Début==================================================================================
	;;;========Header============
	$Top_Header = $GUI_MARGE
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $GUI_HEADER_HAUTEUR)
	GUICtrlSetColor($Header, 0x696A65)
	GUICtrlSetState($Header, $GUI_DISABLE)
	$Text = GUICtrlCreateLabel("Numéro d'Inscription du Candidat :", $Left_Header + $GUI_MARGE, $Top_Header + 3, ($WidthHeader - 2 * $GUI_MARGE)/2, $GUI_HEADER_HAUTEUR - 6)
	GUICtrlSetColor($Text, $GUI_COLOR_CENTER_HEADERS_TEXT)
	GUICtrlSetBkColor($Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetTip($Text, "Veuillez vérifier ce numéro avant de commencer la récupération.", "Numéro d'Inscription du Candidat", 1,1)
	;;;========Content============
	Local $GUI_Largeur_Label = 121
	Local $GUI_Hateur_Label = 35
	Global $GUI_NumeroCandidat = GUICtrlCreateInput("", 1.5 * $GUI_LARGEUR_PARTIE - $GUI_Largeur_Label / 2, $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE, $GUI_Largeur_Label, $GUI_Hateur_Label, $ES_NUMBER, BitOR($WS_EX_WINDOWEDGE, $WS_EX_CLIENTEDGE))
	GUICtrlSetFont($GUI_NumeroCandidat, 22, 900, 0, "Arial") ; Bold
	GUICtrlSetColor($GUI_NumeroCandidat, $GUI_COLOR_CENTER)
	GUICtrlSetTip($GUI_NumeroCandidat, "Veuillez vérifier ce numéro avant de commencer la récupération.", "Numéro d'Inscription du Candidat", 1,1)
	;;;Numéro du Candidat - Fin====================================================================================

	;;;Matière - Début==================================================================================
	;;;========Header============
	$Top_Header = $GUI_MARGE
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader / 2, $GUI_HEADER_HAUTEUR + 2 * $GUI_MARGE + $GUI_Hateur_Label + 2 * $GUI_MARGE + 1)
	GUICtrlSetColor($Header, 0x696A65)
	GUICtrlSetState($Header, $GUI_DISABLE)
	$Text = GUICtrlCreateLabel("Matière :", $Left_Header + $WidthHeader / 2 + $GUI_MARGE, $Top_Header + 3, ($WidthHeader - 2 * $GUI_MARGE)/2, $GUI_HEADER_HAUTEUR - 6)
	GUICtrlSetColor($Text, $GUI_COLOR_CENTER_HEADERS_TEXT)
	GUICtrlSetBkColor($Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetTip($Text, "- Info/Prog." & @CRLF & "- STI.", "- Matière:", 1,1)
	;;;========Content============
	Local $TmpTop = $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE / 2
	Local $TmpLeft = $Left_Header + $WidthHeader / 2 + $GUI_MARGE
	Global $lblMatiere = GUICtrlCreateLabel("Matière...", $TmpLeft, $TmpTop, $WidthHeader / 2 - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR + 10, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetFont(-1, 16, 100, 4, "Segoe UI Light")
	GUICtrlSetTip($lblMatiere, "Cliquez ici pour Actualiser tout.", "Actualiser.", 0,1)

	;;;***** Matière Info || Prog
	$TmpTop = $TmpTop + $GUI_HEADER_HAUTEUR + 2 * $GUI_MARGE
	$TmpLeft = $Left_Header + $WidthHeader / 2 + $GUI_MARGE + $GUI_MARGE + $GUI_MARGE
	Global $rInfoProg = GUICtrlCreateRadio("", $TmpLeft, $TmpTop, 10, 20) ;Bouton Radio Sc./M./Tech.
	Global $tInfoProg = GUICtrlCreateLabel("Info/Prog", $TmpLeft + 15, $TmpTop + 3, 65, 20)
	GUICtrlSetColor(-1, 0xEEEEEE)
	GUICtrlSetTip($rInfoProg, "Informatique ou Algorithmique et Programmation.", "Matière", 1,1)
	GUICtrlSetTip($tInfoProg, "Informatique ou Algorithmique et Programmation.", "Matière", 1,1)


	;;;***** Matière STI
	$TmpLeft = $TmpLeft + $WidthHeader / 4
	Global $rTic = GUICtrlCreateRadio("", $TmpLeft, $TmpTop, 10, 20) ;Bouton Radio Sc./M./Tech.
	Global $tTic = GUICtrlCreateLabel("STI", $TmpLeft + 15, $TmpTop + 3, 35, 20)
	GUICtrlSetColor(-1, 0xEEEEEE)
	GUICtrlSetTip($rTic, "STI - Systèmes & Technologies de l'Informatique", "Matière", 1,1)
	GUICtrlSetTip($tTic, "STI - Systèmes & Technologies de l'Informatique", "Matière", 1,1)
;~ 	GUICtrlSetState($rTic, $GUI_DISABLE) ;Voir aussi _InitialParams
;~ 	GUICtrlSetState($tTic, $GUI_DISABLE) ;Voir aussi _InitialParams


	;;;Matière - Fin====================================================================================

	;;;Contenu des Dossiers Bac*2* - Début=============================================================================================
	;;;========Header============
	$Top_Header = $Top_Header + $GUI_HEADER_HAUTEUR + 2 * $GUI_MARGE + $GUI_Hateur_Label + 2 * $GUI_MARGE
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $GUI_HEADER_HAUTEUR)
	GUICtrlSetColor($Header, 0x696A65)
	GUICtrlSetState($Header, $GUI_DISABLE)
	Global $HeadContenuDossiersBac = GUICtrlCreateLabel("Contenu des Dossiers Bac*2* :", $Left_Header + $GUI_MARGE, $Top_Header + 3, $WidthHeader - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR - 6)
	GUICtrlSetColor($HeadContenuDossiersBac, $GUI_COLOR_CENTER_HEADERS_TEXT)
	GUICtrlSetBkColor($HeadContenuDossiersBac, $GUI_BKCOLOR_TRANSPARENT)
	;;;========Content============
;~ 	GUICtrlCreateListView ( "text", left, top [, width [, height [, style = -1 [, exStyle = -1]]]] )
	Local $TmpListLeft = $Left_Header + $GUI_MARGE
	Local $TmpListTop = $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE
	Local $TmpListWidth = $WidthHeader - 2 * $GUI_MARGE
	Local $TmpListHeight = 125 - 2 * $GUI_MARGE
	Global $GUI_AfficherLeContenuDesDossiersBac = GUICtrlCreateListView("Dossiers & Fichiers                   |Taille|[Créé à]  [Modif. à]|FullPath", $TmpListLeft, $TmpListTop, $TmpListWidth, $TmpListHeight, $LVS_REPORT, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))
	_GUICtrlListView_SetColumnWidth($GUI_AfficherLeContenuDesDossiersBac, 0, 185)
	_GUICtrlListView_SetColumnWidth($GUI_AfficherLeContenuDesDossiersBac, 1, 40)
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesDossiersBac, 3) ;Hide FullPath column
	;;;Contenu des Dossiers Bac*20* - Fin=============================================================================================


	;;;Contenu des Autres Dossiers  - Début=============================================================================================
	;;;========Header============
	$Top_Header = $TmpListTop + $TmpListHeight + 1.5 * $GUI_MARGE
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $GUI_HEADER_HAUTEUR)
	GUICtrlSetColor($Header, 0x696A65)
	GUICtrlSetState($Header, $GUI_DISABLE)
	Global $HeadContenuAutresDossiers = GUICtrlCreateLabel("Autres Fichiers dans ""Mes Documents"", ""Bureau"", ""C:\"" && ""D:\"":", $Left_Header + $GUI_MARGE, $Top_Header + 3, $WidthHeader - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR - 6)
	GUICtrlSetColor($HeadContenuAutresDossiers, $GUI_COLOR_CENTER_HEADERS_TEXT)
	GUICtrlSetBkColor($HeadContenuAutresDossiers, $GUI_BKCOLOR_TRANSPARENT)
	;;;========Content============
	Local $TmpListLeft = $Left_Header + $GUI_MARGE
	Local $TmpListTop = $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE
	Local $TmpListWidth = $WidthHeader - 2 * $GUI_MARGE
	Local $TmpListHeight = 125 - 2 * $GUI_MARGE
	Global $GUI_AfficherLeContenuDesAutresDossiers = GUICtrlCreateListView("Dossiers & Fichiers                   |Taille|[Créé à]  [Modif. à]|FullPath", $TmpListLeft, $TmpListTop, $TmpListWidth, $TmpListHeight, $LVS_REPORT, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT))
	_GUICtrlListView_SetColumnWidth($GUI_AfficherLeContenuDesAutresDossiers, 0, 185)
	_GUICtrlListView_SetColumnWidth($GUI_AfficherLeContenuDesAutresDossiers, 1, 40)
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesAutresDossiers, 3)
	;;;Contenu des Autres Dossiers  - Fin=============================================================================================

	;;;Log  - Début=============================================================================================
	$TmpListLeft = $Left_Header + 1 ;+ $GUI_MARGE
	$TmpListTop = $TmpListTop + $TmpListHeight + 1.5 * $GUI_MARGE ;$Top_Header; + $GUI_HEADER_HAUTEUR ;+ $GUI_MARGE
	$TmpListWidth = $WidthHeader - 2 ;- 2 * $GUI_MARGE
	$TmpListHeight = 165 + $GUI_HEADER_HAUTEUR

	Global $GUI_Log = _GUICtrlRichEdit_Create($hMainGUI, "", $TmpListLeft, $TmpListTop, $TmpListWidth, $TmpListHeight, BitOR($ES_MULTILINE, $ES_READONLY, $WS_VSCROLL, $WS_HSCROLL, $ES_AUTOVSCROLL))
	Local $hToolTip = _GUIToolTip_Create(0, BitOR($_TT_ghTTDefaultStyle, $TTS_BALLOON)) ; default style tooltip
	_GUIToolTip_AddTool($hToolTip, 0, "Journale des opérations (Double-clic pour l'ouvrir avec Wordpad)", $GUI_Log) ; Multiline ToolTip
	;~ GUICtrlSetTip($GUI_Log, "- Double-clic pour ouvrir ce journal avec Wordpad", "- Journale des opérations", 0,1)
	_GUICtrlRichEdit_SetEventMask($GUI_Log, $ENM_MOUSEEVENTS)
	_GUICtrlRichEdit_SetBkColor($GUI_Log, 0x080808)
	_GUICtrlRichEdit_SetSel($GUI_Log, 0, -1, True) ; select all
	_GUICtrlRichEdit_SetFont($GUI_Log, 8, "Tahoma") ;, "Consolas")
	_GUICtrlRichEdit_Deselect($GUI_Log) ; deselect all

	;;;=======ContextMenu==========

	;;;Log  - Fin=============================================================================================


	;;;Partie à gauche  - Début=============================================================================================

	Global $lblComputerID = GUICtrlCreateLabel(_GetUUID(), 1, 1, $GUI_LARGEUR_PARTIE - 4)
	GUICtrlSetColor($lblComputerID, 0xFFFFFF)
	GUICtrlSetBkColor($lblComputerID, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 8, 100, 0, "Segoe UI Light")
	GUICtrlSetTip($lblComputerID, "L'Identifiant Unique de ce PC." & @CRLF & @CRLF & "Cliquer pour copier.", GUICtrlRead($lblComputerID), 0,1)

	Local $TmpButtonWidth = $GUI_LARGEUR_PARTIE - 4 * $GUI_MARGE ; * 2/3
	Local $TmpButtonHeight = $GUI_HEADER_HAUTEUR * 2
	Global $lblBac = GUICtrlCreateLabel("", $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $GUI_MARGE, $TmpButtonWidth, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($lblBac, 0xFFFFFF)
	GUICtrlSetBkColor($lblBac, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 18, 100, 4, "Segoe UI Light")
	GUICtrlSetTip($lblBac, "Pour les matières Informatique & Algorithmique et programmation", "Dossier de travail du Candidat", 1,1)

	Local $TmpButtonWidth = $GUI_LARGEUR_PARTIE - 4 * $GUI_MARGE ; * 2/3
	Local $TmpButtonHeight = $GUI_HEADER_HAUTEUR * 2
	Global $bRecuperer = GUICtrlCreateButton("Récupérer", $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $TmpButtonHeight + 2 * $GUI_MARGE, $TmpButtonWidth, $TmpButtonHeight)
	GUICtrlSetColor(-1, 0xffffff)
	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
	GUICtrlSetFont(-1, 10)
	GUICtrlSetTip($bRecuperer, @CRLF & "Cette commande permet de:" & @CRLF 	& @CRLF & "1. Sauvegarder le travail du candidat vers un dossier verrouillé sur ce PC." & @CRLF & "2. Copier les dossiers & fichiers du candidat vers la clé USB." & @CRLF & "3. Supprimer les travaux du candidat, pour les matières Info & Prog.", "Copier les fichiers du candidat vers la clé USB", 0,1)

	$Top_Header = 4 * $GUI_MARGE + 2 * $TmpButtonHeight
	$Left_Header = $GUI_MARGE
	$WidthHeader = $GUI_LARGEUR_PARTIE - 2 * $GUI_MARGE
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $GUI_HAUTEUR - $Top_Header - $GUI_MARGE)
	GUICtrlSetColor($Header, 0x66c7fc)
	GUICtrlSetState($Header, $GUI_DISABLE)

	Local $TmpHeaderHauteur = $GUI_HEADER_HAUTEUR + 6
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $TmpHeaderHauteur)
	GUICtrlSetColor($Header, 0x66c7fc)
	GUICtrlSetState($Header, $GUI_DISABLE)

	$Text = GUICtrlCreateLabel("Dossiers Récupérés", $Left_Header + $GUI_MARGE, $Top_Header + 3, $WidthHeader - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($Text, 0xFFFFFF)
	GUICtrlSetBkColor($Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 9.5, 500)
	GUICtrlSetTip($Text, "Veuillez créer un dossier pour chaque candidat absent, et y mettre un sous-dossier intitulé ""Absent""", "Dossiers récupérés", 0,1)

	Global $TextDossiersRecuperes = GUICtrlCreateLabel("", $Left_Header + $GUI_MARGE, $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE + 6, $WidthHeader - 2 * $GUI_MARGE, $GUI_HAUTEUR - 310)
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 10, 200)
	GUICtrlSetTip($TextDossiersRecuperes, "Veuillez créer un dossier pour chaque candidat absent, et y mettre un sous-dossier intitulé ""Absent""", "Dossiers récupérés", 0,1)

	$Top_Header = $GUI_HAUTEUR - 100 - $GUI_MARGE
	$Left_Header = $GUI_MARGE
	$WidthHeader = $GUI_LARGEUR_PARTIE - 2 * $GUI_MARGE

	Global $TextApps_Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $TmpHeaderHauteur)
	GUICtrlSetColor($TextApps_Header, 0x66c7fc)
	GUICtrlSetState($TextApps_Header, BitOR($GUI_HIDE, $GUI_DISABLE))

	Global $TextApps_Text = GUICtrlCreateLabel("Logiciels Ouverts", $Left_Header + $GUI_MARGE, $Top_Header + 3, $WidthHeader - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($TextApps_Text, 0xFFFFFF)
	GUICtrlSetBkColor($TextApps_Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 9.5, 500)
	GUICtrlSetState($TextApps_Text, $GUI_HIDE)
	GUICtrlSetTip(-1, @CRLF & "Veuillez sauvegarder le travail et quitter ces applications." & @CRLF & @CRLF & "Certaines applications (VSCode, Sublime Text...) ne demandent pas de confirmation de fermeture," & @CRLF & "même si des fichiers ne sont pas encore enregistrés:", "Applications à fermer", 0,1)

	Global $TextApps = GUICtrlCreateLabel("", $Left_Header + $GUI_MARGE, $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE + 6, $WidthHeader - 2 * $GUI_MARGE, 100 - $GUI_HEADER_HAUTEUR - 2 * $GUI_MARGE)
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 9)
	GUICtrlSetState($TextApps, $GUI_HIDE)
	GUICtrlSetTip(-1, @CRLF & "Veuillez sauvegarder le travail et de quitter ces applications." & @CRLF & @CRLF & "Certaines applications (VSCode, Sublime Text...) ne demandent pas de confirmation de fermeture," & @CRLF & "même si des fichiers ne sont pas encore enregistrés:", "Applications à fermer", 0,1)
	;;;Partie à gauche  - Fin=============================================================================================

	;;;Partie à Droite  - Début=============================================================================================
	Local $GuiTmpLeft = 3 * $GUI_LARGEUR_PARTIE
	Local $GuiTmpTop = $GUI_MARGE
	Global $lblLabo = GUICtrlCreateLabel("", $GuiTmpLeft + $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $GuiTmpTop, $TmpButtonWidth, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($lblLabo, 0xFFFFFF)
	GUICtrlSetBkColor($lblLabo, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 18, 100, 4, "Segoe UI Light")
	GUICtrlSetTip($lblLabo, " ", "Laboratoire d'Informatique", 0,1)

	$GuiTmpTop = $GuiTmpTop + $GUI_MARGE + $TmpButtonHeight
	Global $lblSeance = GUICtrlCreateLabel("", $GuiTmpLeft + $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $GuiTmpTop, $TmpButtonWidth, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($lblSeance, 0xFFFFFF)
	GUICtrlSetBkColor($lblSeance, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 18, 100, 4, "Segoe UI Light")
	GUICtrlSetTip($lblSeance, " ", "Numéro de la Séance", 1,1)


	$GuiTmpLeft = ($GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2) + 3 * $GUI_LARGEUR_PARTIE
	$GuiTmpTop = $GuiTmpTop + $GUI_MARGE + $TmpButtonHeight
;~ 	Global $bCreerSauvegarde = GUICtrlCreateButton("Créer Sauvegarde sur PC", $GuiTmpLeft, $GuiTmpTop, $TmpButtonWidth, $TmpButtonHeight)
	Global $bCreerSauvegarde = GUICtrlCreateButton("Sauve sur SERVEUR", $GuiTmpLeft, $GuiTmpTop, $TmpButtonWidth, $TmpButtonHeight)
	GUICtrlSetColor(-1, 0xffffff)
	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
	GUICtrlSetFont(-1, 10)
	GUICtrlSetTip($bCreerSauvegarde, @CRLF & "Cette commande permet de:" & @CRLF & @CRLF & "1. Sauvegarder les travaux des candidats dans un dossier verrouillé." & @CRLF & "2. Générer un état sur les travaux des candidats." & @CRLF & "3. Générer la grille d'évaluation correspondante (Excel).", "Sauvegarder les travaux sur serveur/poste réserve", 0,1)

	$GuiTmpLeft = $GuiTmpLeft
	$GuiTmpTop = $GuiTmpTop + $GUI_MARGE + $TmpButtonHeight
	Global $bOpenBackupFldr = GUICtrlCreateButton("Ouvrir Dossier de Sauve", $GuiTmpLeft, $GuiTmpTop, $TmpButtonWidth, $TmpButtonHeight)
	GUICtrlSetColor(-1, 0xffffff)
	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
	GUICtrlSetFont(-1, 10)
	GUICtrlSetTip($bOpenBackupFldr, @CRLF & " ", "Ouvrir le dossier de sauvegarde", 0,1)

	Global $bAide = GUICtrlCreateButton("Aide", $GuiTmpLeft, $GUI_HAUTEUR - $GUI_MARGE - $TmpButtonHeight, $TmpButtonWidth, $TmpButtonHeight)
	GUICtrlSetColor(-1, 0xffffff)
	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
	GUICtrlSetFont(-1, 10)
	GUICtrlSetTip($bAide, " ", "Aide & Fonctionnalités de BacCollector.", 1,1)

	Global $lblBacBackup = GUICtrlCreateLabel("BacBackup est installé", 3 * $GUI_LARGEUR_PARTIE + $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $GUI_HAUTEUR - 2 * $GUI_MARGE - 2 * $TmpButtonHeight, $TmpButtonWidth, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($lblBacBackup, 0xFFFFFF)
	GUICtrlSetBkColor($lblBacBackup, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 11, 100, 0, "Segoe UI Light")
;~ 	GUICtrlSetState($lblBacBackup, $GUI_HIDE)
	GUICtrlSetTip($lblBacBackup, "Cliquez ici pour ouvrir BacBackup", "BacBackup surveille le PC.", 0,1)

	;;; Cadre
	$Top_Header = $GuiTmpTop + $TmpButtonHeight + 2 * $GUI_MARGE
	$Left_Header = $GUI_MARGE + 3 * $GUI_LARGEUR_PARTIE
	$WidthHeader = $GUI_LARGEUR_PARTIE - 2 * $GUI_MARGE
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $GUI_HAUTEUR - $Top_Header - 2 * $GUI_MARGE - $TmpButtonHeight)
	GUICtrlSetColor($Header, 0x66c7fc)
	GUICtrlSetState($Header, $GUI_DISABLE)
	;;; Paramètres
	$Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $TmpHeaderHauteur)
	GUICtrlSetColor($Header, 0x66c7fc)
	GUICtrlSetState($Header, $GUI_DISABLE)

	$Text = GUICtrlCreateLabel("Paramètres", $Left_Header + $GUI_MARGE, $Top_Header + 3, $WidthHeader - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($Text, 0xFFFFFF)
	GUICtrlSetBkColor($Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 9.5, 500)
	GUICtrlSetTip(-1, "Veuillez les mettre à jour chaque séance", "Paramètres de BacCollector.", 1,1)
	;;; Bac:
	Local $TmpLeft = $Left_Header + $GUI_MARGE
	Local $TmpTop = $Top_Header + $TmpHeaderHauteur + $GUI_MARGE
	Local $TmpWidth = $WidthHeader / 2 - 2 * $GUI_MARGE
	Local $TmpHeight = $GUI_HEADER_HAUTEUR
	$Text = GUICtrlCreateLabel("Bac :", $TmpLeft, $TmpTop + 4, $TmpWidth, $TmpHeight, $SS_RIGHT) ; , BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($Text, 0xFFFFFF)
	GUICtrlSetBkColor($Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 10, 550)
	GUICtrlSetTip(-1, "Cette valeur est utilisée dans le nom du dossier de travail des candidats", "Baccalauréat", 1,1)

	Global $cBac = GUICtrlCreateCombo($ANNEES_BAC[0], $TmpLeft + $TmpWidth + $GUI_MARGE, $TmpTop, $TmpWidth + $GUI_MARGE, $TmpHeight, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_DROPDOWNLIST))
	GUICtrlSetData(-1, _ArrayToString ($ANNEES_BAC, "|", 1))
	GUICtrlSetFont(-1, 9, 500)
	GUICtrlSetTip(-1, "Cette valeur est utilisée dans le nom du dossier de travail des candidats", "Baccalauréat", 1,1)
	;;; Séance:
	$TmpLeft = $Left_Header + $GUI_MARGE
	$TmpTop = $TmpTop + $TmpHeight + $GUI_MARGE ;2 * $GUI_MARGE
	$TmpHeight = $GUI_HEADER_HAUTEUR
	$Text = GUICtrlCreateLabel("Séance :", $TmpLeft, $TmpTop + 4, $TmpWidth, $TmpHeight, $SS_RIGHT) ; , BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($Text, 0xFFFFFF)
	GUICtrlSetBkColor($Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 10, 550)
	GUICtrlSetTip(-1, "Valeur utilisée dans la grille d'évaluation générée par BacCollector", "Numéro de la séance", 1,1)

	Global $cSeance = GUICtrlCreateCombo("Séance-1", $TmpLeft + $TmpWidth + $GUI_MARGE, $TmpTop, $TmpWidth + $GUI_MARGE, $TmpHeight, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_DROPDOWNLIST))
	GUICtrlSetData(-1, "Séance-2|Séance-3|Séance-4|Séance-5|Séance-6")
	GUICtrlSetFont(-1, 9, 500)
	GUICtrlSetTip(-1, "Valeur utilisée dans la grille d'évaluation générée par BacCollector", "Numéro de la séance", 1,1)

	;;; Labo:
	$TmpLeft = $Left_Header + $GUI_MARGE
	$TmpTop = $TmpTop + $TmpHeight + $GUI_MARGE ;2 * $GUI_MARGE
	$TmpHeight = $GUI_HEADER_HAUTEUR
	$Text = GUICtrlCreateLabel("Labo :", $TmpLeft, $TmpTop + 4, $TmpWidth, $TmpHeight, $SS_RIGHT) ; , BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($Text, 0xFFFFFF)
	GUICtrlSetBkColor($Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 10, 550)
	GUICtrlSetTip(-1, "Valeur utilisée comme nom de la Clé USB," & @CRLF & "et elle est utilisée dans la grille d'évaluation générée par BacCollector", "Laboratoire", 1,1)

	Global $cLabo = GUICtrlCreateCombo("Labo-1", $TmpLeft + $TmpWidth + $GUI_MARGE, $TmpTop, $TmpWidth + $GUI_MARGE, $TmpHeight, BitOR($GUI_SS_DEFAULT_COMBO, $CBS_DROPDOWNLIST))
	GUICtrlSetData(-1, "Labo-2|Labo-3|Labo-4|Labo-5|Labo-6|Labo-7")
	GUICtrlSetFont(-1, 9, 500)
	GUICtrlSetTip(-1, "Valeur utilisée comme nom de la Clé USB," & @CRLF & "et elle est utilisée dans la grille d'évaluation générée par BacCollector", "Laboratoire", 1,1)

	GUICtrlSetState($bRecuperer, $GUI_FOCUS)

	;;;Partie à Droite  - Fin=============================================================================================
	GUISetState(@SW_SHOW, $hMainGUI)

EndFunc   ;==>_CreateGui
#EndRegion Function "_CreateGui" -----------------------------------------------------------------------------

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func Recuperer()
	Local $NumeroCandidat = GUICtrlRead($GUI_NumeroCandidat)
	_ClearGuiLog()
	_Logging("______", 2, 0)
	_Logging("Début de la récupération...", 4, 1)

	;;;_Logging($sText, $iSuccess = 1, $iGuiLog = 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>Info&Blanc, 5>Info&Red

	;===========================================================
;~ //Début Vérification de la validité du format du numéro d'inscription du candidat
	If StringRegExp($NumeroCandidat, "([0-9]{6})", 0) = 0 Or $NumeroCandidat = "000000" Then
		_Logging("""" & $NumeroCandidat & """ n'est pas un numéro d'inscription valide", 5, 1)
		_Logging("Récupération annulée", 5, 1)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, """" & $NumeroCandidat & """ n'est pas un numéro d'inscription valide", 0)
		Return 0
	EndIf
;~ //Fin Vérification de la validité du format du numéro d'inscription du candidat


	_Logging("Numéro d'inscription du candidat """ & $NumeroCandidat & """", 2, 0)
	_Logging("Matière """ & $Matiere & """", 2, 0)
	_Logging("Lecture et mise à jour des paramètres de " & $PROG_TITLE, 2, 0)
	_SaveParams()
	_InitialParams()

	;===========================================================
;~    Début-Vérification des logiciels Ouverts
	_Logging("Recherche d'applications ouvertes...", 2, 0)
	Local $sApps = _ListeDApplicationsOuvertes()
;~ 	WinActivate ($hMainGUI)
	If $sApps <> "" Then
		Local $TmpSpaces = "                                           "
		_Logging("Applications ouvertes: " & StringReplace($sApps, @CRLF, @CRLF & $TmpSpaces), 5, 1)
;~ 		_Logging("Applications ouvertes: " & StringReplace($sApps, @CRLF, " "), 2, 0)
		;;;_Logging($sText, $iSuccess = 1, $iGuiLog = 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>Info&Blanc, 5>Info&Red
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		GUICtrlSetState($TextApps_Header, BitOR($GUI_SHOW, $GUI_DISABLE))
		GUICtrlSetState($TextApps_Text, $GUI_SHOW)
		GUICtrlSetState($TextApps, $GUI_SHOW)
		GUICtrlSetData($TextApps, $sApps)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Avant de ""Récupérer"" le travail du candidat, veuillez fermer ce(s) logiciel(s) : " & @CRLF & @CRLF & $sApps, 0)
		Return 1
	Else
;~ 		_Logging("Aucune application ouverte", 2, 0)
		GUICtrlSetState($TextApps_Header, BitOR($GUI_HIDE, $GUI_DISABLE))
		GUICtrlSetState($TextApps_Text, $GUI_HIDE)
		GUICtrlSetState($TextApps, $GUI_HIDE)
		GUICtrlSetData($TextApps, "")
	EndIf
;~    Fin-Vérification de logiciels Ouverts

	If $Matiere = "InfoProg" Then
		RecupererInfo($NumeroCandidat)
	Else
		RecupererTic($NumeroCandidat)
	EndIf
	Return 2
EndFunc   ;==>Recuperer

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func RecupererInfo($NumeroCandidat)

	;===========================================================
;~    Début-Vérification de l'existence d'un Dossier Bac*20*
	SplashTextOn("Sans Titre", "Recherche des dossiers ""Bac*20*"" sous la racine ""C:""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("Recherche des dossiers ""Bac*20*"" sous la racine ""C:""", 2, 0)
	Local $Bac = DossiersBac() ;
;~ 	WinActivate ($hMainGUI)

	If $Bac[0] = 0 Then
		SplashOff()
		_Logging("Aucun dossier ""Bac*20*"" trouvé sous la racine ""C:""", 5, 1)
;~ 		_Logging("Aucun dossier ""Bac*20*"" trouvé sous la racine ""C:""", 2, 0)
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Aucun dossier ""Bac20xx"" sous la racine ""C:"" !!!" & @CRLF & "", 0)
		Return
	EndIf
;~    Fin-Vérification de l'existence d'un Dossier Bac*20*

	;===========================================================
;~    Début-Vérification si au moins l'un des dossier Bac*20* n'est pas vide (Vide=aucun Fichier)
	SplashTextOn("Sans Titre", "Vérification des dossiers ""Bac*20*"" sous la racine ""C:""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("Vérification du contenue du(es) dossier(s) : " & _ArrayToString($Bac, ", ", 1), 2, 0) ;

	Local $NbDossiersBacNonVides = 0
	Local $NomsDossiers = ""
	For $i = 1 To $Bac[0]
		$NomsDossiers = $NomsDossiers & " """ & $Bac[$i] & """"
		$sizefldr1 = DirGetSize($Bac[$i], 1)
		If Not @error And $sizefldr1[1] Then
			$NbDossiersBacNonVides += 1
		EndIf
	Next
;~ 	WinActivate ($hMainGUI)
	If Not $NbDossiersBacNonVides Then
		_Logging("Aucun fichier trouvé dans le(s) dossier(s) : " & _ArrayToString($Bac, ",", 1), 5, 1) ;
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		SplashOff()
		If $Bac[0] = 1 Then
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Le dossier " & $NomsDossiers & " ne contient aucun fichier à Récupérer !!!" & @CRLF _
					 & "Si le candidat n'a rien fait, ou qu'il a enregistré sont travail en déhors de ce dossier," & @CRLF _
					 & "veuillez y créer un fichier texte contenant la remarque suivante: " & @CRLF _
					 & @TAB & """Le dossier est vide""" & @CRLF _
					 & "puis cliquez à nouveau sur le bouton ""Récupérer""", 0)
		Else
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Les dossiers " & $NomsDossiers & " ne contiennent aucun fichier à Récupérer !!!" & @CRLF _
					 & "Si le candidat n'a rien fait,  ou qu'il a enregistré sont travail en déhors de ces dossiers," & @CRLF _
					 & "veuillez lui créer un fichier texte, dans l'un des dossiers ""Bac*20*"", contenant la remarque suivante: " & @CRLF _
					 & @TAB & """Le dossier est vide""" & @CRLF _
					 & "puis cliquez à nouveau sur le bouton ""Récupérer""", 0)
		EndIf
		Return
	EndIf
;~    Fin-Vérification si au moins l'un des dossier Bac*20* n'est pas vide

	;===========================================================
;~ //Début  Vérif l'existence du dossier dans le Flash
	_Logging("Vérification si un dossier de même nom """ & $NumeroCandidat & """ existe déjà dans la Clé Usb", 2, 0) ;
	SplashTextOn("Sans Titre", "Vérification si un dossier de même nom """ & $NumeroCandidat & """ existe déjà dans la Clé Usb. " & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	$ScriptDir = @ScriptDir
	If StringRight($ScriptDir, 1) <> "\" Then
		$ScriptDir = $ScriptDir & "\"
	EndIf

	Global $Dest1FlashUSB = $ScriptDir & $NumeroCandidat
	Global $Dest2LocalFldr = $Lecteur & $DossierSauve & "\" & $DossierBacCollector & "\" & $NumeroCandidat
	If FileExists($Dest1FlashUSB) Then
		_Logging("Un dossier de même nom """ & $NumeroCandidat & """ existe déjà sur la Clé USB!!", 5, 1) ;
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		SplashOff()
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Un dossier de même nom """ & $NumeroCandidat & """ existe déjà sur la Clé USB!!", 0)
		Return
	EndIf
;~ //Fin  Vérif l'existence du dossier dans le Flash

	SplashTextOn("Sans Titre", "Création des dossiers de Sauvegarde." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	;===========================================================
	Local $Error0_CreationDossier = 1 ;Création du dossier dans le flash 1:Success 0:Failure
	Local $Error1_CopyBacFldr = 1 ;
	Local $Error2_DirRemove = 1 ;
	Local $Error3_DirCreate = 1 ;

	$Error0_CreationDossier = DirCreate($Dest1FlashUSB)
	Local $TmpError = DirCreate($Dest2LocalFldr)

	If $Error0_CreationDossier = 0 Then ;1:Success 0:Failure
		SplashOff()
		_Logging("Création du Dossier: " & $Dest1FlashUSB, $Error0_CreationDossier)
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Échec lors de la crétaion du dosier sur la Clé USB" & @CRLF _
				 & "L'opération de récupération est annulée" & @CRLF _
				 & "Veuillez vérifier si la Clé USB n'est pas pleine ou protégée en écriture," & @CRLF _
				 & "ou si un Antivirus bloque " & $PROG_TITLE & $PROG_VERSION & @CRLF _
				 & "puis relancer l'opération.", 0)
		Return
	EndIf
	Local $iNberreurs = 0
	;;;Log+++++++
	_Logging("Création du Dossier: " & $Dest1FlashUSB)
	_Logging("Création du Dossier: " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
	$iNberreurs = $iNberreurs - ($TmpError - 1)
	;;;Log+++++++
	SplashOff()
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Copie des dossiers...", "", Default, Default, 1)

	For $i = 1 To $Bac[0]

;~ 		SplashTextOn("Sans Titre", "Copie du dossier """ & $Bac[$i] & """." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
		ProgressSet(Round($i / $Bac[0] * 100), "[" & Round($i / $Bac[0] * 100) & "%] " & "Vérif. du dossier : " & StringRegExpReplace($Bac[$i], "^.*\\", ""))
		$Error1_CopyBacFldr = DirCopy($Bac[$i], $Dest1FlashUSB & StringTrimLeft($Bac[$i], 2), $FC_OVERWRITE)
		$TmpError = DirCopy($Bac[$i], $Dest2LocalFldr & StringTrimLeft($Bac[$i], 2), $FC_OVERWRITE)

		If $Error1_CopyBacFldr = 0 Then ;0: error et 1:Success
;~ 			SplashOff()

			ProgressOff()
			_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest2LocalFldr, $Error1_CopyBacFldr)
			_Logging("Récupération annulée", 5, 1)
			_Logging("______", 2, 0)
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Échec lors de la copie du dossier """ & $Bac[$i] & """" & @CRLF _
					 & "L'opération de sauvegarde est annulée", 0)
			Return
		EndIf

		;;;Log+++++++
		_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest1FlashUSB)
		_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
		$iNberreurs = $iNberreurs - ($TmpError - 1)
		;;;Log+++++++
;~ 		SplashTextOn("Sans Titre", "Suppression de """ & $Bac[$i] & """" & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
		ProgressSet(Round($i / $Bac[0] * 100), "[" & Round($i / $Bac[0] * 100) & "%] " & "Suppression du dossier : " & StringRegExpReplace($Bac[$i], "^.*\\", ""))

		If $Error1_CopyBacFldr = 1 Then
			$Error2_DirRemove = DirRemove($Bac[$i], 1)
			;;;Log+++++++
			_Logging("Suppression du dossier """ & $Bac[$i] & """", $Error2_DirRemove) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
			;;;Log+++++++
		EndIf
	Next
	SplashOff()
	Local $sMaskAutresFichiers = "*"
	$iNberreurs = $iNberreurs + _CopierLeContenuDesAutresDossiers($sMaskAutresFichiers)
	ProgressOff()

	Local $Bac20xx = 'C:\Bac' & GUICtrlRead($cBac)
	SplashTextOn("Sans Titre", "Création du Dossier : """ & $Bac20xx & """" & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	$Error3_DirCreate = DirCreate($Bac20xx) ;----------------
	;;;Log+++++++
	_Logging("Création du Dossier : """ & $Bac20xx & """", $Error3_DirCreate) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
	$iNberreurs = $iNberreurs - ($Error3_DirCreate - 1)
	;;;Log+++++++
	SplashTextOn("Sans Titre", "Recherche de BacBackup." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	If _NouvelleSessionBacBackup() Then
		_Logging("Création d'un dossion pour une nouvelle session pour BacBackup") ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
	EndIf

	SplashOff()
	If $iNberreurs = 0 Then
		;;;_Logging($sText, $iSuccess = 1, $iGuiLog = 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green
		_Logging("Récupération terminée  avec succès", 3, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_CENTER, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "La récupération du travail du candidat a été effectuée avec succès!", 0)

	ElseIf $iNberreurs = 1 Then
		_Logging("Récupération terminée, avec une erreur non critique", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "La récupération du travail du candidat est terminée." & @CRLF _
				 & "Une erreur non critique est produite lors de cette opération." & @CRLF _
				 & "Veuillez lire le log et prendre les mesures nécessaires.", 0)
	Else
		_Logging("Récupération terminée, avec " & $iNberreurs & " erreurs non critiques", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "La récupération du travail du candidat est terminée." & @CRLF _
				 & $iNberreurs & " erreurs non critiques sont produites lors de cette opération." & @CRLF _
				 & "Veuillez lire le log et prendre les mesures nécessaires.", 0)
	EndIf

EndFunc   ;==>RecupererInfo

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func RecupererTic($NumeroCandidat)

	Local $iNberreurs = 0
;~ ;===========================================================
;~ ;=================== WWW ===================================
;~ ;===========================================================
;~    Début-Vérification de l'existence d'un Dossier WWW
	Local $YaSiteWeb = 1
	SplashTextOn("Sans Titre", "Recherche des dossiers d'hébergement locaux d'Apache." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("Recherche des dossiers d'hébergement locaux d'Apache.", 2, 0)
	Local $Www = DossiersEasyPHPwww() ;

	Local $NbWebsiteNonVide = 0
	Local $ListeSites[1] = [0]

	If IsArray($Www) = 0 Then
		$YaSiteWeb = 0

		SplashOff()
		_Logging($PROG_TITLE & " ne trouve aucun dossier d'hébergement (www)", 5, 1)
		_Logging("MsgBox: Poursuivre à la recherche des BD (Oui/Non)?", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & $PROG_VERSION, $PROG_TITLE & " ne trouve aucun dossier d'hébergement (www)" & @CRLF & @CRLF _
				 & "Voulez-vous poursuivre à la recherche et récupération de la base de données?" & @CRLF & @CRLF, 0)
		If $Rep = 1 Then
			_Logging("Non", 2, 0)
			_Logging("Récupération annulée.", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
			_Logging("______", 2, 0)
			Return
		EndIf
		_Logging("Oui", 2, 0)
	Else
;~    Fin-Vérification de l'existence d'un Dossier WWW

		;===========================================================
;~    Début-Vérification si au moins l'un des dossier WWW n'est pas vide (Vide=aucun Fichier)
;~ 		SplashTextOn("Sans Titre", "Scan des sites web dans les dossiers d'hébergement locaux d'Apache." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
		_Logging("Scan des sites web dans : " & _ArrayToString($Www, ", ", 1), 2, 0) ;
		SplashOff()
		ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des sites web...", "", Default, Default, 1)

		For $i = 1 To $Www[0]
			ProgressSet(Round($i / $Www[0] * 100), "[" & Round($i / $Www[0] * 100) & "%] ")
			$TmpSites = _FileListToArrayRec($Www[$i], "*||", 2, 0, 2, 2) ;Dossiers, Non Réc, 0, FullPath
			If IsArray($TmpSites) Then
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "wampthemes"))
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "wamplangues"))
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "img")) ;xampp
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "dashboard")) ;xampp
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "webalizer")) ;xampp
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "xampp")) ;xampp
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "forbidden")) ;xampp
				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "restricted")) ;xampp
;~ 				_ArrayDelete($TmpSites, _ArraySearch($TmpSites, $Www[$i] & "\" & "restricted")) ;xampp
				$TmpSites[0] = UBound($TmpSites) - 1
				For $j = 1 To $TmpSites[0]
					$sizefldr1 = DirGetSize($TmpSites[$j], 1)
					If Not @error And $sizefldr1[1] Then
						$NbWebsiteNonVide += 1
					EndIf
				Next
				$ListeSites[0] += $TmpSites[0]
				_ArrayDelete($TmpSites, 0)
				_ArrayAdd($ListeSites, $TmpSites) ;
			EndIf

		Next

		ProgressOff()

		If Not $NbWebsiteNonVide Then
			$YaSiteWeb = 0

			_Logging("Aucun site web valide trouvé dans le(s) dossier(s) : " & _ArrayToString($Www, ",", 1), 5, 1) ;
			_Logging("MsgBox: Poursuivre à la recherche des BD (Oui/Non)?", 2, 0)
;~ 			SplashOff()
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & $PROG_VERSION, "Aucun site web valide trouvé dans le(s) dossier(s) : " & _ArrayToString($Www, ",", 1) & @CRLF & @CRLF _
					 & "Il se peut que le candidat a mis ses fichiers directement sous le dossier d'hébergement ('www' ou 'htdocs')," &  @CRLF _
					 & "si c'est le cas, veuillez lui créer un dossier et y mettre ses fichiers de travail, puis réessayer la récupération." &  @CRLF&  @CRLF _
					 & "Poursuivre la recherche et récupération de la base de données?" & @CRLF & @CRLF, 0)
			If $Rep = 1 Then
				_Logging("Non", 2, 0)
				_Logging("Récupération annulée.", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
				_Logging("______", 2, 0)
				Return
			EndIf
			_Logging("Oui", 2, 0)

		EndIf
	EndIf
;~    Fin-Vérification si au moins l'un des dossier WWW n'est pas vide (Vide=aucun Fichier)


;~ ;===========================================================
;~ ;=================== DATA  =================================
;~ ;===========================================================
;~    Début-Vérification de l'existence d'un Dossier **DATA**
	Local $YaDatabase = 1
	SplashTextOn("Sans Titre", "Recherche des dossiers de stockage de BD MySql." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("Recherche des dossiers de stockage de BD MySql.", 2, 0)
	Local $Data = DossiersEasyPHPdata() ;
	SplashOff()

	If IsArray($Data) = 0 Then
		$YaDatabase = 0
		_Logging($PROG_TITLE & " ne trouve pas le dossier de stockage des Bases de données MySql.", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("MsgBox: Poursuivre la récupération (Oui/Non)?", 2, 0)
		If $YaSiteWeb Then
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & $PROG_VERSION, $PROG_TITLE & " ne trouve pas le dossier de stockage des Bases de données MySql." & @CRLF & @CRLF _
					 & "Voulez-vous poursuivre la récupération des sites web trouvés?" & @CRLF & @CRLF, 0)
			If $Rep = 1 Then
				_Logging("Non", 2, 0)
				_Logging("Récupération annulée.", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
				_Logging("______", 2, 0)
				Return
			EndIf
			_Logging("Oui", 2, 0)
		Else
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, $PROG_TITLE & " ne trouve pas le dossier de stockage des Bases de données MySql." & @CRLF & @CRLF _
					 & "Récupération annulée.", 0)
			_Logging("Récupération annulée.", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
			_Logging("______", 2, 0)
			Return
		EndIf

	Else
;~    Fin-Vérification de l'existence d'un Dossier DATA

		;===========================================================
;~    Début-Vérification si au moins l'un des dossier DATA n'est pas vide (Vide=aucun Fichier)
;~ 		SplashTextOn("Sans Titre", "Scan des Bases de données dans les dossiers MySql." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
		SplashOff()
		_Logging("Scan des Bases de données dans les dossiers MySql dans : " & _ArrayToString($Data, ", ", 1), 2, 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red

		ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers de BD MySql", "", Default, Default, 1)
		Local $ListeBD[1] = [0]

		For $i = 1 To $Data[0]
			ProgressSet(Round($i / $Data[0] * 100), "[" & Round($i / $Data[0] * 100) & "%] ")
			$TmpBD = _FileListToArrayRec($Data[$i], "*||phpmyadmin;mysql;performance_schema;sys;cdcol;webauth", 2, 0, 2, 2) ;Dossiers, Non Réc, 0, FullPath

			If IsArray($TmpBD) Then
				_ArrayDelete($TmpBD, _ArraySearch($TmpBD, $Data[$i] & "\" & "phpmyadmin"))
				_ArrayDelete($TmpBD, _ArraySearch($TmpBD, $Data[$i] & "\" & "mysql"))
				_ArrayDelete($TmpBD, _ArraySearch($TmpBD, $Data[$i] & "\" & "performance_schema"))
				_ArrayDelete($TmpBD, _ArraySearch($TmpBD, $Data[$i] & "\" & "sys"))
				_ArrayDelete($TmpBD, _ArraySearch($TmpBD, $Data[$i] & "\" & "cdcol"))
				_ArrayDelete($TmpBD, _ArraySearch($TmpBD, $Data[$i] & "\" & "webauth"))
				$TmpBD[0] = UBound($TmpBD) - 1
				;Ajout des BD
				$ListeBD[0] += $TmpBD[0]
				_ArrayDelete($TmpBD, 0)
				_ArrayAdd($ListeBD, $TmpBD) ;
				;Ajout de ib0...
				$TmpBD = _FileListToArrayRec($Data[$i], "ibdata*||", 1, 0, 0, 2) ;Fichier, Non Réc, 0, FullPath
				If IsArray($TmpBD) Then
					;Ajout des BD
					$ListeBD[0] += $TmpBD[0]
					_ArrayDelete($TmpBD, 0)
					_ArrayAdd($ListeBD, $TmpBD) ;
				EndIf

			EndIf
		Next

		ProgressOff()
		If IsArray($ListeBD) = 0 Then
			$YaDatabase = 0
			$iNberreurs &= 1
			_Logging("Aucune base de données trouvée dans : " & _ArrayToString($Data, ", ", 1), 5, 1) ;
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			If $YaSiteWeb Then
				_Logging("MsgBox: Poursuivre la récupération (Oui/Non)?", 2, 0)
				Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & $PROG_VERSION, "Aucune base de données trouvée." & @CRLF & @CRLF _
						 & "Voulez-vous poursuivre la récupération?" & @CRLF & @CRLF, 0)
				If $Rep = 1 Then
					_Logging("Non", 2, 0)
					_Logging("Récupération annulée.", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
					_Logging("______", 2, 0)
					Return
				EndIf
				_Logging("Oui", 2, 0)
			Else
				_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
				_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, $PROG_TITLE & " ne trouve pas le dossier de stockage des Bases de données MySql." & @CRLF & @CRLF _
						 & "Récupération annulée.", 0)
				_Logging("Récupération annulée.", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
				_Logging("______", 2, 0)
				Return
			EndIf
		EndIf
;~    Fin-Vérification si au moins l'un des dossier ***DATA*** n'est pas vide (Vide=aucun Fichier)
	EndIf
;~ ;===========================================================
;~ ;============Vérif Dossier dans l'USB Drive ================
;~ ;===========================================================

;~ //Début  Vérif l'existence du dossier dans le Flash
	_Logging("Vérification si un dossier de même nom """ & $NumeroCandidat & """ existe déjà dans la Clé Usb", 2, 0) ;
	SplashTextOn("Sans Titre", "Vérification si un dossier de même nom """ & $NumeroCandidat & """ existe déjà dans la Clé Usb. " & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	$ScriptDir = @ScriptDir
	If StringRight($ScriptDir, 1) <> "\" Then
		$ScriptDir = $ScriptDir & "\"
	EndIf

	Global $Dest1FlashUSB = $ScriptDir & $NumeroCandidat
	Global $Dest2LocalFldr = $Lecteur & $DossierSauve & "\" & $DossierBacCollector & "\" & $NumeroCandidat
	If FileExists($Dest1FlashUSB) Then
		_Logging("Un dossier de même nom """ & $NumeroCandidat & """ existe déjà sur la Clé USB!!", 5, 1) ;
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		SplashOff()
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Un dossier de même nom """ & $NumeroCandidat & """ existe déjà sur la Clé USB!!", 0)
		Return
	EndIf
;~ //Fin  Vérif l'existence du dossier dans le Flash

	SplashTextOn("Sans Titre", "Création des dossiers de Sauvegarde." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	;===========================================================
	Local $Error0_CreationDossier = 1 ;Création du dossier dans le flash 1:Success 0:Failure
	Local $Error1_CopyBacFldr = 1 ;
	Local $Error2_DirRemove = 1 ;
	Local $Error3_DirCreate = 1 ;

	$Error0_CreationDossier = DirCreate($Dest1FlashUSB)
	Local $TmpError = DirCreate($Dest2LocalFldr)

	If $Error0_CreationDossier = 0 Then ;1:Success 0:Failure
		SplashOff()
		_Logging("Création du Dossier: " & $Dest1FlashUSB, $Error0_CreationDossier)
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Échec lors de la crétaion du dossier sur la Clé USB" & @CRLF _
				 & "L'opération de récupération est annulée" & @CRLF _
				 & "Veuillez vérifier si la Clé USB n'est pas pleine ou protégée en écriture," & @CRLF _
				 & "ou si un Antivirus bloque " & $PROG_TITLE & $PROG_VERSION & @CRLF _
				 & "puis relancer l'opération.", 0)
		Return
	EndIf
	;;;Log+++++++
	_Logging("Création du Dossier: " & $Dest1FlashUSB)
	_Logging("Création du Dossier: " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
	$iNberreurs = $iNberreurs - ($TmpError - 1)
	;;;Log+++++++
	Local $KesPos, $KesFldr
;~ ;===========================================================
;~ ;============Copie des Sites Web            ================
;~ ;===========================================================
	SplashOff()
;~ 	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers: [Bac*20*]", "", Default, Default, 1)
;~ 	ProgressSet(Round($n/$Liste[0]*100), "[" & Round($n/$Liste[0]*100) & "%] " & "Vérif. de : " & StringRegExpReplace($File_Name, "^.*\\", ""))

	If $YaSiteWeb Then
		ProgressOn($PROG_TITLE & $PROG_VERSION, "Copie des sites web...", "", Default, Default, 1)
		For $i = 1 To $ListeSites[0]
			ProgressSet(Round($i / $ListeSites[0] * 100), "[" & Round($i / $ListeSites[0] * 100) & "%] " & "Site web """ & StringUpper(StringRegExpReplace($ListeSites[$i], "^.*\\", "")) & """")
;~ 			SplashTextOn("Sans Titre", "Copie du site web """ & StringUpper(StringRegExpReplace($ListeSites[$i], "^.*\\", "")) & """." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
			;;;---------------------------------------------------
			;;;Version 01 donne un chemin complet dans le dossier du Candidat Ex: F:\123456\Program files (86)\Easyphh\www\sw123456
			;$Error1_CopyBacFldr = DirCopy($ListeSites[$i], $Dest1FlashUSB & StringTrimLeft($ListeSites[$i], 2), $FC_OVERWRITE)
			;$TmpError = DirCopy($ListeSites[$i], $Dest2LocalFldr & StringTrimLeft($ListeSites[$i], 2), $FC_OVERWRITE)
			;;;---------------------------------------------------
			$KesPos = StringInStr($ListeSites[$i], "\", 0, -1)
			$KesFldr = StringTrimRight($ListeSites[$i], StringLen($ListeSites[$i]) - $KesPos + 1)
			$KesFldr = StringRegExpReplace($KesFldr, "^.*\\", "")
			$KesFldr &= "\" & StringRegExpReplace($ListeSites[$i], "^.*\\", "")
			$Error1_CopyBacFldr = DirCopy($ListeSites[$i], $Dest1FlashUSB & "\" & $KesFldr, $FC_OVERWRITE)
			$TmpError = DirCopy($ListeSites[$i], $Dest2LocalFldr & "\" & $KesFldr, $FC_OVERWRITE)
			If $Error1_CopyBacFldr = 0 Then ;0: error et 1:Success
				ProgressOff()
;~ 				SplashOff()
				_Logging("Copie du site """ & StringUpper(StringRegExpReplace($ListeSites[$i], "^.*\\", "")) & """ vers " & $Dest2LocalFldr & "\" & $KesFldr, $Error1_CopyBacFldr)
				_Logging("Récupération annulée", 5, 1)
				_Logging("______", 2, 0)
				_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
				_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Échec lors de la copie du dossier """ & $ListeSites[$i] & """" & @CRLF _
						 & "L'opération de sauvegarde est annulée", 0)
				Return
			EndIf
			;;;Log+++++++
			_Logging("Copie du site """ & StringUpper(StringRegExpReplace($ListeSites[$i], "^.*\\", "")) & """ vers " & $Dest1FlashUSB)
			_Logging("Copie du site """ & StringUpper(StringRegExpReplace($ListeSites[$i], "^.*\\", "")) & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++
		Next
	EndIf

;~ ;===========================================================
;~ ;============Copie des Bases de données     ================
;~ ;===========================================================
	If $YaDatabase Then
		ProgressOn($PROG_TITLE & $PROG_VERSION, "Copie des bases de données...", "", Default, Default, 1)
		For $i = 1 To $ListeBD[0]
			$KesPos = StringInStr($ListeBD[$i], "\", 0, -1)
			$KesFldr = StringTrimRight($ListeBD[$i], StringLen($ListeBD[$i]) - $KesPos + 1)
			$KesFldr = StringRegExpReplace($KesFldr, "^.*\\", "")
			$KesFldr &= "\" & StringRegExpReplace($ListeBD[$i], "^.*\\", "")
			$Error1_CopyBacFldr = DirCopy($ListeBD[$i], $Dest1FlashUSB & "\" & $KesFldr, $FC_OVERWRITE)
			$TmpError = DirCopy($ListeBD[$i], $Dest2LocalFldr & "\" & $KesFldr, $FC_OVERWRITE)

			If FileGetAttrib($ListeBD[$i]) = 'D' Then
				ProgressSet(Round($i / $ListeBD[0] * 100), "[" & Round($i / $ListeBD[0] * 100) & "%] " & "Base de données """ & StringUpper(StringRegExpReplace($ListeBD[$i], "^.*\\", "")) & """")
;~ 				SplashTextOn("Sans Titre", "Copie de la BD """ & StringUpper(StringRegExpReplace($ListeBD[$i], "^.*\\", "")) & """." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
				$Error1_CopyBacFldr = DirCopy($ListeBD[$i], $Dest1FlashUSB & "\" & $KesFldr, $FC_OVERWRITE)
				$TmpError = DirCopy($ListeBD[$i], $Dest2LocalFldr & "\" & $KesFldr, $FC_OVERWRITE)
			Else
				ProgressSet(Round($i / $ListeBD[0] * 100), "[" & Round($i / $ListeBD[0] * 100) & "%] " & "Fichiers """ & StringUpper(StringRegExpReplace($ListeBD[$i], "^.*\\", "")) & """")
;~ 				SplashTextOn("Sans Titre", "Copie du fichier """ & StringUpper(StringRegExpReplace($ListeBD[$i], "^.*\\", "")) & """." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
				$Error1_CopyBacFldr = FileCopy($ListeBD[$i], $Dest1FlashUSB & "\" & $KesFldr, $FC_OVERWRITE + $FC_CREATEPATH)
				$TmpError = FileCopy($ListeBD[$i], $Dest2LocalFldr & "\" & $KesFldr, $FC_OVERWRITE + $FC_CREATEPATH)
			EndIf

			If $Error1_CopyBacFldr = 0 Then ;0: error et 1:Success
				ProgressOff()
;~ 				SplashOff()
				_Logging("Copie de la BD """ & StringUpper(StringRegExpReplace($ListeBD[$i], "^.*\\", "")) & """ vers " & $Dest2LocalFldr, $Error1_CopyBacFldr)
				_Logging("Récupération annulée", 5, 1)
				_Logging("______", 2, 0)
				_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
				_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Échec lors de la copie du dossier """ & $ListeBD[$i] & """" & @CRLF _
						 & "L'opération de sauvegarde est annulée", 0)
				Return
			EndIf
			;;;Log+++++++
			_Logging("Copie de la BD """ & StringUpper(StringRegExpReplace($ListeBD[$i], "^.*\\", "")) & """ vers " & $Dest1FlashUSB)
			_Logging("Copie de la BD """ & StringUpper(StringRegExpReplace($ListeBD[$i], "^.*\\", "")) & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++
		Next
	EndIf

	ProgressOff()
;~ 	SplashOff()
	Local $sMaskAutresFichiers = "*"
	$iNberreurs = $iNberreurs + _CopierLeContenuDesAutresDossiersNoRemove($sMaskAutresFichiers)

	SplashTextOn("Sans Titre", "Recherche de BacBackup." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	If _NouvelleSessionBacBackup() Then
		_Logging("Création d'un dossion pour une nouvelle session pour BacBackup") ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
	EndIf

	SplashOff()
	If $iNberreurs = 0 And $YaDatabase And $YaSiteWeb Then
		;;;_Logging($sText, $iSuccess = 1, $iGuiLog = 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green
		_Logging("Récupération terminée  avec succès", 3, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_CENTER, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "La récupération du travail du candidat a été effectuée avec succès!" & @CRLF, 0)

	Else
		_Logging("Récupération terminée, avec " & $iNberreurs - ($YaDatabase - 1) - ($YaSiteWeb - 1) & " erreurs.", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "La récupération du travail du candidat est terminée." & @CRLF _
				 & $iNberreurs - ($YaDatabase - 1) - ($YaSiteWeb - 1) & " erreurs sont produites lors de cette opération." & @CRLF _
				 & "Veuillez lire attentivement le log et prendre les mesures nécessaires." & @CRLF, 0)
	EndIf

EndFunc   ;==>RecupererTic

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _AfficherLeContenuDesDossiersBac()
;~ 	SplashTextOn("Sans Titre", "Recherche des dossiers ""Bac*20*"" sous la racine ""C:""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers: [Bac*20*]", "", Default, Default, 1)
	_GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesDossiersBac)
	GUICtrlSetTip($GUI_AfficherLeContenuDesDossiersBac, "Double-clic pour afficher l'élément dans l'Explorateur.", "Contenu des dossiers ?Bac*2*?", 1,1)
	Local $NombreDeFichiers = 0
	Local $Bac = DossiersBac() ;
	Local $Liste[1] = [0] ;
	If $Bac[0] <> 0 Then
		For $i = 1 To $Bac[0]
			_ArrayAdd($Liste, $Bac[$i]) ;
			$Liste[0] += 1
			$Kes = _FileListToArrayRec($Bac[$i], "*", 0, 1, 0, 2)
			If IsArray($Kes) Then
				$Liste[0] += $Kes[0]
				_ArrayDelete($Kes, 0)
				_ArrayAdd($Liste, $Kes) ;
			EndIf
		Next
	EndIf
	Local $TmpMsgForLogging = ""
	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
			$File_Name = $Liste[$N]
			ProgressSet(Round($N / ($Liste[0]) * 100), "[" & Round($N / ($Liste[0]) * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($File_Name, "^.*\\", ""))
			If FileGetAttrib($File_Name) = 'D' Then
				If StringInStr(StringTrimLeft($File_Name, 3), "\") = 0 Then
					$KesTmp = DirGetSize($File_Name, 1)
					$File_size = $KesTmp[1] & ' fichier(s)'
				Else
					ContinueLoop
				EndIf
			Else
				$File_size = _FineSize(FileGetSize($File_Name))
				$NombreDeFichiers += 1
			EndIf
			$File_t = FileGetTime($File_Name, 1) ; creation Time
			$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "]"
			$File_t = FileGetTime($File_Name, 0) ; Modif Time
			$File_time &= "   [" & $File_t[3] & ":" & $File_t[4] & "]"
			$item = GUICtrlCreateListViewItem($File_Name & "|" & $File_size & "|" & $File_time & "|" & $File_Name, $GUI_AfficherLeContenuDesDossiersBac)
			$TmpMsgForLogging &= @CRLF & """" & $File_Name & """" & "    (" & $File_size & "-" & $File_time & ")"
			If StringInStr($File_size, ' fichier(s)') Then
				GUICtrlSetColor(-1, 0x0000FF)
			EndIf
		Next
	EndIf
	Local $TmpSpaces = "                            "
	If $TmpMsgForLogging <> "" Then _Logging("Liste de dossiers/fichiers de travail dans ""Bac20xx"" : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)

	ProgressOff()
;~ 	SplashOff()
EndFunc   ;==>_AfficherLeContenuDesDossiersBac

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _AfficherLeContenuDesDossiersTic()

	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers ""www"", ""htdocs""...", "", Default, Default, 1)
;~ 	SplashTextOn("Sans Titre", "Scan des dossiers d'hébergement locaux d'Apache." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	;===================================================================
	;======= Sites Web =================================================
	;===================================================================
	_GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesDossiersBac)
	GUICtrlSetTip($GUI_AfficherLeContenuDesDossiersBac, "Double-clic pour afficher l'élément dans l'Explorateur.", "Contenu des Dossiers d'hébergement des serveurs Web (www & htdocs...)", 1,1)
	Local $NombreDeFichiers = 0
	Local $Www = DossiersEasyPHPwww() ;
	Local $Liste[1] = [0] ;
	Local $WebSiteName, $FileName, $WebSite_Full_Path
	If $Www[0] <> 0 Then
		For $i = 1 To $Www[0]
			ProgressSet(Round($i / $Www[0] * 100), "[" & Round($i / $Www[0] * 100) & "%]")
			$Kes = _FileListToArrayRec($Www[$i], "*||", 2, 0, 2, 2) ;Dossiers, Non Réc, 0, FullPath

			If IsArray($Kes) Then
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "wampthemes"))
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "wamplangues"))
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "img")) ;xampp
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "dashboard")) ;xampp
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "webalizer")) ;xampp
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "xampp")) ;xampp
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "forbidden")) ;xampp
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "restricted")) ;xampp
				_ArrayDelete($Kes, _ArraySearch($Kes, $Www[$i] & "\" & "restricted")) ;xampp
				$Kes[0] = UBound($Kes) - 1
				$Liste[0] += $Kes[0]
				_ArrayDelete($Kes, 0)
				_ArrayAdd($Liste, $Kes) ;
			EndIf
		Next
	EndIf
	Local $TmpMsgForLogging = ""
	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
			$WebSite_Full_Path = $Liste[$N]
			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%]" & "Analyse du site web : " & StringRegExpReplace($WebSite_Full_Path, "^.*\\", ""))
			$Kes = _FileListToArrayRec($WebSite_Full_Path, "*", 1, 1, 2, 1) ;Fichiers, Réc, 2, Relative Path
			$File_t = FileGetTime($WebSite_Full_Path, 1) ; Creation Time
			$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "] "
			$WebSiteName = StringUpper(StringRegExpReplace($WebSite_Full_Path, "^.*\\", ""))

			If Not IsArray($Kes) Then
				$item = GUICtrlCreateListViewItem("[" & $WebSiteName & "]" & "|" & "Site Web" & "|" & $File_time & " Site Web vide" & "|" & $WebSite_Full_Path, $GUI_AfficherLeContenuDesDossiersBac)
				$TmpMsgForLogging &= @CRLF & """" & $WebSite_Full_Path & """" & "    (" & "Site Web" & "-" & $File_time & "Site Web vide )"
				GUICtrlSetColor(-1, 0x0000FF)
			Else
				$item = GUICtrlCreateListViewItem("[" & $WebSiteName & "]" & "|" & "Site Web" & "|" & $File_time & $Kes[0] & " fichier(s)" & "|" & $WebSite_Full_Path, $GUI_AfficherLeContenuDesDossiersBac)
				$TmpMsgForLogging &= @CRLF & """" & $WebSite_Full_Path & """" & "    (" & "Site Web" & "-" & $File_time & $Kes[0] & " fichier(s) )"
				GUICtrlSetColor(-1, 0x0000FF)
				For $i = 1 To $Kes[0]
					ProgressSet(Round($i / $Kes[0] * 100), "[" & Round($i / $Kes[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($Kes[$i], "^.*\\", ""))
					$File_t = FileGetTime($WebSite_Full_Path & "\" & $Kes[$i], 1) ; Creation Time
					$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "] "
					$File_t = FileGetTime($WebSite_Full_Path & "\" & $Kes[$i], 0) ; Modif Time
					$File_time &= "[" & $File_t[3] & ":" & $File_t[4] & "] "
					$FileName = $Kes[$i]
					$FileExtension = StringUpper(StringRegExpReplace($FileName, "^.*\.", ""))
					$item = GUICtrlCreateListViewItem("[" & $WebSiteName & "] " & $FileName & "|" & $FileExtension & "|" & $File_time & "|" & $WebSite_Full_Path & "\" & $Kes[$i], $GUI_AfficherLeContenuDesDossiersBac)
					$TmpMsgForLogging &= @CRLF & """" & $WebSite_Full_Path & "\" & $Kes[$i] & "\" & $FileName & """" & "    (" & $File_time & ")"
				Next
			EndIf
		Next
	EndIf

	Local $TmpSpaces = "                            "
	If $TmpMsgForLogging <> "" Then _Logging("Liste des sites/fichiers web : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)

	;===================================================================
	;======= Bases de Données ==========================================
	;===================================================================
;~ 	SplashTextOn("Sans Titre", "Recherche des Bases de Données MySql." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des bases de données", "[0%] Veuillez patienter un moment, initialisation...", Default, Default, 1)
	_GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesAutresDossiers)
	GUICtrlSetTip($GUI_AfficherLeContenuDesAutresDossiers, "Double-clic pour afficher l'élément dans l'Explorateur.", "Bases de données & Tables", 1,1)

	Local $NombreDeFichiers = 0
	Local $Data = DossiersEasyPHPdata() ;
	Local $Liste[1] = [0] ;
	Local $DB_Name, $Table_Name, $DB_FullPath
	If $Data[0] <> 0 Then
		For $i = 1 To $Data[0]
			ProgressSet(Round($i / $Data[0] * 100), "[" & Round($i / $Data[0] * 100) & "%] ")
			$Kes = _FileListToArrayRec($Data[$i], "*||phpmyadmin;mysql;performance_schema;sys;cdcol;webauth", 2, 0, 2, 2)

			If IsArray($Kes) Then

				_ArrayDelete($Kes, _ArraySearch($Kes, $Data[$i] & "\" & "phpmyadmin"))
				_ArrayDelete($Kes, _ArraySearch($Kes, $Data[$i] & "\" & "mysql"))
				_ArrayDelete($Kes, _ArraySearch($Kes, $Data[$i] & "\" & "performance_schema"))
				_ArrayDelete($Kes, _ArraySearch($Kes, $Data[$i] & "\" & "sys"))
				_ArrayDelete($Kes, _ArraySearch($Kes, $Data[$i] & "\" & "cdcol"))
				_ArrayDelete($Kes, _ArraySearch($Kes, $Data[$i] & "\" & "webauth"))

				$Kes[0] = UBound($Kes) - 1

				$Liste[0] += $Kes[0]
				_ArrayDelete($Kes, 0)
				_ArrayAdd($Liste, $Kes) ;
			EndIf
		Next
	EndIf
	$TmpMsgForLogging = ""
	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
			$DB_FullPath = $Liste[$N]
			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Analyse de la BD : " & StringRegExpReplace($DB_FullPath, "^.*\\", ""))
			$Kes = _FileListToArray($DB_FullPath, "*.frm", 1) ;Les tables

			$File_t = FileGetTime($DB_FullPath, 1) ; Creation Time
			$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "] "
			$DB_Name = StringUpper(StringRegExpReplace($DB_FullPath, "^.*\\", ""))
			If Not IsArray($Kes) Then
				$item = GUICtrlCreateListViewItem("[" & $DB_Name & "]" & "|" & "BD" & "|" & $File_time & "BD vide" & "|" & $DB_FullPath, $GUI_AfficherLeContenuDesAutresDossiers)
				$TmpMsgForLogging &= @CRLF & """" & $DB_FullPath & """" & "    (" & "BD" & "-" & $File_time & "BD vide )"
				GUICtrlSetColor(-1, 0x0000FF)
			Else
				$item = GUICtrlCreateListViewItem("[" & $DB_Name & "]" & "|" & "BD" & "|" & $File_time & $Kes[0] & " table(s)" & "|" & $DB_FullPath, $GUI_AfficherLeContenuDesAutresDossiers)
				$TmpMsgForLogging &= @CRLF & """" & $DB_FullPath & """" & "    (" & "BD" & "-" & $File_time & $Kes[0] & " table(s) )"
				GUICtrlSetColor(-1, 0x0000FF)
				For $i = 1 To $Kes[0]
					ProgressSet(Round($i / $Kes[0] * 100), "[" & Round($i / $Kes[0] * 100) & "%] " & "Analyse de la table : " & StringRegExpReplace($Kes[$i], "^.*\\", ""))
					$File_t = FileGetTime($DB_FullPath & "\" & $Kes[$i], 1) ; Creation Time
					$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "] "
					$File_t = FileGetTime($DB_FullPath & "\" & $Kes[$i], 0) ; Modif Time
					$File_time &= "[" & $File_t[3] & ":" & $File_t[4] & "] "
					$Table_Name = StringUpper(StringTrimRight($Kes[$i], 4))
					$item = GUICtrlCreateListViewItem("[" & $DB_Name & "] " & $Table_Name & "|" & "Table" & "|" & $File_time & "|" & $DB_FullPath & "\" & $Kes[$i], $GUI_AfficherLeContenuDesAutresDossiers)
					$TmpMsgForLogging &= @CRLF & """" & $DB_FullPath & "\" & $Kes[$i] & """" & "    (Table" & " - " & $File_time & ")"
				Next
			EndIf
		Next
	EndIf
	If $TmpMsgForLogging <> "" Then _Logging("Liste de bases de données / tables : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)

	ProgressOff()

;~ 	SplashOff()
EndFunc   ;==>_AfficherLeContenuDesDossiersTic

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, $SubFldr, $RealPath = 0)
	Local $OtherFilesDestFlashUSB = ""
	Local $OtherFilesDestLocalFldr = ""
	Local $FileName = ""

	Local $iNberreurs = 0
	Local $TmpError
	Local $ListeFiltered[1] = [0]
	For $N = 1 To $Liste[0]
		$File_Name = $Liste[$N]
;~ 		ProgressSet(Round($i/$Bac[0]*100), "[" & Round($i/$Bac[0]*100) & "%] " & "Vérif. de : " & StringRegExpReplace($Bac[$i], "^.*\\", ""))
		ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringTrimRight($Liste[$N], 1))
		If AgeDuFichierEnMinutesModification($Dossier & $File_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
			$ListeFiltered[0] += 1
			_ArrayAdd($ListeFiltered, $File_Name)
		EndIf
	Next

	If $ListeFiltered[0] > 0 Then
		$OtherFilesDestFlashUSB = $Dest1FlashUSB & $SubFldr
		$OtherFilesDestLocalFldr = $Dest2LocalFldr & $SubFldr
		$TmpError = DirCreate($OtherFilesDestFlashUSB)
		;;;Log+++++++
		_Logging("Création du Dossier: " & $OtherFilesDestFlashUSB, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
		$iNberreurs = $iNberreurs - ($TmpError - 1)
		;;;Log+++++++

		$TmpError = DirCreate($OtherFilesDestLocalFldr)
		;;;Log+++++++
		_Logging("Création du Dossier: " & $OtherFilesDestLocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
		$iNberreurs = $iNberreurs - ($TmpError - 1)
		;;;Log+++++++
		Local $Kes = ""
		For $N = 1 To $ListeFiltered[0]
			ProgressSet(Round($N / $ListeFiltered[0] * 100), "[" & Round($N / $ListeFiltered[0] * 100) & "%] " & "Copie de : " & StringTrimRight($ListeFiltered[$N], 1))
			$File_Name = $ListeFiltered[$N]
			$TmpError = FileCopy($Dossier & $File_Name, $OtherFilesDestFlashUSB & $File_Name, 8)

			If $RealPath = 1 Then
				$Kes = $Dossier
			Else
				$Kes = $Dossier
			EndIf
			;;;Log+++++++
			_Logging("Copie du Fichier: """ & $Kes & $File_Name & """ vers """ & $OtherFilesDestFlashUSB & $File_Name & """", $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++
			$TmpError = FileCopy($Dossier & $File_Name, $OtherFilesDestLocalFldr & $File_Name, 8)
			;;;Log+++++++
			_Logging("Copie du Fichier: """ & $Kes & $File_Name & """ vers """ & $OtherFilesDestLocalFldr & $File_Name & """", $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++
			ProgressSet(Round($N / $ListeFiltered[0] * 100), "[" & Round($N / $ListeFiltered[0] * 100) & "%] " & "Suppression de : " & StringTrimRight($ListeFiltered[$N], 1))
			$TmpError = FileDelete($Dossier & $File_Name)
			;;;Log+++++++
			_Logging("Suppression du Fichier : """ & $Kes & $File_Name & """", $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++
		Next
	EndIf

	Return $iNberreurs
EndFunc   ;==>_SubCopierLeContenuDesAutresDossiers

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _SubCopierLeContenuDesAutresDossiersNoRemove($Liste, $Dossier, $SubFldr, $RealPath = 0)
	Local $OtherFilesDestFlashUSB = ""
	Local $OtherFilesDestLocalFldr = ""
	Local $FileName = ""
	Local $iNberreurs = 0
	Local $TmpError
	Local $ListeFiltered[1] = [0]
	For $N = 1 To $Liste[0]
		$File_Name = $Liste[$N]
		ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringTrimRight($Liste[$N], 1))
		If AgeDuFichierEnMinutesModification($Dossier & $File_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
			$ListeFiltered[0] += 1
			_ArrayAdd($ListeFiltered, $File_Name)
		EndIf
	Next

	If $ListeFiltered[0] > 0 Then
		$OtherFilesDestFlashUSB = $Dest1FlashUSB & $SubFldr
		$OtherFilesDestLocalFldr = $Dest2LocalFldr & $SubFldr
		$TmpError = DirCreate($OtherFilesDestFlashUSB)
		;;;Log+++++++
		_Logging("Création du Dossier: " & $OtherFilesDestFlashUSB, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
		$iNberreurs = $iNberreurs - ($TmpError - 1)
		;;;Log+++++++

		$TmpError = DirCreate($OtherFilesDestLocalFldr)
		;;;Log+++++++
		_Logging("Création du Dossier: " & $OtherFilesDestLocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
		$iNberreurs = $iNberreurs - ($TmpError - 1)
		;;;Log+++++++
		Local $Kes = ""
		For $N = 1 To $ListeFiltered[0]
			ProgressSet(Round($N / $ListeFiltered[0] * 100), "[" & Round($N / $ListeFiltered[0] * 100) & "%] " & "Copie de : " & StringTrimRight($ListeFiltered[$N], 1))
			$File_Name = $ListeFiltered[$N]
			$TmpError = FileCopy($Dossier & $File_Name, $OtherFilesDestFlashUSB & $File_Name, 8)

			If $RealPath = 1 Then
				$Kes = $Dossier
			Else
				$Kes = $Dossier
			EndIf
			;;;Log+++++++
			_Logging("Copie du Fichier: """ & $Kes & $File_Name & """ vers """ & $OtherFilesDestFlashUSB & $File_Name & """", $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++
			$TmpError = FileCopy($Dossier & $File_Name, $OtherFilesDestLocalFldr & $File_Name, 8)
			;;;Log+++++++
			_Logging("Copie du Fichier: """ & $Kes & $File_Name & """ vers """ & $OtherFilesDestLocalFldr & $File_Name & """", $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++
		Next
	EndIf

	Return $iNberreurs
EndFunc   ;==>_SubCopierLeContenuDesAutresDossiersNoRemove

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _SubCopyFoldersUserFolders($sSrc, $sDest_RelativePath)
	Local $iNberreurs = 0
	Local $Dossier = $sSrc
	Local $iFilesCountInFldr = 0, $iFldrSize = 0, $aFldr_info
	$Liste = _FileListToArrayRec($Dossier, "*||", 30, 0, 2, 1)

	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
;~ 			ProgressSet(Round($n/$Liste[0]*100), "[" & Round($n/$Liste[0]*100) & "%] " & "Vérif. de : " & StringRegExpReplace($Liste[$n], "^.*\\", ""))
			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringTrimRight($Liste[$N], 1))
			$Fldr_Name = $Dossier & StringTrimRight($Liste[$N], 1)
			$Fldr_Name_Relative_Path = $sDest_RelativePath & $Liste[$N]
			If AgeDuFichierEnMinutesCreation($Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
				$aFldr_info = DirGetSize($Fldr_Name, 1)
				$iFldrSize = $aFldr_info[0]
				$iFilesCountInFldr = $aFldr_info[1]
				If $iFilesCountInFldr = 0 And $iFldrSize = 0 Then
					$Error2_DirRemove = DirRemove($Fldr_Name, 1)
					;;;Log+++++++
					_Logging("Suppression du [dossier vide] """ & $Fldr_Name_Relative_Path & """", $Error2_DirRemove) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
					$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
					;;;Log+++++++
				Else
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Copie de : " & StringTrimRight($Liste[$N], 1))
					$Error1_CopyBacFldr = DirCopy($Fldr_Name, $Dest1FlashUSB & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
					$TmpError = DirCopy($Fldr_Name, $Dest2LocalFldr & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)

					;;;Log+++++++
					_Logging("Copie de """ & $Fldr_Name_Relative_Path & """ vers " & $Dest1FlashUSB, $Error1_CopyBacFldr) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
					$iNberreurs = $iNberreurs - ($Error1_CopyBacFldr - 1)

					_Logging("Copie de """ & $Fldr_Name_Relative_Path & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
					$iNberreurs = $iNberreurs - ($TmpError - 1)
					;;;Log+++++++

					If $Error1_CopyBacFldr = 1 Then
						ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Suppression de : " & StringTrimRight($Liste[$N], 1))
						$Error2_DirRemove = DirRemove($Fldr_Name, 1)
						;;;Log+++++++
						_Logging("Suppression du dossier """ & $Fldr_Name_Relative_Path & """", $Error2_DirRemove) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
						$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
						;;;Log+++++++
					EndIf
				EndIf
			EndIf

		Next
	EndIf
	Return $iNberreurs

EndFunc   ;==>_SubCopyFoldersUserFolders

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
Func _SubCopyFoldersUserFoldersNoRemove($sSrc, $sDest_RelativePath)
	Local $iNberreurs = 0
	Local $Dossier = $sSrc
	Local $iFilesCountInFldr = 0, $iFldrSize = 0, $aFldr_info
	$Liste = _FileListToArrayRec($Dossier, "*||", 30, 0, 2, 1)

	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringTrimRight($Liste[$N], 1))
			$Fldr_Name = $Dossier & StringTrimRight($Liste[$N], 1)
			$Fldr_Name_Relative_Path = $sDest_RelativePath & $Liste[$N]
			If AgeDuFichierEnMinutesCreation($Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
				$aFldr_info = DirGetSize($Fldr_Name, 1)
				$iFldrSize = $aFldr_info[0]
				$iFilesCountInFldr = $aFldr_info[1]
				If $iFilesCountInFldr = 0 And $iFldrSize = 0 Then
					$Error2_DirRemove = DirRemove($Fldr_Name, 1)
					;;;Log+++++++
					_Logging("Suppression du [dossier vide] """ & $Fldr_Name_Relative_Path & """", $Error2_DirRemove) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
					$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
					;;;Log+++++++
				Else
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Copie de : " & StringTrimRight($Liste[$N], 1))
					$Error1_CopyBacFldr = DirCopy($Fldr_Name, $Dest1FlashUSB & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
					$TmpError = DirCopy($Fldr_Name, $Dest2LocalFldr & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)

					;;;Log+++++++
					_Logging("Copie de """ & $Fldr_Name_Relative_Path & """ vers " & $Dest1FlashUSB, $Error1_CopyBacFldr) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
					$iNberreurs = $iNberreurs - ($Error1_CopyBacFldr - 1)

					_Logging("Copie de """ & $Fldr_Name_Relative_Path & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
					$iNberreurs = $iNberreurs - ($TmpError - 1)
					;;;Log+++++++
				EndIf
			EndIf

		Next
	EndIf
	Return $iNberreurs

EndFunc   ;==>_SubCopyFoldersUserFoldersNoRemove

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦


Func _CopierLeContenuDesAutresDossiers($Mask)
	Local $iNberreurs = 0

;~ ///////**********************************************************************
;~ ///////**********      Copy - Dossiers créés sur le Bureau           ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche et Copie des dossiers récents sur le" & @CRLF & """Bureau""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	SplashOff()
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Bureau...", "", Default, Default, 1)
;~ 	ProgressSet(Round($i/$Bac[0]*100), "[" & Round($i/$Bac[0]*100) & "%] " & "Vérif. de : " & StringRegExpReplace($Bac[$i], "^.*\\", ""))
	$iNberreurs += _SubCopyFoldersUserFolders(@DesktopDir & "\", "Bureau\")

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers Modifiés sur le Bureau        ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des fichiers récents sur le" & @CRLF & """Bureau""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Bureau...", "", Default, Default, 1)
	Local $Liste[1] = [0] ;
	Local $Dossier = @DesktopDir & '\' ;
	$Liste = _FileListToArrayRec($Dossier, $Mask & "|*.lnk|", 1, 1, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\Bureau\')
	EndIf

;~ ///////**********************************************************************
;~ ///////***** Copy - Dossiers créés dans le dossier de profil utilisateur ****
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des dossiers récents dans" & @CRLF & """Mes documents""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

;~ 	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Profil Utilisateur...", "", Default, Default, 1)
;~ 	$iNberreurs += _SubCopyFoldersUserFolders(@UserProfileDir & "\", "ProfilU\")

;~ ///////**********************************************************************
;~ ///////*** Copy - Fichiers Modifiés dans le dossier de profil utilisateur ***
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des fichiers récents sur le" & @CRLF & """Bureau""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Profil Utilisateur...", "", Default, Default, 1)
	Local $Liste[1] = [0] ;
	Local $Dossier = @UserProfileDir & '\' ;
	$Liste = _FileListToArrayRec($Dossier, "*.py;*.ipynb;*.ui;*.accdb;*.xls*;*.csv;*.doc*;*.ppt*||", 1, 0, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\ProfilU\')
	EndIf

;~ ///////**********************************************************************
;~ ///////**********      Copy - Dossiers créés dans Mes documents      ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des dossiers récents dans" & @CRLF & """Mes documents""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Mes documents...", "", Default, Default, 1)
	$iNberreurs += _SubCopyFoldersUserFolders(@MyDocumentsDir & "\", "Mes documents\")

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers Modifiés dans Mes documents   ********
;~ ///////**********************************************************************

;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des fichiers récents dans" & @CRLF & """Mes documents""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Mes documents...", "", Default, Default, 1)
	;;; Files in MyDocuments Recursive
	Local $Liste[1] = [0] ;
	Local $Dossier = @MyDocumentsDir & '\' ;
	$Liste = _FileListToArrayRec($Dossier, $Mask & "||", 1, 1, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\Mes documents\')
	EndIf


;~ ///////**********************************************************************
;~ ///////**********      Copy Dossiers récemment créés sous les lecteurs    ********
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers>> disques locaux...", "", Default, Default, 1)
	Local $aDrive = DriveGetDrive('FIXED')
	Local $MaskExclude = ""
;~ 	_ArraySort($aDrive,1,1)
	For $i = 1 To $aDrive[0]
		$MaskExclude = ""
		If $aDrive[$i] = "C:" Then
			$MaskExclude = "Program Files*;Users;Windows;Intel;PerfLogs;EasyPhp*;Bac*2*;Documents and Settings;Config.Msi;Recovery;PySchool"
		EndIf

		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; pour Exclure les hdd externes
				And _WinAPI_IsWritable($aDrive[$i]) _
				Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
			$Liste = _FileListToArrayRec($Dossier, "*|" & $MaskExclude & "|", 30, 0, 2, 1)
			If IsArray($Liste) Then
				For $N = 1 To $Liste[0]
					$Fldr_Name = StringTrimRight($Liste[$N], 1)
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & $Fldr_Name)
					If StringRegExp($Fldr_Name, "^(?i)bac\s*\d*2\d*\s*$", $STR_REGEXPMATCH, 1) or AgeDuFichierEnMinutesCreation($Dossier & $Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
						$Fldr_info = DirGetSize($Dossier & $Fldr_Name, 1)
						$Fldr_size = $Fldr_info[0]
						$FldrFilesCount = $Fldr_info[1]
						If $Fldr_size = 0 And $FldrFilesCount = 0 Then
							$Error2_DirRemove = DirRemove($Dossier & $Fldr_Name, 1)
							;;;Log+++++++
							_Logging("Suppression du [dossier vide] """ & $Dossier & $Fldr_Name & """", $Error2_DirRemove) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
							$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
							;;;Log+++++++
						Else
							If $Fldr_size < $TAILLE_MAX_DU_DOSSIER_SOUS_LECTEUR * 1024 * 1024 Then
								$Fldr_Name_Relative_Path = StringLeft($Dossier, 1) & "_2pts" & "\" & $Fldr_Name
								ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Copie de : " & $Fldr_Name)
								$Error1_CopyBacFldr = DirCopy($Dossier & $Fldr_Name, $Dest1FlashUSB & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
								$TmpError = DirCopy($Dossier & $Fldr_Name, $Dest2LocalFldr & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
								;;;Log+++++++
								_Logging("Copie de """ & $Dossier & $Fldr_Name & """ vers " & $Dest1FlashUSB, $Error1_CopyBacFldr) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
								$iNberreurs = $iNberreurs - ($Error1_CopyBacFldr - 1)

								_Logging("Copie de """ & $Dossier & $Fldr_Name & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
								$iNberreurs = $iNberreurs - ($TmpError - 1)
								;;;Log+++++++

								If $Error1_CopyBacFldr = 1 Then
									ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Suppression de : " & $Fldr_Name)
									$Error2_DirRemove = DirRemove($Dossier & $Fldr_Name, 1)
									;;;Log+++++++
									_Logging("Suppression du dossier """ & $Fldr_Name & """", $Error2_DirRemove) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
									$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
									;;;Log+++++++
								EndIf
							EndIf
						EndIf
					EndIf

				Next
			EndIf
		EndIf
	Next

;~ ///////**********************************************************************
;~ ///////*******      Copy - Fichiers créés sur les Lecteurs du hdd     *******
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des disques locaux [dossiers]...", "", Default, Default, 1)
	Local $aDrive = DriveGetDrive('FIXED')
	Local $MaskExclude = "*.bat;*.sys;*.log;*.rtf"
	For $i = 1 To $aDrive[0]
		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; pour Exclure les hdd externes
				And _WinAPI_IsWritable($aDrive[$i]) _
				Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
;~ 			SplashTextOn("Sans Titre", "Recherche/Copie des fichiers récents sous la racine " & $Dossier & "." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

			$Liste = _FileListToArrayRec($Dossier, $Mask & "|" & $MaskExclude & "|", 29, 0, 0, 1)
			If IsArray($Liste) Then
				$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\' & StringLeft($Dossier, 1) & '_2pts\', 1)
			EndIf
		EndIf
	Next



	Return $iNberreurs
EndFunc   ;==>_CopierLeContenuDesAutresDossiers

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _CopierLeContenuDesAutresDossiersNoRemove($Mask)
	Local $iNberreurs = 0
	SplashOff()

;~ ///////**********************************************************************
;~ ///////**********      Copy - Dossiers créés sur le Bureau           ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche et Copie des dossiers récents sur le" & @CRLF & """Bureau""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	SplashOff()
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Bureau...", "", Default, Default, 1)
	$iNberreurs += _SubCopyFoldersUserFoldersNoRemove(@DesktopDir & "\", "Bureau\")

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers Modifiés sur le Bureau        ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des fichiers récents sur le" & @CRLF & """Bureau""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Bureau...", "", Default, Default, 1)
	Local $Liste[1] = [0] ;
	Local $Dossier = @DesktopDir & '\' ;
	$Liste = _FileListToArrayRec($Dossier, $Mask & "|*.lnk|", 1, 1, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiersNoRemove($Liste, $Dossier, '\Bureau\')
	EndIf

;~ ///////**********************************************************************
;~ ///////**********      Copy - Dossiers créés dans Mes documents      ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des dossiers récents dans" & @CRLF & """Mes documents""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Mes documents...", "", Default, Default, 1)
	$iNberreurs += _SubCopyFoldersUserFoldersNoRemove(@MyDocumentsDir & "\", "Mes documents\")

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers Modifiés dans Mes documents   ********
;~ ///////**********************************************************************

;~ 	SplashTextOn("Sans Titre", "Recherche/Copie des fichiers récents dans" & @CRLF & """Mes documents""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Mes documents...", "", Default, Default, 1)
	;;; Files in MyDocuments Recursive
	Local $Liste[1] = [0] ;
	Local $Dossier = @MyDocumentsDir & '\' ;
	$Liste = _FileListToArrayRec($Dossier, $Mask & "||", 1, 1, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiersNoRemove($Liste, $Dossier, '\Mes documents\')
	EndIf


;~ ///////**********************************************************************
;~ ///////**********      Copy Dossiers récemment créés sous les lecteurs    ********
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> disques locaux...", "", Default, Default, 1)
	Local $aDrive = DriveGetDrive('FIXED')
	Local $MaskExclude = ""
;~ 	_ArraySort($aDrive,1,1)
	For $i = 1 To $aDrive[0]
		$MaskExclude = ""
		If $aDrive[$i] = "C:" Then
			$MaskExclude = "Program Files*;Users;Windows;Intel;PerfLogs;EasyPhp*;Bac*2*;Documents and Settings"
		EndIf

		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; pour Exclure les hdd externes
				And _WinAPI_IsWritable($aDrive[$i]) _
				Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
			$Liste = _FileListToArrayRec($Dossier, "*|" & $MaskExclude & "|", 30, 0, 2, 1)
			If IsArray($Liste) Then
				For $N = 1 To $Liste[0]
					$Fldr_Name = StringTrimRight($Liste[$N], 1)
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & $Fldr_Name)
					If StringRegExp($Fldr_Name, "^(?i)bac\s*\d*2\d*\s*$", $STR_REGEXPMATCH, 1) or AgeDuFichierEnMinutesCreation($Dossier & $Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
						$Fldr_info = DirGetSize($Dossier & $Fldr_Name, 1)
						$Fldr_size = $Fldr_info[0]
						$FldrFilesCount = $Fldr_info[1]
						If $Fldr_size = 0 And $FldrFilesCount = 0 Then
							$Error2_DirRemove = DirRemove($Dossier & $Fldr_Name, 1)
							;;;Log+++++++
							_Logging("Suppression du [dossier vide] """ & $Dossier & $Fldr_Name & """", $Error2_DirRemove) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
							$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
							;;;Log+++++++
						Else
							If $Fldr_size < $TAILLE_MAX_DU_DOSSIER_SOUS_LECTEUR * 1024 * 1024 Then
								$Fldr_Name_Relative_Path = StringLeft($Dossier, 1) & "_2pts" & "\" & $Fldr_Name
								ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Copie de : " & $Fldr_Name)
								$Error1_CopyBacFldr = DirCopy($Dossier & $Fldr_Name, $Dest1FlashUSB & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
								$TmpError = DirCopy($Dossier & $Fldr_Name, $Dest2LocalFldr & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
								;;;Log+++++++
								_Logging("Copie de """ & $Dossier & $Fldr_Name & """ vers " & $Dest1FlashUSB, $Error1_CopyBacFldr) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
								$iNberreurs = $iNberreurs - ($Error1_CopyBacFldr - 1)

								_Logging("Copie de """ & $Dossier & $Fldr_Name & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
								$iNberreurs = $iNberreurs - ($TmpError - 1)
								;;;Log+++++++

							EndIf
						EndIf
					EndIf

				Next
			EndIf
		EndIf
	Next

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers créés sur les Lecteurs du hdd     ********
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des disques locaux [dossiers]...", "", Default, Default, 1)
	Local $aDrive = DriveGetDrive('FIXED')
	Local $MaskExclude = "autoexec.bat;config.sys;*.log;*.rtf"
	For $i = 1 To $aDrive[0]
		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; pour Exclure les hdd externes
				And _WinAPI_IsWritable($aDrive[$i]) _
				Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
;~ 			SplashTextOn("Sans Titre", "Recherche/Copie des fichiers récents sous la racine " & $Dossier & "." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

			$Liste = _FileListToArrayRec($Dossier, $Mask & "|" & $MaskExclude & "|", 29, 0, 0, 1)
			If IsArray($Liste) Then
				$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\' & StringLeft($Dossier, 1) & '_2pts\', 1)
			EndIf
		EndIf
	Next


;~ ///////**********************************************************************
;~ ///////**********      Copy Dossiers Bac*2* *********************************
;~ ///////***********  Pour la matière TIC         *****************************
;~ ///////**********************************************************************

	ProgressOn($PROG_TITLE & $PROG_VERSION, "Copie des dossiers Bac*20*...", "", Default, Default, 1)
	Local $Bac = DossiersBac()
	If IsArray($Bac) Then
		For $i = 1 To $Bac[0]

			ProgressSet(Round($i / $Bac[0] * 100), "[" & Round($i / $Bac[0] * 100) & "%] " & "Copie de : " & $Bac[$i])
			$Fldr_info = DirGetSize($Bac[$i], 1)
			$Fldr_size = $Fldr_info[0]
			$FldrFilesCount = $Fldr_info[1]
			If $Fldr_size > 0 Or $FldrFilesCount > 0 Then

;~ 				SplashTextOn("Sans Titre", "Copie du dossier """ & $Bac[$i] & """." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
				$Error1_CopyBacFldr = DirCopy($Bac[$i], $Dest1FlashUSB & StringTrimLeft($Bac[$i], 2), $FC_OVERWRITE)
				$TmpError = DirCopy($Bac[$i], $Dest2LocalFldr & StringTrimLeft($Bac[$i], 2), $FC_OVERWRITE)
				;;;Log+++++++
				$iNberreurs = $iNberreurs - ($Error1_CopyBacFldr - 1)
				_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest1FlashUSB, $Error1_CopyBacFldr)
				_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest2LocalFldr, $TmpError) ; _Logging($sText, $iSuccess = 1, $iGuiLog = 1)
				$iNberreurs = $iNberreurs - ($TmpError - 1)
				;;;Log+++++++

			EndIf
		Next
	EndIf
	ProgressOff()
;~ 	SplashOff()

	Return $iNberreurs
EndFunc   ;==>_CopierLeContenuDesAutresDossiersNoRemove

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _AfficherLeContenuDesAutresDossiers($Mask = "*")
	_GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesAutresDossiers)
	GUICtrlSetTip($GUI_AfficherLeContenuDesAutresDossiers, "- Bureau" & @CRLF & "- Mes documents" & @CRLF & "- Dossier du profil de l'utilisateur" & @CRLF & "- Racines des lecteurs C:, D: ..." & @CRLF & @CRLF &  "(Double-clic sur un élément pour l'afficher dans l'Explorateur.)", "Éléments récemment créés/modifiés dans :", 1,1)

;~ ///////**********************************************************************
;~ ///////**********      Dossiers créés sur le Bureau        ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche des dossiers récents sur le bureau." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Bureau...", "", Default, Default, 1)
	Local $Dossier = @DesktopDir & '\' ;
	Local $TmpMsgForLogging = ""
	Local $TmpSpaces = "                            "
	$Liste = _FileListToArrayRec($Dossier, "*||", 30, 0, 2, 1)
	If IsArray($Liste) Then
		_Logging("Scan des dossiers sur le bureau : " & $Liste[0] & " dossier(s)", 2, 0)
		For $N = 1 To $Liste[0]
			$Fldr_Name = @DesktopDir & '\' & StringTrimRight($Liste[$N], 1)
;~ 			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			ProgressSet(Round($N / $Liste[0] * 100), "[" & $N & "/" & $Liste[0] & "] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			$Fldr_Name_Relative_Path = "Bureau\" & $Liste[$N]
;~ 			$Fldr_Name = StringTrimRight($Liste[$n],1)
			If AgeDuFichierEnMinutesCreation($Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
				$Fldr_Name_Relative_Path = "Bureau\" & $Liste[$N]
				$Fldr_info = DirGetSize($Fldr_Name, 1)
				$Fldr_size = $Fldr_info[0]
				If $Fldr_info[1] = 0 Then
					$Fldr_filesCount = "  [dossier vide]"
				ElseIf $Fldr_info[1] = 1 Then
					$Fldr_filesCount = "  [1 fichier]"
				Else
					$Fldr_filesCount = "  [" & $Fldr_info[1] & " fichiers]"
				EndIf
				$Fldr_t = FileGetTime($Fldr_Name, 1) ; creation Time
				$Fldr_time = "[" & $Fldr_t[3] & ":" & $Fldr_t[4] & "]"
				$item = GUICtrlCreateListViewItem($Fldr_Name_Relative_Path & "|" & _FineSize($Fldr_size) & "|" & $Fldr_time & " " & $Fldr_filesCount & "|" & @DesktopDir & '\' & $Liste[$N], $GUI_AfficherLeContenuDesAutresDossiers)
				$TmpMsgForLogging &= @CRLF & """" & $Fldr_Name_Relative_Path & """    (" & _FineSize($Fldr_size) & "-" & $Fldr_time & " " & $Fldr_filesCount & ")"
				If $Fldr_size > $TAILLE_MAX_DU_DOSSIER_SOUS_LECTEUR * 1024 * 1024 Then
					GUICtrlSetColor(-1, 0xFF0000)
				EndIf
			EndIf

		Next
		If $TmpMsgForLogging <> "" Then
			_Logging("  -->  Liste de dossiers récents sur le bureau : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
		Else
			_Logging("  -->  Aucun dossie récent.", 2, 0)
		EndIf
	EndIf
;~ ///////**********************************************************************
;~ ///////**********      Fichiers Modifiés sur le bureau     ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche des fichiers récents sur le ""Bureau""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressSet(0, "[0%] Veuillez patienter un moment, initialisation...", "Scan des fichiers >> Bureau...")
	Local $Dossier = @DesktopDir & '\' ;
	$TmpMsgForLogging = ""
	Local $Liste[1] = [0] ;
	$Liste = _FileListToArrayRec($Dossier, $Mask & "|*.lnk|", 1, 1, 2, 1)

	If IsArray($Liste) Then
		_Logging("Scan des fichiers sur le bureau : " & $Liste[0] & " fichiers(s)", 2, 0)
		For $N = 1 To $Liste[0]
;~ 			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			ProgressSet(Round($N / $Liste[0] * 100), "[" & $N & "/" & $Liste[0] & "] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			$File_Name = @DesktopDir & '\' & $Liste[$N]
			$File_Name_Relative_Path = "Bureau\" & $Liste[$N]
			If AgeDuFichierEnMinutesModification($File_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
				$File_size = _FineSize(FileGetSize($File_Name))

				$File_t = FileGetTime($File_Name, 1) ; creation Time
				$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "]"
				$File_t = FileGetTime($File_Name, 0) ; Modif Time
				$File_time &= "   [" & $File_t[3] & ":" & $File_t[4] & "]"

				$item = GUICtrlCreateListViewItem($File_Name_Relative_Path & "|" & $File_size & "|" & $File_time & "|" & @DesktopDir & '\' & $Liste[$N], $GUI_AfficherLeContenuDesAutresDossiers)
				$TmpMsgForLogging &= @CRLF & """" & $File_Name_Relative_Path & """    (" & $File_size & "-" & $File_time & ")"

;~ 				GUICtrlSetColor(-1, 0xFF0000)
			EndIf

		Next

		If $TmpMsgForLogging <> "" Then
			_Logging("  -->  Liste de fichiers récents sur le bureau : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
		Else
			_Logging("  -->  Aucun fichier récent.", 2, 0)
		EndIf
	EndIf



;~ ///////**********************************************************************
;~ ///////**********      Dossiers créés dans profil utilisateur        ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche des dossiers récents dans Mes documents." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
;~ 	ProgressSet(0, "[0%] Veuillez patienter un moment, initialisation...", "Scan des dossiers >> profil utilisateur...")
;~ 	Local $Dossier = @UserProfileDir & '\' ;
;~ 	$TmpMsgForLogging = ""
;~ 	$Liste = _FileListToArrayRec($Dossier, "*||", 30, 0, 2, 1)
;~ 	If IsArray($Liste) Then
;~ 		_Logging("Scan des dossiers dans profil utilisateur : " & $Liste[0] & " dossier(s)", 2, 0)
;~ 		For $N = 1 To $Liste[0]
;~ 			ProgressSet(Round($N / $Liste[0] * 100), "[" & $N & "/" & $Liste[0] & "] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
;~ 			$Fldr_Name = @UserProfileDir & '\' & StringTrimRight($Liste[$N], 1)
;~ 			$Fldr_Name_Relative_Path = "ProfilU\" & $Liste[$N]
;~ 			If AgeDuFichierEnMinutesCreation($Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
;~ 				$Fldr_Name_Relative_Path = "ProfilU\" & $Liste[$N]
;~ 				$Fldr_info = DirGetSize($Fldr_Name, 1)
;~ 				$Fldr_size = $Fldr_info[0]
;~ 				If $Fldr_info[1] = 0 Then
;~ 					$Fldr_filesCount = "  [dossier vide]"
;~ 				ElseIf $Fldr_info[1] = 1 Then
;~ 					$Fldr_filesCount = "  [1 fichier]"
;~ 				Else
;~ 					$Fldr_filesCount = "  [" & $Fldr_info[1] & " fichiers]"
;~ 				EndIf
;~ 				$Fldr_t = FileGetTime($Fldr_Name, 1) ; creation Time
;~ 				$Fldr_time = "[" & $Fldr_t[3] & ":" & $Fldr_t[4] & "]" ; detailed
;~ 				$item = GUICtrlCreateListViewItem($Fldr_Name_Relative_Path & "|" & _FineSize($Fldr_size) & "|" & $Fldr_time & " " & $Fldr_filesCount & "|" & @MyDocumentsDir & '\' & $Liste[$N], $GUI_AfficherLeContenuDesAutresDossiers)
;~ 				$TmpMsgForLogging &= @CRLF & """" & $Fldr_Name_Relative_Path & """    (" & _FineSize($Fldr_size) & "-" & $Fldr_time & " " & $Fldr_filesCount & ")"
;~ 				If $Fldr_size > 10 * 1024 * 1024 Then
;~ 					GUICtrlSetColor(-1, 0xFF0000)
;~ 				EndIf
;~ 			EndIf

;~ 		Next
;~ 		If $TmpMsgForLogging <> "" Then
;~ 			_Logging("  -->  Liste de dossiers récents dans profil utilisateur : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
;~ 		Else
;~ 			_Logging("  -->  Aucun dossier récent.", 2, 0)
;~ 		EndIf
;~ 	EndIf

;~ ///////**********************************************************************
;~ ///////***   Fichiers Modifiés dans le dossier de profil utilisateur   ******
;~ ///////**********************************************************************

;~ 	SplashTextOn("Sans Titre", "Recherche des fichiers récents dans ""Mes documents""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressSet(0, "[0%] Veuillez patienter un moment, initialisation...", "Scan des fichiers >> profil utilisateur...")
	Local $Dossier = @UserProfileDir & '\' ;
	$TmpMsgForLogging = ""
	Local $Liste[1] = [0] ;
	$Liste = _FileListToArrayRec($Dossier, "*.py;*.ipynb;*.ui;*.accdb;*.xls*;*.csv;*.doc*;*.ppt*||", 1, 0, 2, 1)  ; Non récursif..

	If IsArray($Liste) Then
		_Logging("Scan des fichiers dans Profil Utilisateur : " & $Liste[0] & " fichier(s)", 2, 0)
		For $N = 1 To $Liste[0]
;~ 			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			ProgressSet(Round($N / $Liste[0] * 100), "[" & $N & "/" & $Liste[0] & "] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			$File_Name = @UserProfileDir & '\' & $Liste[$N]
			$File_Name_Relative_Path = "ProfilU\" & $Liste[$N]
			If AgeDuFichierEnMinutesModification($File_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
				$File_size = _FineSize(FileGetSize($File_Name))
				$File_t = FileGetTime($File_Name, 1) ; creation Time
				$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "]"
				$File_t = FileGetTime($File_Name, 0) ; Modif Time
				$File_time &= "   [" & $File_t[3] & ":" & $File_t[4] & "]"
;~ 				$item = GUICtrlCreateListViewItem($File_Name & "|" & $File_size & "|" & $File_time, $GUI_AfficherLeContenuDesAutresDossiers)
				$item = GUICtrlCreateListViewItem($File_Name_Relative_Path & "|" & $File_size & "|" & $File_time & "|" & @MyDocumentsDir & '\' & $Liste[$N], $GUI_AfficherLeContenuDesAutresDossiers)
				$TmpMsgForLogging &= @CRLF & """" & $File_Name_Relative_Path & """    (" & $File_size & "-" & $File_time & ")"
;~ 				GUICtrlSetColor(-1, 0xFF0000)
			EndIf

		Next

		If $TmpMsgForLogging <> "" Then
			_Logging("  -->  Liste de fichiers récents dans Profil Utilisateur : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
		Else
			_Logging("  -->  Aucun fichier récent.", 2, 0)
		EndIf
	EndIf



;~ ///////**********************************************************************
;~ ///////**********      Dossiers créés dans Mes documents   ********
;~ ///////**********************************************************************
;~ 	SplashTextOn("Sans Titre", "Recherche des dossiers récents dans Mes documents." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressSet(0, "[0%] Veuillez patienter un moment, initialisation...", "Scan des dossiers >> Mes documents...")
	Local $Dossier = @MyDocumentsDir & '\' ;
	$TmpMsgForLogging = ""
	$Liste = _FileListToArrayRec($Dossier, "*||", 30, 0, 2, 1)
	If IsArray($Liste) Then
		_Logging("Scan des dossiers dans Mes documents : " & $Liste[0] & " dossier(s)", 2, 0)
		For $N = 1 To $Liste[0]
;~ 			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			ProgressSet(Round($N / $Liste[0] * 100), "[" & $N & "/" & $Liste[0] & "] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			$Fldr_Name = @MyDocumentsDir & '\' & StringTrimRight($Liste[$N], 1)
			$Fldr_Name_Relative_Path = "Mes documents\" & $Liste[$N]
;~ 			$Fldr_Name = StringTrimRight($Liste[$n],1)
			If AgeDuFichierEnMinutesCreation($Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
				$Fldr_Name_Relative_Path = "Mes documents\" & $Liste[$N]
				$Fldr_info = DirGetSize($Fldr_Name, 1)
				$Fldr_size = $Fldr_info[0]
				If $Fldr_info[1] = 0 Then
					$Fldr_filesCount = "  [dossier vide]"
				ElseIf $Fldr_info[1] = 1 Then
					$Fldr_filesCount = "  [1 fichier]"
				Else
					$Fldr_filesCount = "  [" & $Fldr_info[1] & " fichiers]"
				EndIf
				$Fldr_t = FileGetTime($Fldr_Name, 1) ; creation Time
				$Fldr_time = "[" & $Fldr_t[3] & ":" & $Fldr_t[4] & "]" ; detailed
				$item = GUICtrlCreateListViewItem($Fldr_Name_Relative_Path & "|" & _FineSize($Fldr_size) & "|" & $Fldr_time & " " & $Fldr_filesCount & "|" & @MyDocumentsDir & '\' & $Liste[$N], $GUI_AfficherLeContenuDesAutresDossiers)
				$TmpMsgForLogging &= @CRLF & """" & $Fldr_Name_Relative_Path & """    (" & _FineSize($Fldr_size) & "-" & $Fldr_time & " " & $Fldr_filesCount & ")"
				If $Fldr_size > 10 * 1024 * 1024 Then
					GUICtrlSetColor(-1, 0xFF0000)
				EndIf
			EndIf

		Next
		If $TmpMsgForLogging <> "" Then
			_Logging("  -->  Liste de dossiers récents dans Mes documents : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
		Else
			_Logging("  -->  Aucun dossier récent.", 2, 0)
		EndIf
	EndIf

;~ ///////**********************************************************************
;~ ///////**********      Fichiers Modifiés dans Mes documents   ********
;~ ///////**********************************************************************

;~ 	SplashTextOn("Sans Titre", "Recherche des fichiers récents dans ""Mes documents""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressSet(0, "[0%] Veuillez patienter un moment, initialisation...", "Scan des fichiers >> Mes documents...")
	Local $Dossier = @MyDocumentsDir & '\' ;
	$TmpMsgForLogging = ""
	Local $Liste[1] = [0] ;
	$Liste = _FileListToArrayRec($Dossier, $Mask & "||", 1, 1, 2, 1)

	If IsArray($Liste) Then
		_Logging("Scan des fichiers dans Mes documents : " & $Liste[0] & " fichier(s)", 2, 0)
		For $N = 1 To $Liste[0]
;~ 			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			ProgressSet(Round($N / $Liste[0] * 100), "[" & $N & "/" & $Liste[0] & "] " & "Vérif. de : " & StringRegExpReplace($Liste[$N], "^.*\\", ""))
			$File_Name = @MyDocumentsDir & '\' & $Liste[$N]
			$File_Name_Relative_Path = "Mes documents\" & $Liste[$N]
			If AgeDuFichierEnMinutesModification($File_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
				$File_size = _FineSize(FileGetSize($File_Name))
				$File_t = FileGetTime($File_Name, 1) ; creation Time
				$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "]"
				$File_t = FileGetTime($File_Name, 0) ; Modif Time
				$File_time &= "   [" & $File_t[3] & ":" & $File_t[4] & "]"
;~ 				$item = GUICtrlCreateListViewItem($File_Name & "|" & $File_size & "|" & $File_time, $GUI_AfficherLeContenuDesAutresDossiers)
				$item = GUICtrlCreateListViewItem($File_Name_Relative_Path & "|" & $File_size & "|" & $File_time & "|" & @MyDocumentsDir & '\' & $Liste[$N], $GUI_AfficherLeContenuDesAutresDossiers)
				$TmpMsgForLogging &= @CRLF & """" & $File_Name_Relative_Path & """    (" & $File_size & "-" & $File_time & ")"
;~ 				GUICtrlSetColor(-1, 0xFF0000)
			EndIf

		Next

		If $TmpMsgForLogging <> "" Then
			_Logging("  -->  Liste de fichiers récents dans Mes documents : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
		Else
			_Logging("  -->  Aucun fichier récent.", 2, 0)
		EndIf
	EndIf
;~ ///////**********************************************************************
;~ ///////**********      Dossiers récemment créés sous les lecteurs    ********
;~ ///////**********************************************************************
	Local $aDrive = DriveGetDrive('FIXED')

;~ 	_ArraySort($aDrive,1,1)
	Local $MaskExclude = ""
	For $i = 1 To $aDrive[0]
		$MaskExclude = "" ;init
		If $aDrive[$i] = "C:" Then
			$MaskExclude = "Program Files*;Users;Windows;Intel;PerfLogs;EasyPhp*;Bac*2*;Documents and Settings"
		EndIf

		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; pour Exclure les hdd externes
				And _WinAPI_IsWritable($aDrive[$i]) _
				Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
;~ 			SplashTextOn("Sans Titre", "Recherche des dossiers récents sous la racine " & $Dossier & "." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
			ProgressSet(0, "[0%] Veuillez patienter un moment, initialisation...", "Scan des dossiers >> lecteur " & $Dossier)

			$Liste = _FileListToArrayRec($Dossier, "*|" & $MaskExclude & "|", 30, 0, 2, 2)
			If IsArray($Liste) Then
				$TmpMsgForLogging = ""
				_Logging("Scan des dossiers sous le lecteur """ & $Dossier & """ : " & $Liste[0] & " dossier(s)", 2, 0)
				For $N = 1 To $Liste[0]
					$Fldr_Name = StringTrimRight($Liste[$N], 1)
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($Fldr_Name, "^.*\\", ""))
					If StringRegExp($Fldr_Name, "^(?i)bac\s*\d*2\d*\s*$", $STR_REGEXPMATCH, 1) or AgeDuFichierEnMinutesCreation($Fldr_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
						$Fldr_info = DirGetSize($Fldr_Name, 1)
						$Fldr_size = $Fldr_info[0]
						If $Fldr_info[1] = 0 Then
							$Fldr_filesCount = "  [dossier vide]"
						ElseIf $Fldr_info[1] = 1 Then
							$Fldr_filesCount = "  [1 fichier]"
						Else
							$Fldr_filesCount = "  [" & $Fldr_info[1] & " fichiers]"
						EndIf
						$Fldr_t = FileGetTime($Fldr_Name, 1) ; creation Time
						$Fldr_time = "[" & $Fldr_t[3] & ":" & $Fldr_t[4] & "]" ; detailed
						$item = GUICtrlCreateListViewItem($Fldr_Name & "\" & "|" & _FineSize($Fldr_size) & "|" & $Fldr_time & " " & $Fldr_filesCount & "|" & $Fldr_Name & "\", $GUI_AfficherLeContenuDesAutresDossiers)
						$TmpMsgForLogging &= @CRLF & """" & $Fldr_Name & "\" & """    (" & _FineSize($Fldr_size) & "-" & $Fldr_time & " " & $Fldr_filesCount & ")"
						If $Fldr_size > 10 * 1024 * 1024 Then
							GUICtrlSetColor(-1, 0xFF0000)
						EndIf
					EndIf

				Next
			EndIf
			If $TmpMsgForLogging <> "" Then
				_Logging("  -->  Liste de dossiers récents sous le lecteur """ & $Dossier & """ : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
			Else
				_Logging("  -->  Aucun dossier récent.", 2, 0)
			EndIf
		EndIf
	Next


;~ ///////**********************************************************************
;~ ///////**********      Fichiers récemment Modifiés sous les lecteurs    ********
;~ ///////**********************************************************************

	Local $aDrive = DriveGetDrive('FIXED')
	Local $MaskExclude = "autoexec.bat;config.sys;*.log;*.rtf"
	For $i = 1 To $aDrive[0]
		If (DriveGetType($aDrive[$i], $DT_DRIVETYPE) = "Fixed") _ ; les hdd
				And (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; pour Exclure les hdd externes
				And _WinAPI_IsWritable($aDrive[$i]) _
				Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
;~ 			SplashTextOn("Sans Titre", "Recherche des fichiers récents sous la racine " & $Dossier & "." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
			ProgressSet(0, "[0%] Veuillez patienter un moment, initialisation...", "Scan des dossiers >> lecteur " & $Dossier)

			$Liste = _FileListToArrayRec($Dossier, $Mask & "|" & $MaskExclude & "|", 29, 0, 2, 2)
			If IsArray($Liste) Then
				$TmpMsgForLogging = ""
				_Logging("Scan des fichiers sous le lecteur """ & $Dossier & """ : " & $Liste[0] & " fichier(s)", 2, 0)
				For $N = 1 To $Liste[0]
					$File_Name = $Liste[$N]
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringRegExpReplace($File_Name, "^.*\\", ""))
					If AgeDuFichierEnMinutesModification($File_Name) < $AGE_DU_FICHIER_EN_MINUTES Then
						$File_size = _FineSize(FileGetSize($File_Name))
						$File_t = FileGetTime($File_Name, 1) ; creation Time
						$File_time = "[" & $File_t[3] & ":" & $File_t[4] & "]"
						$File_t = FileGetTime($File_Name, 0) ; Modif Time
						$File_time &= "   [" & $File_t[3] & ":" & $File_t[4] & "]"
						$item = GUICtrlCreateListViewItem($File_Name & "|" & $File_size & "|" & $File_time & "|" & $File_Name, $GUI_AfficherLeContenuDesAutresDossiers)
						$TmpMsgForLogging &= @CRLF & """" & $File_Name & """    (" & $File_size & "-" & $File_time & ")"
;~ 						GUICtrlSetColor(-1, 0xFF0000)
					EndIf

				Next
				If $TmpMsgForLogging <> "" Then
					_Logging("  -->  Liste de fichiers récents sous le lecteur """ & $Dossier & """ : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $TmpSpaces), 2, 0)
				Else
					_Logging("  -->  Aucun fichier récent.", 2, 0)
				EndIf
			EndIf
		EndIf
	Next


	ProgressOff()
	SplashOff()
EndFunc   ;==>_AfficherLeContenuDesAutresDossiers

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func Sauvegarder()
	_ClearGuiLog()
	_Logging("______", 2, 0)
	_Logging("Début de la sauvegarde des dossiers collectés vers le disque dur...", 4, 1)
	_Logging("Lecture et mise à jour des paramètres de " & $PROG_TITLE, 2, 0)
	_SaveParams()
	_InitialParams()

	Local Const $LIGNE = '----------------------------------------------------------------------------------' & @CRLF
	Local $TmpError = ""
	Local $iNberreurs = 0
	Local $sRapport = ""
	Local $sRapportFileName = "Rapport"
	Local $iTotFiles_DossiersRecup = 0
	Local $iTotSize_DossiersRecup = 0

	Local $sCheminScripDir = @ScriptDir
	If StringRight($sCheminScripDir, 1) <> "\" Then
		$sCheminScripDir = $sCheminScripDir & "\"
	EndIf
	Local $DestLocalFldr = ""
	Local $TmpListe[1] = [0] ;

	Local $Liste[1][5] = [["Candidat", "NbFichiers", "Extensions", "Taille", "Remarque"]] ;

;~ _Logging($sText, $iSuccess = 1, $iGuiLog = 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	_Logging("Préparation de la liste de dossiers", 4, 0)
	$TmpListe = _FileListToArray($sCheminScripDir, "??????", $FLTA_FOLDERS)

	If Not (IsArray($TmpListe)) Or ($TmpListe[0] = 0) Then
		_Logging("Aucun dossier de travail trouvé dans """ & @ScriptDir & """", 5, 1) ;
		_Logging("Sauvegarde annulée", 5, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Aucun dossier à sauvegarder !" & @CRLF _
				 & "Veuillez vérifier que " & $PROG_TITLE & " et les dossiers de travail des candidats" & @CRLF _
				 & "se trouvent dans le même chemin." & @CRLF, 0)
		Return
	EndIf

	SplashTextOn("Sans Titre", "Préparation de la liste de dossiers." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	_ArrayDelete($TmpListe, 0)
	_ArraySort($TmpListe)
	Local $TmpUneLigneMatrice = ""
	For $N = 0 To UBound($TmpListe) - 1
		$sDir = $TmpListe[$N]
		If StringRegExp($sDir, "([0-9]{6})", 0) = 1 Then
			$TmpUneLigneMatrice = ""
			$aFolderInfo = DirGetSize($sCheminScripDir & $sDir, 1)
			$FolderSize = $aFolderInfo[0]
			$FilesCount = $aFolderInfo[1]
			$iTotFiles_DossiersRecup += $FilesCount
			$iTotSize_DossiersRecup += $FolderSize
			$FolderSize = _FineSize($FolderSize)
			$sNbFilesByExt = FilesByExtension($sCheminScripDir & $sDir)
			If ($FilesCount = 0) And ($aFolderInfo[0] = 0) Then ;Nb Fichiers =0 & Taille =0
				$TmpUneLigneMatrice &= $sDir & "|"  ; N° Inscription
				$TmpUneLigneMatrice &= "" & "|"  ;  Nb Fichiers
				$TmpUneLigneMatrice &= "" & "|"  ; Nb Fichiers par Exrension (Limit 6 extensions)
				$TmpUneLigneMatrice &= "" & "|"  ; Taille
				$TmpUneLigneMatrice &= "Abs."  ; Notes/Remarques
			Else
				$TmpUneLigneMatrice &= $sDir & "|"  ; N° Inscription
				$TmpUneLigneMatrice &= $FilesCount & "|"  ;  Nb Fichiers
				$TmpUneLigneMatrice &= $sNbFilesByExt & "|"  ; Nb Fichiers par Exrension (Limit 6 extensions)
				$TmpUneLigneMatrice &= $FolderSize & "|"  ; Taille
				$TmpUneLigneMatrice &= ""  ; Notes/Remarques
			EndIf
			;----
			If ($aFolderInfo[1] = 0) And ($aFolderInfo[0] = 0) Then ;Nb Fichiers =0 & Taille =0
				Local $KesList = _FileListToArray($sCheminScripDir & $TmpListe[$N] & "\", "Abs*", $FLTA_FOLDERS)
				If IsArray($KesList) = 0 Or $KesList[0] = 0 Then
					$DossiersNonConformesPourCandidatsAbsents &= @TAB & "» " & $TmpListe[$N] & @CRLF
					$TmpUneLigneMatrice &= "¹"
				EndIf
			EndIf
			_ArrayAdd($Liste, $TmpUneLigneMatrice)
			;----
		EndIf
	Next
	If UBound($Liste, 1)-1 = 0 Then
		_Logging("Aucun dossier de travail trouvé dans """ & @ScriptDir & """", 5, 1) ;
		_Logging("Sauvegarde annulée", 5, 1)
		_Logging("______", 2, 0)
		SplashOff()
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Aucun dossier à sauvegarder !" & @CRLF _
				 & "Veuillez vérifier que " & $PROG_TITLE & " et les dossiers de travail des candidats" & @CRLF _
				 & "sont dans le même chemin." & @CRLF, 0)
		Return
	EndIf

	; Si le nombre de dossiers > 15
	If UBound($Liste, 1)-1 > 15 Then
		_Logging("Nombre de dossiers de travail > 15 : " & UBound($Liste, 1) - 1 & " dossiers", 4) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("MsgBox: Continuer (Oui/Non)?", 2, 0)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "Non|~Oui", $PROG_TITLE & $PROG_VERSION, "Le Nombre de dossiers de travail > 15 : " & UBound($Liste, 1) - 1 & " dossiers"  & @CRLF _
				 & "Veuillez vérifier la liste de candidats dans ce Labo pour cette séance." & @CRLF _
				 & "Si vous décider de poursuivre cette opération, la grille d'évaluation, comprendrait uniquement les 15 premiers dossiers!" & @CRLF & @CRLF _
				 & "voulez-vous continuer quand même?" & @CRLF, 0)

		If $Rep = 2 Then
			_Logging("Oui. Continuer avec  " & UBound($Liste, 1) - 1 & " dossiers.", 5, 0)
		Else
			_Logging("Non.", 2, 0)
			SplashOff()
			return
		EndIf
	EndIf


	Local $TmpSpaces = "                            "
	_Logging("Liste prête à sauvegarder : " & UBound($Liste, 1) - 1 & " dossier(s)", 4) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	_Logging("Liste : " & _ArrayToString($Liste, "", 1, -1, @CRLF & $TmpSpaces, 0, 0), 2, 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	;ici - Au moins un dossier sauvegardé

	SplashTextOn("Sans Titre", "Création du dossier de sauvegarde." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)


	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;~ Création du Sous-dossier de sauvegarde
	Local $DestLocalFldr = IniRead($Lecteur & $DossierSauve & "\0-BacCollector\BacCollector.ini", "Params", "SousDossierSauve", "-1")
	If $DestLocalFldr = -1 Or StringRegExp(StringLeft($DestLocalFldr, 2), "([0-9]{2})", 0) = 0 Then
		$Tmp = "01"
	Else
		$Tmp = StringLeft($DestLocalFldr, 2)
		$Tmp = $Tmp + 1
		If $Tmp > 99 Then
			$Tmp = "01"
		EndIf
		$Tmp = "0" & $Tmp
		$Tmp = StringRight($Tmp, 2)
	EndIf

	$DestLocalFldr = $Tmp
;~ 	$DestLocalFldr &= '-' & "L" & StringRight(GUICtrlRead($cLabo), 1) & "S" & StringRight(GUICtrlRead($cSeance), 1) ;L1S3: Labo-1 Séance-3
;~ 	$DestLocalFldr &= '-' & @MDAY & _MonthName(@MON) & @YEAR
	$DestLocalFldr &= '__' & "Labo" & StringRight(GUICtrlRead($cLabo), 1) & "_Séance" & StringRight(GUICtrlRead($cSeance), 1) ;L1S3: Labo-1 Séance-3
;~ 	$DestLocalFldr &= '__' &GUICtrlRead($cLabo) & "_" & GUICtrlRead($cSeance)
	$DestLocalFldr &= '__' & @MDAY & "-" & @MON & "-" & @YEAR
	$DestLocalFldr &= '__' & @HOUR & "h" & @MIN
	IniWrite($Lecteur & $DossierSauve & "\0-BacCollector\BacCollector.ini", "Params", "SousDossierSauve", $DestLocalFldr)

	$DestLocalFldr = $Lecteur & $DossierSauve & "\" & $DossierBacCollector & "\" & $DestLocalFldr & "\"
	$TmpError = DirCreate($DestLocalFldr)
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
	;;;Log+++++++
	_Logging("Création du dossier de sauvegarde: " & @CRLF & "         """ & $DestLocalFldr & """", $TmpError) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	If $TmpError = 0 Then ;1:Success 0:Failure
		_Logging("Sauvegarde annulée", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("______", 2, 0)
		SplashOff()
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Échec lors de la création du dossier de sauvegarde: " & @CRLF _
				 & "   » " & $DestLocalFldr & @CRLF _
				 & "L'opération de sauvegarde est annulée" & @CRLF, 0)
		Return
	EndIf
	;;;Log+++++++
	SplashTextOn("Sans Titre", "Copie des dossiers." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	If $Matiere = "InfoProg" Then
		$KesMat = "Informatique / Algorithmique et Programmation"
	Else
		$KesMat = "STI - Systèmes & Technologies de l'Informatique"
	EndIf
	SplashOff()
	Local $N = UBound($Liste, 1) - 1
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Opération de sauvegarde [" & $N & " dossiers]", "", Default, Default, 1)

	For $i = 1 To $N
		ProgressSet(Round($i / $N * 100), "[" & Round($i / $N * 100) & "%] " & "Copie du dossier : [" & $Liste[$i][0] & "]")
		$TmpError = DirCopy($sCheminScripDir & $Liste[$i][0], $DestLocalFldr & $Liste[$i][0], $FC_OVERWRITE)
		;;;Log+++++++
		_Logging("Copie du dossier: """ & $sCheminScripDir & $Liste[$i][0] & """", $TmpError) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		If $TmpError = 0 Then ;1:Success 0:Failure
			$iNberreurs += 1
			$Liste[$i][4] = "Pb. copie"
		EndIf
		;;;Log+++++++
	Next
	ProgressSet(Round(($i-1) / $N * 100), "[" & Round(($i-1) / $N * 100) & "%] " & "Création du Rapport et de la Grille d'Évaluation.")
	;;;;;;;;;;;;;;;------------------
	Local $TmpListeCaniats =  $Liste
	_ArrayColDelete($TmpListeCaniats, 1)
	_ArrayColDelete($TmpListeCaniats, 1)
	_ArrayColDelete($TmpListeCaniats, 1)
	_ArrayDelete($TmpListeCaniats , 0)
	$KesNomGrilleEval = _GenererGrilleEvaluation($TmpListeCaniats)
	If $KesNomGrilleEval Then
		_Logging("Création de la grille d'évaluation: " & @CRLF & "         """ & $KesNomGrilleEval & """", 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	Else
		_Logging("Création de la grille d'évaluation.", 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	EndIf

	Local $sCheminScripDir = @ScriptDir
	If StringRight($sCheminScripDir, 1) <> "\" Then
		$sCheminScripDir = $sCheminScripDir & "\"
	EndIf

	If FileExists($sCheminScripDir & $KesNomGrilleEval) Then
		FileCopy($sCheminScripDir & $KesNomGrilleEval, $DestLocalFldr)
	EndIf
   ;;;;;;;;;;;;;;;------------------

	ProgressSet(Round(($i-1) / $N * 100), "[" & Round(($i-1) / $N * 100) & "%] " & "Génération du Rapport.")
	$sRapportFileName &= '__' & GUICtrlRead($cSeance) & "_" &  GUICtrlRead($cLabo)
	$sRapportFileName &= '__' & JourDeLaSemaine() & "-" & @MDAY & "-" & @MON & "-" & @YEAR
	$sRapportFileName &= '__' & @HOUR & "h" & @MIN

	_GenererRapportPdf($sRapportFileName, $Liste, $DestLocalFldr, $KesMat)

	If FileExists($sCheminScripDir & $sRapportFileName & ".pdf") Then
		FileCopy($sCheminScripDir & $sRapportFileName & ".pdf", $DestLocalFldr)
	EndIf

	ProgressOff()
	If $iNberreurs = 0 Then
		_Logging("Sauvegarde terminée avec succès", 3, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("______", 2, 0)
		_CopierLog($DestLocalFldr)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_CENTER, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(64, "Ok", $PROG_TITLE & $PROG_VERSION, "La Sauvegarde est terminée avec succès." & @CRLF, 0)
	ElseIf $iNberreurs = 1 Then
		_Logging("Sauvegarde terminée avec une seule erreur", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("______", 2, 0)
		_CopierLog($DestLocalFldr)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "La sauvegarde est terminée, mais il y a eu une erreur." & @CRLF, 0)
	Else
		_Logging("Récupération terminée, avec " & $iNberreurs & " erreurs", 5, 1)
		_Logging("______", 2, 0)
		_CopierLog($DestLocalFldr)
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & $PROG_VERSION, "La sauvegarde est terminée, mais il y a eu " & $iNberreurs & " erreurs." & @CRLF, 0)

	EndIf
	WinActivate($hMainGUI)
EndFunc   ;==>Sauvegarder

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _FormatCol($sString, $iReservedSpace)
	$sString &= "                                    "
	Return StringLeft($sString, $iReservedSpace)
EndFunc   ;==>_FormatCol

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func FilesByExtension($sSearchDir)
	Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
	$aArray = _FileListToArrayRec($sSearchDir, "*.sql;*.htm*;*.css;*.js;*.php;*.accdb;*.csv;*.py;*.ipynb;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.ui;*.dat;*.txt||", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
	If @error Or Not IsArray($aArray) Then
		Return ""
	EndIf
	For $i = 1 To $aArray[0]
		$aPathSplit = _PathSplit($aArray[$i], $sDrive, $sSearchDir, $sFileName, $sExtension)
		$aArray[$i] = StringLower($sExtension)
	Next
	Local $sd = ObjCreate("Scripting.Dictionary"), $col = 2 ; col to search in
	; get the number of occurences for each unique element in col $col
	For $i = 1 To $aArray[0]
		If StringLen($aArray[$i]) > 0 And StringLen($aArray[$i]) < 7 Then
			$sd.Item($aArray[$i]) = ($sd.Exists($aArray[$i]) = 0) ? 1 : $sd.Item($aArray[$i]) + 1
		EndIf
	Next
	$aArray = _SDtoArray($sd)
	Local $sNbFilesByExtension = ""
	Local $N = $aArray[0][0]
	If $N > 6 Then  ; Top 6
		$N = 6
	EndIf

	For $i = 1 To $N - 1
		; si nb extension > 9 -> 9
		$sNbFilesByExtension &= ($aArray[$i][1] > 9 ? 9 : $aArray[$i][1])& " " & $aArray[$i][0] & " • "
	Next
	$sNbFilesByExtension &= ($aArray[$i][1] > 9 ? 9 : $aArray[$i][1])& " " & $aArray[$i][0]
;~ 	ConsoleWrite($sNbFilesByExtension)
;~ 	MsgBox($MB_SYSTEMMODAL, "Info", $sNbFilesByExtension)
	Return $sNbFilesByExtension
EndFunc   ;==>FilesByExtension
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
Func _SDtoArray($_sd)
	Local $count = $_sd.Count
	Local $ret[$count + 1][2]
	$ret[0][0] = $count
	For $i = 1 To $count
		$ret[$i][0] = $_sd.Keys[$i - 1]
		$ret[$i][1] = $_sd.Items[$i - 1]
	Next
	_ArraySort($ret, 1, 1, 0, 1)
	Return $ret
EndFunc   ;==>_SDtoArray

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _GenererRapportPdf($sRapportFileName, $Liste, $DestLocalFldr, $sExamMatiere)
	_SetTitle($sRapportFileName)
	_SetSubject("Bac Pratique " & GUICtrlRead($cBac) & " - " & JourDeLaSemaine() & " " & @MDAY & "-" & @MON & "-" & @YEAR)
	_SetAuthor("moez.romdhane@tarbia.tn")
	_SetCreator($PROG_TITLE & $PROG_VERSION)
	_SetKeywords("baccalauréat, informatique, examen, pratique, baccollector, sti")
	_OpenAfter(True);open after generation
	_SetUnit($PDF_UNIT_CM)
	_SetPaperSize("A4")
	_SetZoomMode($PDF_ZOOM_FULLPAGE)
	_SetOrientation($PDF_ORIENTATION_PORTRAIT)
	_SetLayoutMode($PDF_LAYOUT_CONTINOUS)


	Local $sCheminScripDir = @ScriptDir
	If StringRight($sCheminScripDir, 1) <> "\" Then
		$sCheminScripDir = $sCheminScripDir & "\"
	EndIf
	;initialize the pdf
	_InitPDF($sCheminScripDir & $sRapportFileName & ".pdf")

; === load used font(s) ===
	_LoadFontTT("_CalibriB", $PDF_FONT_CALIBRI, $PDF_FONT_BOLD)
	_LoadFontTT("_CalibriI", $PDF_FONT_CALIBRI, $PDF_FONT_ITALIC)
	_LoadFontTT("_Calibri", $PDF_FONT_CALIBRI)
;~ Colours:
	Local $_iColorBorder = 0x669bbc ; Bleu clair
	Local $_iColorText1 = 0x003566 ; Bleu foncé
	Local $_iColorText2 = 0x212529 ; gris
	Local $_iColorText3 = 0x9e0059 ; Rouge

	Local $pHight = 29.7

	_BeginPage() ; begin page
		_SetColourStroke($_iColorBorder)
		_Draw_Rectangle(1.5, $pHight - 4.2, 18, 2.8, $PDF_STYLE_STROKED, 0.5, 0xFFFFFF, 0.05)
		_SetColourStroke(0)
		_InsertRenderedText(10.5, $pHight - 2.4, "Bac Pratique " & GUICtrlRead($cBac), "_Calibri", 20, 100, $PDF_ALIGN_CENTER, $_iColorText1, $_iColorText1)
		_InsertRenderedText(10.5, $pHight - 3.1, $sExamMatiere, "_Calibri", 13, 100, $PDF_ALIGN_CENTER, $_iColorText1, $_iColorText1)
		_InsertRenderedText(10.5, $pHight - 3.8, JourDeLaSemaine() & " " & @MDAY & "-" & @MON & "-" & @YEAR, "_CalibriB", 13, 100, $PDF_ALIGN_CENTER, $_iColorText2, $_iColorText2)

		_InsertRenderedText(10.5, $pHight - 5.4, StringUpper(GUICtrlRead($cSeance) & " • " & GUICtrlRead($cLabo)), "_CalibriB", 22, 100, $PDF_ALIGN_CENTER, $_iColorText2, $_iColorText2)
		Local $iRowH = 0.6
		Local $iY = 27.3
		Local $iX = 1.5
		Local $iStepUp = 0.2
		Local $_FontSize = 9.5
		Local $CellFillColour = 0xFFFFFF
		Local $CellBorderColor = $_iColorBorder
		Local $fThickness = 0.01
		Local $aCellsWidth[2] = [5, 13]
		Local $aCellsContent[2] = ['Id de l''Ordinateur de Sauvegarde', _GetUUID()]
		_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_Calibri" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_LEFT , 0x1e1e1e, 0x1e1e1e )
		_InsertRenderedText(10.5, 1.5, $PROG_TITLE & $PROG_VERSION, "_Calibri", 8, 100, $PDF_ALIGN_CENTER, $_iColorText2, $_iColorText2)
		$iY += $iRowH
		Local $aCellsWidth[2] = [5, 13]
		Local $aCellsContent[2] = ['Dossier de Sauvegarde', $DestLocalFldr]
		_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_Calibri" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_LEFT , 0x1e1e1e, 0x1e1e1e )
;~ 		_InsertRenderedText(10.5, 1.5, $PROG_TITLE & $PROG_VERSION, "_Calibri", 8, 100, $PDF_ALIGN_CENTER, $_iColorText2, $_iColorText2)

		_DrawText(19.8, 1.2, $PROG_TITLE & $PROG_VERSION, "_Calibri", 7, $PDF_ALIGN_LEFT, 90)
		_SetColourFill(0x0000ee)
		_InsertLink(20.1, 1.2, "https://github.com/romoez/BacCollector", "_Calibri", 7, $PDF_ALIGN_LEFT, 90)
		_SetColourFill(0)

		$iY = 6
		$fThickness = 0.01
		$_FontSize = 11.5
		$iRowH = .7
		$iStepUp = 0.2
		Local $aCellsWidth[6] = [.7, 1.8, 1.2, 9.1, 2, 3.2]
		Local $aCellsContent[6] = ['#', "N° Ins.", "Files", "Nombre de Fichiers / Type (Top 6)", "Taille", "Rques."]
		_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_CalibriB" , $iRowH, $iStepUp, 0xefefef, $CellBorderColor, $fThickness, $_FontSize  + 1.5, 100, $PDF_ALIGN_CENTER , $_iColorText1, $_iColorText1 )
		Local $N = UBound($Liste, 1) - 1
		$N = $N <= 15 ? $N : 15
		$_FontSize = 11.5
		$iY += $iRowH
		For $i = 1 To $N
			Local $array1D[UBound($Liste, 2) + 1]

			; Copie de la ligne N°1 dans le tableau 1D
			$array1D[0] = $i
			$array1D[1] = StringLeft($Liste[$i][0], 3) & "-" & StringRight($Liste[$i][0], 3)
			For $j = 1 To UBound($Liste, 2) - 1
				$array1D[$j + 1] = $Liste[$i][$j]
			Next

			If Mod($i, 2) = 0 Then
				$CellFillColour = 0xefefef
			Else
				$CellFillColour = 0xFFFFFF
			EndIf
			_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $array1D, "_Calibri" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_LEFT , 0x000000, 0x000000 )
			$iY += $iRowH
		Next
		$iY += $iRowH
		Local $_sConsigne = "N.B. :"
		_InsertRenderedText(1.5, $pHight - $iY, $_sConsigne, "_CalibriB", 11, 100, $PDF_ALIGN_LEFT, $_iColorText2, $_iColorText2)
		Local $_sConsigne = "Assurez-vous que les fichiers manquants ne se trouvent pas déjà sur l'ordinateur du candidat."
;~ 		Local $_sConsigne = "Écrire devant le numéro du candidat correspondant, les types des fichiers manquants dans son travail."
		_Paragraph( $_sConsigne , 1.5 + 1.2, $pHight - $iY, 16.8, "_Calibri", 11)

;~ 		===== La liste des extensions



		If StringInStr($sExamMatiere, "STI") Then
			$iY = 23
			Local $iHeaderWidth = 9
			Local $aCellsWidth[1] = [$iHeaderWidth]
			Local $aCellsContent[1] = ['Types des fichiers demandés pour la matière STI']
			$iRowH = .7
			$iStepUp = 0.2
			$CellFillColour = 0xefefef
			_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_CalibriB" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_CENTER , 0x1e1e1e, 0x1e1e1e )
			$iY += $iRowH
			$CellFillColour = 0xFFFFFF
			Local $aCellsWidth[1] = [9]
			Local $aCellsContent[1] = ['sql • html • css • js • php']
			_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_Calibri" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_CENTER , 0x1e1e1e, 0x1e1e1e )
		Else
			$iY = 22.3
			$iRowH = .7
			$iStepUp = 0.2
			$CellFillColour = 0xefefef
			$iHeaderWidth = 12.4
			Local $aCellsWidth[1] = [$iHeaderWidth]
			Local $aCellsContent[1] = ['Types des fichiers demandés par section']
			_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_CalibriB" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_CENTER , 0x1e1e1e, 0x1e1e1e )
			$iY += $iRowH
			Local $aCellsWidth[5] = [3.4, 2.4, 1.9, 1.8, 2.9]
			Local $aCellsContent[5] = ['Éco & Gest', 'Sc • T • M', 'Lettres','Sport', 'SI']
			_InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_CalibriB" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_CENTER , 0x1e1e1e, 0x1e1e1e )
			$iY += $iRowH
			$CellFillColour = 0xFFFFFF
			Local $aCellsContent1[5] = ['accdb • xlsx • csv', 'py ou ipynb', 'xlsx', 'xlsx', 'py ou ipynb']
			Local $aCellsContent2[5] = ['py ou ipynb', 'ui', 'docx', 'pptx', 'ui • dat • txt']
			_InsertPdf2TableRows(1.5, $iY, $aCellsWidth,  $aCellsContent1, $aCellsContent2, "_Calibri" , $iRowH, $iStepUp, $CellFillColour, $CellBorderColor, $fThickness, $_FontSize, 100, $PDF_ALIGN_CENTER , 0x1e1e1e, 0x1e1e1e )
			$iY += $iRowH
		EndIf

		$iY += $iRowH  + .5
		Local $_sConsigne = "Veuillez noter que cette liste est indicative et sujette à modification. Toute modification ou mise à jour sera communiquée par le coordinateur."
		_Paragraph( $_sConsigne , 1.5, $pHight - $iY, $iHeaderWidth, "_Calibri", 9)
	_EndPage()

	_ClosePDFFile()
EndFunc

Func _InsertPdfTableRow($iX, $iY, $aCellsWidth,  $aCellsContent, $sAlias = "_Calibri" , $iRowH = 0.7, $iStepUp = 0.2, $CellFillColour = 0xFFFFFF, $CellBorderColour = 0x000000, $fThickness = 0.01, $_FontSize = 13, $iScale = 100, $sAlign = $PDF_ALIGN_LEFT , $iTextFillColour = 0x000000, $iTextOutlineColour = 0x000000 )
    Local $iPgW = Round(_GetPageWidth() / _GetUnit(), 1)
    Local $iPgH = Round(_GetPageHeight() / _GetUnit(), 1)
	Local $iYbuttom = $iPgH - ($iY + $iRowH)
    For $i = 0 To UBound($aCellsWidth) - 1
		_SetColourStroke($CellBorderColour)
        _Draw_Rectangle($iX, $iYbuttom, $aCellsWidth[$i], $iRowH, $PDF_STYLE_STROKED, 0, $CellFillColour, $fThickness)
		If $sAlign = $PDF_ALIGN_LEFT Then
			_InsertRenderedText($iX + 0.1, $iYbuttom + $iStepUp, $aCellsContent[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
		ElseIf $sAlign = $PDF_ALIGN_CENTER Then
			_InsertRenderedText($iX + $aCellsWidth[$i] / 2, $iYbuttom + $iStepUp, $aCellsContent[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
		Else
			_InsertRenderedText($iX + $aCellsWidth[$i], $iYbuttom + $iStepUp, $aCellsContent[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
		EndIf

		$iX += $aCellsWidth[$i]
    Next
    _SetColourStroke(0)
EndFunc   ;==>_InsertTable

Func _InsertPdf2TableRows($iX, $iY, $aCellsWidth,  $aCellsContent1, $aCellsContent2, $sAlias = "_Calibri" , $iRowH = 0.7, $iStepUp = 0.2, $CellFillColour = 0xFFFFFF, $CellBorderColour = 0x000000, $fThickness = 0.01, $_FontSize = 13, $iScale = 100, $sAlign = $PDF_ALIGN_LEFT , $iTextFillColour = 0x000000, $iTextOutlineColour = 0x000000 )
    Local $iPgW = Round(_GetPageWidth() / _GetUnit(), 1)
    Local $iPgH = Round(_GetPageHeight() / _GetUnit(), 1)
	Local $iYbuttom = $iPgH - ($iY + $iRowH * 2)
    For $i = 0 To UBound($aCellsWidth) - 1
		_SetColourStroke($CellBorderColour)
        _Draw_Rectangle($iX, $iYbuttom, $aCellsWidth[$i], $iRowH * 2, $PDF_STYLE_STROKED, 0, $CellFillColour, $fThickness)
		If $sAlign = $PDF_ALIGN_LEFT Then
			_InsertRenderedText($iX + 0.1, $iYbuttom + $iStepUp + $iRowH, $aCellsContent1[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
			_InsertRenderedText($iX + 0.1, $iYbuttom + $iStepUp, $aCellsContent2[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
		ElseIf $sAlign = $PDF_ALIGN_CENTER Then
			_InsertRenderedText($iX + $aCellsWidth[$i] / 2, $iYbuttom + $iStepUp + $iRowH, $aCellsContent1[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
			_InsertRenderedText($iX + $aCellsWidth[$i] / 2, $iYbuttom + $iStepUp, $aCellsContent2[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
		Else
			_InsertRenderedText($iX + $aCellsWidth[$i], $iYbuttom + $iStepUp + $iRowH, $aCellsContent1[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
			_InsertRenderedText($iX + $aCellsWidth[$i], $iYbuttom + $iStepUp, $aCellsContent2[$i], $sAlias, $_FontSize, 100, $sAlign, $iTextFillColour, $iTextOutlineColour)
		EndIf

		$iX += $aCellsWidth[$i]
    Next
    _SetColourStroke(0)
EndFunc   ;==>_InsertTable

Func _InsertLink($iX, $iY, $sURL, $sFontAlias, $iFontSize, $iAlign, $iRotate)
    __InitObj()
    _DrawText($iX, $iY, $sURL, $sFontAlias, $iFontSize, $iAlign, $iRotate)
    __EndObj()
    __InitObj()
    __ToBuffer("<</URI(" & __ToPdfStr($sURL) & ") /Type /Action /S /URI>>")
    __EndObj()
EndFunc   ;==>_InsertLink

Func JourDeLaSemaine()
	Local $iDayOfWeek = @WDAY ; Récupère le jour de la semaine pour la date d'aujourd'hui

	; Conversion du résultat numérique en nom du jour de la semaine
	Local $sDayOfWeek = ""
	Switch $iDayOfWeek
		Case 1
			$sDayOfWeek = "Dimanche"
		Case 2
			$sDayOfWeek = "Lundi"
		Case 3
			$sDayOfWeek = "Mardi"
		Case 4
			$sDayOfWeek = "Mercredi"
		Case 5
			$sDayOfWeek = "Jeudi"
		Case 6
			$sDayOfWeek = "Vendredi"
		Case 7
			$sDayOfWeek = "Samedi"
	EndSwitch
	Return $sDayOfWeek
EndFunc
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _ListeDeDossiersRecuperes($SearchMask = "??????", $RegEx = "([0-9]{6})")
;~ 	SplashTextOn("Sans Titre", "Préparation de la liste des dossiers récupérés." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers récupérés:", "", Default, Default, 1)
	Local $sListeDossiersRecuperes = ""
	Local $iNombreDossiersRecuperes = 0
	Local $iTailleTotaleDossiersRecuperes = 0
	Local $sDir, $aFolderInfo, $FolderSize, $FilesCount

	Global $DossiersNonConformesPourCandidatsAbsents = ""  ; Sans aucun fichier mais ne contient pas un dossier nommé "Absent"
	Local $TmpList

	Local $sChemin = @ScriptDir ;

	If StringRight($sChemin, 1) <> "\" Then
		$sChemin = $sChemin & "\"
	EndIf

	Local $Liste[1] = [0] ;
	$Liste = _FileListToArray($sChemin, $SearchMask, $FLTA_FOLDERS)

	If IsArray($Liste) Then
		_ArraySort($Liste, 0, 1)
		For $N = 1 To $Liste[0]
			$sDir = $Liste[$N]
			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : [" & StringRegExpReplace($sDir, "^.*\\", "") & "]")
			If StringRegExp($sDir, $RegEx, 0) = 1 Then
				$aFolderInfo = DirGetSize($sChemin & $sDir, 1)
				$iNombreDossiersRecuperes += 1
				$iTailleTotaleDossiersRecuperes += $aFolderInfo[0]
				If $iNombreDossiersRecuperes <= 15 Then
					$FolderSize = _FineSize($aFolderInfo[0])
					$FilesCount = $aFolderInfo[1]
					$sDir = StringLeft($sDir, 3) & " " & StringRight($sDir, 3)
					If ($FilesCount = 0) And ($aFolderInfo[0] = 0) Then ;Nb Fichiers =0 & Taille =0
						$sListeDossiersRecuperes = $sListeDossiersRecuperes & "   » " & $sDir & "  ____Abs.____" & @CRLF
					Else
						$sListeDossiersRecuperes = $sListeDossiersRecuperes & "   » " & $sDir & "  [" & $FolderSize & "]" & @CRLF
					EndIf
				EndIf
				;----
				If ($aFolderInfo[1] = 0) And ($aFolderInfo[0] = 0) Then ;Nb Fichiers =0 & Taille =0
					$TmpList = _FileListToArray($sChemin & $Liste[$N] & "\", "Abs*", $FLTA_FOLDERS)
					If IsArray($TmpList) = 0 Or $TmpList[0] = 0 Then
						$DossiersNonConformesPourCandidatsAbsents &= @TAB & "» " & $Liste[$N] & @CRLF
					EndIf
				EndIf
				;----
			EndIf
		Next
		ProgressOff() ;Set(Round($n/$Liste[0])*100, StringRegExpReplace($sDir, "^.*\\", ""))
		SplashTextOn("Sans Titre", "Affichage de la liste des dossiers récupérés." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
		If $iNombreDossiersRecuperes > 1 Then
			$sListeDossiersRecuperes = $iNombreDossiersRecuperes & " Dossiers  [" & _FineSize($iTailleTotaleDossiersRecuperes) & "] :" _
					 & @CRLF & @CRLF & $sListeDossiersRecuperes
			If $iNombreDossiersRecuperes > 15 Then
				$sListeDossiersRecuperes = $sListeDossiersRecuperes & "... " & ($iNombreDossiersRecuperes - 15) & " autres dossiers !!!"
			EndIf
		EndIf
	EndIf

	SplashOff()
	Return $sListeDossiersRecuperes
EndFunc   ;==>_ListeDeDossiersRecuperes

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _ListeDApplicationsOuvertes()
;~ 	SplashTextOn("Sans Titre", "Recherche de logiciels ouverts." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn("", "Vérif. des logiciels ouverts:", "", Default, Default, 1)

	Local $App = "" ;

	Local $SoftsToClose[][3] = [ _
			["Atom", "atom.exe", 1] _
			, ["Brackets", "Brackets.exe", 1] _
			, ["Dreamweaver", "Dreamweaver.exe", 1] _
			, ["Geany", "geany.exe", 1] _
			, ["GIMP", "GIMP", 0] _
			, ["Gvim", "gvim.exe", 1] _
			, ["LibreOffice/OpenOffice...", "soffice.bin", 2] _
			, ["Microsoft Access", "MSAccess.exe", 1] _
			, ["Microsoft Excel", "Excel.exe", 1] _
			, ["Microsoft Frontpage", "FrontPG.exe", 1] _
			, ["Microsoft PowerPoint", "POWERPNT.exe", 1] _
			, ["Microsoft Publisher", "MSPUB.exe", 1] _
			, ["Microsoft Word", "WinWord.exe", 1] _
			, ["Ms Expression Web 4", "ExpressionWeb.exe", 1] _
			, ["Notepad", "notepad.exe", 1] _
			, ["Notepad++", "notepad++.exe", 1] _
			, ["Paint .Net", "PaintDotNet.exe", 1] _
			, ["Paint", "MSPaint.exe", 1] _
			, ["PyScripter", "PyScripter.exe", 1] _
			, ["PyCharm", "pycharm", 0] _
			, ["Qt Designer", "designer.exe", 1] _
			, ["Rapid PHP", "rapidphp.exe", 1] _
			, ["Sublime Text", "sublime_text.exe", 1] _
			, ["UltraEdit", "uedit", 0] _
			, ["Vim", "vim.exe", 1] _
			, ["Visual Studio Code", "Code.exe", 1] _
			, ["Webuilder", "webuild.exe", 1] _
			, ["Windows Movie Maker", "Moviemk.exe", 1] _
			, ["Wing Python IDE", "wing", 0] _
			]

	$N = UBound($SoftsToClose, 1) - 1
	For $i = 0 To $N
		ProgressSet(Round($i / $N * 100), "[" & Round($i / $N * 100) & "%] " & "Vérif. de : " & $SoftsToClose[$i][0])
		If $SoftsToClose[$i][2] = 1 Then
			If ProcessExists($SoftsToClose[$i][1]) Then $App = $App & "» " & $SoftsToClose[$i][0] & @CRLF
		ElseIf $SoftsToClose[$i][2] = 0 Then
			If _MyProcessByPartName($SoftsToClose[$i][1]) Then $App = $App & "» " & $SoftsToClose[$i][0] & @CRLF
		ElseIf $SoftsToClose[$i][2] = 2 Then ;Openoffice /LibreOffice Process = soffice.bin
			If ProcessExists($SoftsToClose[$i][1]) Then
				Local $aWinList = WinList("[REGEXPTITLE:(?i)Office]")
				_ArrayDelete($aWinList, 0)
;~ 				_ArraySort($aWinList)
				If _ArraySearch($aWinList, ".*LibreOffice Base", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "LibreOffice Base" & @CRLF
				If _ArraySearch($aWinList, ".*LibreOffice Calc", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "LibreOffice Calc" & @CRLF
				If _ArraySearch($aWinList, ".*LibreOffice Draw", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "LibreOffice Draw" & @CRLF
				If _ArraySearch($aWinList, ".*LibreOffice Impress", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "LibreOffice Impress" & @CRLF
				If _ArraySearch($aWinList, ".*LibreOffice Math", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "LibreOffice Math" & @CRLF
				If _ArraySearch($aWinList, ".*LibreOffice Writer", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "LibreOffice Writer" & @CRLF
				If _ArraySearch($aWinList, ".*OpenOffice Base", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "OpenOffice Base" & @CRLF
				If _ArraySearch($aWinList, ".*OpenOffice Calc", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "OpenOffice Calc" & @CRLF
				If _ArraySearch($aWinList, ".*OpenOffice Draw", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "OpenOffice Draw" & @CRLF
				If _ArraySearch($aWinList, ".*OpenOffice Impress", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "OpenOffice Impress" & @CRLF
				If _ArraySearch($aWinList, ".*OpenOffice Math", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "OpenOffice Math" & @CRLF
				If _ArraySearch($aWinList, ".*OpenOffice Writer", 0, 0, 0, 3) <> -1 Then $App = $App & "» " & "OpenOffice Writer" & @CRLF

			EndIf
		EndIf
	Next

	$aPythonwProcess = ProcessList("pythonw.exe")
	Local $Liste[1] = [0], $temp, $sTitle
	$N = UBound($aPythonwProcess, 1) - 1
	For $i = 1 To $N
		$temp = _WinGetByPID($aPythonwProcess[$i][1])
		$sTitle = WinGetTitle($temp)
		if $sTitle <> "" Then
			$App = $App & "» " & $sTitle & @CRLF
		EndIf
	Next

	$aPythonwProcess = ProcessList("python.exe")
	Local $Liste[1] = [0], $temp, $sTitle
	$N = UBound($aPythonwProcess, 1) - 1
	For $i = 1 To $N
		$temp = _WinGetByPID($aPythonwProcess[$i][1])
		$sTitle = WinGetTitle($temp)
		if $sTitle <> "" Then
			$App = $App & "» " & $sTitle & @CRLF
		EndIf
	Next

	ProgressOff()
;~ 	SplashOff()
	Return $App

EndFunc   ;==>_ListeDApplicationsOuvertes

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _NumeroCandidat()
	SplashTextOn("Sans Titre", "Détermination du numéro d'inscription du candidat." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	Local $Bac = DossiersBac() ;
	Local $Liste[1] = [0] ;
	Local $ReturnPath = False
	Local $Flag = 0 ; 0:All, 1:Files only , 2:Fldrs only
	Local $NomFrom = "Bac"
	_Logging("Recherche du numéro d'inscription du candidat dans les dossiers ""Bac20xx"".", 2, 0)
	If $Bac[0] <> 0 Then
		For $i = 1 To $Bac[0]
			$Kes = _FileListToArray($Bac[$i], "*", $Flag, $ReturnPath)
			If IsArray($Kes) Then
				$Liste[0] += $Kes[0]
				_ArrayDelete($Kes, 0)
				_ArrayAdd($Liste, $Kes) ;
			EndIf
		Next
	EndIf
	Local $sDrive = "", $sDir = "", $FileName = "", $sExtension = ""
	Local $sNumCandidat = '000000'
	Local $Trouve = False
	Local $aTmp[1] = [0]
	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
			$aPathSplit = _PathSplit($Liste[$N], $sDrive, $sDir, $FileName, $sExtension)
			$aTmp = StringRegExp($FileName, "([0-9]{6})", $STR_REGEXPARRAYMATCH)
			If @error = 0 Then
				$sNumCandidat = $aTmp[0]
				$Trouve = True
				_Logging("Numéro trouvé: """ & $Liste[$N] & """.", 2, 0)
				ExitLoop
			EndIf
		Next
	EndIf

	If Not $Trouve Then _Logging("Aucun numéro trouvé.", 2, 0)
	SplashOff()
;~ 	WinActivate ($hMainGUI)
	Return $sNumCandidat
EndFunc   ;==>_NumeroCandidat

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _NumeroCandidatTic()
	SplashTextOn("Sans Titre", "Détermination du numéro d'inscription du candidat." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("Recherche du numéro d'inscription du candidat à partir des noms des sites web.", 2, 0)

	Local $Www = DossiersEasyPHPwww() ;
	Local $Liste[1] = [0] ;
	Local $ReturnPath = False
	Local $Flag = 2 ; 0:All, 1:Files only , 2:Fldrs only
	If $Www[0] <> 0 Then
		For $i = 1 To $Www[0]
			$Kes = _FileListToArray($Www[$i], "*", $Flag, $ReturnPath)
			If IsArray($Kes) Then
				$Liste[0] += $Kes[0]
				_ArrayDelete($Kes, 0)
				_ArrayAdd($Liste, $Kes) ;
			EndIf
		Next
	EndIf
	Local $FldrName
	Local $sNumCandidat = '000000'
	Local $Trouve = False
	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
			$FldrName = $Liste[$N]
			If StringRegExp($FldrName, "(.)*([0-9]{6}$)", 0) = 1 Then
				$sNumCandidat = StringRight($FldrName, 6)
				$Trouve = True
				_Logging("Numéro trouvé: """ & $Liste[$N] & """.", 2, 0)
				ExitLoop
			EndIf
		Next
	EndIf

	If Not $Trouve Then
		_Logging("Recherche du numéro d'inscription du candidat à partir des noms des bases de données.", 2, 0)
		Local $Data = DossiersEasyPHPdata() ;
		Local $Liste[1] = [0] ;
		Local $ReturnPath = False
		Local $Flag = 2 ; 0:All, 1:Files only , 2:Fldrs only
		If $Data[0] <> 0 Then
			For $i = 1 To $Data[0]
				$Kes = _FileListToArray($Data[$i], "*", $Flag, $ReturnPath)
				If IsArray($Kes) Then
					$Liste[0] += $Kes[0]
					_ArrayDelete($Kes, 0)
					_ArrayAdd($Liste, $Kes) ;
				EndIf
			Next
		EndIf
		Local $FldrName
		Local $sNumCandidat = '000000'
		If IsArray($Liste) Then
			For $N = 1 To $Liste[0]
				$FldrName = $Liste[$N]
				If StringRegExp($FldrName, "(.)*([0-9]{6}$)", 0) = 1 Then
					$sNumCandidat = StringRight($FldrName, 6)
					$Trouve = True
					_Logging("Numéro trouvé: """ & $Liste[$N] & """.", 2, 0)
					ExitLoop
				EndIf
			Next
		EndIf

	EndIf

	If Not $Trouve Then _Logging("Aucun numéro trouvé.", 2, 0)

	SplashOff()
;~ 	WinActivate ($hMainGUI)
	Return $sNumCandidat
EndFunc   ;==>_NumeroCandidatTic

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _GUICtrlRichEdit_WriteLine($hWnd, $sText, $iColor = -1, $iIncrement = 0, $sAttrib = "")

	; Count the @CRLFs
	StringReplace(_GUICtrlRichEdit_GetText($hWnd, True), @CRLF, "")
	Local $iLines = @extended
	; Adjust the text char count to account for the @CRLFs
	Local $iEndPoint = _GUICtrlRichEdit_GetTextLength($hWnd, True, True) - $iLines

	; Add new text
	_GUICtrlRichEdit_AppendText($hWnd, $sText & @CRLF)

	; Select text between old and new end points
	_GUICtrlRichEdit_SetSel($hWnd, $iEndPoint, -1)
	; Convert colour from RGB to BGR
	$iColor = Hex($iColor, 6)
	$iColor = '0x' & StringMid($iColor, 5, 2) & StringMid($iColor, 3, 2) & StringMid($iColor, 1, 2)
	; Set colour
	If $iColor <> -1 Then _GUICtrlRichEdit_SetCharColor($hWnd, $iColor)
	; Set size
	If $iIncrement <> 0 Then _GUICtrlRichEdit_ChangeFontSize($hWnd, $iIncrement)
	; Set weight
	If $sAttrib <> "" Then _GUICtrlRichEdit_SetCharAttributes($hWnd, $sAttrib)
	; Clear selection
	_GUICtrlRichEdit_Deselect($hWnd)
	_GUICtrlRichEdit_ScrollLines($hWnd, 1) ;$iLines)

EndFunc   ;==>_GUICtrlRichEdit_WriteLine

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _InitLogging()
	Global $hLogFile = FileOpen(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".log", 1)
	Local $sDateTime = "[" & _Now() & "]"
	FileWriteLine($hLogFile, '')
	FileWriteLine($hLogFile, '==============================================================================')
	FileWriteLine($hLogFile, "______" & $sDateTime & "______Démarrage de " & $PROG_TITLE & $PROG_VERSION)
	FileWriteLine($hLogFile, '')
	FileWriteLine($hLogFile, _SystemInfo())

	Local $DossierSauve = "Sauvegardes"
	Local $aDrive = DriveGetDrive('FIXED')

;~ 		_ArraySort($aDrive, 0, 1)
	$Lecteur = @HomeDrive ; "C:" ; $aDrive[1] ; $aDrive[1] peut être A: !!
	For $i = 1 To $aDrive[0]
		If $aDrive[$i] = @HomeDrive Then ContinueLoop
		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; to Exclude external Hdd(s)
				And _WinAPI_IsWritable($aDrive[$i]) _ ;writable
				And DriveSpaceFree($aDrive[$i] & "\") > $FREE_SPACE_DRIVE_BACKUP _
				Then
			$Lecteur = $aDrive[$i]
			ExitLoop
		EndIf
	Next
	$Lecteur = StringUpper($Lecteur) & "\"
;~ 	MsgBox(0,"",$Lecteur & $DossierSauve & "\0-CapturesEcran\BacBackup.ini")
	If _IsRegistryExist("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{498AA8A4-2CBE-4368-BFA0-E0CF3F338536}_is1", "DisplayName") _
			And ProcessExists('BacBackup.exe') _
			And FileExists($Lecteur & $DossierSauve & "\BacBackup\BacBackup.ini") Then
		Local $DossierSessionBacBackup = IniRead($Lecteur & $DossierSauve & "\BacBackup\BacBackup.ini", "Params", "DossierSession", "")
		If $DossierSessionBacBackup <> "" Then
			FileWriteLine($hLogFile, "BacBackup   : """ & $Lecteur & $DossierSauve & "\BacBackup\" & $DossierSessionBacBackup & """")
		EndIf
	EndIf


	FileWriteLine($hLogFile, '')
EndFunc   ;==>_InitLogging

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _CopierLog($DestLocalFldr)
	FileFlush($hLogFile)
	FileCopy(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".log", $DestLocalFldr)
EndFunc   ;==>_CopierLog

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _FinLogging()
	Local $sDateTime = "[" & _Now() & "]"
	FileWriteLine($hLogFile, "______" & $sDateTime & "______Fermeture de " & $PROG_TITLE & $PROG_VERSION)
	FileClose($hLogFile)
EndFunc   ;==>_FinLogging

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _Logging($sText, $iSuccess = 1, $iGuiLog = 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	If Not FileExists(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".log") Then
		_InitLogging()
	EndIf

	Local $sPreLogFile = ""
	If $iSuccess = 0 Then
		$sPreLogFile = "_Échec - "
	ElseIf $iSuccess = 1 Then
		$sPreLogFile = "Succès - "
	ElseIf $iSuccess = 2 Or $iSuccess = 3 Or $iSuccess = 4 Or $iSuccess = 5 Then
		$sPreLogFile = "______ - " ;info
	EndIf

	If $iGuiLog = 1 Then
		Local $sPreLogGui = "»»» "
		Local $Color = -1
		If $iSuccess = 1 Then
			$sPreLogGui = "- "
		EndIf
		Switch $iSuccess
			Case 0, 5
				$Color = 0xEE3E34 ;Red
			Case 1, 4
				$Color = -1 ;Normal
			Case 2
				$Color = 0x01A3FA ; Blue
			Case 3
				$Color = 0x60AC00 ;Green
		EndSwitch
		_GUICtrlRichEdit_WriteLine($GUI_Log, $sPreLogGui & $sText, $Color)
	EndIf
	Local $sTime = "[" & _NowTime(5) & "]"
	FileWriteLine($hLogFile, $sTime & " " & $sPreLogFile & $sText)
EndFunc   ;==>_Logging

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _ClearGuiLog()
	_GUICtrlRichEdit_SetText($GUI_Log, "")
EndFunc   ;==>_ClearGuiLog

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _UnLockFoldersBC()
;~ 	_UnLockFolder($Lecteur & $DossierSauve)
;~ 	_UnLockFolder($Lecteur & $DossierSauve & "\" & $DossierBacCollector)
;~ 	_LockFolder($Lecteur & $DossierSauve)
EndFunc   ;==>_UnLockFoldersBC

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _LockFoldersBC()
;~ 	_UnLockFolder($Lecteur & $DossierSauve)
;~ 	_LockFolder($Lecteur & $DossierSauve & "\" & $DossierBacCollector)
	_LockFolder($Lecteur & $DossierSauve)
EndFunc   ;==>_LockFoldersBC

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
	#forceref $iMsg, $wParam
	Local $tMsgFilter

	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
	Local $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
	Local $iCode = DllStructGetData($tNMHDR, "Code")
	;;;; Test
	Local $tagNMHDR = DllStructCreate("int;int;int", $lParam)
	Local $code = DllStructGetData($tagNMHDR, 3)
	Local $hListViewItem, $aListLine

	If $wParam = $GUI_AfficherLeContenuDesDossiersBac And $code = -3 Then
		$hListViewItem = GUICtrlRead($GUI_AfficherLeContenuDesDossiersBac)
		If $hListViewItem Then
			$aListLine = StringSplit(GUICtrlRead($hListViewItem), "|")
			_ShowInExplorer($aListLine[4])
		EndIf

	ElseIf $wParam = $GUI_AfficherLeContenuDesAutresDossiers And $code = -3 Then
		$hListViewItem = GUICtrlRead($GUI_AfficherLeContenuDesAutresDossiers)
		If $hListViewItem Then
			$aListLine = StringSplit(GUICtrlRead($hListViewItem), "|")
			_ShowInExplorer($aListLine[4])
		EndIf
	Else
		;;;; Test
		Switch $hWndFrom
			Case $GUI_Log
				Select
					Case $iCode = $EN_MSGFILTER
						$tMsgFilter = DllStructCreate($tagMSGFILTER, $lParam)
						If DllStructGetData($tMsgFilter, "msg") = $WM_LBUTTONDBLCLK Then
							_OpenTempLog()
						EndIf

				EndSelect
		EndSwitch
	EndIf

	Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _ShowInExplorer($sFileFolder)
	If Not FileExists($sFileFolder) Then
		MsgBox(16 + 262144, $PROG_TITLE & $PROG_VERSION, "Cet élément a été déplacé ou supprimé:" & @CRLF & @CRLF & """" & $sFileFolder & """", 0, $hMainGUI)
		Return
	EndIf

	If FileGetAttrib($sFileFolder) = 'D' Then ;c'est un dossier
		Run("explorer.exe /n, /e, " & '"' & $sFileFolder & '"')
	Else
		Run("explorer.exe /n, /e, /select, " & '"' & $sFileFolder & '"')
	EndIf
EndFunc   ;==>_ShowInExplorer

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _OpenTempLog()
	Local $Text = _GUICtrlRichEdit_GetText($GUI_Log)
	If StringLen($Text) = 0 Then Return
	Local $FilePath = @TempDir & "\" & $PROG_TITLE & "_tmp_log.rtf"
	Local $hTmpFile = FileOpen($FilePath, 2)
	Local $Kes = _GUICtrlRichEdit_StreamToVar($GUI_Log)
	$Kes = StringReplace($Kes, "\red255\green255\blue255", "\red0\green0\blue0")
	$Kes = StringReplace($Kes, "MS Shell Dlg", "Segoe UI")
	$Kes = StringReplace($Kes, "\fs17\", "\fs21\")

	FileWriteLine($hTmpFile, $Kes)
	FileClose($hTmpFile)
	_RunDos('Start Wordpad "' & $FilePath & '"')

EndFunc   ;==>_OpenTempLog


Func _CheckRunningFromUsbDrive()
	If FileExists(@ScriptDir & "\__dev.txt") Then Return
	Local $Drive = StringLeft(@ScriptDir, 2)

;~ _ExtMsgBox ($vIcon, $vButton, $sTitle, $sText, [$iTimeout, [$hWin, [$iVPos, [$bMain = True]]]])
	If (DriveGetType($Drive, $DT_BUSTYPE) <> "USB") Then
;~ 		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBoxSet(1, 0, 0x004080, 0xFFFF00, 9, "Comic Sans MS")
		_ExtMsgBox(128, "Ok", $PROG_TITLE & $PROG_VERSION, $PROG_TITLE & " est exécuté à partir du disque local." & @CRLF & @CRLF _
				 & "Le bouton ""Récupérer"" sera désactivé." & @CRLF, 12, $hMainGUI)
		GUICtrlSetState($bRecuperer, $GUI_DISABLE)
	ElseIf Not _WinAPI_IsWritable($Drive) Then
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS")
		_ExtMsgBox(128, "Ok", $PROG_TITLE & $PROG_VERSION, "Le lecteur """ & $Drive & "\"" n'est pas enregistrable, veuillez exécutr " & $PROG_TITLE & " à partir d'une Clé USB." & @CRLF & @CRLF _
				 & "Le bouton ""Récupérer"" sera désactivé." & @CRLF, 12, $hMainGUI)
		GUICtrlSetState($bRecuperer, $GUI_DISABLE)
	ElseIf DriveSpaceFree($Drive & "\") < 100 Then ;100 Mo
		_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS")
		_ExtMsgBox(128, "Ok", $PROG_TITLE & $PROG_VERSION, "L'espace disque dans la Clé USB est insuffisant pour la récupération des dossiers de travail des candidats.." & @CRLF & @CRLF _
				 & "Le bouton ""Récupérer"" sera désactivé." & @CRLF, 12, $hMainGUI)
		GUICtrlSetState($bRecuperer, $GUI_DISABLE)
	EndIf
EndFunc   ;==>_CheckRunningFromUsbDrive



Func _UnLockAll()
	Local $DossierSauve = "Sauvegardes"
	Local $Lecteur = ""

	;Contrôle ajouté après la Demo au Crefoc de Tunis 1
	Local $aDrive = DriveGetDrive('FIXED')
	Local $TmpMsg = ""

	For $i = 1 To $aDrive[0]
		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; to Exclude external Hdd(s)
				And _WinAPI_IsWritable($aDrive[$i]) _ ;writable
				Then
			$Lecteur = StringUpper($aDrive[$i] & "\")
			If FileExists($Lecteur & $DossierSauve) Then
				$TmpMsg &= "Déverrouillage du dossier [" & $Lecteur & $DossierSauve & "]" & @CRLF
				_UnLockFolder($Lecteur & $DossierSauve)
				FileSetAttrib($Lecteur & $DossierSauve, "-RASH")
			EndIf
		EndIf
	Next
	If RegDelete("HKCU\SOFTWARE\BacBackup") = 1 Then $TmpMsg &= @CRLF & "Suppression de l'entrée de BacCollector dans la base de registre (UUID)."
	$TmpMsg &= @CRLF & "Exécution terminée."
	MsgBox(262144, $PROG_TITLE & $PROG_VERSION, $TmpMsg)
EndFunc   ;==>_UnLockAll

Func _CheckBacCollectorExists()
	If Not FileExists(@ScriptFullPath) Then
		Local $sScriptDrive = StringLeft(@ScriptFullPath, StringLen(@ScriptFullPath) - StringLen(@ScriptName))
		If StringLen($sScriptDrive) = 3 And StringRegExp($sScriptDrive, "^([a-zA-Z]{1,1}:)?(\\|\/)$", 0) = 1 And Not FileExists($sScriptDrive) Then
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "Le lecteur """ & $sScriptDrive & """ où se trouve l'exécutable """ & @ScriptName & """ n'existe plus." & @CRLF _
					 & "Il est très important de ne pas retirer la Clé USB avant de quitter " & $PROG_TITLE & ", car cela pourrait endommager les dossiers déjà récupérés." & @CRLF  & @CRLF _
					 & $PROG_TITLE & " va se fermer immédiatement.", 0)
		Else
			_ExtMsgBoxSet(1, 0, 0x660000, 0xFFFFFF, 9, "Comic Sans MS", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & $PROG_VERSION, "L'exécutable """ & @ScriptName & """ n'existe plus dans son emplacement d'origine." & @CRLF _
					 & $PROG_TITLE & " va se fermer immédiatement.", 0)
		EndIf
		Exit
	EndIf
EndFunc

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

#Region "Générer Grille d'évaluation" -----------------------------------------------------------------------------------
Func _GenererGrilleEvaluation($aNumCandiats)
	If Not IsArray($aNumCandiats) Then Return 0
	Local $sSheet = ''
	Local $sFldr = _TmpGrilleVide()


	$sSheet = $sSheet & '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	$sSheet = $sSheet &  @CRLF & '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{482217F5-CA01-4A07-A820-71D429B33AFE}">'
	$sSheet = $sSheet &  @CRLF & '<dimension ref="A2:N24"/>'
	$sSheet = $sSheet &  @CRLF & '<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="85" zoomScaleNormal="85" workbookViewId="0"><selection activeCell="F2" sqref="F2:I4"/></sheetView></sheetViews>'
	$sSheet = $sSheet &  @CRLF & '<sheetFormatPr baseColWidth="10" defaultRowHeight="12.75" x14ac:dyDescent="0.2"/><cols><col min="1" max="1" width="3.85546875" customWidth="1"/><col min="2" max="2" width="11.7109375" customWidth="1"/><col min="3" max="12" width="11.28515625" customWidth="1"/><col min="13" max="13" width="7.42578125" customWidth="1"/><col min="14" max="14" width="11.28515625" customWidth="1"/></cols>'
	$sSheet = $sSheet &  @CRLF & '<sheetData>'
	$sSheet = $sSheet &  @CRLF & '<row r="2" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="B2" s="26" t="s"><v>5</v></c><c r="C2" s="16"><v>' & StringRight(GUICtrlRead($cSeance),1)& '</v></c><c r="D2" s="16"/><c r="E2" s="2"/><c r="F2" s="17"/><c r="G2" s="18"/><c r="H2" s="18"/><c r="I2" s="19"/><c r="J2" s="10"/><c r="K2" s="17"/><c r="L2" s="18"/><c r="M2" s="19"/><c r="N2" s="3"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="3" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="B3" s="26" t="s"><v>7</v></c><c r="C3" s="16"><v>' & StringRight(GUICtrlRead($cLabo),1)& '</v></c><c r="D3" s="16"/><c r="E3" s="3"/><c r="F3" s="20"/><c r="G3" s="21"/><c r="H3" s="21"/><c r="I3" s="22"/><c r="J3" s="9"/><c r="K3" s="20"/><c r="L3" s="21"/><c r="M3" s="22"/><c r="N3" s="3"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="4" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="B4" s="26" t="s"><v>6</v></c><c r="C4" s="15"><f ca="1">TODAY()</f><v>43625</v></c><c r="D4" s="15"/><c r="E4" s="3"/><c r="F4" s="23"/><c r="G4" s="24"/><c r="H4" s="24"/><c r="I4" s="25"/><c r="J4" s="9"/><c r="K4" s="23"/><c r="L4" s="24"/><c r="M4" s="25"/><c r="N4" s="3"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="5" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="B5" s="34"/><c r="C5" s="35"/><c r="D5" s="35"/><c r="E5" s="3"/><c r="F5" s="13"/><c r="G5" s="13"/><c r="H5" s="13"/><c r="I5" s="13"/><c r="J5" s="9"/><c r="K5" s="13"/><c r="L5" s="13"/><c r="M5" s="13"/><c r="N5" s="3"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="6" spans="1:14" ht="91.5" customHeight="1" x14ac:dyDescent="0.2"><c r="B6" s="36" t="s"><v>2</v></c><c r="C6" s="33"/><c r="D6" s="33"/><c r="E6" s="33"/><c r="F6" s="33"/><c r="G6" s="33"/><c r="H6" s="33"/><c r="I6" s="33"/><c r="J6" s="33"/><c r="K6" s="33"/><c r="L6" s="33"/><c r="M6" s="32" t="s"><v>0</v></c><c r="N6" s="4"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="7" spans="1:14" s="30" customFormat="1" ht="23.1" customHeight="1" x14ac:dyDescent="0.2"><c r="A7" s="27"/><c r="B7" s="28" t="s"><v>3</v></c><c r="C7" s="28"/><c r="D7" s="28"/><c r="E7" s="28"/><c r="F7" s="28"/><c r="G7" s="28"/><c r="H7" s="28"/><c r="I7" s="28"/><c r="J7" s="28"/><c r="K7" s="28"/><c r="L7" s="28"/><c r="M7" s="28" t="str"><f>IF(SUM(C7:L7)=0,"",SUM(C7:L7))</f><v/></c><c r="N7" s="29"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="8" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="A8" s="1"/><c r="B8" s="6"/><c r="C8" s="5"/><c r="D8" s="5"/><c r="E8" s="5"/><c r="F8" s="5"/><c r="G8" s="5"/><c r="H8" s="5"/><c r="I8" s="5"/><c r="J8" s="5"/><c r="K8" s="5"/><c r="L8" s="5"/><c r="M8" s="6"/><c r="N8" s="4"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="9" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="B9" s="11" t="s"><v>4</v></c><c r="C9" s="6"/><c r="D9" s="6"/><c r="E9" s="6"/><c r="F9" s="6"/><c r="G9" s="6"/><c r="H9" s="6"/><c r="I9" s="6"/><c r="J9" s="6"/><c r="K9" s="6"/><c r="L9" s="6"/><c r="M9" s="12" t="s"><v>1</v></c><c r="N9" s="11" t="str"><f t="shared" ref="N9:N24" si="0">B9</f><v>N° Inscri</v></c></row>'


	Local $Ligne, $CandidNum, $ExcelLineNumber, $NumAutoCandid, $StyleCells
	Local $CellsErrorXL_Range =''
	Local $CellsErrorXL_AsText =''
	$ExcelLineNumber = 9 ;Le numéro de la première

	For $i = 0 To UBound($aNumCandiats)-1
		$CandidNum = $aNumCandiats[$i][0]
		$NumAutoCandid = $i + 1
		$ExcelLineNumber = $ExcelLineNumber + 1

		If StringInStr($aNumCandiats[$i][1], "Abs") Then
			$Note = "<v>88.88</v>"
			$CellsErrorXL_AsText = $CellsErrorXL_AsText & "M" & $ExcelLineNumber & " "
			$StyleCells = "37"
		Else
			$Note = '<f>IF(SUM(C%ExcelLineNumber:K%ExcelLineNumber)=0,"",SUM(C%ExcelLineNumber:K%ExcelLineNumber))</f><v/>'
			$CellsErrorXL_Range = $CellsErrorXL_Range & "M" & $ExcelLineNumber & " "
			$StyleCells = "7"
		EndIf

		$Ligne ='<row r="%ExcelLineNumber" spans="1:14" ht="23.1" customHeight="1" x14ac:dyDescent="0.2"><c r="A%ExcelLineNumber" s="31"><v>' & $NumAutoCandid & '</v></c><c r="B%ExcelLineNumber" s="14"><v>' & $CandidNum & '</v></c><c r="C%ExcelLineNumber" s="' & $StyleCells & '"/><c r="D%ExcelLineNumber" s="' & $StyleCells & '"/><c r="E%ExcelLineNumber" s="' & $StyleCells & '"/><c r="F%ExcelLineNumber" s="' & $StyleCells & '"/><c r="G%ExcelLineNumber" s="' & $StyleCells & '"/><c r="H%ExcelLineNumber" s="' & $StyleCells & '"/><c r="I%ExcelLineNumber" s="' & $StyleCells & '"/><c r="J%ExcelLineNumber" s="' & $StyleCells & '"/><c r="K%ExcelLineNumber" s="' & $StyleCells & '"/><c r="L%ExcelLineNumber" s="' & $StyleCells & '"/><c r="M%ExcelLineNumber" s="8" t="str">' & $Note & '</c><c r="N%ExcelLineNumber" s="14"><f t="shared" si="0"/><v>' & $CandidNum & '</v></c></row>'
		$sSheet = $sSheet & @CRLF & StringReplace($Ligne, "%ExcelLineNumber", $ExcelLineNumber)
		If $i = 14 Then ExitLoop
	Next

	$CellsErrorXL_AsText = StringStripWS($CellsErrorXL_AsText, 2) ;$STR_STRIPTRAILING (2) = strip trailing white space
	$CellsErrorXL_Range = StringStripWS($CellsErrorXL_Range, 2) ;$STR_STRIPTRAILING (2) = strip trailing white space

	$sSheet = $sSheet &  @CRLF & '</sheetData>'
	$sSheet = $sSheet &  @CRLF & '<mergeCells count="5"><mergeCell ref="C2:D2"/><mergeCell ref="F2:I4"/><mergeCell ref="K2:M4"/><mergeCell ref="C3:D3"/><mergeCell ref="C4:D4"/></mergeCells>'
	$sSheet = $sSheet &  @CRLF & '<printOptions horizontalCentered="1"/>'
	$sSheet = $sSheet &  @CRLF & '<pageMargins left="0.19685039370078741" right="0.19685039370078741" top="0.19685039370078741" bottom="0.19685039370078741" header="3.937007874015748E-2" footer="0.19685039370078741"/>'
	$sSheet = $sSheet &  @CRLF & '<pageSetup paperSize="9" orientation="landscape" r:id="rId1"/>'
	$sSheet = $sSheet &  @CRLF & '<headerFooter><oddFooter>&amp;R&amp;"Cambria,Normal"&amp;8Grille d''évaluation générée par &amp;"Cambria,Gras"' & $PROG_TITLE & '&amp;"Cambria,Normal"' & $PROG_VERSION & '</oddFooter></headerFooter>'
	$sSheet = $sSheet &  @CRLF & '<ignoredErrors><ignoredError sqref="' & $CellsErrorXL_Range & '" formulaRange="1"/><ignoredError sqref="' & $CellsErrorXL_AsText & '" numberStoredAsText="1" formulaRange="1"/></ignoredErrors>'
	$sSheet = $sSheet &  @CRLF & '</worksheet>'

	Local $FileSheet = FileOpen(@TempDir & "\BacCollector\" & $sFldr & "\xl\worksheets\sheet1.xml", 256 + 2) ;$FO_UTF8_NOBOM + $FO_OVERWRITE
	FileWrite($FileSheet, $sSheet)
	FileClose($FileSheet) ; Close the open file after writing

	$sSheetName = GUICtrlRead($cSeance) & ' ' & GUICtrlRead($cLabo)

	Local $sWorkbook = ''
	$sWorkbook = $sWorkbook &  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	$sWorkbook = $sWorkbook &  @CRLF & '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="21001"/><workbookPr/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="0" documentId="13_ncr:1_{33728A07-704C-4A9F-86E5-EBF9C54C3A43}" xr6:coauthVersionLast="38" xr6:coauthVersionMax="38" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/><bookViews><workbookView xWindow="120" yWindow="75" windowWidth="13995" windowHeight="10230" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews><sheets><sheet name="' & $sSheetName & '" sheetId="12" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">''' & $sSheetName & '''!$A$1:$N$24</definedName></definedNames><calcPr calcId="191029"/></workbook>'
	Local $FileWorkbook = FileOpen(@TempDir & "\BacCollector\" & $sFldr & "\xl\workbook.xml", 256 + 2) ;$FO_UTF8_NOBOM + $FO_OVERWRITE
	FileWrite($FileWorkbook, $sWorkbook)
	FileClose($FileWorkbook) ; Close the open file after writing

	Local $sApp = ''
	$sApp = $sApp &  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	$sApp = $sApp &  @CRLF & '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Feuilles de calcul</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Plages nommées</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>' & $sSheetName & '</vt:lpstr><vt:lpstr>''' & $sSheetName & '''!Zone_d_impression</vt:lpstr></vt:vector></TitlesOfParts><Manager>Moez Romdhane</Manager><Company>La Communauté des Enseignants d''Informatique en Tunisie</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinkBase>https://www.facebook.com/groups/InfoTun</HyperlinkBase><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>'
	Local $FileApp = FileOpen(@TempDir & "\BacCollector\" & $sFldr & "\docProps\app.xml", 256 + 2) ;$FO_UTF8_NOBOM + $FO_OVERWRITE
	FileWrite($FileApp, $sApp)
	FileClose($FileApp) ; Close the open file after writing

;~ 	Run("explorer.exe /n, /e, " & '"' & @TempDir & "\BacCollector\" & $sFldr & '"')
;~ ************** Zip avec zipfldr.dll
	Local $Zip = _Zip_Create(@TempDir & "\BacCollector\Grille.zip") ;Create The Zip File. Returns a Handle to the zip File
	_Zip_AddFolderContents($Zip, @TempDir & "\BacCollector\" & $sFldr) ;Add a folder's content in the zip file


;~ ************** Zip avec external 7-zip32.dll/7-zip64.dll
;~ 	Local $ArcFile = @TempDir & "\BacCollector\Grille.zip"
;~ 	Local $FolderName = @TempDir & "\BacCollector\" & $sFldr & "\*"

;~ 	If _7ZipStartup() Then ; To test dll opening (can be omitted for some functions)
;~ 		$retResult = _7ZipAdd(0, $ArcFile, $FolderName)
;~ 		If @error Then
;~ 			MsgBox(16, "Erreur Grille d'évaluation", "Erreur lors de la préparation de la Grille d'évaluation")
;~ 		EndIf
;~ 	Else
;~ 		MsgBox(16, "Erreur Grille d'évaluation", "Erreur lors de la préparation de la Grille d'évaluation")
;~ 	EndIf
;~ 	_7ZipShutdown()



	DirRemove(@TempDir & "\BacCollector\" & $sFldr, 1)
	If @error Then
		FileDelete(@TempDir & "\BacCollector\Grille.zip")
		Return 0
	EndIf

	Local $sCheminScripDir = @ScriptDir
	If StringRight($sCheminScripDir, 1) <> "\" Then
		$sCheminScripDir = $sCheminScripDir & "\"
	EndIf

	Local $KesNomGrille ="GrilleEval_" & GUICtrlRead($cSeance) & "_" &  GUICtrlRead($cLabo) & ".xlsx"
	Local $fileNumber = 1
	While FileExists($sCheminScripDir & $KesNomGrille)
		$KesNomGrille ="GrilleEval_" & GUICtrlRead($cSeance) & "_" &  GUICtrlRead($cLabo) & "_(" & $fileNumber & ")_" & @HOUR & "h" & @MIN & ".xlsx"
		$fileNumber += 1
	WEnd

	If Not FileMove( @TempDir & "\BacCollector\Grille.zip", $sCheminScripDir & $KesNomGrille , 1 ) Then
		MsgBox(16, "Erreur Grille d'évaluation", "Erreur lors de la création de la Grille d'évaluation")
		Return 0
	Else
		Return $KesNomGrille
	EndIf
EndFunc;==>_GenererGrilleEvaluation

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦

Func _TmpGrilleVide()
    $sFldr = StringFormat('%04X_%04X', _
            Random(0, 0xffff), _
            BitOR(Random(0, 0x3fff), 0x8000) _
        )
	DirCreate(@TempDir & "\BacCollector\" & $sFldr)
	FileInstall(".\Grille\[Content_Types].xml", @TempDir & "\BacCollector\" & $sFldr & "\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\_rels")
	FileInstall(".\Grille\_rels\.rels", @TempDir & "\BacCollector\" & $sFldr & "\_rels\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\docProps")
;~ 	FileInstall(".\Grille\docProps\app.xml", @TempDir & "\BacCollector\" & $sFldr & "\docProps\")
	FileInstall(".\Grille\docProps\core.xml", @TempDir & "\BacCollector\" & $sFldr & "\docProps\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\xl")
	FileInstall(".\Grille\xl\calcChain.xml", @TempDir & "\BacCollector\" & $sFldr & "\xl\")
	FileInstall(".\Grille\xl\sharedStrings.xml", @TempDir & "\BacCollector\" & $sFldr & "\xl\")
	FileInstall(".\Grille\xl\styles.xml", @TempDir & "\BacCollector\" & $sFldr & "\xl\")
;~ 	FileInstall(".\Grille\xl\workbook.xml", @TempDir & "\BacCollector\" & $sFldr & "\xl\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\xl\_rels")
	FileInstall(".\Grille\xl\_rels\workbook.xml.rels", @TempDir & "\BacCollector\" & $sFldr & "\xl\_rels\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\xl\printerSettings")
	FileInstall(".\Grille\xl\printerSettings\printerSettings1.bin", @TempDir & "\BacCollector\" & $sFldr & "\xl\printerSettings\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\xl\theme")
	FileInstall(".\Grille\xl\theme\theme1.xml", @TempDir & "\BacCollector\" & $sFldr & "\xl\theme\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\xl\worksheets")
;~ 	FileInstall(".\Grille\xl\worksheets\sheet1.xml", @TempDir & "\BacCollector\" & $sFldr & "\xl\worksheets\")

	DirCreate(@TempDir & "\BacCollector\" & $sFldr & "\xl\worksheets\_rels")
	FileInstall(".\Grille\xl\worksheets\_rels\sheet1.xml.rels", @TempDir & "\BacCollector\" & $sFldr & "\xl\worksheets\_rels\")

	Return $sFldr

EndFunc   ;==>_TmpGrilleVide

#EndRegion "Générer Grille d'évaluation" --------------------------------------------------------------------------------

;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦
;~ _ArrayDisplay($TmpList, $PROG_TITLE ,"",32,Default ,"Liste de Dossiers/Fichiers Surveillés")
;¦¦¦
;~ MsgBox ( 0, "", $sDir  )
;¦¦¦
