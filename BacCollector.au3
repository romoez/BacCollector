#NoTrayIcon
#RequireAdmin

#Region ;**** AutoIt3Wrapper ****
#AutoIt3Wrapper_Run_AU3Check=n
#AutoIt3Wrapper_Tidy_Stop_OnError=n
#AutoIt3Wrapper_Run_Au3Stripper=y
#AutoIt3Wrapper_Res_Field=URL|https://github.com/romoez/BacCollector
#AutoIt3Wrapper_Res_Field=Email|moez.romdhane@tarbia.tn
#AutoIt3Wrapper_Res_Language=1036

#AutoIt3Wrapper_Run_After=copy "%out%" ".\Tests\Labo00"
#AutoIt3Wrapper_Run_After=copy "%out%" ".\Tests\Labo01"
#AutoIt3Wrapper_Run_After=copy "%out%" ".\Tests\Labo02"
#AutoIt3Wrapper_Run_After=copy "%out%" "D:\14 SharedFolder\00-BacCollector"
#EndRegion ;**** AutoIt3Wrapper ****

#Region ;**** pragma ****
#pragma compile(Icon, BacCollector.ico)
#pragma compile(Out, BacCollector.exe)
;~ #pragma compile(UPX, False) ;
;~ #pragma compile(Compression, 9)
#pragma compile(FileDescription, Collecte des travaux lors des examens pratiques d'informatique.)
#pragma compile(ProductName, BacCollector)
#pragma compile(ProductVersion, 1.0.26.205)
#pragma compile(FileVersion, 1.0.26.205)
#pragma compile(LegalCopyright, 2018-2026 © Communauté Tunisienne des Enseignants d'Informatique)
#pragma compile(Comments, BacCollector: Collecte des travaux lors des examens pratiques d'informatique.)
#pragma compile(CompanyName, Communauté Tunisienne des Enseignants d'Informatique)
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
#include <GUIToolTip.au3>
#include <ListViewConstants.au3>
#include <Misc.au3> ; Pour _IsPressed()
#include <Process.au3>
#include <StaticConstants.au3>
#include <WinAPIFiles.au3>
#include <WinAPIProc.au3>
#include <WinAPIShellEx.au3>
#include <WindowsConstants.au3>
#include "Utils.au3"
#include "Include\7Zip.au3"  ; https://www.autoitscript.fr/forum/viewtopic.php?f=21&t=1943
#include "Include\ExtMsgBox.au3"  ; https://www.autoitscript.com/forum/topic/109096-extended-message-box-new-version-19-nov-21/
#include "Include\GUIExtender.au3"  ; https://www.autoitscript.com/forum/topic/117909-guiextender-original-version/
#include "Include\MPDF_UDF.au3" ; https://www.autoitscript.com/forum/topic/118827-create-pdf-from-your-application/
#include "Include\StringSize.au3"  ; https://www.autoitscript.com/forum/topic/114034-stringsize-m23-new-version-16-aug-11/
;~ #include "include\Zip.au3"  ; https://www.autoitscript.com/forum/topic/73425-zipau3-udf-in-pure-autoit/
;~ #include "CheckSumVerify2.a3x"  ; https://www.autoitscript.com/forum/topic/164148-checksumverify-verify-integrity-of-the-compiled-exe/

#include <Globals.au3>
#include <Modules\Export.au3>
#include <Modules\App_GUI.au3>
#EndRegion ;**** include ****

_KillOtherScript()


If $CMDLINE[0] Then
	Local $sPattern = '(?i)^/(?:unlock|deblo|clea|nett|deverr|déblo|déverr)'   ;/Debloquer /Deblocage /Unlock /Clean /CleanUp /Clear /Unlock Deverrouillage Deverrouiller Nettoyage Nettoyer...
	Select
		Case StringRegExp($CMDLINE[1], $sPattern)
			_UnLockAll()
			Exit
		Case StringInStr(StringLower($CMDLINE[1]), "help") Or StringInStr(StringLower($CMDLINE[1]), "aide") _
				Or StringInStr($CMDLINE[1], "?")
			MsgBox(262144, $PROG_TITLE & $PROG_VERSION, "Syntaxe de ligne de commande: " & @CRLF & @CRLF _
					 & "» Aifficher l'aide : " & @CRLF _
					 & @TAB & "- BacCollector /?, ou:" & @CRLF _
					 & @TAB & "- BacCollector /help, ou:" & @CRLF _
					 & @TAB & "- BacCollector /aide, ou:" & @CRLF _
					 & "» Déverrouillage des dossiers ""X:\Sauvegardes\"" : " & @CRLF _
					 & @TAB & "- BacCollector /déverrouiller, ou:" & @CRLF _
					 & @TAB & "- BacCollector /unlock, ou encore:" & @CRLF _
					 & @TAB & "- BacCollector /débloquer" & @CRLF & @CRLF _
					 & "» Récupération automatique des dossiers de travail du candidat: " & @CRLF _
					 & @TAB & "- [Cette fonctionnalité n'est pas encore implémentée.]") ; & @CRLF _
			Exit
		Case StringInStr(StringLower($CMDLINE[1]), "run_as_admin")
			If $CMDLINE[0] > 1 Then
				$sUserName = $CMDLINE[2]
			EndIf
		Case Else
			MsgBox(262144, $PROG_TITLE & $PROG_VERSION, "Syntaxe de ligne de commande: " & @CRLF & @CRLF _
					 & "» Aifficher l'aide : " & @CRLF _
					 & @TAB & "- BacCollector /?, ou:" & @CRLF _
					 & @TAB & "- BacCollector /help, ou:" & @CRLF _
					 & @TAB & "- BacCollector /aide, ou:" & @CRLF _
					 & "» Déverrouillage des dossiers ""X:\Sauvegardes\"" : " & @CRLF _
					 & @TAB & "- BacCollector /déverrouiller, ou:" & @CRLF _
					 & @TAB & "- BacCollector /unlock, ou encore:" & @CRLF _
					 & @TAB & "- BacCollector /débloquer" & @CRLF & @CRLF _
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
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
WinActivate($hMainGUI)

While 1

	$aMsg = GUIGetMsg(1)
	Switch $aMsg[0]
		Case $GUI_EVENT_CLOSE
			_FinLogging()
			_GUICtrlRichEdit_Destroy($GUI_Log)
			Exit
		Case $bTogglePart3
			_CheckBacCollectorExists()
			_TogglePart3()
		Case $bRecuperer
			_CheckBacCollectorExists()
			GUISetState(@SW_DISABLE, $hMainGUI) ;
			$Ret = Recuperer()
			GUISetState(@SW_ENABLE, $hMainGUI) ;
			If $Ret <> 0 Then
				_Initialisation(0) ; ->0:Ne pas initialiser le numero de candidat ->1:init
			Else
				_Initialisation(1) ; ->0:Ne pas initialiser le numero de candidat ->1:init
			EndIf
			WinActivate($hMainGUI)
		Case $bCreerSauvegarde
			_CheckBacCollectorExists()
			GUISetState(@SW_DISABLE, $hMainGUI) ;
			Sauvegarder()
			GUISetState(@SW_ENABLE, $hMainGUI) ;
			WinActivate($hMainGUI)
;~ 			WinActivate($hMainGUI) déplacé vers la fonction, si affich Rapport, on n'acive pas la fenêtre
		Case $bOpenBackupFldr
			_CheckBacCollectorExists()
			_InitialParams()
			WinActivate($hMainGUI)
			If Not FileExists($DossierBacCollector) Then _InitialParams()
			Run("explorer.exe /e, " & '"' & $DossierBacCollector & '"')
			_Logging("Ouverture du dossier de sauvegarde : " & @CRLF & "                            " _
					 & """" & $DossierBacCollector & """", 2, 0)

		Case $lblBacBackup
			_CheckBacCollectorExists()
			_OpenBacBackupInterface()
		Case $lblUsbCleaner
			_CheckBacCollectorExists()
			_OpenUsbCleanerUrl()
		Case $bAPropos
			_CheckBacCollectorExists()
			_ShowAPropos()
		Case $rInfoProg, $tInfoProg
			_CheckBacCollectorExists()
			If $Matiere <> 'InfoProg' Then
				_Logging("Changement de la matière : ""STI"" --> ""Info/Prog""", 2, 0)
				GUICtrlSetState($rInfoProg, $GUI_CHECKED)
				$Matiere = 'InfoProg'
				_InitialisationInfo()
			EndIf

		Case $rSti, $tSti
			_CheckBacCollectorExists()
			If $Matiere <> 'STI' Then
				_Logging("Changement de la matière : ""Info/Prog"" --> ""STI""", 2, 0)
				GUICtrlSetState($rSti, $GUI_CHECKED)
				$Matiere = 'STI'
				_InitialisationSti()
			EndIf

		Case $lblMatiere
			_CheckBacCollectorExists()
			_Logging("Actualisation des données suite au clic sur le label de la matière : " & $Matiere, 2, 0)
			_Initialisation()

		Case $cBac
			_CheckBacCollectorExists()
			_SaveData_Bacxx()
		Case $bCreerDossierCandidatAbs
			If _OpenDialogAbsent() Then
				_AfficherDossierRecuperes()
			EndIf
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
;~ 		Case $TextDossiersRecuperes
;~ 			Run("explorer.exe " & '"' & @ScriptDir & '"')

	EndSwitch
    ; Détection du clic molette (MBUTTON) sur n'importe où
    If _IsPressed("04") Then ; "04" = Code virtuel de la molette (sans WinAPIvK.au3)
        If Not $g_bDragging Then
            $g_bDragging = True
			$g_iGUITransparence = 153
			WinSetTrans($hMainGUI, "", $g_iGUITransparence) ; 60% de transparence (153/255)
            Local $aMousePos = MouseGetPos()
            Local $aWinPos = WinGetPos($hMainGUI)
            $g_iOffsetX = $aMousePos[0] - $aWinPos[0]
            $g_iOffsetY = $aMousePos[1] - $aWinPos[1]
        EndIf
        ; Déplacement fluide
        Local $aNewMousePos = MouseGetPos()
        WinMove($hMainGUI, "", $aNewMousePos[0] - $g_iOffsetX, $aNewMousePos[1] - $g_iOffsetY)
    Else
		If $g_iGUITransparence <> 255 Then
			$g_iGUITransparence = 255
			WinSetTrans($hMainGUI, "", $g_iGUITransparence)
		EndIf
        $g_bDragging = False
    EndIf

    Sleep(10) ; Réduire l'utilisation du CPU

WEnd

;=========================================================

Func _Initialisation($initNumeroCandidat = 1)
	If $Matiere = "STI" Then
		_InitialisationSti($initNumeroCandidat)
	Else
		_InitialisationInfo($initNumeroCandidat)
	EndIf
EndFunc   ;==>_Initialisation

;=========================================================

#Region Function "_InitialisationInfo" ------------------------------------------------------------------------
Func _InitialisationInfo($initNumeroCandidat = 1)
	SplashTextOn("Sans Titre", "Initialisation de l'opération." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	GUICtrlSetData($HeadContenuDossiersBac, "Contenu des dossiers 'Bac*2*':")
	GUICtrlSetData($HeadContenuAutresDossiers, "Autres Fichiers dans ""Mes Documents"", ""Bureau"", racine des lecteurs du hdd :")

	GUICtrlSetTip($HeadContenuDossiersBac, " ", "Contenu des dossiers 'Bac*2*'", 1,1)
	GUICtrlSetTip($HeadContenuAutresDossiers, "- Bureau" & @CRLF & "- Mes documents" & @CRLF & "- Téléchargements" & @CRLF & "- Dossier du profil de l'utilisateur" & @CRLF & "- Racines des lecteurs C:, D: ..." & @CRLF & @CRLF &  "(Double-clic sur un élément pour l'afficher dans l'Explorateur.)", "Éléments récemment créés/modifiés dans :", 1,1)

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
			GUICtrlSetState($rSti, $GUI_UNCHECKED)
			GUICtrlSetData($lblMatiere, "Info/Prog")
			$sFiltreFichiersAChercher = "*"
		Case 'STI'
			GUICtrlSetState($rInfoProg, $GUI_UNCHECKED)
			GUICtrlSetState($rSti, $GUI_CHECKED)
			GUICtrlSetData($lblMatiere, "STI")
			$sFiltreFichiersAChercher = "*"
	EndSwitch
	;;;Section - Fin ===================================
	Local $sTmpSpaces = "                            "
	Local $iStart = TimerInit()
	$sApps = _ListeDApplicationsOuvertes()
	If $sApps <> "" Then
		_Logging("Applications ouvertes: " & StringReplace(@CRLF & $sApps, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
		GUICtrlSetData($TextApps, $sApps)
	Else
		_Logging("Aucune application ouverte.", 2, 0, TimerDiff($iStart))
		GUICtrlSetData($TextApps, "-")
	EndIf

	Local $sDossiersRecuperes = _ListeDeDossiersRecuperes()
	_AfficherDossierRecuperes(1) ; $sLog = 0 / 1
	_AfficherLeContenuDesDossiersBac()
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesDossiersBac, 3) ;Hide FullPath column
	_AfficherLeContenuDesAutresDossiers($sFiltreFichiersAChercher)
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesAutresDossiers, 3) ;Hide FullPath column

	GUISetState(@SW_ENABLE, $hMainGUI) ;
	WinActivate ($hMainGUI);
	_VerifierDossiersNonConformesPourCandidatsAbsents()
EndFunc   ;==>_InitialisationInfo
#EndRegion Function "_InitialisationInfo" ------------------------------------------------------------------------

;=========================================================

#Region Function "_InitialisationSti" ------------------------------------------------------------------------
Func _InitialisationSti($initNumeroCandidat = 1)
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
		GUICtrlSetData($GUI_NumeroCandidat, _NumeroCandidatSti())
	EndIf

	IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "Matiere", $Matiere)
	Switch $Matiere
		Case 'InfoProg'
			GUICtrlSetState($rInfoProg, $GUI_CHECKED)
			GUICtrlSetState($rSti, $GUI_UNCHECKED)
			GUICtrlSetData($lblMatiere, "Info/Prog")
			$sFiltreFichiersAChercher = "*"
		Case 'STI'
			GUICtrlSetState($rInfoProg, $GUI_UNCHECKED)
			GUICtrlSetState($rSti, $GUI_CHECKED)
			GUICtrlSetData($lblMatiere, "STI")
			$sFiltreFichiersAChercher = "*"
	EndSwitch
	Local $iStart = TimerInit()
	$sApps = _ListeDApplicationsOuvertes()
	If $sApps <> "" Then
		Local $sTmpSpaces = "                            "
		_Logging("Applications ouvertes: " & StringReplace(@CRLF & $sApps, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
		GUICtrlSetData($TextApps, $sApps)
	Else
		_Logging("Aucune application ouverte.", 2, 0, TimerDiff($iStart))
		GUICtrlSetData($TextApps, "-")
	EndIf

	_AfficherDossierRecuperes()
	_AfficherLeContenuDesDossiersSti()
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesDossiersBac, 3) ;Hide FullPath column
	_GUICtrlListView_HideColumn($GUI_AfficherLeContenuDesAutresDossiers, 3) ;Hide FullPath column
	GUISetState(@SW_ENABLE, $hMainGUI) ;
	WinActivate ($hMainGUI);
	_VerifierDossiersNonConformesPourCandidatsAbsents()
EndFunc   ;==>_InitialisationSti
#EndRegion Function "_InitialisationSti" ------------------------------------------------------------------------

;=========================================================

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
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", """" & $NumeroCandidat & """ n'est pas un numéro d'inscription valide", 0)
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
	Local $iStart = TimerInit()
	Local $sApps = _ListeDApplicationsOuvertes()
	If $sApps <> "" Then
		Local $sTmpSpaces = "                            "
		_Logging("Applications ouvertes: " & StringReplace(@CRLF & $sApps, @CRLF, @CRLF & $sTmpSpaces), 5, 1, TimerDiff($iStart))
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		GUICtrlSetData($TextApps, $sApps)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Avant de ""Récupérer"" le travail du candidat, veuillez fermer ce(s) logiciel(s) : " & @CRLF & @CRLF & $sApps, 0)
		Return 1
	Else
		_Logging("Aucune application.", 2, 0, TimerDiff($iStart))
		GUICtrlSetData($TextApps, "-")
	EndIf

	If $Matiere = "InfoProg" Then
		Local $iRes = RecupererInfo($NumeroCandidat)
	Else
		Local $iRes = RecupererSti($NumeroCandidat)
	EndIf
	Return $iRes
EndFunc   ;==>Recuperer

; =========================================================
; Fonction: _VerifierExistenceDossierDestination
; Description: Vérifie si un dossier de destination existe déjà
; Paramètres:
;   $NumeroCandidat - Numéro du candidat
; Retour: True si le dossier n'existe pas (OK pour continuer), False sinon
; =========================================================
Func _VerifierExistenceDossierDestination($NumeroCandidat)
	_Logging("Vérification si un dossier de même nom """ & $NumeroCandidat & """ existe déjà dans la Clé Usb", 2, 0)
	SplashTextOn("Sans Titre", "Vérification si un dossier de même nom """ & $NumeroCandidat & """ existe déjà dans la Clé Usb. " & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	$ScriptDir = @ScriptDir
	If StringRight($ScriptDir, 1) <> "\" Then
		$ScriptDir = $ScriptDir & "\"
	EndIf

	Global $Dest1FlashUSB = $ScriptDir & $NumeroCandidat
	Global $Dest2LocalFldr = $DossierBacCollector & "\" & $NumeroCandidat & _
				"_____" & @YEAR & "-" & @MON & "-" & @MDAY & "__" & @HOUR & "-" & @MIN

	If FileExists($Dest1FlashUSB) Then
		_Logging("Un dossier de même nom """ & $NumeroCandidat & """ existe déjà sur la Clé USB!!", 5, 1)
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		SplashOff()
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Un dossier de même nom """ & $NumeroCandidat & """ existe déjà sur la Clé USB!!", 0)
		Return False
	EndIf

	SplashOff()
	Return True
EndFunc   ;==>_VerifierExistenceDossierDestination

; =========================================================
; Fonction: _CreerDossiersDestination
; Description: Crée les dossiers de destination sur USB et disque local
; Retour: Nombre d'erreurs (0 si succès, -1 si échec critique)
; =========================================================
Func _CreerDossiersDestination()
	SplashTextOn("Sans Titre", "Création des dossiers de Sauvegarde." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	Local $Error0_CreationDossier = DirCreate($Dest1FlashUSB)
	Local $TmpError = DirCreate($Dest2LocalFldr)

	If $Error0_CreationDossier = 0 Then
		SplashOff()
		_Logging("Création du Dossier: " & $Dest1FlashUSB, $Error0_CreationDossier)
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec lors de la création du dossier sur la Clé USB" & @CRLF _
				 & "L'opération de récupération est annulée" & @CRLF _
				 & "Veuillez vérifier si la Clé USB n'est pas pleine ou protégée en écriture," & @CRLF _
				 & "ou si un Antivirus bloque " & $PROG_TITLE & $PROG_VERSION & @CRLF _
				 & "puis relancer l'opération.", 0)
		Return -1
	EndIf

	Local $iNberreurs = 0
	_Logging("Création du Dossier: " & $Dest1FlashUSB)
	_Logging("Création du Dossier: " & $Dest2LocalFldr, $TmpError)
	$iNberreurs = $iNberreurs - ($TmpError - 1)

	SplashOff()
	Return $iNberreurs
EndFunc   ;==>_CreerDossiersDestination

; =========================================================
; Fonction: _VerifierDossiersBac
; Description: Vérifie l'existence et le contenu des dossiers Bac*2*
; Retour: Array des dossiers Bac ou False en cas d'échec/annulation
; =========================================================
Func _VerifierDossiersBac()
	SplashTextOn("Sans Titre", "Recherche des dossiers ""Bac*2*"" sous la racine ""C:""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("Recherche des dossiers ""Bac*2*"" sous la racine ""C:""", 2, 0)
	Local $Bac = DossiersBac()

	If $Bac[0] = 0 Then
		SplashOff()
		_Logging("Aucun dossier ""Bac*2*"" trouvé sous la racine ""C:""", 5, 1)
		_Logging("Récupération annulée", 5, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Aucun dossier ""Bac20xx"" sous la racine ""C:"" !" & @CRLF & "", 0)
		Return False
	EndIf

	SplashTextOn("Sans Titre", "Vérification des dossiers ""Bac*2*"" sous la racine ""C:""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("Vérification du contenu du(es) dossier(s) : " & _ArrayToString($Bac, ", ", 1), 2, 0)

	Local $NbDossiersBacNonVides = 0
	Local $NomsDossiers = ""
	For $i = 1 To $Bac[0]
		$NomsDossiers = $NomsDossiers & " """ & $Bac[$i] & """"
		$sizefldr1 = DirGetSize($Bac[$i], 1)
		If Not @error And $sizefldr1[1] Then
			$NbDossiersBacNonVides += 1
		EndIf
	Next

	If Not $NbDossiersBacNonVides Then
		_Logging("Aucun fichier trouvé dans le(s) dossier(s) bac : " & _ArrayToString($Bac, ",", 1), 5, 1)
		SplashOff()

		Local $message
		If $Bac[0] = 1 Then
			$message = _
				"DOSSIER BAC VIDE" & @CRLF & @CRLF & _
				"Le dossier " & $NomsDossiers & " ne contient aucun travail." & @CRLF & @CRLF & _
				"OPTIONS :" & @CRLF & _
				"1. Récupération manuelle :" & @CRLF & _
				"   • Déplacer les fichiers du candidat vers " & $NomsDossiers & @CRLF & _
				"   • OU créer un fichier d'alerte dans ce dossier" & @CRLF & @CRLF & _
				"2. Génération automatique :" & @CRLF & _
				"   ► [Générer et Continuer] : Crée le fichier d'alerte et poursuit l'extraction" & @CRLF
		Else
			$message = _
				"TOUS LES DOSSIERS BACS SONT VIDES" & @CRLF & @CRLF & _
				"Dossiers vides : " & $NomsDossiers & @CRLF & @CRLF & _
				"OPTIONS :" & @CRLF & _
				"1. Récupération manuelle :" & @CRLF & _
				"   • Déplacer les fichiers vers l'un des dossiers listés" & @CRLF & _
				"   • OU créer un fichier d'alerte (_NOTE_AU_CORRECTEUR.txt) dans un de ces dossiers" & @CRLF & @CRLF & _
				"2. Génération automatique :" & @CRLF & _
				"   ► [Générer et Continuer] : Crée le fichier d'alerte et poursuit l'extraction" & @CRLF
		EndIf

		_Logging("Demande : générer fichier fichier d'alerte et continuer ou annuler la récupération?", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_WARNING, 0x000000, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $reponse = _ExtMsgBox(64, "         Générer et Continuer        |Annuler", $PROG_TITLE & " - Avertissement", $message, 0)

		Switch $reponse
			Case 1
				Local $Bac20xx = 'C:\Bac' & GUICtrlRead($cBac)
				If Not FileExists($Bac20xx) Then DirCreate($Bac20xx)
				Local $sFichierAlerte = $Bac20xx & "\_NOTE_AU_CORRECTEUR.txt"

				Local $sContenuRapport = $PROG_TITLE & $PROG_VERSION & " [" & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "]" & @CRLF & _
										 "---" & @CRLF & @CRLF & _
										 "Note pour le correcteur :" & @CRLF & _
										 " • Le(s) dossier(s) attendu(s) pour les travaux est/sont vide(s)." & @CRLF & _
										 " • Aucun fichier rendu par le candidat n'a été trouvé à l'emplacement prévu." & @CRLF

				If FileWrite($sFichierAlerte, $sContenuRapport) Then
					_Logging("Fichier d'alerte créé : " & $sFichierAlerte, 2, 1)
				Else
					_Logging("Échec création du fichier d'alerte : " & $sFichierAlerte, 5, 1)
					_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", _
						"Erreur : impossible de créer '_NOTE_AU_CORRECTEUR.txt' dans '" & $Bac20xx & "'.", 0)
					_Logging("Récupération annulée.", 5, 1)
					Return False
				EndIf
			Case 0, 2
				_Logging("Récupération annulée.", 5, 1)
				Return False
		EndSwitch
		_Logging("______", 2, 0)
	EndIf

	SplashOff()
	Return $Bac
EndFunc   ;==>_VerifierDossiersBac

; =========================================================
; Fonction: _CopierDossiersBac
; Description: Copie les dossiers Bac*2* vers les destinations
; Paramètres:
;   $Bac - Array des dossiers Bac à copier
; Retour: Nombre d'erreurs rencontrées
; =========================================================
Func _CopierDossiersBac($Bac)
	Local $iNberreurs = 0
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Copie des dossiers...", "", Default, Default, 1)

	For $i = 1 To $Bac[0]
		$sizefldr1 = DirGetSize($Bac[$i], 1)
		If @error Or $sizefldr1[1] = 0 Then
			_Logging("Ignorer le dossier vide """ & $Bac[$i] & """")
			Local $ErrorEmptyFolderRemove = DirRemove($Bac[$i], 1)
			_Logging("Suppression du dossier vide """ & $Bac[$i] & """", $ErrorEmptyFolderRemove)
			ContinueLoop
		EndIf

		ProgressSet(Round($i / $Bac[0] * 100), "[" & Round($i / $Bac[0] * 100) & "%] " & "Vérif. du dossier : " & StringRegExpReplace($Bac[$i], "^.*\\", ""))
		$Error1_CopyBacFldr = DirCopy($Bac[$i], $Dest1FlashUSB & StringTrimLeft($Bac[$i], 2), $FC_OVERWRITE)
		$TmpError = DirCopy($Bac[$i], $Dest2LocalFldr & StringTrimLeft($Bac[$i], 2), $FC_OVERWRITE)

		If $Error1_CopyBacFldr = 0 Then
			ProgressOff()
			_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest2LocalFldr, $Error1_CopyBacFldr)
			_Logging("Récupération annulée", 5, 1)
			_Logging("______", 2, 0)
			_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec lors de la copie du dossier """ & $Bac[$i] & """" & @CRLF _
					 & "L'opération de sauvegarde est annulée", 0)
			Return -1
		EndIf

		_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest1FlashUSB)
		_Logging("Copie de """ & $Bac[$i] & """ vers " & $Dest2LocalFldr, $TmpError)
		$iNberreurs = $iNberreurs - ($TmpError - 1)

		ProgressSet(Round($i / $Bac[0] * 100), "[" & Round($i / $Bac[0] * 100) & "%] " & "Suppression du dossier : " & StringRegExpReplace($Bac[$i], "^.*\\", ""))

		If $Error1_CopyBacFldr = 1 Then
			$Error2_DirRemove = DirRemove($Bac[$i], 1)
			_Logging("Suppression du dossier """ & $Bac[$i] & """", $Error2_DirRemove)
			$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
		EndIf
	Next

	ProgressOff()
	Return $iNberreurs
EndFunc   ;==>_CopierDossiersBac

; =========================================================
; Fonction: _CopierCapturesEcranFraude
; Description: Copie les captures d'écran en cas de détection de fraude
; Retour: Aucun
; =========================================================
Func _CopierCapturesEcranFraude()
    Local $DossierSession = IniRead($DossierBase & "\BacBackup\BacBackup.ini", "Params", "DossierSession", "/^_^\")

    ; Vérifier si le dossier de preuves UsbWatcher existe
    If FileExists($DossierBase & "\BacBackup\" & $DossierSession & "\_UsbWatcher") Then
        ; --- Logging interne ---
        _Logging($__g_sWarnIcon & " Fraude USB détectée. Création du rapport et copie des preuves.", 4)

        ; --- Chemins pour le suivi complet (Poste Candidat) ---
        Local $CheminDossierSession = $DossierBase & "\BacBackup\" & $DossierSession
        Local $CheminCaptures = $CheminDossierSession & "\_CapturesEcran"
        Local $CheminJournal = $CheminDossierSession & "\_journal_presse_papier.log"

        ; --- Construction du message ---
        Local $sContenu = _
            "========================================================" & @CRLF & _
            "      UTILISATION NON AUTORISÉE D'UN PÉRIPHÉRIQUE USB" & @CRLF & _
            "========================================================" & @CRLF & @CRLF & _
            "Date et Heure : " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & @CRLF & @CRLF & _
            "DÉTAILS :" & @CRLF & _
            "Le système a détecté le branchement d'un support amovible durant l'épreuve." & @CRLF & _
            "Le dossier ""_UsbWatcher"" contient les preuves immédiates collectées." & @CRLF & @CRLF & _
            "SUIVI COMPLET (Sur le poste du candidat) :" & @CRLF & _
						"Pour une analyse approfondie, consulter également sur le poste du candidat :" & @CRLF & _
            "   • Dossier de session : " & $CheminDossierSession & @CRLF & _
            "   • Captures d'écran : " & $CheminCaptures & @CRLF & _
            "   • Journal presse-papier : " & $CheminJournal & @CRLF & @CRLF & _
            "ACTIONS REQUISES (Surveillant / Coordinateur) :" & @CRLF & _
            "   1. Examiner le dossier ""_UsbWatcher"" joint." & @CRLF & _
            "   2. Consulter les éléments de suivi complet ci-dessus." & @CRLF & _
            "   3. Rédiger un rapport de fraude / tentative de fraude." & @CRLF & _
            "   4. Joindre toutes les preuves au rapport." & @CRLF & @CRLF & _
            "NOTE AU CORRECTEUR :" & @CRLF & _
            "Le travail du candidat doit être évalué normalement," & @CRLF & _
            "sans tenir compte de ce signalement." & @CRLF & @CRLF & _
            "========================================================" & @CRLF & _
            "Signalement généré automatiquement par BacCollector."

        ; --- Création du fichier d'alerte ---
        Local $hInfoFile = FileOpen($Dest1FlashUSB & "\" & $__g_sNomFichierAlerteFraude, $FO_OVERWRITE + $FO_UTF8)
        If $hInfoFile <> -1 Then
            FileWrite($hInfoFile, $sContenu)
            FileClose($hInfoFile)
            _Logging("Création du fichier """ & $__g_sNomFichierAlerteFraude & """")  ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
        EndIf

        ; --- Copie des dossiers de preuves ---
        _DirCopyWithProgress($DossierBase & "\BacBackup\" & $DossierSession & "\_UsbWatcher", $Dest1FlashUSB & "\_UsbWatcher", 1, "Copie des preuves de fraude...")
        _DirCopyWithProgress($DossierBase & "\BacBackup\" & $DossierSession & "\_UsbWatcher", $Dest2LocalFldr & "\_UsbWatcher", 1, "Copie des preuves de fraude...")
    EndIf
EndFunc   ;==>_CopierCapturesEcranFraude

; =========================================================
; Fonction: _FinaliserRecuperation
; Description: Finalise la récupération (BacBackup, initialisation)
; Retour: Aucun
; =========================================================
Func _FinaliserRecuperation()
	Local $Bac20xx = 'C:\Bac' & GUICtrlRead($cBac)
	SplashTextOn("Sans Titre", "Création du Dossier : """ & $Bac20xx & """" & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	Local $iNberreurs = 0
	$Error3_DirCreate = DirCreate($Bac20xx)
	_Logging("Création du Dossier : """ & $Bac20xx & """", $Error3_DirCreate)
	$iNberreurs = $iNberreurs - ($Error3_DirCreate - 1)

	SplashTextOn("Sans Titre", "Recherche de BacBackup." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	If _NouvelleSessionBacBackup() Then
		_Logging("Création d'une nouvelle session BacBackup")
	Else
		_Logging("BacBackup est introuvable ou la création d'une nouvelle session de surveillance a échoué.", 5)
	EndIf

	SplashOff()

	Return $iNberreurs
EndFunc   ;==>_FinaliserRecuperation

; =========================================================
; Fonction: _AfficherMessageFinal
; Description: Affiche le message final selon le résultat de la récupération
; Paramètres:
;   $NumeroCandidat - Numéro du candidat
;   $iNberreurs - Nombre d'erreurs rencontrées
; Retour: Aucun
; =========================================================
Func _AfficherMessageFinal($NumeroCandidat, $iNberreurs)
	If $iNberreurs = 0 Then
		_Logging("Récupération terminée avec succès", 3, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_SUCCESS, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONINFO, "Ok", $PROG_TITLE & " - Succès", "La récupération du travail du candidat N° " & $NumeroCandidat & " a été effectuée avec succès!", 0)
	ElseIf $iNberreurs = 1 Then
		_Logging("Récupération terminée, avec une erreur non critique", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Avertissement", "La récupération du travail du candidat N° " & $NumeroCandidat & " est terminée." & @CRLF _
				 & "Une erreur non critique est produite lors de cette opération." & @CRLF _
				 & "Veuillez lire le log et prendre les mesures nécessaires.", 0)
	Else
		_Logging("Récupération terminée, avec " & $iNberreurs & " erreurs non critiques", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Avertissement", "La récupération du travail du candidat N° " & $NumeroCandidat & " est terminée." & @CRLF _
				 & $iNberreurs & " erreurs non critiques sont produites lors de cette opération." & @CRLF _
				 & "Veuillez lire le log et prendre les mesures nécessaires.", 0)
	EndIf
EndFunc   ;==>_AfficherMessageFinal

; =========================================================
; Fonction: RecupererInfo
; Description: Fonction principale de récupération pour les examens standards
; Paramètres:
;   $NumeroCandidat - Numéro du candidat
; Retour: Aucun
; =========================================================
Func RecupererInfo($NumeroCandidat)
	; Vérification de l'existence du dossier destination
	If Not _VerifierExistenceDossierDestination($NumeroCandidat) Then Return 3

	; Vérification des dossiers Bac
	Local $Bac = _VerifierDossiersBac()
	If Not IsArray($Bac) Then Return 4

	; Création des dossiers de destination
	Local $iNberreurs = _CreerDossiersDestination()
	If $iNberreurs = -1 Then Return 5

	; Copie des dossiers Bac
	Local $iErreursCopie = _CopierDossiersBac($Bac)
	If $iErreursCopie = -1 Then Return 6
	$iNberreurs += $iErreursCopie

	; Copie des captures d'écran en cas de fraude
	_CopierCapturesEcranFraude()

	; Copie des autres fichiers (Bureau, Documents, etc.)
	Local $sMaskAutresFichiers = "*"
	$iNberreurs += _CopierLeContenuDesAutresDossiers($sMaskAutresFichiers)

	Local $aDossiers = DossiersRessources()
	$iNberreurs += _CopierEtSupprimerDossiers($aDossiers, "Res*ource*")

	; Finalisation
	$iNberreurs += _FinaliserRecuperation()

	; Message final
	_AfficherMessageFinal($NumeroCandidat, $iNberreurs)
	Return 0
EndFunc   ;==>RecupererInfo

; #FUNCTION# ====================================================================================================================
; Name...........: _WwwFolderHasContent
; Description....: Analyse la racine d'un serveur Web pour détecter du contenu utilisateur personnalisé.
; Syntax.........: _WwwFolderHasContent($sApacheDocumentRoot)
; Parameters.....: $sApacheDocumentRoot - Chemin complet du répertoire à analyser (ex: C:\wamp64\www).
; Return values..: Success - True si un fichier/dossier (autre que les exclus) est trouvé.
;                  Failure - False si le dossier est vide, n'existe pas ou ne contient que des éléments exclus.
; Author.........: BacCollector Team
; Modified.......:
; Remarks........: Ignore les fichiers cachés, "index.php" et les dossiers systèmes (wampthemes, wamplangues, img, etc.).
; Related........:
; Example........: If _WwwFolderHasContent("C:\wamp64\www") Then MsgBox(0, "Info", "Contenu détecté")
; ===============================================================================================================================
Func _WwwFolderHasContent($sApacheDocumentRoot)
    ; Nettoyage : suppression des slashs de fin pour uniformiser
    $sApacheDocumentRoot = StringRegExpReplace($sApacheDocumentRoot, "[\\/]+$", "")

    ; 1. Vérification de l'existence du dossier
    Local $sRootAttribs = FileGetAttrib($sApacheDocumentRoot)
    If @error Or Not StringInStr($sRootAttribs, "D") Then Return False

    ; 2. Initialisation de la recherche
    Local $hSearch = FileFindFirstFile($sApacheDocumentRoot & "\*")
    If $hSearch = -1 Then Return False

    Local $sFileName, $sFullFilePath, $sFileAttribs
    While 1
        $sFileName = FileFindNextFile($hSearch)
        If @error Then ExitLoop

        ; A. Ignorer les entrées système relatives au parcours
        If $sFileName = "." Or $sFileName = ".." Then ContinueLoop

        ; B. Ignorer index.php (insensible à la casse)
        If StringLower($sFileName) = "index.php" Then ContinueLoop

        ; C. Ignorer les dossiers spécifiques (insensible à la casse)
        Switch StringLower($sFileName)
            Case "wampthemes", "wamplangues", "img", "dashboard", "webalizer", "xampp", "forbidden", "restricted", "phpmyadmin"
                ContinueLoop
        EndSwitch

        ; D. Récupérer les attributs pour vérifier si caché
        $sFullFilePath = $sApacheDocumentRoot & "\" & $sFileName
        $sFileAttribs = FileGetAttrib($sFullFilePath)

        ; E. Ignorer si le fichier ou dossier est caché (Attribut H)
        If StringInStr($sFileAttribs, "H") Then ContinueLoop

        ; --- SI ON ARRIVE ICI ---
        ; C'est qu'on a trouvé quelque chose qui n'est ni index.php,
        ; ni dans la liste d'exclusion, ni caché.
        FileClose($hSearch)
        Return True
    WEnd

    FileClose($hSearch)
    Return False ; Rien trouvé de pertinent
EndFunc   ;==>_WwwFolderHasContent

; #FUNCTION# ====================================================================================================================
; Name...........: _TraiterDossierWww
; Description....: Copie un dossier www/htdocs avec gestion simplifiée d'index.php
; Syntax.........: _TraiterDossierWww($sWwwPath[, $bRemoveAfter = True])
; Parameters.....: $sWwwPath      - Chemin du dossier www/htdocs
;                  $bRemoveAfter   - Supprimer le contenu après copie (défaut: True)
; Return values..: Integer - Nombre d'erreurs (0 = succès, -1 = échec critique, -2 = ignoré)
; Remarks........: Logique :
;                  1. Copie intégrale vers les deux destinations
;                  2. Détermine si index.php est "spécial" (caché/système/ancien)
;                  3. Supprime index.php des destinations si spécial (toujours)
;                  4. Si $bRemoveAfter = True : supprime/recrée le dossier source
;                  5. Si $bRemoveAfter = True ET index.php spécial : restaure index.php depuis une destination
; ===============================================================================================================================
Func _TraiterDossierWww($sWwwPath, $bRemoveAfter = True)
    Local $hTimer = TimerInit()
    Local $iNberreurs = 0

    ; Vérifier s'il y a du contenu à sauvegarder
    If Not _WwwFolderHasContent($sWwwPath) Then
        _Logging("Le dossier """ & $sWwwPath & """ est vide - Ignoré", 2, 0)
        Return -2
    EndIf

    _Logging("🕗 Traitement du dossier """ & $sWwwPath & """" , 2, 0)

    ; Déterminer si index.php est "spécial"
    Local $sIndexPhpPath = $sWwwPath & "\index.php"
    Local $bIndexPhpExiste = FileExists($sIndexPhpPath)
    Local $bIndexPhpSpecial = False

    If $bIndexPhpExiste Then
        Local $sAttrib = FileGetAttrib($sIndexPhpPath)
        Local $bIsHiddenOrSystem = StringInStr($sAttrib, "H") Or StringInStr($sAttrib, "S")
        Local $bIsOld = False

        ; Calculer l'âge du fichier
        Local $aTime = FileGetTime($sIndexPhpPath, $FT_MODIFIED, $FT_ARRAY)
        If Not @error Then
            Local $sFileTime = $aTime[0] & "/" & $aTime[1] & "/" & $aTime[2] & " " & $aTime[3] & ":" & $aTime[4] & ":" & $aTime[5]
            Local $iAgeMinutes = _DateDiff("n", $sFileTime, _NowCalc())
            $bIsOld = ($iAgeMinutes > $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES_STI)
        EndIf

        $bIndexPhpSpecial = $bIsHiddenOrSystem Or $bIsOld
        _Logging("index.php détecté - Spécial=" & $bIndexPhpSpecial & " (Attributs: " & $sAttrib & ", Âge: " & $iAgeMinutes & " min)", 1, 0)
    EndIf

    ; Copier le dossier complet vers les deux destinations
    Local $sDestRelative = StringReplace(StringTrimLeft($sWwwPath, 3), "\", ".")
    Local $sDest1 = $Dest1FlashUSB & "\" & $sDestRelative
    Local $sDest2 = $Dest2LocalFldr & "\" & $sDestRelative

    Local $iError1 = DirCopy($sWwwPath, $sDest1, $FC_OVERWRITE)
    Local $iError2 = DirCopy($sWwwPath, $sDest2, $FC_OVERWRITE)

    ; Gestion des erreurs critiques
    If $iError1 = 0 Then
        _Logging("Échec de copie vers """ & $sDest1 & """", 5, 1, TimerDiff($hTimer))
        _ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
        _ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec lors de la copie du dossier """ & $sWwwPath & """" & @CRLF & _
                "L'opération de sauvegarde est annulée.", 0)
        _Logging("Récupération annulée", 5, 1)
        _Logging("______", 2, 0)
        Return -1
    EndIf

    _Logging("Copie réussie vers """ & $sDest1 & """", 1, 0)
    _Logging("Copie vers """ & $sDest2 & """", $iError2, 0)
    $iNberreurs += (1 - $iError2)  ; Incrémente si erreur

    ; Suppression et recréation du dossier source si demandé
    If $bRemoveAfter Then
        _Logging("Suppression et recréation de """ & $sWwwPath & """", 2, 0)

        ; Supprimer le dossier source
        Local $iDel = DirRemove($sWwwPath, 1)
        If $iDel = 0 Then
            _Logging("Échec suppression de """ & $sWwwPath & """", 5, 1)
            $iNberreurs += 1
        EndIf

        ; Recréer le dossier
        Local $iCreate = DirCreate($sWwwPath)
        If $iCreate = "" Then
            _Logging("Échec recréation de """ & $sWwwPath & """", 5, 1)
            $iNberreurs += 1
        EndIf

        ; Restaurer index.php si spécial
        If $bIndexPhpSpecial And $iDel And $iCreate Then
            ; Utiliser une des destinations comme source pour restaurer
            Local $sIndexSrc = $sDest1 & "\index.php"
            ; Vérifier si le fichier existe toujours (il a pu être supprimé par d'autres processus)
            If Not FileExists($sIndexSrc) Then $sIndexSrc = $sDest2 & "\index.php"
            Local $iRestore = FileCopy($sIndexSrc, $sIndexPhpPath, $FC_OVERWRITE)
        EndIf
    EndIf
    ; Supprimer index.php des destinations si spécial (TOUJOURS, indépendamment de $bRemoveAfter)
    If $bIndexPhpSpecial Then
        ; Supprimer index.php des destinations
		FileSetAttrib($sDest1 & "\index.php", "-RASH")
        FileDelete($sDest1 & "\index.php")
		FileSetAttrib($sDest2 & "\index.php", "-RASH")
        FileDelete($sDest2 & "\index.php")
    EndIf

	Local $aFolderInfo = DirGetSize($sDest1, 1)
	If ($aFolderInfo[1] = 0) And ($aFolderInfo[0] = 0) Then ;Nb Fichiers =0 & Taille =0

		Local $sNoteCorrecteur = $PROG_TITLE & $PROG_VERSION & " [" & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "]" & @CRLF & _
				"---" & @CRLF & @CRLF & _
				"Note pour le correcteur:" & @CRLF & _
				"  • Aucun fichier trouvé dans le dossier racine (www/htdocs)." & @CRLF & _
				"  • Ceci n'est pas dû à une erreur de copie :" & @CRLF & _
				"    le dossier ne contient que des sous-dossiers (aucun fichier)."  & @CRLF
		; --- Création du fichier d'alerte dans les deux destinations ---
        Local $sNomFichierRapport = "\_NOTE_AU_CORRECTEUR.txt"

        ; Destination 1 : Clé USB
		Local $hFile1 = FileOpen($sDest1 & $sNomFichierRapport, $FO_OVERWRITE + $FO_UTF8) ; 2 = Mode écriture (écrase l'ancien)
		If $hFile1 <> -1 Then
			FileWrite($hFile1, $sNoteCorrecteur)
			FileClose($hFile1)
			_Logging("Fichier d'alerte créé : " & $sDest1 & $sNomFichierRapport, 2, 1)
		EndIf
        ; Destination 2 : Dossier Local
		Local $hFile2 = FileOpen($sDest2 & $sNomFichierRapport, $FO_OVERWRITE + $FO_UTF8)
		If $hFile2 <> -1 Then
			FileWrite($hFile2, $sNoteCorrecteur)
			FileClose($hFile2)
			_Logging("Fichier d'alerte créé : " & $sDest2 & $sNomFichierRapport, 2, 1)
		EndIf
	EndIf

    _Logging("Traitement terminé : """ & $sWwwPath & """", 2, 0, TimerDiff($hTimer))
    Return $iNberreurs
EndFunc   ;==>_TraiterDossierWww

; #FUNCTION# ====================================================================================================================
; Name...........: _CopierEtSupprimerDossiers
; Description....: Copie un tableau de dossiers vers les destinations globales puis les supprime
; Syntax.........: _CopierEtSupprimerDossiers($aDossiers, $sTypeDossier[, $bRemoveAfter = True])
; Parameters.....: $aDossiers     - Tableau de chemins de dossiers à traiter (format [0]=nb, [1..n]=chemins)
;                  $sMsgTypeDossiers  - Nom descriptif pour la barre de progression (ex: "Bac*2*")
;                  $bRemoveAfter - Booléen pour supprimer après copie (défaut: True)
; Return values..: Nombre d'erreurs rencontrées pendant l'opération
; Author.........: BacCollector Team
; Remarks........: Utilise les variables globales $Dest1FlashUSB et $Dest2LocalFldr pour les destinations
; Example........: $nbErreurs = _CopierEtSupprimerDossiers($aDossiersBac, "Bac*2*")
; ===============================================================================================================================
Func _CopierEtSupprimerDossiers($aDossiers, $sMsgTypeDossiers, $bRemoveAfter = True)
    Local $iNberreurs = 0

    ; Affichage de la barre de progression avec le type de dossier personnalisé
    ProgressOn($PROG_TITLE & $PROG_VERSION, "Copie des dossiers " & $sMsgTypeDossiers & "...", "", Default, Default, 1)

    ; Vérification que le paramètre est bien un tableau valide
    If IsArray($aDossiers) And $aDossiers[0] > 0 Then
        ; Boucle sur chaque dossier du tableau
        For $i = 1 To $aDossiers[0]
            ; Vérification que le dossier existe et n'est pas vide
            Local $sizefldr1 = DirGetSize($aDossiers[$i], 1)
            If @error Or $sizefldr1[1] = 0 Then
                _Logging("Ignorer le dossier vide """ & $aDossiers[$i] & """")
            Else
                ; Mise à jour de la barre de progression
                ProgressSet(Round($i / $aDossiers[0] * 100), "[" & Round($i / $aDossiers[0] * 100) & "%] " & "Copie de : " & $aDossiers[$i])

                ; Récupération des informations sur le dossier
                Local $Fldr_info = DirGetSize($aDossiers[$i], 1)
                Local $Fldr_size = $Fldr_info[0]
                Local $FldrFilesCount = $Fldr_info[1]

                ; Copie uniquement si le dossier contient des fichiers
                If $Fldr_size > 0 Or $FldrFilesCount > 0 Then
                    ; Construction des chemins de destination
                    Local $sDestPath1 = $Dest1FlashUSB & StringTrimLeft($aDossiers[$i], 2)
                    Local $sDestPath2 = $Dest2LocalFldr & StringTrimLeft($aDossiers[$i], 2)

                    ; Création des dossiers parents si nécessaire
                    DirCreate(StringLeft($sDestPath1, StringInStr($sDestPath1, "\", 0, -1) - 1))
                    DirCreate(StringLeft($sDestPath2, StringInStr($sDestPath2, "\", 0, -1) - 1))

                    ; Copie vers les destinations globales
                    Local $Error1_CopyFldr = DirCopy($aDossiers[$i], $sDestPath1, $FC_OVERWRITE)
                    Local $TmpError = DirCopy($aDossiers[$i], $sDestPath2, $FC_OVERWRITE)

                    ; Journalisation et comptage des erreurs
                    $iNberreurs += (1 - $Error1_CopyFldr)
                    _Logging("Copie de """ & $aDossiers[$i] & """ vers " & $sDestPath1, $Error1_CopyFldr)
                    _Logging("Copie de """ & $aDossiers[$i] & """ vers " & $sDestPath2, $TmpError)
                    $iNberreurs += (1 - $TmpError)
                EndIf
            EndIf

            ; Suppression du dossier source si demandé
            If $bRemoveAfter Then
                Local $Error2_DirRemove = DirRemove($aDossiers[$i], 1)
                _Logging("Suppression du dossier """ & $aDossiers[$i] & """", $Error2_DirRemove)
                $iNberreurs += (1 - $Error2_DirRemove)
            EndIf
        Next
    Else
        _Logging("Aucun dossier à copier pour le format : " & $sMsgTypeDossiers, 5, 0)
    EndIf

    ProgressOff()
    Return $iNberreurs
EndFunc   ;==>_CopierEtSupprimerDossiers

; =========================================================
; Fonction: RecupererSti
; Description: Fonction principale de récupération pour les examens STI
;              (Sites web et bases de données)
; Paramètres:
;   $NumeroCandidat - Numéro du candidat
;   $bRemoveAfter - Supprimer les sources après copie (défaut: True)
; Retour: Aucun
; =========================================================
Func RecupererSti($NumeroCandidat, $bRemoveAfter = True)
    Local $iNberreurs = 0
    Local $YaSiteWeb = 1
    Local $YaDatabase = 1
    Local $bDossierWwwCompletCopie = False

    ; ===== Vérification et scan des sites web =====
    SplashTextOn("Sans Titre", "Recherche des dossiers d'hébergement locaux d'Apache." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
    _Logging("Début de recherche des dossiers d'hébergement locaux d'Apache.", 2, 0)
	Local $iStart = TimerInit()
    Local $Www = DossiersEasyPHPwww()

	If Not IsArray($Www) or $Www[0] = 0 Then
        $YaSiteWeb = 0
        SplashOff()
        _Logging($PROG_TITLE & " ne trouve aucune racine des documents (www/htdocs)", 5, 1, TimerDiff($iStart))
		Local $sConseilsSpecifiques = ""

		If $g_bBacBackupDetected Then
			$sConseilsSpecifiques = "• BacBackup est détecté : Utilisez ses sauvegardes périodiques" & @CRLF & _
									"  pour restaurer le dossier et consultez les captures d'écran" & @CRLF & _
									"  afin de comprendre l'historique des actions du candidat." & @CRLF & @CRLF
		Else
			$sConseilsSpecifiques = "• Vérifiez si le dossier www/htdocs n'est pas dans la Corbeille ou déplacé." & @CRLF & _
									"• Recherchez manuellement des dossiers www/htdocs." & @CRLF & @CRLF
		EndIf

		Local $sMsg = $PROG_TITLE & " ne trouve aucune racine des documents www/htdocs." & @CRLF & @CRLF & _
				"CAUSE(S) POSSIBLE(S) :" & @CRLF & _
				"• Le serveur local (XAMPP-Lite/XAMPP/WAMP) n'est pas installé." & @CRLF & _
				"• Le dossier www/htdocs a été déplacé, renommé ou supprimé." & @CRLF & @CRLF & _
				"ACTIONS RECOMMANDÉES :" & @CRLF & _
				$sConseilsSpecifiques & _
				"La récupération va être abandonnée."
        _ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
        _ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", $sMsg, 0)
		_Logging("Récupération abandonnée.", 5, 1)
		_Logging("______", 2, 0)
        Return 3
    Else
        _Logging("Scan des sites web dans : " & _ArrayToString($Www, ", ", 1), 2, 0)
        SplashOff()
        ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des sites web...", "", Default, Default, 1)

        ; Analyser chaque dossier www/htdocs
		Local $sTmpListeFormateeWww = ""
        Local $bAuMoinsUnContenu = False
        For $i = 1 To $Www[0]
            ProgressSet(Round($i / $Www[0] * 100), "[" & Round($i / $Www[0] * 100) & "%] Analyse de " & $Www[$i])
			$sTmpListeFormateeWww &= @CRLF & @TAB & "» " & $Www[$i]
            If _WwwFolderHasContent($Www[$i]) Then
                $bAuMoinsUnContenu = True
                ExitLoop
            EndIf
        Next
        ProgressOff()
		_Logging("Recherche complétée.", 2, 0, TimerDiff($iStart))
        If Not $bAuMoinsUnContenu Then
            $YaSiteWeb = 0
			Local $sTmpSpaces = "                            "
            _Logging("Aucun contenu à récupérer dans les dossiers www/htdocs:" & StringReplace($sTmpListeFormateeWww, @CRLF, @CRLF & $sTmpSpaces), 5, 1, TimerDiff($iStart))
            _Logging("MsgBox: Poursuivre à la recherche des BD (Oui/Non)?", 2, 0)
            _ExtMsgBoxSet(1, 0, $GUI_COLOR_WARNING, 0x000000, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
            Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & " - Avertissement", _
							"Aucun contenu à récupérer dans les dossiers www/htdocs :"  & _
							$sTmpListeFormateeWww & @CRLF & @CRLF _
                    & "Poursuivre la recherche et récupération de la base de données?" & @CRLF, 0)
            If $Rep <= 1 Then
                _Logging("Non", 2, 0)
                _Logging("Récupération annulée.", 5, 1)
                _Logging("______", 2, 0)
                Return 4
            EndIf
            _Logging("Oui", 2, 0)
        EndIf
    EndIf

    ; ===== Vérification et scan des bases de données =====
    SplashTextOn("Sans Titre", "Recherche des dossiers de stockage de BD MySql." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
    _Logging("Recherche des dossiers de stockage de BD MySql.", 2, 0)
	Local $iStart = TimerInit()
    Local $Data = DossiersEasyPHPdata()
    SplashOff()

    If Not IsArray($Data)  or $Data[0] = 0 Then
        $YaDatabase = 0
        _Logging($PROG_TITLE & " ne trouve aucun dossier de stockage de BD MySql/MariaDB.", 5, 1, TimerDiff($iStart))
		Local $sConseilsSpecifiques = ""
		If $g_bBacBackupDetected Then
			$sConseilsSpecifiques = "• BacBackup est détecté : Utilisez ses sauvegardes périodiques" & @CRLF & _
									"  pour restaurer le dossier et consultez les captures d'écran" & @CRLF & _
									"  afin de comprendre l'historique des actions du candidat." & @CRLF & @CRLF
		Else
			$sConseilsSpecifiques = "• Vérifiez si le dossier '\mysql\data' n'est pas dans la Corbeille ou déplacé." & @CRLF & _
									"• Recherchez manuellement des dossiers data." & @CRLF & @CRLF
		EndIf

		Local $sMsg = $PROG_TITLE & " ne trouve aucun dossier de stockage de BD MySql/MariaDB." & @CRLF & @CRLF & _
				"CAUSE(S) POSSIBLE(S) :" & @CRLF & _
				"• Le dossier 'data' a été déplacé, renommé ou supprimé." & @CRLF & @CRLF & _
				"ACTIONS RECOMMANDÉES :" & @CRLF & _
				$sConseilsSpecifiques & _
				"La récupération va être abandonnée."
        _ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
        _ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", $sMsg, 0)
		_Logging("Récupération abandonnée.", 5, 1)
		_Logging("______", 2, 0)
        Return 5
    Else
        _Logging("Scan des Bases de données dans les dossiers MySql dans : " & _ArrayToString($Data, ", ", 1), 2, 0)
        SplashOff()
        ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers de BD MySql", "", Default, Default, 1)
        Local $ListeBD[1] = [0]
        Local $ListeDataFolders[1] = [0]
        For $i = 1 To $Data[0]
            ProgressSet(Round($i / $Data[0] * 100), "[" & Round($i / $Data[0] * 100) & "%] ")
            $TmpBD = _FileListToArrayRec($Data[$i], "*|phpmyadmin;mysql;performance_schema;sys;cdcol;webauth|", 2, 0, 2, 2)
            If IsArray($TmpBD) Then
                $ListeDataFolders[0] += 1
                _ArrayAdd($ListeDataFolders, $Data[$i])
            EndIf
        Next
        ProgressOff()
		_Logging("Recherche complétée.", 2, 0, TimerDiff($iStart))
        If $ListeDataFolders[0] = 0 Then
            $YaDatabase = 0
            _Logging("Aucune base de données trouvée dans : " & _ArrayToString($Data, ", ", 1), 5, 1)
            If $YaSiteWeb Then
                _Logging("MsgBox: Poursuivre la récupération (Oui/Non)?", 2, 0)
				_ExtMsgBoxSet(1, 0, $GUI_COLOR_WARNING, 0x000000, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
                Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & " - Avertissement", "Aucune base de données trouvée." & @CRLF & @CRLF _
                        & "Voulez-vous poursuivre la récupération?" & @CRLF & @CRLF, 0)
                If $Rep <= 1 Then
                    _Logging("Non", 2, 0)
                    _Logging("Récupération annulée.", 5, 1)
                    _Logging("______", 2, 0)
                    Return 7
                EndIf
                _Logging("Oui", 2, 0)
            Else
                _ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
				_ExtMsgBox(64, "Ok", $PROG_TITLE & " - Aucun travail trouvé", _
						"Aucun travail n'a été trouvé pour ce candidat." & @CRLF & @CRLF & _
						"Si le candidat n'a rendu aucun travail, veuillez suivre cette procédure manuelle :" & @CRLF & @CRLF & _
						"1. Créez un dossier portant le numéro d'inscription du candidat." & @CRLF & _
						"2. À l'intérieur, ajoutez un fichier texte nommé 'AUCUN_TRAVAIL.txt'." & @CRLF & _
						"3. Contenu du fichier : " & @CRLF & _
						"   « Le candidat n'a produit aucun travail lors de l'épreuve. »" & @CRLF & @CRLF & _
						"Cette procédure permet de justifier l'absence de production." & @CRLF & @CRLF & _
						@TAB & "» Récupération annulée.", 0)
				_Logging("Récupération annulée.", 5, 1)
                _Logging("______", 2, 0)
                Return 8
            EndIf
        EndIf
    EndIf

    ; ===== Vérification de l'existence du dossier destination =====
    If Not _VerifierExistenceDossierDestination($NumeroCandidat) Then Return

    ; ===== Création des dossiers de destination =====
    Local $iErreurCreation = _CreerDossiersDestination()
    If $iErreurCreation = -1 Then Return 9
    $iNberreurs += $iErreurCreation

    ; ===== Copie des sites web =====
    If $YaSiteWeb Then
        ProgressOn($PROG_TITLE & $PROG_VERSION, "Copie des sites web...", "", Default, Default, 1)
        For $i = 1 To $Www[0]
            ProgressSet(Round($i / $Www[0] * 100), "[" & Round($i / $Www[0] * 100) & "%] Traitement de " & $Www[$i])
            Local $iResultat = _TraiterDossierWww($Www[$i], $bRemoveAfter)
            If $iResultat = -1 Then
                ; Erreur critique
                Return 10
            ElseIf $iResultat = -2 Then
                ; Rien à copier (dossier vide ou seulement index.php)
                ContinueLoop
            Else
                ; Succès, ajouter les erreurs éventuelles
                $iNberreurs += $iResultat
            EndIf
        Next
        ProgressOff()
	Else
        ; --- Formulation du message pour le rapport ---
		Local $sNoteCorrecteur = $PROG_TITLE & $PROG_VERSION & " [" & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "]" & @CRLF & _
				"---" & @CRLF & @CRLF & _
				"Note pour le correcteur:" & @CRLF & _
				"  • Les racines des documents (www/htdocs) ont été identifiées, mais elles ne contiennent aucun fichier ou dossier." & @CRLF & _
				"  • Ceci n'est pas dû à une erreur de copie ou de transfert" & @CRLF & _
				"  • Le professeur surveillant a été averti de cette situation et a confirmé la poursuite de la récupération."

        ; --- Création du fichier d'alerte dans les deux destinations ---
        Local $sNomFichierRapport = "\_NOTE_AU_CORRECTEUR.txt"

        ; Destination 1 : Clé USB
		Local $hFile1 = FileOpen($Dest1FlashUSB & $sNomFichierRapport, $FO_OVERWRITE + $FO_UTF8) ; 2 = Mode écriture (écrase l'ancien)
		If $hFile1 <> -1 Then
			FileWrite($hFile1, $sNoteCorrecteur)
			FileClose($hFile1)
			_Logging("Fichier d'alerte créé : " & $Dest1FlashUSB & $sNomFichierRapport, 2, 1)
		EndIf
        ; Destination 2 : Dossier Local
		Local $hFile2 = FileOpen($Dest2LocalFldr & $sNomFichierRapport, $FO_OVERWRITE + $FO_UTF8)
		If $hFile2 <> -1 Then
			FileWrite($hFile2, $sNoteCorrecteur)
			FileClose($hFile2)
			_Logging("Fichier d'alerte créé : " & $Dest2LocalFldr & $sNomFichierRapport, 2, 1)
		EndIf
    EndIf
	; ===== Copie des bases de données =====
	If $YaDatabase Then
		Local $hTimerGlobal = TimerInit()
		ProgressOn($PROG_TITLE & $PROG_VERSION, "Compression et Copie des dossiers Data...", "", Default, Default, 1)
		Local $sFldrName = StringFormat('%04X_%04X', Random(0, 0xFFFF), BitOR(Random(0, 0x3FFF), 0x8000))
		Local $sTempDir = @TempDir & "\BacCollector\" & $sFldrName & "\"
		DirCreate($sTempDir)

        For $i = 1 To $ListeDataFolders[0]
            ConsoleWrite($ListeDataFolders[$i] & @CRLF)
            ProgressSet(Round($i / $ListeDataFolders[0] * 100), "[" & Round($i / $ListeDataFolders[0] * 100) & "%] " & "Dossier """ & $ListeDataFolders[$i] & """")
            Local $ArcFileName = StringReplace(StringTrimLeft($ListeDataFolders[$i], 3), "\", ".") & ".zip"
            Local $ArcFile = $sTempDir & $ArcFileName

            ; --- Détection version (MariaDB en priorité) ---
            Local $sVersion = "Version non détectée"
            Local $sExePath = ""
            ; Normalisation du chemin: supprime \data final (avec gestion \ optionnel et insensible à la casse)
            Local $sParentDir = StringRegExpReplace($ListeDataFolders[$i], "(?i)\\data\\?$", "")
            Local $sBinPath = $sParentDir & "\bin"
            Local $sVersionFile = ""

            ; 1. Tester MariaDB d'abord
            If FileExists($sBinPath & "\mariadb.exe") Then
                $sExePath = $sBinPath & "\mariadb.exe"
                $sVersion = FileGetVersion($sExePath)
                $sVersion = (@error ? "MariaDB (version non lisible)" : "MariaDB " & $sVersion)
            ; 2. Ensuite MySQL
            ElseIf FileExists($sBinPath & "\mysql.exe") Then
                $sExePath = $sBinPath & "\mysql.exe"
                $sVersion = FileGetVersion($sExePath)
                $sVersion = (@error ? "MySQL (version non lisible)" : "MySQL " & $sVersion)
            Else
                $sVersion = "ERREUR: Aucun exécutable dans " & $sBinPath & @CRLF & _
                           "Recherché: mariadb.exe (prioritaire) puis mysql.exe"
            EndIf

            ; Création fichier version (uniquement si chemin valide)
            If $sParentDir <> $ListeDataFolders[$i] Then
                Local $sVersionFileName = StringReplace(StringTrimLeft($sParentDir, 3), "\", ".") & "_version.txt"
                $sVersionFile = $sTempDir & $sVersionFileName
                Local $hFile = FileOpen($sVersionFile, $FO_OVERWRITE + $FO_UTF8)
                If $hFile <> -1 Then
                    ; Contenu structuré et pédagogique pour le correcteur
                    FileWrite($hFile, _
								"===============================================" & @CRLF & _
								"BacCollector - Rapport de version MySQL/MariaDB" & @CRLF & _
								"===============================================" & @CRLF & _
								"Généré le: " & @MDAY & "/" & @MON & "/" & @YEAR & " à " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF & _
								@CRLF & _
								"Version détectée   : " & $sVersion & @CRLF & _
								"Chemin analysé     : " & $sParentDir & "\bin" & @CRLF & _
								"Exécutable trouvé  : " & ($sExePath ? $sExePath : "Aucun") & @CRLF & _
								@CRLF & _
								"-----------------------------------------------" & @CRLF & _
								@CRLF & _
								"PROCÉDURE DE RÉCUPÉRATION DES BASES DE DONNÉES :" & @CRLF & _
								@CRLF & _
								"En cas de nécessité de restauration des bases de données :" & @CRLF & _
								"  • Étape 1 : Arrêter le service MySQL/MariaDB via le panneau de contrôle." & @CRLF & _
								"  • Étape 2 : Renommer le dossier 'data' existant (sauvegarde préventive)." & @CRLF & _
								"  • Étape 3 : Extraire le fichier compressé '" & $ArcFileName & "' (joint) dans le dossier :" & @CRLF & _
								"       - XAMPP-Lite : C:\xampp_lite_x_x\apps\mysql\" & @CRLF & _
								"       - XAMPP      : C:\xampp\mysql\" & @CRLF & _
								"       - WampServer : C:\wamp64\bin\mysql\mysqlx.x.x\" & @CRLF & _
								"  • Étape 4 : Redémarrer le service MySQL/MariaDB." & @CRLF & _
								@CRLF & _
								"IMPORTANT :" & @CRLF & _
								"  • Il est essentiel d'utiliser la même version du SGBD (indiquée ci-dessus)." & @CRLF & _
								"  • Cette méthode de sauvegarde (dossier data complet) garantit la" & @CRLF & _
								"    récupération des bases même en l'absence de fichiers .sql exportés." & @CRLF & _
								"  • Le dossier 'data' doit être placé au même niveau que le dossier 'bin'." & @CRLF)
                    FileClose($hFile)
                Else
                    _Logging("ÉCHEC création fichier version: " & $sVersionFile, 0, 0, 0)
                    $sVersionFile = "" ; Désactive les copies ultérieures
                EndIf
            EndIf

            ; --- DÉBUT MESURE COMPRESSION ---
            Local $hTimerComp = TimerInit()
            If _7ZipStartup() Then
                $retResult = _7ZipAdd(0, $ArcFile, $ListeDataFolders[$i], 0, 1, 1)
                If @error Then
                    _Logging("Compression """ & $ListeDataFolders[$i] & """ vers " & $ArcFile, 0, 0, TimerDiff($hTimerComp))
                Else
                    _Logging("Compression réussie: """ & $ListeDataFolders[$i] & """", 1, 0, TimerDiff($hTimerComp))
                EndIf
            Else
                _Logging("Échec chargement 7Zip dll pour """ & $ListeDataFolders[$i] & """", 0, 0, TimerDiff($hTimerComp))
            EndIf
            _7ZipShutdown()
            ; --- FIN MESURE COMPRESSION ---

            Local $iError1_CopyBacFldr = 0
            Local $iTmpError = 0

            If FileExists($ArcFile) Then
                ; --- COPIE USB (mesure stricte pour le ZIP uniquement) ---
                Local $hTimerCopyUSB = TimerInit()
                $iError1_CopyBacFldr = FileCopy($ArcFile, $Dest1FlashUSB, $FC_OVERWRITE + $FC_CREATEPATH)
                Local $iDurationUSB = TimerDiff($hTimerCopyUSB)
                ; Copie version APRÈS mesure (hors chronométrage)
                If $sVersionFile Then FileCopy($sVersionFile, $Dest1FlashUSB, $FC_OVERWRITE + $FC_CREATEPATH)

                ; --- COPIE LOCALE (mesure stricte pour le ZIP uniquement) ---
                Local $hTimerCopyLocal = TimerInit()
                $iTmpError = FileCopy($ArcFile, $Dest2LocalFldr, $FC_OVERWRITE + $FC_CREATEPATH)
                Local $iDurationLocal = TimerDiff($hTimerCopyLocal)
                ; Copie version APRÈS mesure (hors chronométrage)
                If $sVersionFile Then FileCopy($sVersionFile, $Dest2LocalFldr, $FC_OVERWRITE + $FC_CREATEPATH)

                ; Nettoyage systématique
                FileDelete($ArcFile)
                If $sVersionFile Then FileDelete($sVersionFile)
            Else
                ; Nettoyage si compression échouée
                If $sVersionFile Then FileDelete($sVersionFile)
                $iError1_CopyBacFldr = 0
            EndIf

            If $iError1_CopyBacFldr = 0 Then
                ProgressOff()
                _Logging("ÉCHEC COPIE USB: """ & $ListeDataFolders[$i] & """ vers " & $Dest1FlashUSB & "\" & $ArcFileName, $iError1_CopyBacFldr, 0, $iDurationUSB)
                _Logging("Récupération annulée", 5, 1)
                _Logging("______", 2, 0)
                _ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
                _ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec lors de la copie du dossier """ & $ListeDataFolders[$i] & """" & @CRLF _
                        & "L'opération de sauvegarde est annulée", 0)
                Return 11
            EndIf

            _Logging("Copie USB: """ & $ListeDataFolders[$i] & """", 1, 0, $iDurationUSB)
            _Logging("Copie locale: """ & $ListeDataFolders[$i] & """", $iTmpError, 0, $iDurationLocal)
            $iNberreurs = $iNberreurs - ($iTmpError - 1)
        Next

		DirRemove($sTempDir, 1)
		ProgressOff()

		; --- LOG FINAL AVEC DURÉE TOTALE ---
		Local $iTotalDuration = TimerDiff($hTimerGlobal)
		_Logging("Toutes les opérations de compression et copie des BD terminées.", 2, 0, $iTotalDuration)
	Else
        ; --- Formulation du message pour le rapport ---
		Local $sNoteCorrecteur = $PROG_TITLE & $PROG_VERSION & " [" & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "]" & @CRLF & _
				"---" & @CRLF & @CRLF & _
				"Note pour le correcteur:" & @CRLF & _
				"  • Les dossiers data des serveurs MySQL/MariaDB ont été identifiés, mais ils ne contiennent aucune base de données utilisateur." & @CRLF & _
				"  • Ceci n'est pas dû à une erreur de copie ou de transfert." & @CRLF & _
				"  • Le professeur surveillant a été averti de cette situation et a confirmé la poursuite de la récupération."
        ; --- Création du fichier d'alerte dans les deux destinations ---
        Local $sNomFichierRapport = "\_NOTE_AU_CORRECTEUR.txt"

        ; Destination 1 : Clé USB
		Local $hFile1 = FileOpen($Dest1FlashUSB & $sNomFichierRapport, $FO_OVERWRITE + $FO_UTF8) ; 2 = Mode écriture (écrase l'ancien)
		If $hFile1 <> -1 Then
			FileWrite($hFile1, $sNoteCorrecteur)
			FileClose($hFile1)
			_Logging("Fichier d'alerte créé : " & $Dest1FlashUSB & $sNomFichierRapport, 2, 1)
		EndIf
        ; Destination 2 : Dossier Local
		Local $hFile2 = FileOpen($Dest2LocalFldr & $sNomFichierRapport, $FO_OVERWRITE + $FO_UTF8)
		If $hFile2 <> -1 Then
			FileWrite($hFile2, $sNoteCorrecteur)
			FileClose($hFile2)
			_Logging("Fichier d'alerte créé : " & $Dest2LocalFldr & $sNomFichierRapport, 2, 1)
		EndIf
	EndIf
    ; ===== Copie des captures d'écran en cas de fraude =====
    _CopierCapturesEcranFraude()

    ; ===== Copie des autres fichiers (sans suppression pour STI) =====
    Local $sMaskAutresFichiers = "*"
    $iNberreurs += _CopierLeContenuDesAutresDossiersSti($sMaskAutresFichiers)

    ; ===== Finalisation =====
    _FinaliserRecuperation()

	; ===== Message final =====

	If $iNberreurs <> 0 Then
		; --- CAS 1 : ERREURS CRITIQUES ---
		Local $iErreursAjustees = $iNberreurs - ($YaDatabase - 1) - ($YaSiteWeb - 1)
		_Logging("Récupération terminée, avec " & $iErreursAjustees & " erreurs.", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Erreur", "La récupération du travail du candidat N° " & $NumeroCandidat & " est terminée." & @CRLF _
				& $iErreursAjustees & " erreur(s) est/sont produite(s) lors de cette opération." & @CRLF _
				& "Veuillez lire attentivement le log et prendre les mesures nécessaires." & @CRLF, 0)
	ElseIf $YaDatabase And $YaSiteWeb Then
		; --- CAS 2 : TOUT EST OK ---
		_Logging("Récupération terminée avec succès", 3, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_SUCCESS, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONINFO, "Ok", $PROG_TITLE & " - Succès", "La récupération du travail du candidat N° " & $NumeroCandidat & " a été effectuée avec succès!" & @CRLF, 0)

	ElseIf Not $YaDatabase Then
		; --- CAS 3 : PAS DE BASE DE DONNÉES ---
		_Logging("Récupération terminée (Base de données manquante)", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_WARNING, 0x000000, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Avertissement", "La récupération du candidat N° " & $NumeroCandidat & " est terminée, mais la BASE DE DONNÉES est absente." & @CRLF, 0)

	ElseIf Not $YaSiteWeb Then
		; --- CAS 4 : PAS DE SITE WEB ---
		_Logging("Récupération terminée (Site Web manquant)", 2, 1)
		_Logging("______", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_WARNING, 0x000000, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Avertissement", "La récupération du candidat N° " & $NumeroCandidat & " est terminée, mais le SITE WEB est absent." & @CRLF, 0)
	EndIf
	Return 0
EndFunc   ;==>RecupererSti
;=========================================================

Func _AfficherLeContenuDesDossiersBac()
;~ 	SplashTextOn("Sans Titre", "Recherche des dossiers ""Bac*2*"" sous la racine ""C:""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers: [Bac*2*]", "", Default, Default, 1)
	_GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesDossiersBac)
	GUICtrlSetTip($GUI_AfficherLeContenuDesDossiersBac, "Double-clic pour afficher l'élément dans l'Explorateur.", "Contenu des dossiers 'Bac*2*'", 1,1)
	Local $NombreDeFichiers = 0
	Local $iStart = TimerInit()
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
	_GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesDossiersBac)
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
	_GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesDossiersBac)

	If $TmpMsgForLogging <> "" Then
		Local $sTmpSpaces = "                            "
		_Logging("Liste de dossiers/fichiers de travail dans ""Bac*2*"" : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
	Else
		_Logging("Aucun dossiers/fichiers de travail dans ""Bac*2*""", 2, 0, TimerDiff($iStart))
	EndIf
	ProgressOff()
;~ 	SplashOff()
EndFunc   ;==>_AfficherLeContenuDesDossiersBac

;=========================================================

Func _AfficherLeContenuDesDossiersSti()
    ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers ""www"", ""htdocs""...", "", Default, Default, 1)

    ;===================================================================
    ;======= Sites Web =================================================
    ;===================================================================
    _GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesDossiersBac)
    GUICtrlSetTip($GUI_AfficherLeContenuDesDossiersBac, "Double-clic pour afficher l'élément dans l'Explorateur.", "Contenu des Dossiers d'hébergement des serveurs Web (www & htdocs...)", 1, 1)

    Local $iStart = TimerInit()
    Local $aWww = DossiersEasyPHPwww()
    Local $aListeSites[1] = [0]
    Local $sWebSiteName, $sFileName, $sFileExtension, $sWebSite_Full_Path
    Local $sExclusions = "wampthemes;wamplangues;img;dashboard;webalizer;xampp;forbidden;restricted"

    If $aWww[0] <> 0 Then
        For $i = 1 To $aWww[0]
            ProgressSet(Round($i / $aWww[0] * 100), "[" & Round($i / $aWww[0] * 100) & "%] Scan racine...")
            ; OPTIMISATION: Exclusion directe dans la recherche
            Local $aDossiersSites = _FileListToArrayRec($aWww[$i], "*|" & $sExclusions, 2, 0, 2, 2)

            If IsArray($aDossiersSites) Then
                $aDossiersSites[0] = UBound($aDossiersSites) - 1
                $aListeSites[0] += $aDossiersSites[0]
                _ArrayDelete($aDossiersSites, 0)
                _ArrayAdd($aListeSites, $aDossiersSites)
            EndIf
        Next
    EndIf

    _GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesDossiersBac)
    Local $sTmpMsgForLogging = ""
    Local $iOldPercent = -1, $iPercent = 0

    If IsArray($aListeSites) Then
        For $N = 1 To $aListeSites[0]
            $sWebSite_Full_Path = $aListeSites[$N]

            $iPercent = Int(($N / $aListeSites[0]) * 100)
            If $iPercent <> $iOldPercent Then
                ProgressSet($iPercent, "[" & $iPercent & "%] Analyse : " & StringRegExpReplace($sWebSite_Full_Path, "^.*\\", ""))
                $iOldPercent = $iPercent
            EndIf

            Local $aFichiersSite = _FileListToArrayRec($sWebSite_Full_Path, "*", 1, 1, 2, 1)
            Local $aFileTimeCreation = FileGetTime($sWebSite_Full_Path, 1)
            Local $sFileTime = "[" & $aFileTimeCreation[3] & ":" & $aFileTimeCreation[4] & "] "
            $sWebSiteName = StringUpper(StringRegExpReplace($sWebSite_Full_Path, "^.*\\", ""))

            If Not IsArray($aFichiersSite) Then
                GUICtrlCreateListViewItem("[" & $sWebSiteName & "]" & "|" & "Site Web" & "|" & $sFileTime & " Site Web vide" & "|" & $sWebSite_Full_Path, $GUI_AfficherLeContenuDesDossiersBac)
                $sTmpMsgForLogging &= @CRLF & """" & $sWebSite_Full_Path & """" & "    (" & "Site Web" & "-" & $sFileTime & "Site Web vide )"
                GUICtrlSetColor(-1, 0x0000FF)
            Else
                GUICtrlCreateListViewItem("[" & $sWebSiteName & "]" & "|" & "Site Web" & "|" & $sFileTime & $aFichiersSite[0] & " fichier(s)" & "|" & $sWebSite_Full_Path, $GUI_AfficherLeContenuDesDossiersBac)
                $sTmpMsgForLogging &= @CRLF & """" & $sWebSite_Full_Path & """" & "    (" & "Site Web" & "-" & $sFileTime & $aFichiersSite[0] & " fichier(s) )"
                GUICtrlSetColor(-1, 0x0000FF)

                For $i = 1 To $aFichiersSite[0]
                    $aFileTimeCreation = FileGetTime($sWebSite_Full_Path & "\" & $aFichiersSite[$i], 1)
                    $sFileTime = "[" & $aFileTimeCreation[3] & ":" & $aFileTimeCreation[4] & "] "
                    Local $aFileTimeModif = FileGetTime($sWebSite_Full_Path & "\" & $aFichiersSite[$i], 0)
                    $sFileTime &= "[" & $aFileTimeModif[3] & ":" & $aFileTimeModif[4] & "] "
                    $sFileName = $aFichiersSite[$i]
                    $sFileExtension = StringUpper(StringRegExpReplace($sFileName, "^.*\.", ""))

                    GUICtrlCreateListViewItem("[" & $sWebSiteName & "] " & $sFileName & "|" & $sFileExtension & "|" & $sFileTime & "|" & $sWebSite_Full_Path & "\" & $aFichiersSite[$i], $GUI_AfficherLeContenuDesDossiersBac)
                    $sTmpMsgForLogging &= @CRLF & """" & $sWebSite_Full_Path & "\" & $aFichiersSite[$i] & "\" & $sFileName & """" & "    (" & $sFileTime & ")"
                Next
            EndIf
        Next
    EndIf
    _GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesDossiersBac)

    Local $sTmpSpaces = "                            "
    If $sTmpMsgForLogging <> "" Then
        _Logging("Liste des sites/fichiers web : " & StringReplace($sTmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
    Else
        _Logging("Aucun site ou fichier web trouvé.", 2, 0, TimerDiff($iStart))
    EndIf

    ;===================================================================
    ;======= Bases de Données ==========================================
    ;===================================================================
    ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des bases de données", "[0%] Initialisation...", Default, Default, 1)
    _GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesAutresDossiers)
    GUICtrlSetTip($GUI_AfficherLeContenuDesAutresDossiers, "Double-clic pour afficher l'élément dans l'Explorateur.", "Bases de données & Tables", 1, 1)

    $iStart = TimerInit()
    Local $aData = DossiersEasyPHPdata()
    Local $aListeBD[1] = [0]
    Local $sDB_Name, $sTable_Name, $sDB_FullPath

    If $aData[0] <> 0 Then
        For $i = 1 To $aData[0]
            ProgressSet(Round($i / $aData[0] * 100), "[" & Round($i / $aData[0] * 100) & "%] ")
            ; Exclusion native via le pipe (identique à l'originale)
            Local $aDossiersBD = _FileListToArrayRec($aData[$i], "*|phpmyadmin;mysql;performance_schema;sys;cdcol;webauth", 2, 0, 2, 2)

            If IsArray($aDossiersBD) Then
                $aDossiersBD[0] = UBound($aDossiersBD) - 1
                $aListeBD[0] += $aDossiersBD[0]
                _ArrayDelete($aDossiersBD, 0)
                _ArrayAdd($aListeBD, $aDossiersBD)
            EndIf
        Next
    EndIf

    _GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesAutresDossiers)
    $sTmpMsgForLogging = ""
    $iOldPercent = -1

    If IsArray($aListeBD) Then
        For $N = 1 To $aListeBD[0]
            $sDB_FullPath = $aListeBD[$N]

            $iPercent = Int(($N / $aListeBD[0]) * 100)
            If $iPercent <> $iOldPercent Then
                ProgressSet($iPercent, "[" & $iPercent & "%] BD : " & StringRegExpReplace($sDB_FullPath, "^.*\\", ""))
                $iOldPercent = $iPercent
            EndIf

            Local $aTablesBD = _FileListToArray($sDB_FullPath, "*.frm", 1)
            Local $aFileTimeCreation = FileGetTime($sDB_FullPath, 1)
            Local $sFileTime = "[" & $aFileTimeCreation[3] & ":" & $aFileTimeCreation[4] & "] "
            $sDB_Name = StringUpper(StringRegExpReplace($sDB_FullPath, "^.*\\", ""))

            If Not IsArray($aTablesBD) Then
                GUICtrlCreateListViewItem("[" & $sDB_Name & "]" & "|" & "BD" & "|" & $sFileTime & "BD vide" & "|" & $sDB_FullPath, $GUI_AfficherLeContenuDesAutresDossiers)
                $sTmpMsgForLogging &= @CRLF & """" & $sDB_FullPath & """" & "    (" & "BD" & "-" & $sFileTime & "BD vide )"
                GUICtrlSetColor(-1, 0x0000FF)
            Else
                GUICtrlCreateListViewItem("[" & $sDB_Name & "]" & "|" & "BD" & "|" & $sFileTime & $aTablesBD[0] & " table(s)" & "|" & $sDB_FullPath, $GUI_AfficherLeContenuDesAutresDossiers)
                $sTmpMsgForLogging &= @CRLF & """" & $sDB_FullPath & """" & "    (" & "BD" & "-" & $sFileTime & $aTablesBD[0] & " table(s) )"
                GUICtrlSetColor(-1, 0x0000FF)

                For $i = 1 To $aTablesBD[0]
                    $aFileTimeCreation = FileGetTime($sDB_FullPath & "\" & $aTablesBD[$i], 1)
                    $sFileTime = "[" & $aFileTimeCreation[3] & ":" & $aFileTimeCreation[4] & "] "
                    Local $aFileTimeModif = FileGetTime($sDB_FullPath & "\" & $aTablesBD[$i], 0)
                    $sFileTime &= "[" & $aFileTimeModif[3] & ":" & $aFileTimeModif[4] & "] "
                    $sTable_Name = StringUpper(StringTrimRight($aTablesBD[$i], 4))

                    GUICtrlCreateListViewItem("[" & $sDB_Name & "] " & $sTable_Name & "|" & "Table" & "|" & $sFileTime & "|" & $sDB_FullPath & "\" & $aTablesBD[$i], $GUI_AfficherLeContenuDesAutresDossiers)
                    $sTmpMsgForLogging &= @CRLF & """" & $sDB_FullPath & "\" & $aTablesBD[$i] & """" & "    (Table" & " - " & $sFileTime & ")"
                Next
            EndIf
        Next
    EndIf
    _GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesAutresDossiers)

    If $sTmpMsgForLogging <> "" Then
        _Logging("Liste de bases de données / tables : " & StringReplace($sTmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
    Else
        _Logging("Aucune base de données trouvée.", 2, 0, TimerDiff($iStart))
    EndIf

    ProgressOff()
EndFunc

Func _AfficherLeContenuDesDossiersSti__()

	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers ""www"", ""htdocs""...", "", Default, Default, 1)

	;===================================================================
	;======= Sites Web =================================================
	;===================================================================
	_GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesDossiersBac)
	GUICtrlSetTip($GUI_AfficherLeContenuDesDossiersBac, "Double-clic pour afficher l'élément dans l'Explorateur.", "Contenu des Dossiers d'hébergement des serveurs Web (www & htdocs...)", 1,1)
	Local $NombreDeFichiers = 0
	Local $iStart = TimerInit()
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
				$Kes[0] = UBound($Kes) - 1
				$Liste[0] += $Kes[0]
				_ArrayDelete($Kes, 0)
				_ArrayAdd($Liste, $Kes) ;
			EndIf
		Next
	EndIf
	_GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesDossiersBac)
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
	_GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesDossiersBac)
	Local $sTmpSpaces = "                            "
	If $TmpMsgForLogging <> "" Then
		_Logging("Liste des sites/fichiers web : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
	Else
		_Logging("Aucun site ou fichier web trouvé.", 2, 0, TimerDiff($iStart))
	EndIf

	;===================================================================
	;======= Bases de Données ==========================================
	;===================================================================
;~ 	SplashTextOn("Sans Titre", "Recherche des Bases de Données MySql." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des bases de données", "[0%] Veuillez patienter un moment, initialisation...", Default, Default, 1)
	_GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesAutresDossiers)
	GUICtrlSetTip($GUI_AfficherLeContenuDesAutresDossiers, "Double-clic pour afficher l'élément dans l'Explorateur.", "Bases de données & Tables", 1,1)

	Local $NombreDeFichiers = 0
	Local $iStart = TimerInit()
	Local $Data = DossiersEasyPHPdata() ;
	Local $Liste[1] = [0] ;
	Local $DB_Name, $Table_Name, $DB_FullPath
	If $Data[0] <> 0 Then
		For $i = 1 To $Data[0]
			ProgressSet(Round($i / $Data[0] * 100), "[" & Round($i / $Data[0] * 100) & "%] ")
			$Kes = _FileListToArrayRec($Data[$i], "*|phpmyadmin;mysql;performance_schema;sys;cdcol;webauth|", 2, 0, 2, 2)

			If IsArray($Kes) Then
				$Kes[0] = UBound($Kes) - 1
				$Liste[0] += $Kes[0]
				_ArrayDelete($Kes, 0)
				_ArrayAdd($Liste, $Kes) ;
			EndIf
		Next
	EndIf
	_GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesAutresDossiers)
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
	_GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesAutresDossiers)
	If $TmpMsgForLogging <> "" Then
		_Logging("Liste de bases de données / tables : " & StringReplace($TmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
	Else
		_Logging("Aucune base de données trouvée.", 2, 0, TimerDiff($iStart))
	EndIf

	ProgressOff()

;~ 	SplashOff()
EndFunc   ;==>_AfficherLeContenuDesDossiersSti

;=========================================================
; =========================================================
; Fonction: _SubCopierLeContenuDesAutresDossiers
; Description: Copie les fichiers récents d'une liste vers deux destinations
; Paramètres:
;   $Liste - Array contenant la liste des fichiers à traiter
;   $Dossier - Chemin source des fichiers
;   $SubFldr - Sous-dossier de destination relatif
;   $RealPath - Si 1, inclut le chemin complet dans les logs (défaut: 0)
;   $bRemoveAfter - Si True, supprime les fichiers sources après copie (défaut: True)
;   $iAgeMinutes - Age maximum des fichiers en minutes (défaut: $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES)
; Retour: Nombre d'erreurs rencontrées
; =========================================================
Func _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, $SubFldr, $RealPath = 0, $bRemoveAfter = True, $iAgeMinutes = $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES)
	Local $OtherFilesDestFlashUSB = ""
	Local $OtherFilesDestLocalFldr = ""
	Local $FileName = ""

	Local $iNberreurs = 0
	Local $TmpError
	Local $ListeFiltered[1] = [0]

	For $N = 1 To $Liste[0]
		$File_Name = $Liste[$N]
		ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringTrimRight($Liste[$N], 1))
		If AgeDuFichierEnMinutesModification($Dossier & $File_Name) < $iAgeMinutes Then
			$ListeFiltered[0] += 1
			_ArrayAdd($ListeFiltered, $File_Name)
		EndIf
	Next

	If $ListeFiltered[0] > 0 Then
		$OtherFilesDestFlashUSB = $Dest1FlashUSB & $SubFldr
		$OtherFilesDestLocalFldr = $Dest2LocalFldr & $SubFldr
		$TmpError = DirCreate($OtherFilesDestFlashUSB)
		;;;Log+++++++
		_Logging("Création du Dossier: " & $OtherFilesDestFlashUSB, $TmpError)
		$iNberreurs = $iNberreurs - ($TmpError - 1)
		;;;Log+++++++

		$TmpError = DirCreate($OtherFilesDestLocalFldr)
		;;;Log+++++++
		_Logging("Création du Dossier: " & $OtherFilesDestLocalFldr, $TmpError)
		$iNberreurs = $iNberreurs - ($TmpError - 1)
		;;;Log+++++++

		Local $Kes = ""
		If $RealPath = 1 Then
			$Kes = $Dossier
		EndIf

		For $N = 1 To $ListeFiltered[0]
			ProgressSet(Round($N / $ListeFiltered[0] * 100), "[" & Round($N / $ListeFiltered[0] * 100) & "%] " & "Copie de : " & StringTrimRight($ListeFiltered[$N], 1))
			$File_Name = $ListeFiltered[$N]
			$TmpError = FileCopy($Dossier & $File_Name, $OtherFilesDestFlashUSB & $File_Name, 8)
			;;;Log+++++++
			_Logging("Copie du Fichier: """ & $Kes & $File_Name & """ vers """ & $OtherFilesDestFlashUSB & $File_Name & """", $TmpError)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++

			$TmpError = FileCopy($Dossier & $File_Name, $OtherFilesDestLocalFldr & $File_Name, 8)
			;;;Log+++++++
			_Logging("Copie du Fichier: """ & $Kes & $File_Name & """ vers """ & $OtherFilesDestLocalFldr & $File_Name & """", $TmpError)
			$iNberreurs = $iNberreurs - ($TmpError - 1)
			;;;Log+++++++

			If $bRemoveAfter Then
				ProgressSet(Round($N / $ListeFiltered[0] * 100), "[" & Round($N / $ListeFiltered[0] * 100) & "%] " & "Suppression de : " & StringTrimRight($ListeFiltered[$N], 1))
				$TmpError = FileDelete($Dossier & $File_Name)
				;;;Log+++++++
				_Logging("Suppression du Fichier : """ & $Kes & $File_Name & """", $TmpError)
				$iNberreurs = $iNberreurs - ($TmpError - 1)
				;;;Log+++++++
			EndIf
		Next
	EndIf

	Return $iNberreurs
EndFunc   ;==>_SubCopierLeContenuDesAutresDossiers

; =========================================================
; Fonction: _SubCopierLeContenuDesAutresDossiersNoRemove
; Description: Wrapper de compatibilité - Copie sans suppression
; Note: Appelle _SubCopierLeContenuDesAutresDossiers avec $bRemoveAfter = False
; =========================================================
Func _SubCopierLeContenuDesAutresDossiersNoRemove($Liste, $Dossier, $SubFldr, $RealPath = 0)
	Return _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, $SubFldr, $RealPath, False)
EndFunc   ;==>_SubCopierLeContenuDesAutresDossiersNoRemove

; =========================================================
; Fonction: _SubCopyFoldersUserFolders
; Description: Copie les dossiers récents d'un répertoire utilisateur
; Paramètres:
;   $sSrc - Chemin source à scanner
;   $sDest_RelativePath - Chemin relatif de destination
;   $bRemoveAfter - Si True, supprime les dossiers sources après copie (défaut: True)
;   $iAgeMinutes - Age maximum des dossiers en minutes (défaut: $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES)
; Retour: Nombre d'erreurs rencontrées
; =========================================================
Func _SubCopyFoldersUserFolders($sSrc, $sDest_RelativePath, $bRemoveAfter = True, $iAgeMinutes = $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES)
	Local $iNberreurs = 0
	Local $Dossier = $sSrc
	Local $iFilesCountInFldr = 0, $iFldrSize = 0, $aFldr_info
	$Liste = _FileListToArrayRec($Dossier, "*||", 30, 0, 2, 1)

	If IsArray($Liste) Then
		For $N = 1 To $Liste[0]
			ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & StringTrimRight($Liste[$N], 1))
			$Fldr_Name = $Dossier & StringTrimRight($Liste[$N], 1)
			$Fldr_Name_Relative_Path = $sDest_RelativePath & $Liste[$N]
			If AgeDuFichierEnMinutesCreation($Fldr_Name) < $iAgeMinutes Then
				$aFldr_info = DirGetSize($Fldr_Name, 1)
				$iFldrSize = $aFldr_info[0]
				$iFilesCountInFldr = $aFldr_info[1]
				If $iFilesCountInFldr = 0 And $iFldrSize = 0 Then
					$Error2_DirRemove = DirRemove($Fldr_Name, 1)
					;;;Log+++++++
					_Logging("Suppression du [dossier vide] """ & $Fldr_Name_Relative_Path & """", $Error2_DirRemove)
					$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
					;;;Log+++++++
				Else
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Copie de : " & StringTrimRight($Liste[$N], 1))
					$Error1_CopyBacFldr = DirCopy($Fldr_Name, $Dest1FlashUSB & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
					$TmpError = DirCopy($Fldr_Name, $Dest2LocalFldr & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)

					;;;Log+++++++
					_Logging("Copie de """ & $Fldr_Name_Relative_Path & """ vers " & $Dest1FlashUSB, $Error1_CopyBacFldr)
					$iNberreurs = $iNberreurs - ($Error1_CopyBacFldr - 1)

					_Logging("Copie de """ & $Fldr_Name_Relative_Path & """ vers " & $Dest2LocalFldr, $TmpError)
					$iNberreurs = $iNberreurs - ($TmpError - 1)
					;;;Log+++++++

					If $bRemoveAfter And $Error1_CopyBacFldr = 1 Then
						ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Suppression de : " & StringTrimRight($Liste[$N], 1))
						$Error2_DirRemove = DirRemove($Fldr_Name, 1)
						;;;Log+++++++
						_Logging("Suppression du dossier """ & $Fldr_Name_Relative_Path & """", $Error2_DirRemove)
						$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
						;;;Log+++++++
					EndIf
				EndIf
			EndIf

		Next
	EndIf
	Return $iNberreurs

EndFunc   ;==>_SubCopyFoldersUserFolders

; =========================================================
; Fonction: _SubCopyFoldersUserFoldersNoRemove
; Description: Wrapper de compatibilité - Copie sans suppression
; Note: Appelle _SubCopyFoldersUserFolders avec $bRemoveAfter = False
; =========================================================
Func _SubCopyFoldersUserFoldersNoRemove($sSrc, $sDest_RelativePath)
	Return _SubCopyFoldersUserFolders($sSrc, $sDest_RelativePath, False)
EndFunc   ;==>_SubCopyFoldersUserFoldersNoRemove

; =========================================================
; Fonction: _CopierLeContenuDesAutresDossiers
; Description: Fonction principale pour copier les contenus des dossiers système
;              Scanne Bureau, Mes documents, Téléchargements, Profil utilisateur,
;              mu_code et les lecteurs locaux
; Paramètres:
;   $Mask - Masque de fichiers à rechercher
;   $bRemoveAfter - Si True, supprime les sources après copie (défaut: True)
;   $iAgeMinutes - Age maximum en minutes (défaut: $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES)
; Retour: Nombre d'erreurs rencontrées
; =========================================================
Func _CopierLeContenuDesAutresDossiers($Mask, $bRemoveAfter = True, $iAgeMinutes = $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES)
	Local $iNberreurs = 0

;~ ///////**********************************************************************
;~ ///////**********      Copy - Dossiers créés sur le Bureau           ********
;~ ///////**********************************************************************
	SplashOff()
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Bureau...", "", Default, Default, 1)
	$iNberreurs += _SubCopyFoldersUserFolders(@DesktopDir & "\", "Bureau\", $bRemoveAfter, $iAgeMinutes)

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers Modifiés sur le Bureau        ********
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Bureau...", "", Default, Default, 1)
	Local $Liste[1] = [0]
	Local $Dossier = @DesktopDir & '\'
	$Liste = _FileListToArrayRec($Dossier, $Mask & "|*.lnk|", 1, 1, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\Bureau\', 0, $bRemoveAfter, $iAgeMinutes)
	EndIf

;~ ///////**********************************************************************
;~ ///////*** Copy - Fichiers Modifiés dans le dossier de profil utilisateur ***
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Profil Utilisateur...", "", Default, Default, 1)
	Local $Liste[1] = [0]
	Local $Dossier = @UserProfileDir & '\'
	$Liste = _FileListToArrayRec($Dossier, "*.py;*.ipynb;*.ui;*.accdb;*.xls*;*.csv;*.doc*;*.ppt*||", 1, 0, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\ProfilU\', 0, $bRemoveAfter, $iAgeMinutes)
	EndIf

;~ ///////**********************************************************************
;~ ///////******** Copy - Fichiers Modifiés dans le dossier mu_code      *******
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> mu_code...", "", Default, Default, 1)
	Local $Liste[1] = [0]
	Local $Dossier = @UserProfileDir & '\mu_code\'
	$Liste = _FileListToArrayRec($Dossier, "*.py||", 1, 0, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\mu_code\', 0, $bRemoveAfter, $iAgeMinutes)
	EndIf

;~ ///////**********************************************************************
;~ ///////**********      Copy - Dossiers créés dans Mes documents      ********
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers >> Mes documents...", "", Default, Default, 1)
	$iNberreurs += _SubCopyFoldersUserFolders(@MyDocumentsDir & "\", "Mes documents\", $bRemoveAfter, $iAgeMinutes)

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers Modifiés dans Mes documents   ********
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Mes documents...", "", Default, Default, 1)
	Local $Liste[1] = [0]
	Local $Dossier = @MyDocumentsDir & '\'
	$Liste = _FileListToArrayRec($Dossier, $Mask & "||", 1, 1, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\Mes documents\', 0, $bRemoveAfter, $iAgeMinutes)
	EndIf

;~ ///////**********************************************************************
;~ ///////**********      Copy - Fichiers Modifiés dans Téléchargements   ******
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des fichiers >> Téléchargements...", "", Default, Default, 1)
	Local $Liste[1] = [0]
	Local $Dossier = _WinAPI_ShellGetKnownFolderPath($FOLDERID_Downloads) & '\'
	$Liste = _FileListToArrayRec($Dossier, $Mask & "||", 1, 1, 0, 1)
	If IsArray($Liste) Then
		$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\Téléchargements\', 0, $bRemoveAfter, $iAgeMinutes)
	EndIf

;~ ///////**********************************************************************
;~ ///////**********      Copy Dossiers récemment créés sous les lecteurs    ********
;~ ///////**********************************************************************
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Scan des dossiers>> disques locaux...", "", Default, Default, 1)
	Local $aDrive = DriveGetDrive('FIXED')
	Local $MaskExclude = ""
	For $i = 1 To $aDrive[0]
		$MaskExclude = ""
		If $aDrive[$i] = "C:" Then
			$MaskExclude = "Program Files*;Users;Windows;Intel;PerfLogs;EasyPhp*;xampp*;wamp*;Bac*2*;Documents and Settings;Config.Msi;Recovery;PySchool"
		EndIf

		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _
				And _WinAPI_IsWritable($aDrive[$i]) Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
			$Liste = _FileListToArrayRec($Dossier, "*|" & $MaskExclude & "|", 30, 0, 2, 1)
			If IsArray($Liste) Then
				For $N = 1 To $Liste[0]
					$Fldr_Name = StringTrimRight($Liste[$N], 1)
					ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Vérif. de : " & $Fldr_Name)
					If StringRegExp($Fldr_Name, "^(?i)bac\s*\d*2\d*\s*$", $STR_REGEXPMATCH, 1) Or AgeDuFichierEnMinutesCreation($Dossier & $Fldr_Name) < $iAgeMinutes Then
						$Fldr_info = DirGetSize($Dossier & $Fldr_Name, 1)
						$Fldr_size = $Fldr_info[0]
						$FldrFilesCount = $Fldr_info[1]
						If $Fldr_size = 0 And $FldrFilesCount = 0 Then
							$Error2_DirRemove = DirRemove($Dossier & $Fldr_Name, 1)
							;;;Log+++++++
							_Logging("Suppression du [dossier vide] """ & $Dossier & $Fldr_Name & """", $Error2_DirRemove)
							$iNberreurs = $iNberreurs - ($Error2_DirRemove - 1)
							;;;Log+++++++
						Else
							If $Fldr_size < $TAILLE_MAX_DU_DOSSIER_SOUS_LECTEUR * 1024 * 1024 Then
								$Fldr_Name_Relative_Path = StringLeft($Dossier, 1) & "_2pts" & "\" & $Fldr_Name
								ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Copie de : " & $Fldr_Name)
								$Error1_CopyBacFldr = DirCopy($Dossier & $Fldr_Name, $Dest1FlashUSB & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
								$TmpError = DirCopy($Dossier & $Fldr_Name, $Dest2LocalFldr & "\" & $Fldr_Name_Relative_Path, $FC_OVERWRITE)
								;;;Log+++++++
								_Logging("Copie de """ & $Dossier & $Fldr_Name & """ vers " & $Dest1FlashUSB, $Error1_CopyBacFldr)
								$iNberreurs = $iNberreurs - ($Error1_CopyBacFldr - 1)

								_Logging("Copie de """ & $Dossier & $Fldr_Name & """ vers " & $Dest2LocalFldr, $TmpError)
								$iNberreurs = $iNberreurs - ($TmpError - 1)
								;;;Log+++++++

								If $bRemoveAfter And $Error1_CopyBacFldr = 1 Then
									ProgressSet(Round($N / $Liste[0] * 100), "[" & Round($N / $Liste[0] * 100) & "%] " & "Suppression de : " & $Fldr_Name)
									$Error2_DirRemove = DirRemove($Dossier & $Fldr_Name, 1)
									;;;Log+++++++
									_Logging("Suppression du dossier """ & $Fldr_Name & """", $Error2_DirRemove)
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
		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _
				And _WinAPI_IsWritable($aDrive[$i]) Then
			$Dossier = StringUpper($aDrive[$i]) & "\"
			$Liste = _FileListToArrayRec($Dossier, $Mask & "|" & $MaskExclude & "|", 29, 0, 0, 1)
			If IsArray($Liste) Then
				$iNberreurs = $iNberreurs + _SubCopierLeContenuDesAutresDossiers($Liste, $Dossier, '\' & StringLeft($Dossier, 1) & '_2pts\', 1, $bRemoveAfter, $iAgeMinutes)
			EndIf
		EndIf
	Next

	Return $iNberreurs
EndFunc   ;==>_CopierLeContenuDesAutresDossiers

; =========================================================
; Fonction: _CopierLeContenuDesAutresDossiersSti
; Description: Fonction principale pour copier sans suppression
;              Inclut également la copie des dossiers Bac*2* pour la matière STI
; Paramètres:
;   $Mask - Masque de fichiers à rechercher
; Retour: Nombre d'erreurs rencontrées
; Note: Version étendue incluant la copie des dossiers Bac*2*
; =========================================================
Func _CopierLeContenuDesAutresDossiersSti($Mask, $bRemoveAfter = True, $iAgeMinutes = $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES_STI)
	Local $iNberreurs = 0

	; Appel de la fonction principale sans suppression
	$iNberreurs = _CopierLeContenuDesAutresDossiers($Mask, $bRemoveAfter, $iAgeMinutes)

;~ ///////**********************************************************************
;~ ///////**********    Copy Dossiers Bac*2* et "Res*ource    ******************
;~ ///////**********************************************************************

	Local $aDossiers = DossiersBac()
	$iNberreurs += _CopierEtSupprimerDossiers($aDossiers, "Bac*2*", $bRemoveAfter)

	Local $aDossiers = DossiersRessources()
	$iNberreurs += _CopierEtSupprimerDossiers($aDossiers, "Res*ource*", $bRemoveAfter)

	Return $iNberreurs
EndFunc   ;==>_CopierLeContenuDesAutresDossiersSti

#Region _AfficherLeContenuDesAutresDossiers
; #FUNCTION# ====================================================================================================================
; Nom ..........: _AfficherLeContenuDesAutresDossiers
; Description ...: Affiche les fichiers et dossiers récents de divers emplacements (Bureau, Documents, Téléchargements, Profil, etc.)
; Syntax ........: _AfficherLeContenuDesAutresDossiers([$sMask = "*"])
; Paramètres ....: $sMask              - [optionnel] Masque de recherche pour les fichiers. Par défaut "*" (tous les fichiers)
; Retour ........: Aucun - Affichage dans l'interface graphique
; Auteur ........: CTEI
; Modifié le ....:
; Remarques .....: Scanne Bureau, Mes documents, Téléchargements, ProfilU, dossiers Ressources et lecteurs fixes
; ===============================================================================================================================
Func _AfficherLeContenuDesAutresDossiers($sMask = "*")
    ; Initialisation
    _GUICtrlListView_DeleteAllItems($GUI_AfficherLeContenuDesAutresDossiers)
    GUICtrlSetTip($GUI_AfficherLeContenuDesAutresDossiers, _
        "- Bureau" & @CRLF & "- Mes documents" & @CRLF & "- Téléchargements" & @CRLF & _
        "- Dossier du profil de l'utilisateur" & @CRLF & "- Racines des lecteurs C:, D: ..." & @CRLF & @CRLF & _
        "(Double-clic sur un élément pour l'afficher dans l'Explorateur.)", _
        "Éléments récemment créés/modifiés dans :", 1, 1)

    ; Configuration des emplacements de scan
    ; Format: [Nom, Chemin, Masque, Récursivité (1=oui, 0=non)]
    Local $aEmplacementsScan[4][4] = [ _
        ["Bureau", @DesktopDir, $sMask & "|*.lnk|", 1], _                      ; Récursif
        ["Mes documents", @MyDocumentsDir, $sMask & "||", 1], _                ; Récursif
        ["Téléchargements", _WinAPI_ShellGetKnownFolderPath($FOLDERID_Downloads), $sMask & "||", 1], _ ; Récursif
        ["ProfilU", @UserProfileDir, "*.py;*.ipynb;*.ui;*.accdb;*.xls*;*.csv;*.doc*;*.ppt*||", 0] _ ; Non récursif
    ]

    ProgressOn($PROG_TITLE & $PROG_VERSION, "Initialisation du scan...", "", Default, Default, 1)

    ; 1. Scan des dossiers récents (Bureau et Mes documents)
    _ScannerDossiersRecents(@DesktopDir & "\", "Bureau", "Scan des dossiers >> Bureau...")
    _ScannerDossiersRecents(@MyDocumentsDir & "\", "Mes documents", "Scan des dossiers >> Mes documents...")

    ; 2. Scan des fichiers récents avec gestion de la récursivité
    For $i = 0 To UBound($aEmplacementsScan) - 1
        _ScannerFichiersRecents($aEmplacementsScan[$i][1] & "\", $aEmplacementsScan[$i][2], _
            $aEmplacementsScan[$i][0], "Scan des fichiers >> " & $aEmplacementsScan[$i][0] & "...", $aEmplacementsScan[$i][3])
    Next

    ; 3. Scan des dossiers ressources
    _ScannerDossiersRessources()

    ; 4. Scan des lecteurs fixes
    _ScannerLecteursFixes($sMask)

    ; Nettoyage
    ProgressOff()
    SplashOff()
EndFunc   ;==>_AfficherLeContenuDesAutresDossiers

; #FUNCTION# ====================================================================================================================
; Nom ..........: _ScannerDossiersRecents
; Description ...: Scanne les dossiers récents dans un emplacement spécifique
; Syntax ........: _ScannerDossiersRecents($sDossier, $sNomEmplacement, $sTitreProgression)
; Paramètres ....: $sDossier           - Chemin du dossier à scanner
;                  $sNomEmplacement    - Nom affiché pour l'emplacement
;                  $sTitreProgression  - Titre pour la barre de progression
; Retour ........: Aucun - Les résultats sont ajoutés à la liste
; Auteur ........: CTEI
; Modifié le ....:
; Remarques .....: Limité aux 30 derniers dossiers, vérifie l'âge depuis la création
; ===============================================================================================================================
Func _ScannerDossiersRecents($sDossier, $sNomEmplacement, $sTitreProgression)
    ProgressSet(0, "[0%] Initialisation...", $sTitreProgression)

    Local $sTmpMsgForLogging = ""
    Local $iStart = TimerInit()
    Local $aListeDossiers = _FileListToArrayRec($sDossier, "*||", 30, 0, 2, 1)
    If Not IsArray($aListeDossiers) Then Return

    _GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesAutresDossiers)
    Local $iOldPercent = -1, $iPercent = 0

    For $i = 1 To $aListeDossiers[0]
        $iPercent = Int(($i / $aListeDossiers[0]) * 100)
        If $iPercent <> $iOldPercent Then
            ProgressSet($iPercent, "[" & $iPercent & "%] Vérif. : " & StringRegExpReplace($aListeDossiers[$i], "^.*\\", ""))
            $iOldPercent = $iPercent
        EndIf

        Local $sCheminComplet = $sDossier & StringTrimRight($aListeDossiers[$i], 1)
        Local $sCheminRelatif = $sNomEmplacement & "\" & $aListeDossiers[$i]

        If AgeDuFichierEnMinutesCreation($sCheminComplet) < $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES Then
            _AjouterDossierAuListView($sCheminComplet, $sCheminRelatif, $sDossier & $aListeDossiers[$i], $sTmpMsgForLogging)
        EndIf
    Next
    _GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesAutresDossiers)

    ; Logging
    _Logging("Scan des dossiers dans " & $sNomEmplacement & " : " & $aListeDossiers[0] & " dossier(s)", 2, 0, TimerDiff($iStart))
    If $sTmpMsgForLogging <> "" Then
        Local $sTmpSpaces = "                            "
        _Logging("  -->  Liste de dossiers récents dans " & $sNomEmplacement & " : " & _
            StringReplace($sTmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0)
    Else
        _Logging("  -->  Aucun dossier récent.", 2, 0)
    EndIf
EndFunc   ;==>_ScannerDossiersRecents

; #FUNCTION# ====================================================================================================================
; Nom ..........: _ScannerFichiersRecents
; Description ...: Scanne les fichiers récents dans un emplacement spécifique
; Syntax ........: _ScannerFichiersRecents($sDossier, $sMask, $sNomEmplacement, $sTitreProgression[, $iRecursif = 1])
; Paramètres ....: $sDossier           - Chemin du dossier à scanner
;                  $sMask              - Masque de recherche pour les fichiers
;                  $sNomEmplacement    - Nom affiché pour l'emplacement
;                  $sTitreProgression  - Titre pour la barre de progression
;                  $iRecursif          - [optionnel] Recherche récursive (1=oui, 0=non). Par défaut 1
; Retour ........: Aucun - Les résultats sont ajoutés à la liste
; Auteur ........: CTEI
; Modifié le ....:
; Remarques .....: Vérifie l'âge depuis la modification, permet de contrôler la récursivité
; ===============================================================================================================================
Func _ScannerFichiersRecents($sDossier, $sMask, $sNomEmplacement, $sTitreProgression, $iRecursif = 1)
    ProgressSet(0, "[0%] Initialisation...", $sTitreProgression)

    Local $sTmpMsgForLogging = ""
    Local $iStart = TimerInit()
    Local $aListeFichiers = _FileListToArrayRec($sDossier, $sMask, 1, $iRecursif, 2, 1)
    If Not IsArray($aListeFichiers) Then Return

    _GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesAutresDossiers)
    Local $iOldPercent = -1, $iPercent = 0

    For $i = 1 To $aListeFichiers[0]
        $iPercent = Int(($i / $aListeFichiers[0]) * 100)
        If $iPercent <> $iOldPercent Then
            ProgressSet($iPercent, "[" & $iPercent & "%] Vérif. : " & StringRegExpReplace($aListeFichiers[$i], "^.*\\", ""))
            $iOldPercent = $iPercent
        EndIf

        Local $sCheminComplet = $sDossier & $aListeFichiers[$i]
        Local $sCheminRelatif = $sNomEmplacement & "\" & $aListeFichiers[$i]

        If AgeDuFichierEnMinutesModification($sCheminComplet) < $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES Then
            _AjouterFichierAuListView($sCheminComplet, $sCheminRelatif, $sDossier & $aListeFichiers[$i], $sTmpMsgForLogging)
        EndIf
    Next
    _GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesAutresDossiers)

    _Logging("Scan des fichiers dans " & $sNomEmplacement & " : " & $aListeFichiers[0] & " fichier(s)", 2, 0, TimerDiff($iStart))
    If $sTmpMsgForLogging <> "" Then
        Local $sTmpSpaces = "                            "
        _Logging("  -->  Liste de fichiers récents dans " & $sNomEmplacement & " : " & _
            StringReplace($sTmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0)
    Else
        _Logging("  -->  Aucun fichier récent.", 2, 0)
    EndIf
EndFunc   ;==>_ScannerFichiersRecents

; #FUNCTION# ====================================================================================================================
; Nom ..........: _ScannerDossiersRessources
; Description ...: Scanne les dossiers dans C:\Ressources
; Syntax ........: _ScannerDossiersRessources()
; Paramètres ....: Aucun
; Retour ........: Aucun - Les dossiers sont ajoutés à la liste
; Auteur ........: CTEI
; Modifié le ....:
; Remarques .....: Utilise la fonction DossiersRessources() pour obtenir la liste
; ===============================================================================================================================
Func _ScannerDossiersRessources()
    ProgressSet(0, "[0%] Initialisation...", "Scan des dossiers C:\Ressources")
    Local $iStart = TimerInit()
    Local $aListeRessources = DossiersRessources()
    If Not IsArray($aListeRessources) Or $aListeRessources[0] = 0 Then Return

    Local $sTmpMsgForLogging = ""
    _GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesAutresDossiers)
    Local $iOldPercent = -1, $iPercent = 0

    For $i = 1 To $aListeRessources[0]
        $iPercent = Int(($i / $aListeRessources[0]) * 100)
        If $iPercent <> $iOldPercent Then
            ProgressSet($iPercent, "[" & $iPercent & "%] Vérif. : " & StringRegExpReplace($aListeRessources[$i], "^.*\\", ""))
            $iOldPercent = $iPercent
        EndIf

        _AjouterDossierAuListView($aListeRessources[$i], $aListeRessources[$i] & "\", _
            $aListeRessources[$i] & "\", $sTmpMsgForLogging, False)
    Next
    _GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesAutresDossiers)

    _Logging("Scan des dossiers C:\Ressources : " & $aListeRessources[0] & " dossier(s)", 2, 0, TimerDiff($iStart))
    If $sTmpMsgForLogging <> "" Then
        Local $sTmpSpaces = "                            "
        _Logging("  -->  Liste des dossiers 'Ressources' : " & _
            StringReplace($sTmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0)
    Else
        _Logging("  -->  Aucun dossier 'Ressources'.", 2, 0)
    EndIf
EndFunc   ;==>_ScannerDossiersRessources

; #FUNCTION# ====================================================================================================================
; Nom ..........: _ScannerLecteursFixes
; Description ...: Scanne les lecteurs fixes (non-USB) pour les dossiers et fichiers récents
; Syntax ........: _ScannerLecteursFixes($sMask)
; Paramètres ....: $sMask              - Masque de recherche pour les fichiers
; Retour ........: Aucun - Les résultats sont ajoutés à la liste via les fonctions de scan spécifiques
; Auteur ........: CTEI
; Modifié le ....:
; Remarques .....: Exclut les lecteurs USB et les lecteurs non inscriptibles
; ===============================================================================================================================
Func _ScannerLecteursFixes($sMask)
    Local $aDrives = DriveGetDrive('FIXED')
    If @error Then Return

    For $i = 1 To $aDrives[0]
        If DriveGetType($aDrives[$i], $DT_BUSTYPE) = "USB" Or Not _WinAPI_IsWritable($aDrives[$i]) Then ContinueLoop

        Local $sDrivePath = StringUpper($aDrives[$i]) & "\"
        Local $sMaskExclude = ""

        If $aDrives[$i] = "C:" Then
            $sMaskExclude = "Program Files*;Users;Windows;Intel;PerfLogs;EasyPhp*;xampp*;wamp*;Bac*2*;Documents and Settings"
        EndIf

        _ScannerDossiersLecteur($sDrivePath, $sMaskExclude)
        _ScannerFichiersLecteur($sDrivePath, $sMask, $sMaskExclude)
    Next
EndFunc   ;==>_ScannerLecteursFixes

; #FUNCTION# ====================================================================================================================
; Nom ..........: _ScannerDossiersLecteur
; Description ...: Scanne les dossiers récents sur un lecteur spécifique
; Syntax ........: _ScannerDossiersLecteur($sDossier, $sMaskExclude)
; Paramètres ....: $sDossier           - Chemin du lecteur à scanner
;                  $sMaskExclude       - Masque d'exclusion pour les dossiers
; Retour ........: Aucun - Les dossiers sont ajoutés à la liste
; Auteur ........: CTEI
; Modifié le ....:
; ===============================================================================================================================
Func _ScannerDossiersLecteur($sDossier, $sMaskExclude)
    ProgressSet(0, "[0%] Initialisation...", "Scan des dossiers >> lecteur " & $sDossier)

    Local $sTmpMsgForLogging = ""
    Local $iStart = TimerInit()
    Local $aListeDossiers = _FileListToArrayRec($sDossier, "*|" & $sMaskExclude & "|", _
							$FLTAR_FOLDERS + $FLTAR_NOHIDDEN + $FLTAR_NOSYSTEM + $FLTAR_NOLINK, 0, 2, 2)
    If Not IsArray($aListeDossiers) Then Return

    _GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesAutresDossiers)
    Local $iOldPercent = -1, $iPercent = 0

    For $i = 1 To $aListeDossiers[0]
        $iPercent = Int(($i / $aListeDossiers[0]) * 100)
        If $iPercent <> $iOldPercent Then
            ProgressSet($iPercent, "[" & $iPercent & "%] Vérif. : " & StringRegExpReplace($aListeDossiers[$i], "^.*\\", ""))
            $iOldPercent = $iPercent
        EndIf

        Local $sCheminComplet = StringTrimRight($aListeDossiers[$i], 1)

        If StringRegExp($sCheminComplet, "^(?i)bac\s*\d*2\d*\s*$", $STR_REGEXPMATCH, 1) Or _
           AgeDuFichierEnMinutesCreation($sCheminComplet) < $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES Then
            _AjouterDossierAuListView($sCheminComplet, $sCheminComplet & "\", $sCheminComplet & "\", $sTmpMsgForLogging)
        EndIf
    Next
    _GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesAutresDossiers)

    _Logging("Scan des dossiers sous le lecteur """ & $sDossier & """ : " & $aListeDossiers[0] & " dossier(s)", 2, 0, TimerDiff($iStart))
    If $sTmpMsgForLogging <> "" Then
        Local $sTmpSpaces = "                            "
        _Logging("  -->  Liste de dossiers récents sous le lecteur """ & $sDossier & """ : " & _
            StringReplace($sTmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0)
    Else
        _Logging("  -->  Aucun dossier récent.", 2, 0)
    EndIf
EndFunc   ;==>_ScannerDossiersLecteur

; #FUNCTION# ====================================================================================================================
; Nom ..........: _ScannerFichiersLecteur
; Description ...: Scanne les fichiers récents sur un lecteur spécifique
; Syntax ........: _ScannerFichiersLecteur($sDossier, $sMask, $sMaskExclude)
; Paramètres ....: $sDossier           - Chemin du lecteur à scanner
;                  $sMask              - Masque de recherche pour les fichiers
;                  $sMaskExclude       - Masque d'exclusion pour les fichiers
; Retour ........: Aucun - Les fichiers sont ajoutés à la liste
; Auteur ........: CTEI
; Modifié le ....:
; Remarques .....: Exclut les fichiers système et cachés
; ===============================================================================================================================
Func _ScannerFichiersLecteur($sDossier, $sMask, $sMaskExclude)
    ProgressSet(0, "[0%] Initialisation...", "Scan des fichiers >> lecteur " & $sDossier)

    Local $sTmpMsgForLogging = ""
    Local $iStart = TimerInit()
    $sMaskExclude &= ";autoexec.bat;config.sys;*.log;*.rtf"

    Local $aListeFichiers = _FileListToArrayRec($sDossier, $sMask & "|" & $sMaskExclude & "|", _
	                        $FLTAR_FILES + $FLTAR_NOHIDDEN + $FLTAR_NOSYSTEM + $FLTAR_NOLINK, 0, 2, 2)
    If Not IsArray($aListeFichiers) Then Return

    _GUICtrlListView_BeginUpdate($GUI_AfficherLeContenuDesAutresDossiers)
    Local $iOldPercent = -1, $iPercent = 0

    For $i = 1 To $aListeFichiers[0]
        $iPercent = Int(($i / $aListeFichiers[0]) * 100)
        If $iPercent <> $iOldPercent Then
            ProgressSet($iPercent, "[" & $iPercent & "%] Vérif. : " & StringRegExpReplace($aListeFichiers[$i], "^.*\\", ""))
            $iOldPercent = $iPercent
        EndIf

        If AgeDuFichierEnMinutesModification($aListeFichiers[$i]) < $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES Then
            _AjouterFichierAuListView($aListeFichiers[$i], $aListeFichiers[$i], $aListeFichiers[$i], $sTmpMsgForLogging)
        EndIf
    Next
    _GUICtrlListView_EndUpdate($GUI_AfficherLeContenuDesAutresDossiers)

    _Logging("Scan des fichiers sous le lecteur """ & $sDossier & """ : " & $aListeFichiers[0] & " fichier(s)", 2, 0, TimerDiff($iStart))
    If $sTmpMsgForLogging <> "" Then
        Local $sTmpSpaces = "                            "
        _Logging("  -->  Liste de fichiers récents sous le lecteur """ & $sDossier & """ : " & _
            StringReplace($sTmpMsgForLogging, @CRLF, @CRLF & $sTmpSpaces), 2, 0)
    Else
        _Logging("  -->  Aucun fichier récent.", 2, 0)
    EndIf
EndFunc   ;==>_ScannerFichiersLecteur
; =========================================================
; Fonction: _AjouterDossierAuListView
; Description: Ajoute un dossier au ListView
; =========================================================
Func _AjouterDossierAuListView($sCheminComplet, $sCheminRelatif, $sCheminAffichage, ByRef $sLogMsg, $bVerifierTaille = True)
    Local $aInfo = DirGetSize($sCheminComplet, 1)
    Local $iTaille = $aInfo[0]
    Local $iNbFichiers = $aInfo[1]

    ; Formater le nombre de fichiers
    Local $sNbFichiers
    Switch $iNbFichiers
        Case 0
            $sNbFichiers = "  [dossier vide]"
        Case 1
            $sNbFichiers = "  [1 fichier]"
        Case Else
            $sNbFichiers = "  [" & $iNbFichiers & " fichiers]"
    EndSwitch

    ; Obtenir le temps de création
    Local $aTime = FileGetTime($sCheminComplet, 1)
    Local $sTime = "[" & $aTime[3] & ":" & $aTime[4] & "]"

    ; Ajouter au ListView
    Local $iItem = GUICtrlCreateListViewItem($sCheminRelatif & "|" & _FineSize($iTaille) & "|" & _
        $sTime & " " & $sNbFichiers & "|" & $sCheminAffichage, $GUI_AfficherLeContenuDesAutresDossiers)

    ; Ajouter au log
    $sLogMsg &= @CRLF & """" & $sCheminRelatif & """    (" & _FineSize($iTaille) & "-" & $sTime & " " & $sNbFichiers & ")"

    ; Colorier en rouge si la taille dépasse la limite
    If $bVerifierTaille And $iTaille > $TAILLE_MAX_DU_DOSSIER_SOUS_LECTEUR * 1024 * 1024 Then
        GUICtrlSetColor($iItem, 0xFF0000)
    EndIf
EndFunc   ;==>_AjouterDossierAuListView

; =========================================================
; Fonction: _AjouterFichierAuListView
; Description: Ajoute un fichier au ListView
; =========================================================
Func _AjouterFichierAuListView($sCheminComplet, $sCheminRelatif, $sCheminAffichage, ByRef $sLogMsg)
    Local $iTaille = FileGetSize($sCheminComplet)
    Local $sTailleFormatee = _FineSize($iTaille)

    ; Obtenir les temps de modification et de création
    Local $aTimeModif = FileGetTime($sCheminComplet, 1)
    Local $aTimeCreation = FileGetTime($sCheminComplet, 0)

    Local $sTime = "[" & $aTimeModif[3] & ":" & $aTimeModif[4] & "]"
    $sTime &= "   [" & $aTimeCreation[3] & ":" & $aTimeCreation[4] & "]"

    ; Ajouter au ListView
    GUICtrlCreateListViewItem($sCheminRelatif & "|" & $sTailleFormatee & "|" & $sTime & "|" & $sCheminAffichage, _
        $GUI_AfficherLeContenuDesAutresDossiers)

    ; Ajouter au log
    $sLogMsg &= @CRLF & """" & $sCheminRelatif & """    (" & $sTailleFormatee & "-" & $sTime & ")"
EndFunc   ;==>_AjouterFichierAuListView
#EndRegion _AfficherLeContenuDesAutresDossiers


; =========================================================
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
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Aucun dossier à sauvegarder !" & @CRLF _
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
				If FileExists($sCheminScripDir & $sDir & "\" & $__g_sNomFichierAlerteFraude) Then
					$TmpUneLigneMatrice &= "/!\ Clé USB ?"  ; Notes/Remarques
				EndIf


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
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Aucun dossier à sauvegarder !" & @CRLF _
				 & "Veuillez vérifier que " & $PROG_TITLE & " et les dossiers de travail des candidats" & @CRLF _
				 & "sont dans le même chemin." & @CRLF, 0)
		Return
	EndIf

	; Si le nombre de dossiers > 15
	If UBound($Liste, 1)-1 > 15 Then
		_Logging("Nombre de dossiers de travail > 15 : " & UBound($Liste, 1) - 1 & " dossiers", 4) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	EndIf


	Local $sTmpSpaces = "                            "
	_Logging("Liste prête à sauvegarder : " & UBound($Liste, 1) - 1 & " dossier(s)", 4) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	_Logging("Liste : " & _ArrayToString($Liste, "", 1, -1, @CRLF & $sTmpSpaces, 0, 0), 2, 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	;ici - Au moins un dossier sauvegardé

	SplashTextOn("Sans Titre", "Création du dossier de sauvegarde." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

;~ Création du Sous-dossier de sauvegarde
	Local $DestLocalFldr = IniRead($DossierBase & "\BacCollector\BacCollector.ini", "Params", "SousDossierSauve", "-1")
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
	$DestLocalFldr &= '__' & "Labo" & StringRight(GUICtrlRead($cLabo), 1) & "_Séance" & StringRight(GUICtrlRead($cSeance), 1) ;L1S3: Labo-1 Séance-3
	$DestLocalFldr &= '__' & @MDAY & "-" & @MON & "-" & @YEAR
	$DestLocalFldr &= '__' & @HOUR & "h" & @MIN
	IniWrite($DossierBase & "\BacCollector\BacCollector.ini", "Params", "SousDossierSauve", $DestLocalFldr)

	$DestLocalFldr = $DossierBacCollector & "\" & $DestLocalFldr & "\"
	$TmpError = DirCreate($DestLocalFldr)

	_Logging("Création du dossier de sauvegarde: " & @CRLF & "         """ & $DestLocalFldr & """", $TmpError) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	If $TmpError = 0 Then ;1:Success 0:Failure
		_Logging("Sauvegarde annulée", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("______", 2, 0)
		SplashOff()
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec lors de la création du dossier de sauvegarde: " & @CRLF _
				 & "   » " & $DestLocalFldr & @CRLF _
				 & "L'opération de sauvegarde est annulée" & @CRLF, 0)
		Return
	EndIf
	;;;Log+++++++
	SplashTextOn("Sans Titre", "Copie des dossiers." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)

	If $Matiere = "InfoProg" Then
		$MatiereLibelle = "Informatique / Algorithmique et Programmation"
	Else
		$MatiereLibelle = "STI - Systèmes & Technologies de l'Informatique"
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
	If $KesNomGrilleEval And FileExists($KesNomGrilleEval) Then
		_Logging("Création de la grille d'évaluation: " & @CRLF & $sTmpSpaces & """" & $KesNomGrilleEval & """", 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		FileCopy($KesNomGrilleEval, $DestLocalFldr)
	Else
		_Logging("Création de la grille d'évaluation.", 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	EndIf
   ;;;;;;;;;;;;;;;------------------

	ProgressSet(Round(($i-1) / $N * 100), "[" & Round(($i-1) / $N * 100) & "%] " & "Génération du Rapport.")
	$sRapportFileName &= '__' & GUICtrlRead($cSeance) & "_" &  GUICtrlRead($cLabo)
	$sRapportFileName &= '__' & JourDeLaSemaine() & "-" & @MDAY & "-" & @MON & "-" & @YEAR
	$sRapportFileName &= '__' & @HOUR & "h" & @MIN

	_GenererRapportPdf($sRapportFileName, $Liste, $DestLocalFldr, $MatiereLibelle)
	Local $sKesNomRapportPdf = $sCheminScripDir & $sRapportFileName & ".pdf"
	If FileExists($sKesNomRapportPdf) Then
		_Logging("Création du rapport PDF: " & @CRLF & $sTmpSpaces & """" & $sKesNomRapportPdf & """", 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		FileCopy($sKesNomRapportPdf, $DestLocalFldr)
	Else
		_Logging("Création du rapport PDF.", 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	EndIf

	ProgressOff()
	If $iNberreurs = 0 Then
		_Logging("Sauvegarde terminée avec succès", 3, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("______", 2, 0)
		_CopierLog($DestLocalFldr)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_CENTER, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(64, "Ok", $PROG_TITLE & " - Succès", "La Sauvegarde est terminée avec succès." & @CRLF, 0)
	ElseIf $iNberreurs = 1 Then
		_Logging("Sauvegarde terminée avec une seule erreur", 5, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
		_Logging("______", 2, 0)
		_CopierLog($DestLocalFldr)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Erreur", "La sauvegarde est terminée, mais il y a eu une erreur." & @CRLF, 0)
	Else
		_Logging("Récupération terminée, avec " & $iNberreurs & " erreurs", 5, 1)
		_Logging("______", 2, 0)
		_CopierLog($DestLocalFldr)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Erreur", "La sauvegarde est terminée, mais il y a eu " & $iNberreurs & " erreurs." & @CRLF, 0)

	EndIf
	WinActivate($hMainGUI)
EndFunc   ;==>Sauvegarder

;=========================================================

#Region Function "_InitialParams" --------------------------------------------------------------------------
Func _InitialParams()
	Local $GuiPart3Extend = IniRead(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "GuiPart3Extend", "-1")
	If StringIsInt($GuiPart3Extend) = 0 Or ($GuiPart3Extend <> 0 And $GuiPart3Extend <> 1) Then
		$GuiPart3Extend = 1
	EndIf
	_GUIExtender_Section_Action($hMainGUI, 1, $GuiPart3Extend) ;2ème paramètre (1) Numéro de la Section  -  3ème paramètre (0) to retract the Section

;~ 	SplashTextOn("Sans Titre", "Lecture des paramètres de ""BacCollector""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn($PROG_TITLE & $PROG_VERSION, "Mise à jour des paramètres...", "[0%] Veuillez patienter un moment, initialisation...", Default, Default, 1)

	$Lecteur = LecteurSauvegarde()
	$DossierBase = $Lecteur & $DossierSauve
	$DossierBacCollector = $DossierBase & "\" & "BacCollector"

	If Not _Directory_Is_Accessible($Lecteur)  And Not IsAdmin() And Not ($CMDLINE[0] > 0 And StringInStr($CMDLINE[1], "run_as_admin")) Then
		Local $sRegKey = 'HKEY_CURRENT_USER\SOFTWARE\Classes\exefile\shell\runas'
		If RegRead($sRegKey & '\command', '') <> '"%1" %*' Then
			RegWrite($sRegKey, 'HasLUAShield', 'REG_SZ', '')
			RegWrite($sRegKey & '\command', '', 'REG_SZ', '"%1" %*')
			RegWrite($sRegKey & '\command', 'IsolatedCommand', 'REG_SZ', '"%1" %*')
		EndIf

		GUISetState(@SW_HIDE, $hMainGUI)
		Local $cleanWorkingDir = @WorkingDir
		If StringRight($cleanWorkingDir, 1) = '\' Then $cleanWorkingDir = StringTrimRight($cleanWorkingDir, 1)
		ShellExecute(@ScriptFullPath, 'run_as_admin ' & @UserName, $cleanWorkingDir, 'runas')
		Exit
	EndIf
	Local $iStart = TimerInit()
	_Logging("🕗 Mise à jour des paramètres ____Début___", 2, 0)
	ProgressSet(20, "[" & 20 & "%] " & "")
	Local $Ok
	If Not FileExists($DossierBase) Then
		$Ok = DirCreate($DossierBase)
		If $Ok Then
			_Logging("Création du dossier de sauvegarde : " & $DossierBase, 1, 0)
		Else
			_Logging("Création du dossier de sauvegarde : " & $DossierBase, 0, 0)
		EndIf
;~ 		_Logging("Création du dossier de sauvegarde : " & $DossierBase, $Ok, 0)
	Else
		_Logging("Dossier de sauvegarde : """ & $DossierBase & """", 2, 0)
	EndIf
	ProgressSet(25, "[" & 25 & "%] " & "")
	Local $Repeated = 0, $Success = 0
	If Not FileExists($DossierBacCollector) Then
		While ($Repeated < 5)
			_UnLockFolder($DossierBase)
			$Success = DirCreate($DossierBacCollector)
			$Repeated += 1
			If $Success Then
				_Logging("Création du dossier  BacCollector : " & $DossierBacCollector, 1, 0)
				ExitLoop
			Else
				Sleep(20)
			EndIf
		WEnd
        If Not $Success Then
			ProgressOff()
			If Not IsAdmin() And Not ($CMDLINE[0] > 0 And StringInStr($CMDLINE[1], "run_as_admin")) Then
				Local $sRegKey = 'HKEY_CURRENT_USER\SOFTWARE\Classes\exefile\shell\runas'
				If RegRead($sRegKey & '\command', '') <> '"%1" %*' Then
					RegWrite($sRegKey, 'HasLUAShield', 'REG_SZ', '')
					RegWrite($sRegKey & '\command', '', 'REG_SZ', '"%1" %*')
					RegWrite($sRegKey & '\command', 'IsolatedCommand', 'REG_SZ', '"%1" %*')
				EndIf

				GUISetState(@SW_HIDE, $hMainGUI)
				Local $cleanWorkingDir = @WorkingDir
				If StringRight($cleanWorkingDir, 1) = '\' Then $cleanWorkingDir = StringTrimRight($cleanWorkingDir, 1)
				ShellExecute(@ScriptFullPath, 'run_as_admin ' & @UserName, $cleanWorkingDir, 'runas')
				Exit
			EndIf
			_Logging("Création du dossier  BacCollector : " & $DossierBacCollector, 0, 0)
			_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec lors de la création du dossier de sauvegarde local, " & @CRLF _
					 & "     """ & $DossierBacCollector & """" & @CRLF _
					 & "ça peut provoquer des erreurs lors de la copie de récupération!", 0)
			ProgressOn($PROG_TITLE & $PROG_VERSION, "Mise à jour des paramètres...", "", Default, Default, 1)
		EndIf
	Else
		_Logging("Dossier BacCollector  : """ & $DossierBacCollector & """", 2, 0)
	EndIf
	ProgressSet(30, "[" & 30 & "%] " & "")

	;;;BacBackup détecté
	If _IsRegistryExist("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{498AA8A4-2CBE-4368-BFA0-E0CF3F338536}_is1", "DisplayName") And ProcessExists('BacBackup.exe') Then
		GUICtrlSetBkColor($lblBacBackup, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetFont($lblBacBackup, 10, 900, 0, "Segoe UI Light")
		GUICtrlSetTip($lblBacBackup, "Cliquez pour ouvrir BacBackup", "BacBackup - Protection active ✔", 0,1)
		GUICtrlSetData($lblBacBackup, "BacBackup est Installé.")
		GUICtrlSetColor($lblBacBackup, 0xFFFFFF)
;~ 		_Logging("BacBackup est installé, et surveille les dossiers de travail des candidats.", 2, 0)
		$g_bBacBackupDetected = True
	Else
		GUICtrlSetBkColor($lblBacBackup, $GUI_BKCOLOR_TRANSPARENT)
		GUICtrlSetFont($lblBacBackup, 10, 900, 0, "Segoe UI Light")
		GUICtrlSetTip($lblBacBackup, "Cliquez pour aller à la page de téléchargement de BacBackup", "BacBackup - Protection inactive " & $__g_sWarnIcon, 0, 1)
		GUICtrlSetData($lblBacBackup, $__g_sWarnIcon & " BacBackup " & $__g_sWarnIcon)
		GUICtrlSetColor($lblBacBackup, 0xFF0000)
;~ 		_Logging("BacBackup n'est pas installé" & $Bac20xx, 2, 0)
		$g_bBacBackupDetected = False
	EndIf


	If Not _IsRegistryExist("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{D50A90DE-3118-4B58-9ABE-FDF795C59970}_is1", "DisplayName") Or Not ProcessExists('UsbCleaner.exe') Then
		GUICtrlSetState($CUsbCleaner, $GUI_SHOW)
		GUICtrlSetState($lblUsbCleaner, $GUI_SHOW)
	EndIf

	ProgressSet(35, "[" & 35 & "%] " & "")

;~ ******************************************
	_LockFolderContents($DossierBacCollector)
	ProgressSet(40, "[" & 40 & "%] " & "")

	If Not $g_bBacBackupDetected Then
		_LockRootFolder($DossierBase)
		ProgressSet(45, "[" & 45 & "%] " & "")
	EndIf

;~ *********************************************
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

	_Logging("Mise à jour des paramètres ____Fin___", 2, 0, TimerDiff($iStart))
	ProgressOff()
;~ 	SplashOff()
;~ 	WinActivate ($hMainGUI)

EndFunc   ;==>_InitialParams
#EndRegion Function "_InitialParams" --------------------------------------------------------------------------

;=========================================================

#Region Functions _SaveParams & _SaveData_xxxx
Func _SaveParams()
	;;; En cas où le fichier ini est effacé après l'ouverture du BacCollector,
	;;; Exemple : après l'ouverture de BacCollector, l'utilisateur a voulu effacer le contenu du Flash USB.
	SplashTextOn("Sans Titre", "Enregistrement des paramètres de ""BacCollector""." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	;;;Save $Matiere From Interface to ini File
	If GUICtrlGetState($rInfoProg) = $GUI_CHECKED Then
		$Matiere = 'InfoProg'
	ElseIf GUICtrlGetState($rSti) = $GUI_CHECKED Then
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
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_INFO, 0xFFFF00, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & " [?]", "Êtes-vous sûr(e) de vouloir modifier : " & @CRLF _
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
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_INFO, 0xFFFF00, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & " [?]", "Êtes-vous sûr(e) de vouloir modifier : " & @CRLF _
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
		_Logging("MsgBox: Modifier le labo : " & $Laboxx_lbl & " --> " & $Laboxx_combo & ". (Oui/Non)?", 2, 0)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_INFO, 0xFFFF00, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		Local $Rep = _ExtMsgBox($EMB_ICONEXCLAM, "~Non|Oui", $PROG_TITLE & " [?]", "Êtes-vous sûr(e) de vouloir modifier : " & @CRLF _
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

;=========================================================
Func _ShowAPropos()
    Local $aCompileDate = FileGetTime(@ScriptFullPath)
    Local $sCompileDate = $aCompileDate[2] & " " & _MonthFullName($aCompileDate[1]) & " " & $aCompileDate[0] & " à " & _
                          $aCompileDate[3] & ":" & $aCompileDate[4] & ":" & $aCompileDate[5]
    Local $GithubLink = "https://github.com/romoez/BacCollector"
    Local $LicenseLink = "https://github.com/romoez/BacCollector/blob/main/LICENSE"
    Local $EmailContact = "moez.romdhane@tarbia.tn"
    Local $iGUIWidth = 480, $iGUIHeight = 260

    ; Création de la fenêtre GUI
    Local $hGUI = GUICreate("À propos", $iGUIWidth, $iGUIHeight, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU), $WS_EX_TOPMOST)
    GUISetBkColor($GUI_COLOR_CENTER)
    GUISetFont(9, $FW_NORMAL, $GUI_FONTNORMAL, "Tahoma")

    ; Titre
    GUICtrlCreateLabel($PROG_TITLE & " v" & $PROG_VERSION, 20, 20, $iGUIWidth - 40, 25, $SS_CENTER)
    GUICtrlSetFont(-1, 17, 200)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Date de compilation
    GUICtrlCreateLabel("Compilé le " & $sCompileDate, 20, 50, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 8, 300, $GUI_FONTNORMAL)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Copyright
    GUICtrlCreateLabel("Copyright © " & $aCompileDate[0] & " La Communauté Tunisienne des Enseignants d'Informatique", 20, 80, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTNORMAL)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Licence
    Local $idLicenseLink = GUICtrlCreateLabel("Licence GPL-3.0", 20, 110, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 300, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; Lien GitHub
    Local $idGitHubLink = GUICtrlCreateLabel($GithubLink, 20, 140, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; Email de contact
    GUICtrlCreateLabel("Contact pour erreurs/suggestions :", 20, 170, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 300)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    Local $idEmailLink = GUICtrlCreateLabel($EmailContact, 20, 190, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; Bouton OK centré
    Local $idOK = GUICtrlCreateButton("OK", ($iGUIWidth - 80) / 2, 220, 80, 30)
    GUICtrlSetCursor(-1, 0)

    GUISetState(@SW_SHOW, $hGUI)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idOK
                ExitLoop
            Case $idGitHubLink
                ShellExecute($GithubLink)
            Case $idLicenseLink
                ShellExecute($LicenseLink)
            Case $idEmailLink
                ShellExecute("mailto:" & $EmailContact & "?subject=" & $PROG_TITLE & " v" & $PROG_VERSION)
        EndSwitch
    WEnd

    GUIDelete($hGUI)
EndFunc

Func _ShowAPropos_gemini()
    Local $aCompileDate = FileGetTime(@ScriptFullPath)
    Local $sCompileDate = $aCompileDate[2] & " " & _MonthFullName($aCompileDate[1]) & " " & $aCompileDate[0] & " à " & _
                          $aCompileDate[3] & ":" & $aCompileDate[4] & ":" & $aCompileDate[5]
    Local $GithubLink = "https://github.com/romoez/BacCollector"
    Local $LicenseLink = "https://github.com/romoez/BacCollector/blob/main/LICENSE"
    Local $sEmail = "votre.email@domaine.com" ; <-- Remplacez par votre adresse

    ; Hauteur augmentée à 270 pour inclure l'email sans tasser le contenu
    Local $iGUIWidth = 480, $iGUIHeight = 270

    Local $hGUI = GUICreate("À propos", $iGUIWidth, $iGUIHeight, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU), $WS_EX_TOPMOST)
    GUISetBkColor($GUI_COLOR_CENTER)
    GUISetFont(9, $FW_NORMAL, $GUI_FONTNORMAL, "Tahoma")

    ; Titre
    GUICtrlCreateLabel($PROG_TITLE & " v" & $PROG_VERSION, 20, 20, $iGUIWidth - 40, 25, $SS_CENTER)
    GUICtrlSetFont(-1, 17, 200)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Date de compilation
    GUICtrlCreateLabel("Compilé le " & $sCompileDate, 20, 50, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 8, 300, $GUI_FONTNORMAL)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Copyright
    GUICtrlCreateLabel("Copyright © " & $aCompileDate[0] & " La Communauté Tunisienne des Enseignants d'Informatique", 20, 80, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTNORMAL)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Licence
    Local $idLicenseLink = GUICtrlCreateLabel("Licence GPL-3.0", 20, 110, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 300, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; Lien GitHub
    Local $idGitHubLink = GUICtrlCreateLabel($GithubLink, 20, 140, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; --- AJOUT DE L'EMAIL ---
    Local $idEmailLink = GUICtrlCreateLabel("Contact, erreurs ou suggestions : " & $sEmail, 20, 170, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; Bouton OK déplacé plus bas (220)
    Local $idOK = GUICtrlCreateButton("OK", ($iGUIWidth - 80)/2, 220, 80, 30)
    GUICtrlSetCursor(-1, 0)

    GUISetState(@SW_SHOW, $hGUI)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idOK
                ExitLoop
            Case $idGitHubLink
                ShellExecute($GithubLink)
            Case $idLicenseLink
                ShellExecute($LicenseLink)
            Case $idEmailLink
                ; Ouvre le client mail avec un objet (Subject) pré-rempli
                ShellExecute("mailto:" & $sEmail & "?subject=Suggestion/Erreur sur " & $PROG_TITLE)
        EndSwitch
    WEnd

    GUIDelete($hGUI)
EndFunc

Func _ShowAPropos__()
    Local $aCompileDate = FileGetTime(@ScriptFullPath)
    Local $sCompileDate = $aCompileDate[2] & " " & _MonthFullName($aCompileDate[1]) & " " & $aCompileDate[0] & " à " & _
                          $aCompileDate[3] & ":" & $aCompileDate[4] & ":" & $aCompileDate[5]
    Local $GithubLink = "https://github.com/romoez/BacCollector"
    Local $LicenseLink = "https://github.com/romoez/BacCollector/blob/main/LICENSE"
    Local $iGUIWidth = 480, $iGUIHeight = 230 ; Hauteur ajustée pour plus d'espace

    ; Création de la fenêtre GUI
    Local $hGUI = GUICreate("À propos", $iGUIWidth, $iGUIHeight, -1, -1, BitOR($WS_CAPTION, $WS_SYSMENU), $WS_EX_TOPMOST)
    GUISetBkColor($GUI_COLOR_CENTER)
	GUISetFont ( 9, $FW_NORMAL , $GUI_FONTNORMAL , "Tahoma"); [, winhandle [, quality]]]]] )

    ; Titre
    GUICtrlCreateLabel($PROG_TITLE & " v" & $PROG_VERSION, 20, 20, $iGUIWidth - 40, 25, $SS_CENTER)
    GUICtrlSetFont(-1, 17, 200)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Date de compilation
    GUICtrlCreateLabel("Compilé le " & $sCompileDate, 20, 50, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 8, 300, $GUI_FONTNORMAL)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Copyright
    GUICtrlCreateLabel("Copyright © " & $aCompileDate[0] & " La Communauté Tunisienne des Enseignants d'Informatique", 20, 80, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTNORMAL)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Licence (ajout)
    Local $idLicenseLink = GUICtrlCreateLabel("Licence GPL-3.0", 20, 110, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 300, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; Site Web
;~     GUICtrlCreateLabel("Github :", 20, 140, $iGUIWidth - 40, 20, $SS_CENTER)
;~     GUICtrlSetFont(-1, 9, 300)
;~     GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)

    ; Lien GitHub
    Local $idGitHubLink = GUICtrlCreateLabel($GithubLink, 20, 140, $iGUIWidth - 40, 20, $SS_CENTER)
    GUICtrlSetFont(-1, 9, 400, $GUI_FONTUNDER)
    GUICtrlSetColor(-1, 0x63C2F5)
    GUICtrlSetCursor(-1, 0)

    ; Bouton OK centré
    Local $idOK = GUICtrlCreateButton("OK", ($iGUIWidth - 80)/2, 180, 80, 30)
    GUICtrlSetCursor(-1, 0)

    GUISetState(@SW_SHOW, $hGUI)

    While 1
        Switch GUIGetMsg()
            Case $GUI_EVENT_CLOSE, $idOK
                ExitLoop
            Case $idGitHubLink
                ShellExecute($GithubLink)
            Case $idLicenseLink
                ShellExecute($LicenseLink)
        EndSwitch
    WEnd

    GUIDelete($hGUI)
EndFunc
;=========================================================

#Region Functions BB
Func _NouvelleSessionBacBackup()
	If $g_bBacBackupDetected Then
		Return _CreerDossierNouvelleSession($Lecteur, $DossierSauve)
	EndIf
EndFunc   ;==>_NouvelleSessionBacBackup

Func _OpenBacBackupInterface()
	Local $sProcessName = "BacBackup.exe"
	Local $sAppId = "{498AA8A4-2CBE-4368-BFA0-E0CF3F338536}_is1"
	Local $sAppPath = @ProgramFilesDir & "\BacBackup\" & $sProcessName
	Local $sProcessPath = _GetProcessPath($sProcessName, $sAppPath, $sAppId)
	If $sProcessPath = "" Then
		_Logging("Ouverture de la page de télépchargement de BacBackup", 2, 0)
		ShellExecute("https://github.com/romoez/BacBackup", "", "", "open")
;~ 		GUICtrlSetState($lblBacBackup, $GUI_HIDE)
		Return 0
	EndIf
	$sBBInterface = _ExtractFolderPath($sProcessPath) & "BacBackup_Interface.exe"
	If FileExists($sBBInterface) Then
		_Logging("Ouverture de la fenêtre de BacBackup", 2, 0)
		Return Run("""" & $sBBInterface & """ baccollector")
	EndIf
	Return -1
EndFunc   ;==>_OpenBacBackupInterface

Func _OpenUsbCleanerUrl()
		_Logging("Ouverture de la page de téléchargement de UsbCleaner", 2, 0)
		ShellExecute("https://github.com/romoez/UsbCleaner", "", "", "open")
EndFunc     ;==>_OpenUsbCleanerUrl
#EndRegion Functions BB

;=========================================================



;=========================================================

;=========================================================

;=========================================================

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
				$sListeDossiersRecuperes = $sListeDossiersRecuperes & "... " & ($iNombreDossiersRecuperes - 15) & " autre(s) dossier(s) !"
			EndIf
		EndIf
	EndIf

	SplashOff()
	Return $sListeDossiersRecuperes
EndFunc   ;==>_ListeDeDossiersRecuperes

;=========================================================

Func _ListeDApplicationsOuvertes()
;~ 	SplashTextOn("Sans Titre", "Recherche de logiciels ouverts." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	ProgressOn("", "Vérif. des logiciels ouverts:", "", Default, Default, 1)

	Local $App = "" ;

	Local $SoftsToClose[][3] = [ _
			  ["Atom", "atom.exe", 1] _
			, ["Dreamweaver", "Dreamweaver.exe", 1] _
			, ["LibreOffice/OpenOffice...", "soffice.bin", 2] _
			, ["Microsoft Access", "MSAccess.exe", 1] _
			, ["Microsoft Excel", "Excel.exe", 1] _
			, ["Microsoft PowerPoint", "POWERPNT.exe", 1] _
			, ["Microsoft Word", "WinWord.exe", 1] _
			, ["Ms Expression Web 4", "ExpressionWeb.exe", 1] _
			, ["Notepad", "notepad.exe", 1] _
			, ["Notepad++", "notepad++.exe", 1] _
			, ["PyScripter", "PyScripter.exe", 1] _
			, ["PyCharm", "pycharm", 0] _
			, ["Qt Designer", "designer.exe", 1] _
			, ["Rapid PHP", "rapidphp.exe", 1] _
			, ["Sublime Text", "sublime_text.exe", 1] _
			, ["UltraEdit", "uedit", 0] _
			, ["Visual Studio Code", "Code.exe", 1] _
			, ["VSCodium", "VSCodium.exe", 1] _
			, ["Webuilder", "webuild.exe", 1] _
			, ["Wing Python IDE", "wing", 0] _
			]
;~ 			, ["Brackets", "Brackets.exe", 1] _
;~ 			, ["Geany", "geany.exe", 1] _
;~ 			, ["GIMP", "GIMP", 0] _
;~ 			, ["Gvim", "gvim.exe", 1] _
;~ 			, ["Microsoft Frontpage", "FrontPG.exe", 1] _
;~ 			, ["Microsoft Publisher", "MSPUB.exe", 1] _
;~ 			, ["Paint .Net", "PaintDotNet.exe", 1] _
;~ 			, ["Paint", "MSPaint.exe", 1] _
;~ 			, ["Vim", "vim.exe", 1] _
;~ 			, ["Windows Movie Maker", "Moviemk.exe", 1] _

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

; #FUNCTION# ====================================================================================================================
; Name...........: _NumeroCandidat
; Description....: Recherche récursive dans les dossiers Bac* d’un numéro de candidat (6 chiffres consécutifs)
; Syntax.........: _NumeroCandidat()
; Return values..: Chaîne – Numéro trouvé (6 chiffres) ou "000000" par défaut
; Remarks........: Extrait le premier match de 6 chiffres consécutifs dans le nom (fichiers ou dossiers) ;
;                  Format du log : "[XXXms] Numéro trouvé: <numéro> (dans <chemin_complet>)"
; ===============================================================================================================================
Func _NumeroCandidat()
    SplashTextOn("Sans Titre", "Détermination du numéro d'inscription du candidat..." & @CRLF & @CRLF & "Analyse des dossiers Bac* en cours...", 330, 120, -1, -1, 49, "Segoe UI", 9)
    _Logging("🕗 Recherche du numéro de candidat dans les dossiers ""Bac*2*""...", 2, 0)
	Local $iStart = TimerInit()
    Local $aBacDirs = DossiersBac()
    Local $sNumCandidat = "000000"


    If IsArray($aBacDirs) And $aBacDirs[0] > 0 Then
        For $i = 1 To $aBacDirs[0]
            Local $sRoot = $aBacDirs[$i]
            _Logging("Analyse récursive de : """ & $sRoot & """", 2, 0)

            ; Recherche récursive : fichiers + dossiers, chemins complets
            Local $aItems = _FileListToArrayRec($sRoot, "*", $FLTAR_FILESFOLDERS, $FLTAR_NORECUR, $FLTAR_SORT, $FLTAR_FULLPATH)
            If IsArray($aItems) Then
                For $j = 1 To $aItems[0]
                    ; Extraction du nom (dernier segment)
                    Local $sNom = StringRegExpReplace($aItems[$j], "^.*[\\/]", "")
                    ; Recherche du premier groupe de 6 chiffres consécutifs
                    Local $aMatch = StringRegExp($sNom, "(\d{6})", $STR_REGEXPARRAYMATCH)
                    If Not @error And IsArray($aMatch) And UBound($aMatch) > 0 Then
                        $sNumCandidat = $aMatch[0]
                        _Logging("Numéro trouvé: " & $sNumCandidat & " (dans """ & $aItems[$j] & """)", 3, 0, TimerDiff($iStart))
                        SplashOff()
                        Return $sNumCandidat
                    EndIf
                Next
            EndIf
        Next
    EndIf

    _Logging("Aucun numéro à 6 chiffres trouvé", 2, 0, TimerDiff($iStart))
    SplashOff()
    Return $sNumCandidat
EndFunc   ;==>_NumeroCandidat

;=========================================================

Func _NumeroCandidatSti()
	SplashTextOn("Sans Titre", "Détermination du numéro d'inscription du candidat." & @CRLF & @CRLF & "Veuillez patienter un moment..." & @CRLF, 330, 120, -1, -1, 49, "Segoe UI", 9)
	_Logging("🕗 Recherche du numéro d'inscription du candidat à partir des noms des sites web.", 2, 0)
	Local $iStart = TimerInit()
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
				_Logging("Numéro trouvé: """ & $Liste[$N] & """.", 2, 0, TimerDiff($iStart))
				ExitLoop
			EndIf
		Next
	EndIf

	If Not $Trouve Then
		_Logging("Aucun numéro trouvé.", 2, 0, TimerDiff($iStart))
		_Logging("🕗 Recherche du numéro d'inscription du candidat à partir des noms des bases de données.", 2, 0)
		Local $iStart = TimerInit()
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
					_Logging("Numéro trouvé: """ & $Liste[$N] & """.", 2, 0, TimerDiff($iStart))
					ExitLoop
				EndIf
			Next
		EndIf
		If Not $Trouve Then _Logging("Aucun numéro trouvé.", 2, 0, TimerDiff($iStart))
	EndIf


	SplashOff()
;~ 	WinActivate ($hMainGUI)
	Return $sNumCandidat
EndFunc   ;==>_NumeroCandidatSti

;=========================================================

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

;=========================================================

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
	$Lecteur = StringLeft(@WindowsDir,2) ; "C:" ; $aDrive[1] ; $aDrive[1] peut être A: !!
	For $i = 1 To $aDrive[0]
		If $aDrive[$i] = StringLeft(@WindowsDir,2) Then ContinueLoop
		If (DriveGetType($aDrive[$i], $DT_BUSTYPE) <> "USB") _ ; to Exclude external Hdd(s)
				And _WinAPI_IsWritable($aDrive[$i]) _ ;writable
				And DriveSpaceFree($aDrive[$i] & "\") > $FREE_SPACE_DRIVE_BACKUP _
				Then
			$Lecteur = $aDrive[$i]
			ExitLoop
		EndIf
	Next
	If $DossierBase = "" Then
		$Lecteur = LecteurSauvegarde()
		$DossierBase = $Lecteur & $DossierSauve
		$DossierBacCollector = $DossierBase & "\" & "BacCollector"
	EndIf
	$Lecteur = StringUpper($Lecteur) & "\"
;~ 	MsgBox(0,"",$DossierBase & "\0-CapturesEcran\BacBackup.ini")
	If $g_bBacBackupDetected And FileExists($DossierBase & "\BacBackup\BacBackup.ini") Then
		Local $DossierSessionBacBackup = IniRead($DossierBase & "\BacBackup\BacBackup.ini", "Params", "DossierSession", "")
		If $DossierSessionBacBackup <> "" Then
			FileWriteLine($hLogFile, "BacBackup   : """ & $DossierBase & "\BacBackup\" & $DossierSessionBacBackup & """")
		EndIf
	EndIf


	FileWriteLine($hLogFile, '')
EndFunc   ;==>_InitLogging

;=========================================================

Func _CopierLog($DestLocalFldr)
	FileFlush($hLogFile)
	FileCopy(@ScriptDir & "\" & StringTrimRight(@ScriptName, 4) & ".log", $DestLocalFldr)
EndFunc   ;==>_CopierLog

;=========================================================

Func _FinLogging()
	Local $sDateTime = "[" & _Now() & "]"
	FileWriteLine($hLogFile, "______" & $sDateTime & "______Fermeture de " & $PROG_TITLE & $PROG_VERSION)
	FileClose($hLogFile)
EndFunc   ;==>_FinLogging

;=========================================================

Func _Logging($sText, $iSuccess = 1, $iGuiLog = 1, $iDuree = 0) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
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
	Local $sDuree = ""
	if $iDuree <> 0 Then
		$sDuree = "[" & $iDuree & "ms] "
	EndIf
	FileWriteLine($hLogFile, $sTime & " " & $sPreLogFile & $sDuree & $sText)
EndFunc   ;==>_Logging

;=========================================================

Func _ClearGuiLog()
	_GUICtrlRichEdit_SetText($GUI_Log, "")
EndFunc   ;==>_ClearGuiLog

;=========================================================
Func VersionWXY()
	Local $parts = StringSplit(FileGetVersion(@ScriptFullPath), '.')
	return $parts[1] & '.' & $parts[2] & '.' & $parts[3]
EndFunc
;=========================================================
;~ Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
;~     #forceref $hWnd, $iMsg, $lParam

;~     Local $iIDFrom = BitAND($wParam, 0xFFFF)
;~     Local $iCode = BitShift($wParam, 16)

;~     If $iIDFrom = $TextDossiersRecuperes And $iCode = 1 Then
;~         Run("explorer.exe " & '"' & @ScriptDir & '"')
;~     EndIf

;~     Return $GUI_RUNDEFMSG
;~ EndFunc

;~ Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
;~ 	#forceref $iMsg, $wParam
;~ 	Local $tMsgFilter

;~ 	Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
;~ 	Local $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
;~ 	Local $iCode = DllStructGetData($tNMHDR, "Code")
;~ 	;;;; Test
;~ 	Local $tagNMHDR = DllStructCreate("int;int;int", $lParam)
;~ 	Local $code = DllStructGetData($tagNMHDR, 3)
;~ 	Local $hListViewItem, $aListLine

;~ 	If $wParam = $GUI_AfficherLeContenuDesDossiersBac And $code = -3 Then
;~ 		$hListViewItem = GUICtrlRead($GUI_AfficherLeContenuDesDossiersBac)
;~ 		If $hListViewItem Then
;~ 			$aListLine = StringSplit(GUICtrlRead($hListViewItem), "|")
;~ 			_ShowInExplorer($aListLine[4])
;~ 		EndIf

;~ 	ElseIf $wParam = $GUI_AfficherLeContenuDesAutresDossiers And $code = -3 Then
;~ 		$hListViewItem = GUICtrlRead($GUI_AfficherLeContenuDesAutresDossiers)
;~ 		If $hListViewItem Then
;~ 			$aListLine = StringSplit(GUICtrlRead($hListViewItem), "|")
;~ 			_ShowInExplorer($aListLine[4])
;~ 		EndIf
;~ 	Else
;~ 		;;;; Test
;~ 		Switch $hWndFrom
;~ 			Case $GUI_Log
;~ 				Select
;~ 					Case $iCode = $EN_MSGFILTER
;~ 						$tMsgFilter = DllStructCreate($tagMSGFILTER, $lParam)
;~ 						If DllStructGetData($tMsgFilter, "msg") = $WM_LBUTTONDBLCLK Then
;~ 							_OpenTempLog()
;~ 						EndIf

;~ 				EndSelect
;~ 		EndSwitch
;~ 	EndIf

;~ 	Return $GUI_RUNDEFMSG
;~ EndFunc   ;==>WM_NOTIFY

;=========================================================

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

;=========================================================

Func _OpenTempLog()
	Local $Editeur = _EditeurFichierLog()
	If $Editeur <> "" Then
		Local $Text = _GUICtrlRichEdit_GetText($GUI_Log)
		If StringLen($Text) = 0 Then Return
		Local $sFilePath = @TempDir & "\" & $PROG_TITLE & "_tmp_log.rtf"
		Local $hTmpFile = FileOpen($sFilePath, $FO_OVERWRITE)
		Local $Kes = _GUICtrlRichEdit_StreamToVar($GUI_Log)
		$Kes = StringReplace($Kes, "\red255\green255\blue255", "\red0\green0\blue0")
		$Kes = StringReplace($Kes, "MS Shell Dlg", "Segoe UI")
		$Kes = StringReplace($Kes, "\fs17\", "\fs21\")
		FileWriteLine($hTmpFile, $Kes)
		FileClose($hTmpFile)
		ShellExecute($Editeur, """" & $sFilePath & """")
;~ 		_RunDos('Start Wordpad "' & $FilePath & '"')
	Else
		Local $sFilePath = @TempDir & "\" & $PROG_TITLE & "_tmp_log.txt"
		Local $Text = _GUICtrlRichEdit_GetText($GUI_Log, True)
		Local $hFileOpen = FileOpen($sFilePath, $FO_OVERWRITE )
		FileWrite($hFileOpen, $Text)
		FileClose($hFileOpen)
		ShellExecute("""" & $sFilePath & """")
	EndIf

EndFunc   ;==>_OpenTempLog

Func _EditeurFichierLog()
    ; Recherche WordPad avec gestion de l'architecture
    Local $sWordPadPath
    If @OSArch = "X64" Then
        $sWordPadPath = EnvGet("ProgramW6432") & "\Windows NT\Accessories\wordpad.exe"
        If Not FileExists($sWordPadPath) Then
            $sWordPadPath = @ProgramFilesDir & "\Windows NT\Accessories\wordpad.exe"
        EndIf
    Else
        $sWordPadPath = @ProgramFilesDir & "\Windows NT\Accessories\wordpad.exe"
    EndIf
    If FileExists($sWordPadPath) Then Return $sWordPadPath

    ; Recherche Word via les versions Office
    Local $aVersions[] = ["16.0", "15.0", "14.0", "12.0", "11.0"]
	Local $sPath
    For $vVersion In $aVersions
        $sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\" & $vVersion & "\Word\InstallRoot"
        $sPath = RegRead($sRegKey, "Path")
        If Not @error Then
            $sPath = StringRegExpReplace($sPath, "\\$", "") & "\WINWORD.EXE"
            If FileExists($sPath) Then Return $sPath
        EndIf
    Next

    ; Recherche dans les chemins standards
    Local $aOfficePaths[] = [ _
        "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", _
        "C:\Program Files\Microsoft Office\Office16\WINWORD.EXE", _
        "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE", _
        "C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE" _
    ]
    For $sPath In $aOfficePaths
        If FileExists($sPath) Then Return $sPath
    Next

    ; Recherche via les App Paths
    Local $aRegPaths[] = ["HKLM64\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe", "HKLM32\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe"]
    For $i = 0 To UBound($aRegPaths) - 1
        Local $sPath = RegRead($aRegPaths[$i], "")
        If Not @error And FileExists($sPath) Then Return $sPath
    Next

    Return ""
EndFunc

Func _CheckRunningFromUsbDrive()
	If FileExists(@ScriptDir & "\__dev.txt") Then Return
	Local $Drive = StringLeft(@ScriptDir, 2)

;~ _ExtMsgBox ($vIcon, $vButton, $sTitle, $sText, [$iTimeout, [$hWin, [$iVPos, [$bMain = True]]]])
	If (DriveGetType($Drive, $DT_BUSTYPE) <> "USB") Then
;~ 		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBoxSet(1, 0, 0x004080, 0xFFFF00, 9, "Segoe UI")
		_ExtMsgBox(128, "Ok", $PROG_TITLE & " - Avertissement", $PROG_TITLE & " est exécuté à partir du disque local." & @CRLF & @CRLF _
				 & "Le bouton ""Récupérer"" sera désactivé." & @CRLF, 12, $hMainGUI)
		GUICtrlSetState($bRecuperer, $GUI_DISABLE)
	ElseIf Not _WinAPI_IsWritable($Drive) Then
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI")
		_ExtMsgBox(128, "Ok", $PROG_TITLE & " - Erreur", "Le lecteur """ & $Drive & "\"" n'est pas enregistrable, veuillez exécuter " & $PROG_TITLE & " à partir d'une Clé USB." & @CRLF & @CRLF _
				 & "Le bouton ""Récupérer"" sera désactivé." & @CRLF, 12, $hMainGUI)
		GUICtrlSetState($bRecuperer, $GUI_DISABLE)
	ElseIf DriveSpaceFree($Drive & "\") < 100 Then ;100 Mo
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI")
		_ExtMsgBox(128, "Ok", $PROG_TITLE & " - Erreur", "L'espace disque dans la Clé USB est insuffisant pour la récupération des dossiers de travail des candidats.." & @CRLF & @CRLF _
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
			If FileExists($DossierBase) Then
				$TmpMsg &= "Déverrouillage du dossier [" & $DossierBase & "]" & @CRLF
				_UnLockFolder($DossierBase)
				FileSetAttrib($DossierBase, "-RASH")
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
			_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Le lecteur " & $sScriptDrive & " contenant l'exécutable de " & @ScriptName & " n'est plus accessible. " & @CRLF _
								& "Ne retirez pas la clé USB avant la fermeture du programme pour éviter toute corruption des données." & @CRLF _
								& $PROG_TITLE & " va se fermer immédiatement.", 0)
		Else
			_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
			_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "L'exécutable """ & @ScriptName & """ n'existe plus dans son emplacement d'origine." & @CRLF _
					 & $PROG_TITLE & " va se fermer immédiatement.", 0)
		EndIf
		Exit
	EndIf
EndFunc

;=========================================================
Func _AfficherDossierRecuperes($sLog = 0)
	Local $iStart = TimerInit()
	Local $sDossiersRecuperes = _ListeDeDossiersRecuperes() ;_ListeDeDossiersRecuperes($SearchMask = "??????", $RegEx = "([0-9]{6})")
	If $sDossiersRecuperes <> "" Then
		GUICtrlSetData($TextDossiersRecuperes, $sDossiersRecuperes)
		If $sLog = 1 Then
			$sTmpSpaces = "                         "
			_Logging("Liste de dossiers de travail déjà récupérés : " & StringReplace($sDossiersRecuperes, @CRLF, @CRLF & $sTmpSpaces), 2, 0, TimerDiff($iStart))
		EndIf
	Else
		GUICtrlSetData($TextDossiersRecuperes, "-")
	EndIf
EndFunc
;=========================================================
Func _OpenDialogAbsent()
	Local Const $GUI_COLOR_ERROR = 0xFFFF00 ; Jaune vif
    ; Création de la boîte de dialogue
    Local $hDialog = GUICreate("Ajout d'un Candidat Absent", 350, 220, -1, -1, -1, $WS_EX_TOPMOST)
    GUISetBkColor($GUI_COLOR_SIDES, $hDialog)

    ; ===== Header =====
    GUICtrlCreateGraphic(0, 0, 350, $GUI_HEADER_HAUTEUR)
    GUICtrlSetBkColor(-1, $GUI_COLOR_SIDES)
    GUICtrlSetState(-1, $GUI_DISABLE)

    ; Titre centré
    Local $HeaderText = GUICtrlCreateLabel("Numéro d'inscription du Candidat absent", 0, 3, 350, $GUI_HEADER_HAUTEUR-6, $SS_CENTER)
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER_HEADERS_TEXT)
    GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
    GUICtrlSetFont(-1, 11, 800) ; Texte plus grand et gras

    ; ===== Zone de saisie =====
    Local $InputWidth = 200
    Local $InputX = (350 - $InputWidth)/2
    Local $hInput = GUICtrlCreateInput("", $InputX, $GUI_HEADER_HAUTEUR + 30, $InputWidth, 40, BitOR($ES_NUMBER, $ES_CENTER))
    GUICtrlSetFont(-1, 22, 900, 0, "Arial")
    GUICtrlSetColor(-1, $GUI_COLOR_CENTER)
    GUICtrlSetLimit(-1, 6)
    GUICtrlSetTip(-1, "Numéro à 6 chiffres - Vérifiez l'existence du dossier avant validation", "Numéro d'Inscription", 1, 1)

    ; ===== Label d'erreur =====
    Local $hErrorLabel = GUICtrlCreateLabel("", $GUI_MARGE, $GUI_HEADER_HAUTEUR + 80, 350 - 2*$GUI_MARGE, 40, $SS_CENTER)
    GUICtrlSetColor(-1, $GUI_COLOR_ERROR)
    GUICtrlSetFont(-1, 10, 800) ; Texte agrandi et gras
    GUICtrlSetState(-1, $GUI_HIDE)

    ; ===== Boutons =====
    Local $BtnWidth = 90
    Local $BtnX = (350 - 2*$BtnWidth - $GUI_MARGE)/2
    Local $hBtnValidate = GUICtrlCreateButton("Valider", $BtnX, 170, $BtnWidth, 35)
    GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
    GUICtrlSetColor(-1, 0xFFFFFF)
    GUICtrlSetFont(-1, 10, 800)

    Local $hBtnCancel = GUICtrlCreateButton("Annuler", $BtnX + $BtnWidth + $GUI_MARGE, 170, $BtnWidth, 35)
    GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
    GUICtrlSetColor(-1, 0xFFFFFF)
    GUICtrlSetFont(-1, 10, 800)

    GUISetState(@SW_SHOW, $hDialog)

    While 1
        Local $nMsg = GUIGetMsg()
        Switch $nMsg
            Case $GUI_EVENT_CLOSE, $hBtnCancel
                GUIDelete($hDialog)
                ExitLoop
            Case $hBtnValidate
                Local $sInput = GUICtrlRead($hInput)
				Local $ScriptDir = @ScriptDir
				If StringRight($ScriptDir, 1) <> "\" Then
					$ScriptDir = $ScriptDir & "\"
				EndIf

                Local $Dest1FlashUSB = $ScriptDir & $sInput ; Chemin du dossier principal

                ; Validation multi-critères
                If StringLen($sInput) <> 6 Then
                    GUICtrlSetData($hErrorLabel, "Le numéro doit contenir exactement 6 chiffres")
                    GUICtrlSetState($hErrorLabel, $GUI_SHOW)
                ElseIf $sInput = "000000" Then
                    GUICtrlSetData($hErrorLabel, """" & $sInput & """ n'est pas un numéro d'inscription valide")
                    GUICtrlSetState($hErrorLabel, $GUI_SHOW)
                ElseIf FileExists($Dest1FlashUSB) Then
                    GUICtrlSetData($hErrorLabel, "Un dossier avec le numéro " & $sInput & " existe déjà !")
                    GUICtrlSetState($hErrorLabel, $GUI_SHOW)
                Else
                    GUICtrlSetState($hErrorLabel, $GUI_HIDE)
                    GUIDelete($hDialog)
                    Return _CreerDossierCandidatAbs($Dest1FlashUSB) ; Appel avec le chemin complet
                    ExitLoop
                EndIf
        EndSwitch
    WEnd
EndFunc
Func 	_VerifierDossiersNonConformesPourCandidatsAbsents()
	If $DossiersNonConformesPourCandidatsAbsents <> "" Then
		If _NbOccurrences("»", $DossiersNonConformesPourCandidatsAbsents) = 1 Then
			; Cas d'un seul dossier
			Local $sMessage = "Dossier non conforme détecté pour un candidat absent :" & @CRLF & @CRLF & _
							  $DossiersNonConformesPourCandidatsAbsents & @CRLF & _
							  "Action nécessaire :" & @CRLF & _
							  "→ Créez un sous-dossier VIDE nommé « Absent » dans ce dossier."
		Else
			; Cas de plusieurs dossiers
			Local $sMessage = "Dossiers non conformes détectés pour des candidats absents :" & @CRLF & @CRLF & _
							  $DossiersNonConformesPourCandidatsAbsents & @CRLF & _
							  "Action nécessaire :" & @CRLF & _
							  "→ Créez un sous-dossier VIDE nommé « Absent » dans chaque dossier listé ci-dessus."
		EndIf
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_WARNING, 0x000000, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox($EMB_ICONEXCLAM, "Ok", $PROG_TITLE & " - Avertissement", $sMessage, 0)
	EndIf
EndFunc

Func _CreerDossierCandidatAbs($sCheminDossier)
    ; Création du dossier principal
    If Not DirCreate($sCheminDossier) Then
		_Logging("Candidat Absent: Échec de la création du dossier """ & $sCheminDossier & """", 5, 1)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec de la création du dossier :" & @CRLF & $sCheminDossier, 0)
        Return 0
    EndIf

    ; Création du sous-dossier "Absent"
    Local $sSousDossier = $sCheminDossier & "\Absent"
    If Not DirCreate($sSousDossier) Then
		_Logging("Candidat Absent: Échec de la création du dossier : """ & $sSousDossier & """", 5, 1)
		_ExtMsgBoxSet(1, 0, $GUI_COLOR_ERROR, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
		_ExtMsgBox(16, "Ok", $PROG_TITLE & " - Erreur", "Échec de la création du dossier :" & @CRLF & $sSousDossier, 0)
        Return 0
    EndIf

    ; Confirmation de création
	_Logging("Candidat Absent: Dossier créé avec succès """ & $sSousDossier & """", 3, 1) ;$iSuccess: 0>Fail&Red , 1>Success&Blanc, 2>Info&Blue , 3>Info&Green, 4>info&Blanc, 5>info&Red
	_ExtMsgBoxSet(1, 0, $GUI_COLOR_CENTER, 0xFFFFFF, 9, "Segoe UI", @DesktopWidth - 25, @DesktopWidth - 25)
	_ExtMsgBox(64, "Ok", $PROG_TITLE & " - Succès", "Dossier créé avec succès :" & @CRLF & $sCheminDossier & _
           @CRLF & @CRLF & "✓ Sous-dossier marqueur :" & @CRLF & $sSousDossier, 0)
	Return 1
EndFunc
;=========================================================




;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦¦
;¦¦¦
;~ _ArrayDisplay($TmpList, $PROG_TITLE ,"",32,Default ,"Liste de Dossiers/Fichiers Surveillés")
;¦¦¦
;~ MsgBox ( 0, "", $sDir  )
;¦¦¦
;~ 🕗
;~ Local $iStart = TimerInit()
;~ TimerDiff($iStart)
