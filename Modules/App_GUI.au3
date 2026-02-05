#include-once
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
#include <Misc.au3>
#include <Process.au3>
#include <StaticConstants.au3>
#include <WinAPIFiles.au3>
#include <WinAPIProc.au3>
#include <WinAPIShellEx.au3>
#include <WindowsConstants.au3>

; Global variables used in GUI
Global $hMainGUI, $GUI_Log, $aMsg
Global $GUI_NumeroCandidat, $bRecuperer, $bCreerSauvegarde, $bOpenBackupFldr, $bTogglePart3, $bAPropos
Global $TextApps, $bCreerDossierCandidatAbs, $lblBacBackup, $lblUsbCleaner
Global $cBac, $cSeance, $cLabo, $lblBac, $lblSeance, $lblLabo, $lblMatiere
Global $rInfoProg, $tInfoProg, $rSti, $tSti
Global $HeadContenuDossiersBac, $HeadContenuAutresDossiers
Global $GUI_AfficherLeContenuDesDossiersBac, $GUI_AfficherLeContenuDesAutresDossiers
Global $TextDossiersRecuperes, $TextApps_Header, $TextApps_Text
Global $g_bDragging = False, $g_iOffsetX, $g_iOffsetY

; Include module files
#include "..\Utils.au3"
#include "..\Globals.au3"
#include "..\include\ExtMsgBox.au3"
#include "..\include\GUIExtender.au3"
#include "..\include\StringSize.au3"

;=====================================
; Create the main GUI interface
; @return void
;=====================================
Func _CreateGui()
	;;; Fenêtre principale  - Début
	Local $sMode = ""
	If IsAdmin() Then
		$sMode = " [Administrateur]"
	EndIf
	Global $hMainGUI = GUICreate($PROG_TITLE & $PROG_VERSION & $sMode, $GUI_LARGEUR, $GUI_HAUTEUR, -1, -1, -1, 0)
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
	GUICtrlSetTip($bTogglePart3, " ", "Cliquez pour afficher/masquer la partie droite de la fenêtre.", 1,1)
	GUICtrlSetCursor($bTogglePart3, $MCID_HAND )
	_GUIExtender_Section_Action($hMainGUI, 1, 0) ;2ème paramètre (1) Numéro de la Section  -  3ème paramètre (0) to retract the Section
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
	GUICtrlSetTip($Text, "- Info/Prog." & @CRLF & "- STI.", "Matière:", 1,1)
	;;;========Content============
	Local $TmpTop = $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE / 2
	Local $TmpLeft = $Left_Header + $WidthHeader / 2 + $GUI_MARGE
	Global $lblMatiere = GUICtrlCreateLabel("Matière...", $TmpLeft, $TmpTop, $WidthHeader / 2 - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR + 10, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetFont(-1, 16, 100, 4, "Segoe UI Light")
	GUICtrlSetTip($lblMatiere, "Cliquez pour Actualiser tout.", "Actualiser.", 1,1)
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
	Global $rSti = GUICtrlCreateRadio("", $TmpLeft, $TmpTop, 10, 20) ;Bouton Radio Sc./M./Tech.
	Global $tSti = GUICtrlCreateLabel("STI", $TmpLeft + 15, $TmpTop + 3, 35, 20)
	GUICtrlSetColor(-1, 0xEEEEEE)
	GUICtrlSetTip($rSti, "STI - Systèmes & Technologies de l'Informatique", "Matière", 1,1)
	GUICtrlSetTip($tSti, "STI - Systèmes & Technologies de l'Informatique", "Matière", 1,1)
;~ 	GUICtrlSetState($rSti, $GUI_DISABLE) ;Voir aussi _InitialParams
;~ 	GUICtrlSetState($tSti, $GUI_DISABLE) ;Voir aussi _InitialParams
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
	GUICtrlSetTip($lblComputerID, "L'Identifiant Unique de ce PC." & @CRLF & @CRLF & "Cliquez pour copier.", GUICtrlRead($lblComputerID), 1,1)
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
	GUICtrlSetTip($bRecuperer, @CRLF & "Cette commande permet de:" & @CRLF 	& @CRLF & "1. Sauvegarder le travail du candidat vers un dossier verrouillé sur ce PC." & @CRLF & "2. Copier les dossiers & fichiers du candidat vers la clé USB." & @CRLF & "3. Supprimer les travaux du candidat, pour les matières Info & Prog.", "Copier les fichiers du candidat vers la clé USB", 1,1)
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
	GUICtrlSetTip($Text, " ", "Dossiers récupérés", 1,1)
	Global $TextDossiersRecuperes = GUICtrlCreateLabel("", $Left_Header + $GUI_MARGE, $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE + 6, $WidthHeader - 2 * $GUI_MARGE, $GUI_HAUTEUR - 310)
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 10, 200)
	GUICtrlSetTip($TextDossiersRecuperes, " " & @CRLF & "Double-clic pour ouvrir l'emplacement des dossiers récupérés.", "Dossiers récupérés", 1,1)
	$Top_Header = $GUI_HAUTEUR - 100 - $GUI_MARGE
	$Left_Header = $GUI_MARGE
	$WidthHeader = $GUI_LARGEUR_PARTIE - 2 * $GUI_MARGE
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~[
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~[
	Local $TmpButtonWidth = $GUI_LARGEUR_PARTIE - 4 * $GUI_MARGE ; * 2/3
	Local $TmpButtonHeight = $GUI_HEADER_HAUTEUR * 2
;~ 	$Top_Header = $GUI_HAUTEUR - 2 * $GUI_MARGE - 3 * $TmpButtonHeight
	Global $bCreerDossierCandidatAbs = GUICtrlCreateButton("Ajouter Candidat Absent", $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $Top_Header - $TmpButtonHeight, $TmpButtonWidth, $TmpButtonHeight * 0.75)
	GUICtrlSetColor(-1, 0xffffff)
	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
	GUICtrlSetFont(-1, 9)
	GUICtrlSetTip($bCreerDossierCandidatAbs, @CRLF & "Ce bouton permet de :" & @CRLF  & @CRLF & @TAB & "1. Créer un dossier portant le numéro du candidat absent." & @CRLF & @TAB & "2. Ajouter à l'intérieur un sous-dossier vide nommé ""Absent""", "Ajouter un candidat absent", 1,1)
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~[
; ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~[
	Global $TextApps_Header = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $TmpHeaderHauteur)
	GUICtrlSetColor($TextApps_Header, 0x66c7fc)
;~ 	GUICtrlSetState($TextApps_Header, BitOR($GUI_HIDE, $GUI_DISABLE))
	GUICtrlSetState($TextApps_Header, BitOR($GUI_SHOW, $GUI_DISABLE))
	Global $TextApps_Text = GUICtrlCreateLabel("Logiciels Ouverts", $Left_Header + $GUI_MARGE, $Top_Header + 3, $WidthHeader - 2 * $GUI_MARGE, $GUI_HEADER_HAUTEUR, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($TextApps_Text, 0xFFFFFF)
	GUICtrlSetBkColor($TextApps_Text, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 9.5, 500)
	GUICtrlSetState($TextApps_Text, $GUI_SHOW)
	Local $sTempMessage = @CRLF & "Enregistrer les travaux et fermer les applications listées." & @CRLF & @CRLF _
					& "🛑  Attention : certaines applications, notamment VSCode et Sublime Text," & @CRLF _
					& "peuvent se fermer sans confirmation, même si des fichiers sont encore non enregistrés."
	GUICtrlSetTip(-1, $sTempMessage , "Fermeture d'applications requise", 2,1)
	Global $TextApps = GUICtrlCreateLabel("", $Left_Header + $GUI_MARGE, $Top_Header + $GUI_HEADER_HAUTEUR + $GUI_MARGE + 6, $WidthHeader - 2 * $GUI_MARGE, 100 - $GUI_HEADER_HAUTEUR - 2 * $GUI_MARGE)
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 9)
	GUICtrlSetState($TextApps, $GUI_SHOW)
	GUICtrlSetTip(-1, $sTempMessage, "Fermeture d'applications requise", 2,1)
	;;;Partie à gauche  - Fin=============================================================================================
	;;;Partie à Droite  - Début=============================================================================================
	Local $GuiTmpLeft = 3 * $GUI_LARGEUR_PARTIE
	Local $GuiTmpTop = $GUI_MARGE
	Global $lblLabo = GUICtrlCreateLabel("", $GuiTmpLeft + $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $GuiTmpTop, $TmpButtonWidth, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($lblLabo, 0xFFFFFF)
	GUICtrlSetBkColor($lblLabo, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 18, 100, 4, "Segoe UI Light")
	GUICtrlSetTip($lblLabo, " ", "Laboratoire d'Informatique", 1,1)
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
	GUICtrlSetTip($bCreerSauvegarde, @CRLF & "Cette commande permet de:" & @CRLF & @CRLF & @TAB & "1. Sauvegarder les travaux des candidats dans un dossier verrouillé." & @CRLF & @TAB & "2. Générer un état sur les travaux des candidats." & @CRLF & @TAB & "3. Générer la grille d'évaluation correspondante (Excel).", "Sauvegarder les travaux sur serveur/poste réserve", 1,1)
	$GuiTmpLeft = $GuiTmpLeft
	$GuiTmpTop = $GuiTmpTop + $GUI_MARGE + $TmpButtonHeight
	Global $bOpenBackupFldr = GUICtrlCreateButton("Ouvrir Dossier de Sauve", $GuiTmpLeft, $GuiTmpTop, $TmpButtonWidth, $TmpButtonHeight)
	GUICtrlSetColor(-1, 0xffffff)
	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
	GUICtrlSetFont(-1, 10)
	GUICtrlSetTip($bOpenBackupFldr, @CRLF & " ", "Ouvrir le dossier de sauvegarde", 1,1)
;~ 	Global $bAide = GUICtrlCreateButton("Aide", $GuiTmpLeft, $GUI_HAUTEUR - $GUI_MARGE - $TmpButtonHeight, $TmpButtonWidth, $TmpButtonHeight)
;~ 	GUICtrlSetColor(-1, 0xffffff)
;~ 	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
;~ 	GUICtrlSetFont(-1, 10)
;~ 	GUICtrlSetTip($bAide, " ", "Aide & Fonctionnalités de BacCollector.", 1,1)
	Global $bAPropos = GUICtrlCreateButton("À propos", $GuiTmpLeft, $GUI_HAUTEUR - $GUI_MARGE - $TmpButtonHeight, $TmpButtonWidth, $TmpButtonHeight)
	GUICtrlSetColor(-1, 0xffffff)
	GUICtrlSetBkColor(-1, $GUI_COLOR_CENTER)
	GUICtrlSetFont(-1, 10)
	;;; Cadre BacBackup
	$Left_Header = $GUI_MARGE + 3 * $GUI_LARGEUR_PARTIE
;~ 	$Top_Header = $GuiTmpTop + $TmpButtonHeight + 2 * $GUI_MARGE
	$Top_Header = $GUI_HAUTEUR - 2 * $GUI_MARGE - 2 * $TmpButtonHeight
	$WidthHeader = $GUI_LARGEUR_PARTIE - 2 * $GUI_MARGE
	Local $CBacBackup = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($CBacBackup, 0x66c7fc)
	GUICtrlSetState($CBacBackup, $GUI_DISABLE)
;~ 	Global $lblBacBackup = GUICtrlCreateLabel("BacBackup est installé", 3 * $GUI_LARGEUR_PARTIE + $GUI_LARGEUR_PARTIE / 2 - $TmpButtonWidth / 2, $GUI_HAUTEUR - 2 * $GUI_MARGE - 2 * $TmpButtonHeight, $TmpButtonWidth, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	Global $lblBacBackup = GUICtrlCreateLabel("BacBackup est installé", $Left_Header, $Top_Header, $WidthHeader, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($lblBacBackup, 0xFFFFFF)
	GUICtrlSetBkColor($lblBacBackup, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 11, 100, 0, "Segoe UI Light")
;~ 	GUICtrlSetState($lblBacBackup, $GUI_HIDE)
	GUICtrlSetTip($lblBacBackup, "Cliquez pour ouvrir BacBackup", "BacBackup surveille le PC.", 1,1)
	GUICtrlSetCursor(-1, 0) ; Curseur main
	;;; Cadre USBCleaner
	$Top_Header = $GUI_HAUTEUR - 2 * $GUI_MARGE - 3 * $TmpButtonHeight
	$WidthHeader = $GUI_LARGEUR_PARTIE - 2 * $GUI_MARGE
	Global $CUsbCleaner = GUICtrlCreateGraphic($Left_Header, $Top_Header, $WidthHeader, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($CUsbCleaner, 0x66c7fc)
	GUICtrlSetState($CUsbCleaner, $GUI_DISABLE)
	Global $lblUsbCleaner = GUICtrlCreateLabel($__g_sWarnIcon & " UsbCleaner", $Left_Header, $Top_Header, $WidthHeader, $TmpButtonHeight, BitOR($SS_CENTER, $SS_CENTERIMAGE))
	GUICtrlSetColor($lblUsbCleaner, 0xFF0000)
	GUICtrlSetBkColor($lblUsbCleaner, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetFont(-1, 10, 900, 0, "Segoe UI Light")
	GUICtrlSetTip($lblUsbCleaner, "Cliquez pour accéder au téléchargement de UsbCleaner", "UsbCleaner - Protection inactive", 2, 1)
	GUICtrlSetCursor(-1, 0) ; Curseur main
	GUICtrlSetState($lblUsbCleaner, $GUI_HIDE)
	GUICtrlSetState($CUsbCleaner, $GUI_HIDE)
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

;=====================================
; Toggle the right panel of the GUI
; @return void
;=====================================
Func _TogglePart3()
	Switch _GUIExtender_Section_State($hMainGUI, 1)
		Case 0
			; Extend to the right (default)
			_GUIExtender_Section_Action($hMainGUI, 1, 1)
			IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "GuiPart3Extend", 1)
			_Logging("Ouverture de la partie droite de l'interface", 2, 0)
		Case 1
			; Retract from the right
			_GUIExtender_Section_Action($hMainGUI, 1, 0)
			IniWrite(StringTrimRight(@ScriptFullPath, 4) & ".ini", "Params", "GuiPart3Extend", 0)
			_Logging("Fermeture de la partie droite de l'interface", 2, 0)
	EndSwitch
EndFunc   ;==>_TogglePart3
;=====================================
; Callback function for GUI notifications
; Handles double-clicks on list items and rich edit events
; @param $hWnd Handle of the window
; @param $iMsg Message ID
; @param $wParam Additional message information
; @param $lParam Additional message information
; @return GUI_RUNDEFMSG or result of operation
;=====================================
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
Func WM_COMMAND($hWnd, $iMsg, $wParam, $lParam)
    #forceref $hWnd, $iMsg, $lParam

    Local $iIDFrom = BitAND($wParam, 0xFFFF)
    Local $iCode = BitShift($wParam, 16)

    If $iIDFrom = $TextDossiersRecuperes And $iCode = 1 Then
        Run("explorer.exe " & '"' & @ScriptDir & '"')
    EndIf

    Return $GUI_RUNDEFMSG
EndFunc
