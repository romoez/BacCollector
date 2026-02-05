#include-once

#Region ;**** Global Const ****
;#Program.Info.Const
Global Const $PROG_TITLE = "BacCollector "
;~ Global Const $PROG_VERSION = VersionWXY() ;"0.8.2.0" --> "0.8.2"
Global Const $PROG_VERSION = FileGetVersion(@ScriptFullPath)

Global $DossierSauve = "Sauvegardes"
Global $Lecteur = ""
Global $DossierBase = ""
Global $DossierBacCollector = ""

Global Const $ANNEES_BAC[5] = ["2026", "2027", "2028", "2029", "2030"]

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
Global Const $GUI_COLOR_SUCCESS = 0x007848
Global Const $GUI_COLOR_ERROR = 0x660000
Global Const $GUI_COLOR_WARNING = 0xFFB300
Global Const $GUI_COLOR_INFO = 0x004080


Global Const $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES = 80
Global Const $AGE_MAX_FICHIERS_A_COPIER__EN_MINUTES_STI = 120
Global Const $TAILLE_MAX_DU_DOSSIER_SOUS_LECTEUR = 10 ;en Mb
Global Const $sUserName = @UserName
#EndRegion ;**** Global Const ****

Global $g_bDragging = False, $g_iOffsetX, $g_iOffsetY
Global $g_iGUITransparence = 255

Global $g_bBacBackupDetected


; Au moins 15 Go d'espace libre pour garantir que Windows peut fonctionner correctement
Global Const $MINIMUM_WINDOWS_FREE_SPACE = 15000 ; en MB
; 5 Go minimum requis pour un lecteur non-système
Global Const $FREE_SPACE_DRIVE_BACKUP = 5000 ; en MB
; Cache global pour les installations XAMPP-LITE/XAMPP/WAMP
Global $__g_aEasyPHPRootsCache = 0
Global $__g_sNomFichierAlerteFraude = "!_FRAUDE_POSSIBLE_USB.txt"

Global $__g_sWarnIcon
If @OSVersion = "WIN_7" Or @OSVersion = "WIN_XP" Or @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_2008" Then
    $__g_sWarnIcon = "[!]" ; Rendu propre pour Windows 7 et antérieurs
Else
    $__g_sWarnIcon = "⚠️"   ; Emoji pour Windows 8, 8.1, 10, 11 et futurs
EndIf