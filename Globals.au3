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
Global Const $GUI_COLOR_CENTER = 0x0F172A
Global Const $GUI_COLOR_SIDES = 0x1E293B
Global Const $GUI_COLOR_CENTER_HEADERS_TEXT = 0xF8FAFC
Global Const $GUI_COLOR_BORDER = 0x777777

Global Const $GUI_COLOR_BTN_MAIN = 0x10B981; 0xF59E0B ; 0x10B981  ; Vert Émeraude pour "Récupérer"
Global Const $GUI_COLOR_BTN_PANEL = 0x475569 ; Gris-Bleu pour le petit bouton volet
;~ ===========================================

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
; ============================================================================
; MODE SIMULATION D'ERREURS — À DÉSACTIVER AVANT COMPILATION FINALE
; Valeurs de $TEST_ERREUR_COPIE :
;   0 = mode normal (production)
;   1 = DirCopy retourne toujours 0 (échec copie total)
;   2 = DirCopy réussit mais intégrité MD5 échoue
;   3 = échec uniquement sur la destination USB (pas locale)
;   4 = échec uniquement sur la destination locale (pas USB)
;   5 = échec au 1er essai, succès au 2ème (teste le retry)
; ============================================================================
Global $TEST_ERREUR_COPIE = 0
Global $TEST_ERREUR_COPIE_COMPTEUR = 0  ; compteur interne pour le mode 5

Global $__g_sWarnIcon
If @OSVersion = "WIN_7" Or @OSVersion = "WIN_XP" Or @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_2008" Then
    $__g_sWarnIcon = "[!]" ; Rendu propre pour Windows 7 et antérieurs
Else
    $__g_sWarnIcon = "⚠️"   ; Emoji pour Windows 8, 8.1, 10, 11 et futurs
EndIf

; Icône du bouton de fermeture forcée des processus
; ⨂ (U+2A02) = N-ARY CIRCLED TIMES OPERATOR — cercle avec X, sobre et lisible
; Garanti affichable dans les contrôles GDI Win32 (plan BMP Unicode)
Global $__g_sKillIcon
If @OSVersion = "WIN_7" Or @OSVersion = "WIN_XP" Or @OSVersion = "WIN_VISTA" Or @OSVersion = "WIN_2008" Then
    $__g_sKillIcon = "[X]" ; ASCII pur pour Windows 7 et antérieurs
Else
    $__g_sKillIcon = "⨂"   ; U+2A02 pour Windows 8, 8.1, 10, 11 et futurs
EndIf