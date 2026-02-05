#include-once
#include <Crypt.au3>
#include <APIDiagConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <File.au3>
#include <Date.au3>
#include <WinAPIProc.au3>
#include <Globals.au3>
;#NoTrayIcon


; ============================================================================
;Nombre d'occurrences de $substring dans $string
Func _NbOccurrences($substring, $string)
	StringReplace($string, $substring, $substring)
	Return @extended
EndFunc   ;==>_NbOccurrences
; ============================================================================
Func _FineSize($iTaille) ;reçoit une taille en Octet >> retourne la taille en "multiple" approprié Ko ou Mo...
	; Tableau des unités de mesure
	Local $aUnits = ["oct.", "Ko", "Mo", "Go", "To", "Po", "Eo", "Zo", "Yo"]

	; Initialiser l'index de l'unité
	Local $iUnitIndex = 0

	; Boucle pour convertir la taille en multiples appropriés
	While $iTaille >= 1024 And $iUnitIndex < UBound($aUnits) - 1
		$iTaille /= 1024
		$iUnitIndex += 1
	WEnd

	; Arrondir la taille à une décimale et retourner avec l'unité correspondante
	Return Round($iTaille, 1) & " " & $aUnits[$iUnitIndex]
EndFunc   ;==>_FineSize
;#########################################################################################
;#########################################################################################
; ============================================================================
Func DossiersBac($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
    Local $aResult = _FileListToArray( _
        StringLeft(@WindowsDir, 2), _
        "bac*2*", _
        $FLTAR_FOLDERS, _
        $FullPath ? $FLTAR_FULLPATH : $FLTAR_RELPATH)

    If Not IsArray($aResult) Then Return _EmptyArray()
    Return $aResult
EndFunc
; ============================================================================
Func DossiersRessources($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
    Local $aResult = _FileListToArray( _
        StringLeft(@WindowsDir, 2), _
        "Res*ource*", _
        $FLTAR_FOLDERS, _
        $FullPath ? $FLTAR_FULLPATH : $FLTAR_RELPATH)

    If Not IsArray($aResult) Then Return _EmptyArray()
    Return $aResult
EndFunc
; ============================================================================
Func _DossiersTravailEleves($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
    Local $aResult = _FileListToArrayRec(StringLeft(@WindowsDir, 2), _
        "1*;2*;3*;4*;7*;8*;9*;dc*;ds*", _
        $FLTAR_FOLDERS + $FLTAR_NOHIDDEN + $FLTAR_NOSYSTEM + $FLTAR_NOLINK, _
        $FLTAR_NORECUR, $FLTAR_FASTSORT, $FullPath ? $FLTAR_FULLPATH : $FLTAR_RELPATH)

    If Not IsArray($aResult) Then Return _EmptyArray()
    Return $aResult
EndFunc

; ============================================================================
Func _DossiersSurBureau($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
    Local $aResult = _FileListToArrayRec(@DesktopDir, _
        "1*;2*;3*;4*;7*;8*;9*;bac*2*;dc*;ds*", _
        $FLTAR_FOLDERS + $FLTAR_NOHIDDEN + $FLTAR_NOSYSTEM + $FLTAR_NOLINK, _
        $FLTAR_NORECUR, $FLTAR_FASTSORT, $FullPath ? $FLTAR_FULLPATH : $FLTAR_RELPATH)

    If Not IsArray($aResult) Then Return _EmptyArray()
    Return $aResult
EndFunc

; ============================================================================
; Initialise le cache des installations WAMP/XAMPP
Func DossiersEasyPHPwww($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
    Local $aEasyPHP

    ; Récupère depuis le cache ou effectue le scan
    If IsArray($__g_aEasyPHPRootsCache) Then
        $aEasyPHP = $__g_aEasyPHPRootsCache
    Else
        $aEasyPHP = _FileListToArrayRec( _
            StringLeft(@WindowsDir, 2), _
            "wamp*;xampp*", _
            $FLTAR_FOLDERS + $FLTAR_NOHIDDEN + $FLTAR_NOSYSTEM + $FLTAR_NOLINK, _
            $FLTAR_NORECUR, _
            $FLTAR_FASTSORT, _
            $FullPath ? $FLTAR_FULLPATH : $FLTAR_RELPATH)

        ; Stocke dans le cache même si vide
        If Not IsArray($aEasyPHP) Then
            $__g_aEasyPHPRootsCache = _EmptyArray()
            Return _EmptyArray()
        EndIf
        $__g_aEasyPHPRootsCache = $aEasyPHP
    EndIf

    If $aEasyPHP[0] = 0 Then Return _EmptyArray()

    ; Construction du tableau résultat
    Local $aResult[$aEasyPHP[0] + 1]
    Local $iCount = 0

    For $i = 1 To $aEasyPHP[0]
        Local $sFolder = $aEasyPHP[$i]

        If FileExists($sFolder & '\www') Then
            $iCount += 1
            $aResult[$iCount] = $sFolder & '\www'
        ElseIf FileExists($sFolder & '\htdocs') Then
            $iCount += 1
            $aResult[$iCount] = $sFolder & '\htdocs'
        EndIf
    Next

    If $iCount = 0 Then Return _EmptyArray()

    ReDim $aResult[$iCount + 1]
    $aResult[0] = $iCount
    Return $aResult
EndFunc

; ============================================================================
; Utilise le cache et l'invalide après usage
Func DossiersEasyPHPdata($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
    Local $aEasyPHP

    ; Récupère le cache préparé par DossiersEasyPHPwww
    If IsArray($__g_aEasyPHPRootsCache) Then
        $aEasyPHP = $__g_aEasyPHPRootsCache
        ; $__g_aEasyPHPRootsCache = 0 ; Invalidation immédiate
    Else
        ; Fallback si appelé sans DossiersEasyPHPwww
        $aEasyPHP = _FileListToArrayRec( _
            StringLeft(@WindowsDir, 2), _
            "wamp*;xampp*", _
            $FLTAR_FOLDERS + $FLTAR_NOHIDDEN + $FLTAR_NOSYSTEM + $FLTAR_NOLINK, _
            $FLTAR_NORECUR, _
            $FLTAR_FASTSORT, _
            $FullPath ? $FLTAR_FULLPATH : $FLTAR_RELPATH)

        If Not IsArray($aEasyPHP) Then Return _EmptyArray()
    EndIf

    If $aEasyPHP[0] = 0 Then Return _EmptyArray()

    ; Pré-allocation pour MySQL + MariaDB par installation
    Local $aResult[$aEasyPHP[0] * 2 + 1]
    Local $iCount = 0

    For $i = 1 To $aEasyPHP[0]
        Local $sBase = $aEasyPHP[$i]
        Local $sDataPath

        ; XAMPP Lite
        $sDataPath = $sBase & '\apps\mysql\data'
        If FileExists($sDataPath) Then
            $iCount += 1
            $aResult[$iCount] = $sDataPath
            ContinueLoop
        EndIf

        ; EasyPHP/XAMPP standard
        $sDataPath = $sBase & '\mysql\data'
        If FileExists($sDataPath) Then
            $iCount += 1
            $aResult[$iCount] = $sDataPath
            ContinueLoop
        EndIf

        ; WampServer MySQL
        If FileExists($sBase & '\bin\mysql') Then
            $sDataPath = _FindDataFldr($sBase & '\bin\mysql')
            If $sDataPath <> "" Then
                $iCount += 1
                $aResult[$iCount] = $sDataPath
            EndIf
        EndIf

        ; WampServer MariaDB
        If FileExists($sBase & '\bin\mariadb') Then
            $sDataPath = _FindDataFldr($sBase & '\bin\mariadb')
            If $sDataPath <> "" Then
                $iCount += 1
                $aResult[$iCount] = $sDataPath
            EndIf
        EndIf
    Next

    If $iCount = 0 Then Return _EmptyArray()

    ReDim $aResult[$iCount + 1]
    $aResult[0] = $iCount
    Return $aResult
EndFunc

; ============================================================================
; Recherche le dossier data dans les installations WampServer versionnées
; Exemple: C:\wamp64\bin\mysql\mysql8.0.34\data
Func _FindDataFldr($PathEasy)
    Local $hSearch = FileFindFirstFile($PathEasy & "\*")
    If $hSearch = -1 Then Return ""

    Local $sEntry, $sCandidate
    While 1
        $sEntry = FileFindNextFile($hSearch)
        If @error Then ExitLoop

        ; Vérifie les dossiers versionnés (mysql*, mariadb*)
        If StringRegExp($sEntry, "(?i)^(mysql|mariadb)", 0) Then
            $sCandidate = $PathEasy & "\" & $sEntry & "\data"
            If FileExists($sCandidate) Then
                FileClose($hSearch)
                Return $sCandidate
            EndIf
        EndIf
    WEnd

    FileClose($hSearch)
    Return ""
EndFunc

; ============================================================================
Func _EmptyArray()
    Local $aEmpty[1] = [0]
    Return $aEmpty
EndFunc
;#########################################################################################
Func LecteurSauvegarde()
    Local $aDrives = DriveGetDrive('FIXED')
    Local $sHomeDrive = StringLeft(@WindowsDir, 2) ; dans certains cas @HomeDrive retourne une chaîne vide

    ; Si aucun lecteur fixe détecté, retourne le lecteur système
    If Not IsArray($aDrives) Then Return StringUpper($sHomeDrive) & "\"

    ; Vérifie d'abord si le lecteur système a assez d'espace
    If DriveSpaceFree($sHomeDrive & "\") > $MINIMUM_WINDOWS_FREE_SPACE Then
        Return StringUpper($sHomeDrive) & "\"
    EndIf

    ; Cherche un autre lecteur avec plus d'espace
    For $i = 1 To $aDrives[0]
        Local $sDrive = $aDrives[$i]

        ; Ignore le lecteur système (déjà testé)
        If $sDrive = $sHomeDrive Then ContinueLoop

        ; Vérifie les critères : non-USB, accessible en écriture, espace suffisant
        If DriveGetType($sDrive, $DT_BUSTYPE) <> "USB" _
            And _WinAPI_IsWritable($sDrive) _
            And DriveSpaceFree($sDrive & "\") > $FREE_SPACE_DRIVE_BACKUP Then
            Return StringUpper($sDrive) & "\"
        EndIf
    Next

    ; Aucun lecteur valide trouvé, retourne le lecteur système par défaut
    Return StringUpper($sHomeDrive) & "\"
EndFunc
;#########################################################################################
Func _KillOtherScript()
	Local $list = ProcessList()
	For $i = 1 To $list[0][0]
		If $list[$i][0] = @ScriptName Then
			If $list[$i][1] <> @AutoItPID Then
				; Kill process
				$r = ProcessClose($list[$i][1])
			EndIf
		EndIf
	Next
EndFunc   ;==>_KillOtherScript
;;************************************************
;MsgBox($MB_SYSTEMMODAL, 'dossier', $www, 360)
;;************************************************
;_ArrayDisplay($Liste, 'BacBackup 1.0.0',"",32,Default ,"Liste de Dossiers/Fichiers Surveillés")

Func AgeDuFichierEnMinutesModification($cFichier)
	Local $cFileDate = FileGetTime($cFichier, 0) ;Get the Last Modified date and time
	If @error Then Return 9999

	Local $iMinutes = _DateDiff("n", $cFileDate[0] & "/" & $cFileDate[1] & "/" & $cFileDate[2] & " " & $cFileDate[3] & ":" & $cFileDate[4] & ":" & $cFileDate[5], _NowCalc())
	If $iMinutes < 0 Then $iMinutes = 9999
	Return $iMinutes
EndFunc   ;==>AgeDuFichierEnMinutesModification

Func AgeDuFichierEnMinutesCreation($cFichier)
	Local $cFileDate = FileGetTime($cFichier, 1) ;Get Creation date and time
	If @error Then Return 9999

	Local $iMinutes = _DateDiff("n", $cFileDate[0] & "/" & $cFileDate[1] & "/" & $cFileDate[2] & " " & $cFileDate[3] & ":" & $cFileDate[4] & ":" & $cFileDate[5], _NowCalc())
	If $iMinutes < 0 Then $iMinutes = 9999
	Return $iMinutes
EndFunc   ;==>AgeDuFichierEnMinutesCreation

;#########################################################################################

; ========== FONCTIONS UTILITAIRES ==========
Func _IsLocked($Dossier)
    ; Vérifie si verrouillé en testant la suppression du marqueur
    Local $sMarker = $Dossier & "\.locked"

    ; Si pas de marqueur, pas verrouillé
    If Not FileExists($sMarker) Then Return False

    ; Tenter de supprimer le marqueur
    ; Si suppression réussit = pas vraiment verrouillé
    ; Si suppression échoue = vraiment verrouillé

    If FileDelete($sMarker) Then
        ; Suppression réussie = le dossier n'est PAS verrouillé
        ; Le marqueur était obsolète, il est maintenant supprimé
        Return False
    Else
        ; Suppression échouée = le dossier EST verrouillé
        Return True
    EndIf
EndFunc

Func _CreateLockMarker($Dossier)
    ; Créer le marqueur AVANT le verrouillage
    Local $sMarker = $Dossier & "\.locked"
    Local $sInfo = "Verrouillé le : " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & @CRLF
    $sInfo &= "Utilisateur : " & @UserName & @CRLF
    $sInfo &= "Ordinateur : " & @ComputerName & @CRLF
    $sInfo &= "Processus : " & @ScriptName

    FileWrite($sMarker, $sInfo)
    FileSetAttrib($sMarker, "+SH")
EndFunc

Func _RemoveLockMarker($Dossier)
    ; Supprimer le marqueur (après déverrouillage)
    Local $sMarker = $Dossier & "\.locked"
    If FileExists($sMarker) Then
        FileSetAttrib($sMarker, "-SH")
        FileDelete($sMarker)
    EndIf
EndFunc

; ========== VERROUILLAGE ==========
Func _LockRootFolder($Dossier)
    ; Verrouille l'accès au dossier racine uniquement
    If Not FileExists($Dossier) Then Return SetError(1, 0, -1)

    ; Vérifier si déjà verrouillé (teste la suppression du marqueur)
    If _IsLocked($Dossier) Then
        Return 0
    EndIf

    ; ÉTAPE 1 : Créer le marqueur AVANT le verrouillage
    _CreateLockMarker($Dossier)

    ; ÉTAPE 2 : Appliquer les attributs
    FileSetAttrib($Dossier, "+SH")

    ; ÉTAPE 3 : Verrouiller avec icacls
    Local $sIcacls = @SystemDir & "\icacls.exe"
    Local $sArgs = '"' & $Dossier & '" /deny *S-1-1-0:(F) /c'
    Local $iPID = Run('"' & $sIcacls & '" ' & $sArgs, "", @SW_HIDE)
    If @error Then
        _RemoveLockMarker($Dossier)
        Return SetError(2, 0, -1)
    EndIf

    ProcessWaitClose($iPID)

    Return 0
EndFunc

Func _LockFolderContents($Dossier)
    ; Verrouille contre la suppression (avec héritage)
    If Not FileExists($Dossier) Then Return SetError(1, 0, -1)

    ; Vérifier si déjà verrouillé
    If _IsLocked($Dossier) Then
        Return 0
    EndIf

    ; ÉTAPE 1 : Créer le marqueur AVANT le verrouillage
    _CreateLockMarker($Dossier)

    ; ÉTAPE 2 : Verrouiller avec icacls
    Local $sIcacls = @SystemDir & "\icacls.exe"
    Local $sArgs = '"' & $Dossier & '" /grant *S-1-1-0:(OI)(CI)M /deny *S-1-1-0:(OI)(CI)(DE,DC) /c'
    Local $iPID = Run('"' & $sIcacls & '" ' & $sArgs, "", @SW_HIDE)
    If @error Then
        _RemoveLockMarker($Dossier)
        Return SetError(2, 0, -1)
    EndIf

    ProcessWaitClose($iPID)

    Return 0
EndFunc

; ========== DÉVERROUILLAGE ==========
Func _UnlockFolder($Dossier, $bRecursive = False)
    ; Déverrouille un dossier (et optionnellement son contenu)
    If Not FileExists($Dossier) Then Return SetError(1, 0, -1)

    ; ÉTAPE 1 : Déverrouiller avec icacls
    Local $sIcacls = @SystemDir & "\icacls.exe"
    Local $sArgs = '"' & $Dossier & '" /reset ' & ($bRecursive ? '/T ' : '') & '/c'
    Local $iPID = Run('"' & $sIcacls & '" ' & $sArgs, "", @SW_HIDE)
    If @error Then Return SetError(2, 0, -1)

    ProcessWaitClose($iPID)

    ; ÉTAPE 2 : Retirer les attributs
    FileSetAttrib($Dossier, "-SH")

    ; ÉTAPE 3 : Supprimer le marqueur APRÈS le déverrouillage
    _RemoveLockMarker($Dossier)

    Return 0
EndFunc

;~ Find out current username when executed as admin
;~ https://www.autoitscript.com/forum/topic/183689-find-out-current-username-when-executed-as-admin/#comment-1319682
Func _GetUsername()
	Local $aResult = DllCall("Wtsapi32.dll", "int", "WTSQuerySessionInformationW", "int", 0, "dword", -1, "int", 5, "dword*", 0, "dword*", 0)
	If @error Or $aResult[0] = 0 Then Return SetError(1, 0, @UserName)
	Local $sUserName = BinaryToString(DllStructGetData(DllStructCreate("byte[" & $aResult[5] & "]", $aResult[4]), 1), 2)
	DllCall("Wtsapi32.dll", "int", "WTSFreeMemory", "ptr", $aResult[4])
	Return $sUserName
EndFunc   ;==>_GetUsername


;This will check if an app with a window title = $strtitle is hosted by wowexec
;Return 1 if exists, 0 if not exists
;Not all 16 bits apps seem to be hosted by wowexec and wowexec is itself hosted by ntvdm
;I have an old game running in a DOS console which is hosted by ntvdm without wowexec
Func WOWAPPEXISTS($strtitle, $winmatchmode)
	Local $arrwow, $arrapp, $i, $j

	AutoItSetOption("WinTitleMatchMode", 4)

	$arrwow = WinList("classname=WOWExecClass")
	If $arrwow[0][0] = 0 Then Return 0

	AutoItSetOption("WinTitleMatchMode", $winmatchmode)

	$arrapp = WinList($strtitle)
	If $arrapp[0][0] = 0 Then Return 0

	For $i = 1 To UBound($arrapp, 1) - 1
		For $j = 1 To UBound($arrwow, 1) - 1
			If WinGetProcess($arrapp[$i][1]) = WinGetProcess($arrwow[$j][1]) Then
				Return 1
			EndIf
		Next
	Next

	Return 0
EndFunc   ;==>WOWAPPEXISTS

Func _MyProcessByPartName($str)
	Local $alist = ProcessList(), $ret
	For $1 = 0 To UBound($alist) - 1
		If StringInStr($alist[$1][0], $str) Then Return $alist[$1][1]
	Next
	Return 0
EndFunc   ;==>_MyProcessByPartName

Func _SystemInfo()
	Local $objWMIService = ObjGet("winmgmts:\\localhost\root\CIMV2")
	Local $Output = ""
	;Get info from WMIC
	If IsObj($objWMIService) Then
		$colCompSysPro = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", "WQL", 0x10 + 0x20)

		;If variable is acceptable
		If IsObj($colCompSysPro) Then
			For $objCompSysPro In $colCompSysPro
				;The Hardware Info
				;Manufacturer/Marque/Vendor
				Local $TmpOrdinateur = $objCompSysPro.Vendor
				$Output &= "Ordinateur  : " & $TmpOrdinateur & @CRLF

				;Model
				If $objCompSysPro.Vendor = "Lenovo" Then
					$objModel = $objCompSysPro.Version
				Else
					$objModel = $objCompSysPro.Name
				EndIf
				$Output &= "Modèle      : " & $objModel & @CRLF
			Next
		EndIf
	EndIf
	;CPU
	$CPUName = RegRead("HKLM64\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString")
	$Output &= "Processeur  : " & $CPUName & @CRLF

	;The RAM
	$RAMStats = MemGetStats()
	$RAMTotal = $RAMStats[1]
	$Output &= "Mémoire     : " & _RAMSuffix($RAMTotal) & @CRLF

	;OS
	$Output &= "Système     : " & "Windows " & StringTrimLeft(@OSVersion, 4) & " " & @OSServicePack & " " & (@OSArch = "X86" ? "32-bit" : "64-bit") & @CRLF
	$Output &= "Computer ID : " & _GetUUID()

	Return $Output


EndFunc   ;==>_SystemInfo

Func _RAMSuffix($Bytes)
	Local $x, $BytesSuffix[6] = ["KB", "MB", "GB", "TB", "PB"]
	While $Bytes > 1000
		$x += 1
		$Bytes /= 1024
	WEnd
	$Bytes = Ceiling(StringFormat('%.2f', $Bytes))
	$Bytes = StringFormat('%.2f', $Bytes)
	Return $Bytes & " " & $BytesSuffix[$x]
EndFunc   ;==>_RAMSuffix

Func _MonthFullName($MonthNum)
	Local $aMonths = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
	If $MonthNum >= 1 And $MonthNum <= 12 Then
		Return $aMonths[$MonthNum - 1]
	EndIf
	Return ""
EndFunc   ;==>_MonthFullName

Func JourDeLaSemaine()
	Local $aDays = ["Dimanche", "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi"]
	Return $aDays[@WDAY - 1]
EndFunc   ;==>JourDeLaSemaine

Func _GetUUID($annee = 2025)
	$Uuid = RegRead("HKCU\SOFTWARE\BacBackup", "UUID")
	If @error = 0 And StringRegExp($Uuid, "^((\d{4}-)(\w{4} ){2}\w{4})$") = 1 And StringLeft($Uuid, 4) >= 2025 Then Return $Uuid
	$Uuid = $annee & "-" & StringFormat('%04X %04X %04X', _
			Random(0, 0xffff), _
			BitOR(Random(0, 0x0fff), 0x4000), _
			BitOR(Random(0, 0x3fff), 0x8000) _
			)
	RegWrite("HKCU\SOFTWARE\BacBackup", "UUID", "REG_SZ", $Uuid)
	Return $Uuid
EndFunc   ;==>_GetUUID

Func _IsRegistryExist($sKeyName, $sValueName)
	$x = RegRead($sKeyName, $sValueName)
	Return @error = 0
EndFunc   ;==>_IsRegistryExist

Func _WinGetByPID($iPID, $iArray = 1) ; 0 Will Return 1 Base Array & 1 Will Return The First Window.
	Local $aError[1] = [0], $aWinList, $sReturn
	If IsString($iPID) Then
		$iPID = ProcessExists($iPID)
	EndIf
	$aWinList = WinList()
	For $A = 1 To $aWinList[0][0]
		If WinGetProcess($aWinList[$A][1]) = $iPID And BitAND(WinGetState($aWinList[$A][1]), 2) Then
			If $iArray Then
				Return $aWinList[$A][1]
			EndIf
			$sReturn &= $aWinList[$A][1] & Chr(1)
		EndIf
	Next
	If $sReturn Then
		Return StringSplit(StringTrimRight($sReturn, 1), Chr(1))
	EndIf
	Return SetError(1, 0, $aError)
EndFunc   ;==>_WinGetByPID

; #FUNCTION# ====================================================================================================================
; Name ..........: _DirCopyWithProgress
; Description ...: Copy one directory and its structure to another directory with a progress bar
; Syntax ........: _DirCopyWithProgress($sSource, $sDest[, [$iOverwrite = 1[, $sMsg = "Copying in progress..."]]])
; Parameters ....: $sSource      - Source path of the folder to copy
;                  $sDest        - Destination path
;                  $iOverwrite   - [optional] 1 to overwrite existing files (default), 0 to not overwrite
;                  $sMsg         - [optional] Message to show in the progress dialog title
; Return values .: Success - Returns 1
;                  Failure - Returns 0 and sets @error to:
;                  |1 - The source folder does not exist
;                  |2 - Unable to create the destination folder
;                  |3 - No files found or unable to list files
;                  |4 - Source and destination are the same
; Author ........: romoez@github
; Modified ......:
; Remarks .......: Does not delete files in the destination that are not in the source, unlike DirCopy with overwrite.
; Related .......: DirCopy, FileCopy
; Link ..........: https://github.com/romoez/BacCollector
; ===============================================================================================================================
Func _DirCopyWithProgress($sourceDir, $destDir, $overwrite = 1, $sMsg = "Copie en cours...")
	If Not FileExists($sourceDir) Then Return SetError(1, 0, 0)
	If $sourceDir = $destDir Then Return SetError(4, 0, 0)

	; Créer le dossier de destination s'il n'existe pas
	If Not FileExists($destDir) Then
		If Not DirCreate($destDir) Then Return SetError(2, 0, 0)
	EndIf

	; Récupérer la liste complète des fichiers dans le dossier source (récursif)
	Local $aFiles = _FileListToArrayRec($sourceDir, "*", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
	If @error Or Not IsArray($aFiles) Then Return SetError(3, 0, 0)

	; Initialiser la barre de progression
	ProgressOn($sMsg, "Initialisation...", "0%", Default, Default, 1)

	; Compter le nombre total de fichiers à copier
	Local $totalFiles = UBound($aFiles) - 1
	If $totalFiles = 0 Then
		ProgressOff()
		Return SetError(3, 0, 0) ; No files to copy
	EndIf
	Local $currentFile = 0

	; Boucle pour copier chaque fichier
	For $i = 1 To $totalFiles
		Local $sourcePath = $aFiles[$i]

		; Construire le chemin de destination correspondant
		Local $relativePath = StringReplace($sourcePath, $sourceDir, "", 1)
		Local $destPath = $destDir & $relativePath

		; Créer le dossier de destination s'il n'existe pas
		Local $destFolder = StringRegExpReplace($destPath, "(^.*\\).*", "$1")
		If Not FileExists($destFolder) Then
			DirCreate($destFolder)
		EndIf

		; Copier le fichier
		FileCopy($sourcePath, $destPath, $overwrite ? $FC_OVERWRITE + $FC_CREATEPATH : $FC_CREATEPATH)

		; Mettre à jour la barre de progression
		$currentFile += 1
		Local $percent = Round(($currentFile / $totalFiles) * 100, 0)
		ProgressSet($percent, "[" & $currentFile & "/" & $totalFiles & "] " & StringRegExpReplace($sourcePath, "^.*\\", ""), $sMsg)
	Next

	; Fermer la barre de progression
	ProgressOff()

	Return 1
EndFunc   ;==>_DirCopyWithProgress


; #FUNCTION# ====================================================================================================================
; Nom............: _GetProcessPath
; Description ...: Récupère le chemin d'un processus en cours d'exécution ou recherche dans des emplacements alternatifs.
; Syntax.........: _GetProcessPath($sProcessName, $sAppId, $sAppPath)
; Paramètres ....: $sProcessName - Nom du processus (ex. "BacBackup.exe")
;                  $sAppId - ID unique de l'application pour la recherche dans la base de registre
;                  $sAppPath - Chemin par défaut où l'exécutable est attendu
; Retour ........: Le chemin complet du processus si trouvé, sinon une chaîne vide avec @error défini.
;                  @error = 1 si le chemin n'est pas trouvé
; Auteur ........: romoez@github
; Lien ..........: https://github.com/romoez/BacCollector
; ===============================================================================================================================

Func _GetProcessPath($sProcessName, $sAppPath = "", $sAppId = "")
	Local $iPID = ProcessExists($sProcessName)
	If $iPID Then
		Local $sPath = _WinAPI_GetProcessFileName($iPID)
		If Not @error Then
			Return $sPath
		EndIf
	EndIf

	; Vérifier le chemin par défaut fourni par $sAppPath
	If FileExists($sAppPath) Then
		Return $sAppPath
	EndIf

	; Chercher dans la base de registre avec $sAppId
	If $sAppId <> "" Then
		Local $sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & $sAppId
		Local $sInstallLocation = RegRead($sRegKey, "InstallLocation")
		If Not @error And $sInstallLocation <> "" Then
			Local $sFullPath = $sInstallLocation & "\" & $sProcessName
			If FileExists($sFullPath) Then
				Return $sFullPath
			EndIf
		EndIf
	EndIf

	; Si aucune méthode ne fonctionne, retourner une erreur
	Return SetError(1, 0, "")
EndFunc   ;==>_GetProcessPath

; #FUNCTION# ====================================================================================================================
; Nom...........: _RunOrShellExecute
; Description ...: Tente de lancer un programme avec Run, et si cela échoue, avec ShellExecute.
; Syntax.........: _RunOrShellExecute($sPath)
; Paramètres ....: $sPath - Chemin complet du programme à lancer
; Retour ........: 1 si le programme est lancé avec succès, sinon 0 avec @error défini.
;                  @error = 2 si Run et ShellExecute échouent
; Auteur ........: romoez@github
; Lien ..........: https://github.com/romoez/BacCollector
; ===============================================================================================================================
Func _RunOrShellExecute($sPath)
	Local $iPID = Run($sPath)
	If $iPID <> 0 Then
		Return 1
	EndIf

	; Si Run échoue, tenter ShellExecute
	ShellExecute($sPath)
	If @error Then
		Return SetError(2, 0, 0)
	Else
		Return 1
	EndIf
EndFunc   ;==>_RunOrShellExecute


; #FUNCTION# ====================================================================================================================
; Nom...........: _ExtractFolderPath
; Description ...: Extrait le chemin du dossier à partir du chemin complet d'un fichier.
;                  Retourne le chemin avec un backslash à la fin.
; Syntax.........: _ExtractFolderPath($sFilePath)
; Paramètres ....: $sFilePath - Chemin complet du fichier (ex. "C:\Program Files (x86)\BacBackup\BacBackup.exe")
; Retour ........: Le chemin du dossier avec un backslash à la fin (ex. "C:\Program Files (x86)\BacBackup\")
;                  Retourne une chaîne vide si aucune barre oblique n'est trouvée dans le chemin.
; Auteur ........: romoez@github
; Lien ..........: https://github.com/romoez/BacCollector
; Exemple .......: Local $sFilePath = "C:\Program Files (x86)\BacBackup\BacBackup.exe"
;                  Local $sFolderPath = _ExtractFolderPath($sFilePath)
;                  ConsoleWrite($sFolderPath & @CRLF)
;                  ; Résultat attendu : "C:\Program Files (x86)\BacBackup\"
; ===============================================================================================================================
Func _ExtractFolderPath($sFilePath)
	Local $iPos = StringInStr($sFilePath, "\", 0, -1) ; Trouve la position de la dernière "\"
	If $iPos > 0 Then
		Return StringLeft($sFilePath, $iPos) ; Extrait le chemin jusqu'à cette position
	Else
		Return "" ; Retourne vide si aucune barre oblique n'est trouvée
	EndIf
EndFunc   ;==>_ExtractFolderPath



Func _CreerDossierNouvelleSession($Lecteur, $DossierSauvegardes)
	; Création du dossier principal
	If Not FileExists($Lecteur & $DossierSauvegardes) Then
		DirCreate($Lecteur & $DossierSauvegardes)
	EndIf
	FileSetAttrib($Lecteur & $DossierSauvegardes, "+SH")

	; Création du sous-dossier BacBackup
	Local $sBacBackup = $Lecteur & $DossierSauvegardes & "\BacBackup"
	If Not FileExists($sBacBackup) Then
		_UnlockFolder($Lecteur & $DossierSauvegardes)
		DirCreate($sBacBackup)
	EndIf

	; Écriture des paramètres dans l'INI
	Local $sIniPath = $sBacBackup & "\BacBackup.ini"
	IniWrite($sIniPath, "Params", "DossierSauvegardes", $DossierSauvegardes)
	IniWrite($sIniPath, "Params", "Lecteur", StringUpper($Lecteur))
	_LockRootFolder($Lecteur & $DossierSauvegardes)

	; Gestion du dossier de session
	Local $DossierSession = IniRead($sIniPath, "Params", "DossierSession", "")
	IniWrite($sIniPath, "Params", "DossierSession", $DossierSession & "_BacCollector")
	Local $Tmp = StringLeft($DossierSession, 3)
	Return 1    ; <<<<<<<<<<<<<<<<<<<<<< RETURN
	; ****************************************************************
	; ****************************************************************
	; ****************************************************************
	; ****************************************************************
	; Logique d'incrémentation du numéro de session
	If StringRegExp($Tmp, "([0-9]{3})", 0) = 0 Then
		$Tmp = "001"
	Else
		$Tmp = Number($Tmp) + 1
		$Tmp = $Tmp > 999 ? "001" : StringFormat("%03d", $Tmp)
	EndIf

	; Création du nom de dossier complet
	$DossierSession = $Tmp & '___' & @MDAY & "_" & @MON & "_" & @YEAR & "___" & @HOUR & "h" & @MIN
	IniWrite($sIniPath, "Params", "DossierSession", $DossierSession)

	; Création du chemin complet et du dossier
	Local $sCheminComplet = StringUpper($Lecteur) & $DossierSauvegardes & "\BacBackup\" & $DossierSession
	If Not FileExists($sCheminComplet) Then
		DirCreate($sCheminComplet)
	EndIf

	Return $sCheminComplet
EndFunc   ;==>_CreerDossierNouvelleSession


Func _Directory_Is_Accessible($sPath)
    If Not StringInStr(FileGetAttrib($sPath), "D", 2) Then Return SetError(1, 0, 0)
    Local $iEnum = 0
    While FileExists($sPath & "\_bc_test_" & $iEnum)
		If DirGetSize($sPath & "\_bc_test_" & $iEnum) = 0 Then
			DirRemove($sPath & "\_bc_test_" & $iEnum)
		EndIf
        $iEnum += 1
    WEnd
    Local $iSuccess = DirCreate($sPath & "\_bc_test_" & $iEnum)
    Switch $iSuccess
        Case 1
            Return DirRemove($sPath & "\_bc_test_" & $iEnum)
        Case Else
            Return False
    EndSwitch
EndFunc   ;==>_Directory_Is_Assesible
