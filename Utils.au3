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
; Détermine si un dossier de base de données MySQL/MariaDB est "vide"
; (c.-à-d. ne contient aucune table utilisateur).
; Les fichiers techniques (db.opt, et fichiers cachés/système) sont ignorés.
; Renvoie True si vide, False sinon.
; ============================================================================
Func _IsMySqlDbEmpty($sDbFolder)
    If Not FileExists($sDbFolder) Then Return True

    ; Liste des fichiers "techniques" à ignorer (créés automatiquement par le SGBD)
    Local Const $aIgnoredFiles[] = ["db.opt"]

    Local $hSearch = FileFindFirstFile($sDbFolder & "\*")
    If $hSearch = -1 Then Return True

    Local $sFileName, $bIsTechnical
    While 1
        $sFileName = FileFindNextFile($hSearch)
        If @error Then ExitLoop
        If $sFileName = "." Or $sFileName = ".." Then ContinueLoop

        ; Ignorer les fichiers techniques connus (insensible à la casse)
        $bIsTechnical = False
        For $sIgnored In $aIgnoredFiles
            If StringLower($sFileName) = StringLower($sIgnored) Then
                $bIsTechnical = True
                ExitLoop
            EndIf
        Next
        If $bIsTechnical Then ContinueLoop

        ; Trouvé un fichier qui n'est pas technique => la base n'est pas vide
        FileClose($hSearch)
        Return False
    WEnd

    FileClose($hSearch)
    Return True
EndFunc   ;==>_IsMySqlDbEmpty


; #FUNCTION# ;===============================================================================
;
; Name...........: _MD5ForFile
; Description ...: Calculates MD5 value for the specific file.
; Syntax.........: _MD5ForFile ($sFile)
; Parameters ....: $sFile - Full path to the file to process.
; Return values .: Success - Returns MD5 value in form of hex string
;                          - Sets @error to 0
;                  Failure - Returns empty string and sets @error:
;                  |1 - CreateFile function or call to it failed.
;                  |2 - CreateFileMapping function or call to it failed.
;                  |3 - MapViewOfFile function or call to it failed.
;                  |4 - MD5Init function or call to it failed.
;                  |5 - MD5Update function or call to it failed.
;                  |6 - MD5Final function or call to it failed.
; Author ........: trancexx
; Link ..........: https://www.autoitscript.com/forum/topic/95558-crc32-md4-md5-sha1-for-files/
;==========================================================================================
Func _MD5ForFile($sFile)

	Local $a_hCall = DllCall("kernel32.dll", "hwnd", "CreateFileW", _
			"wstr", $sFile, _
			"dword", 0x80000000, _ ; GENERIC_READ
			"dword", 3, _ ; FILE_SHARE_READ|FILE_SHARE_WRITE
			"ptr", 0, _
			"dword", 3, _ ; OPEN_EXISTING
			"dword", 0, _ ; SECURITY_ANONYMOUS
			"ptr", 0)

	If @error Or $a_hCall[0] = -1 Then
		Return SetError(1, 0, "")
	EndIf

	Local $hFile = $a_hCall[0]

	$a_hCall = DllCall("kernel32.dll", "ptr", "CreateFileMappingW", _
			"hwnd", $hFile, _
			"dword", 0, _ ; default security descriptor
			"dword", 2, _ ; PAGE_READONLY
			"dword", 0, _
			"dword", 0, _
			"ptr", 0)

	If @error Or Not $a_hCall[0] Then
		DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFile)
		Return SetError(2, 0, "")
	EndIf

	DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFile)

	Local $hFileMappingObject = $a_hCall[0]

	$a_hCall = DllCall("kernel32.dll", "ptr", "MapViewOfFile", _
			"hwnd", $hFileMappingObject, _
			"dword", 4, _ ; FILE_MAP_READ
			"dword", 0, _
			"dword", 0, _
			"dword", 0)

	If @error Or Not $a_hCall[0] Then
		DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
		Return SetError(3, 0, "")
	EndIf

	Local $pFile = $a_hCall[0]
	Local $iBufferSize = FileGetSize($sFile)

	Local $tMD5_CTX = DllStructCreate("dword i[2];" & _
			"dword buf[4];" & _
			"ubyte in[64];" & _
			"ubyte digest[16]")

	DllCall("advapi32.dll", "none", "MD5Init", "ptr", DllStructGetPtr($tMD5_CTX))

	If @error Then
		DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
		DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
		Return SetError(4, 0, "")
	EndIf

	DllCall("advapi32.dll", "none", "MD5Update", _
			"ptr", DllStructGetPtr($tMD5_CTX), _
			"ptr", $pFile, _
			"dword", $iBufferSize)

	If @error Then
		DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
		DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
		Return SetError(5, 0, "")
	EndIf

	DllCall("advapi32.dll", "none", "MD5Final", "ptr", DllStructGetPtr($tMD5_CTX))

	If @error Then
		DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
		DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)
		Return SetError(6, 0, "")
	EndIf

	DllCall("kernel32.dll", "int", "UnmapViewOfFile", "ptr", $pFile)
	DllCall("kernel32.dll", "int", "CloseHandle", "hwnd", $hFileMappingObject)

	Local $sMD5 = Hex(DllStructGetData($tMD5_CTX, "digest"))

	Return SetError(0, 0, $sMD5)

EndFunc   ;==>_MD5ForFile

; ============================================================================
; Vérifie qu'un dossier copié est intègre par rapport à sa source.
; - Compare nombre de fichiers et taille totale (rapide)
; - Compare le MD5 de CHAQUE fichier (volumes faibles : <1s pour 10 Mo)
; Retour : 1 si intégrité parfaite, 0 sinon (avec log détaillé du 1er écart)
; ============================================================================
Func _VerifierIntegriteCopieDossier($sSrc, $sDest)
    If Not FileExists($sSrc) Or Not FileExists($sDest) Then
        _Logging("Vérif intégrité : src ou dest manquante", 5, 0)
        Return 0
    EndIf

    ; --- Comparaison rapide : nb fichiers et taille ---
    ; _DirGetSizeSafe : retourne toujours [0,0,0] si erreur → pas de crash
    Local $aSrcInfo = _DirGetSizeSafe($sSrc)
    Local $aDstInfo = _DirGetSizeSafe($sDest)
    If $aSrcInfo[1] <> $aDstInfo[1] Then
        _Logging("Intégrité ÉCHEC : nb fichiers src=" & $aSrcInfo[1] & " dst=" & $aDstInfo[1], 5, 0)
        Return 0
    EndIf
    If $aSrcInfo[0] <> $aDstInfo[0] Then
        _Logging("Intégrité ÉCHEC : taille src=" & $aSrcInfo[0] & " dst=" & $aDstInfo[0], 5, 0)
        Return 0
    EndIf

    ; --- Comparaison MD5 fichier par fichier ---
    Local $aFiles = _FileListToArrayRec($sSrc, "*", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_RELPATH)
    If Not IsArray($aFiles) Then Return 1 ; aucun fichier (dossier vide cohérent)

    For $i = 1 To $aFiles[0]
        Local $sSrcFile = $sSrc & "\" & $aFiles[$i]
        Local $sDstFile = $sDest & "\" & $aFiles[$i]

        ; Cas fichier vide (taille=0) : _MD5ForFile échoue sur CreateFileMappingW
        ; (comportement documenté Win32 : mapping mémoire interdit sur fichier vide)
        ; → vérifier juste que le fichier existe à destination avec taille=0
        If FileGetSize($sSrcFile) = 0 Then
            If Not FileExists($sDstFile) Or FileGetSize($sDstFile) <> 0 Then
                _Logging("Intégrité ÉCHEC (fichier vide manquant ou taille≠0 à dst) : " & $aFiles[$i], 5, 0)
                Return 0
            EndIf
            ContinueLoop ; fichier vide OK côté src et dst
        EndIf

        ; Fichier non vide : comparaison MD5
        Local $sHashSrc = _MD5ForFile($sSrcFile)
        Local $sHashDst = _MD5ForFile($sDstFile)
        If @error Then
            ; _MD5ForFile a échoué (fichier verrouillé, accès refusé...)
            ; → fallback sur comparaison de taille uniquement
            Local $iSzSrc = FileGetSize($sSrcFile)
            Local $iSzDst = FileGetSize($sDstFile)
            If $iSzSrc <> $iSzDst Then
                _Logging("Intégrité ÉCHEC taille (src=" & $iSzSrc & " dst=" & $iSzDst & ") : " & $aFiles[$i], 5, 0)
                Return 0
            EndIf
            _Logging("Intégrité MD5 non calculable, fallback taille OK : " & $aFiles[$i], 2, 0)
            ContinueLoop
        EndIf
        If $sHashSrc <> $sHashDst Then
            _Logging("Intégrité ÉCHEC MD5 : " & $aFiles[$i], 5, 0)
            Return 0
        EndIf
    Next

    Return 1
EndFunc   ;==>_VerifierIntegriteCopieDossier

; ============================================================================
; Copie un dossier directement vers $sDest avec vérification d'intégrité
; MD5 et retry automatique (3 essais, pause 300ms).
; Retour : 1 si copie ET intégrité OK, 0 sinon.
; ============================================================================
Func _CopierDossierFiable($sSrc, $sDest, $iMaxTry = 3)
    For $iTry = 1 To $iMaxTry
        ; ===== SIMULATION D'ERREURS (mode test uniquement) =====
        Local $iCopy
        Local $bForcerEchecIntegrite = False
        If $TEST_ERREUR_COPIE > 0 Then
            Switch $TEST_ERREUR_COPIE
                Case 1 ; Échec total DirCopy (toutes destinations)
                    _Logging("[TEST] Simulation échec DirCopy : " & $sDest, 2, 0)
                    $iCopy = 0
                Case 2 ; DirCopy OK mais intégrité toujours KO
                    _Logging("[TEST] Simulation intégrité KO : " & $sDest, 2, 0)
                    $iCopy = DirCopy($sSrc, $sDest, $FC_OVERWRITE)
                    $bForcerEchecIntegrite = True
                Case 3 ; Échec USB seulement
                    If StringInStr($sDest, $Dest1FlashUSB) Then
                        _Logging("[TEST] Simulation échec USB uniquement : " & $sDest, 2, 0)
                        $iCopy = 0
                    Else
                        $iCopy = DirCopy($sSrc, $sDest, $FC_OVERWRITE)
                    EndIf
                Case 4 ; Échec local seulement
                    If StringInStr($sDest, $Dest2LocalFldr) Then
                        _Logging("[TEST] Simulation échec LOCAL uniquement : " & $sDest, 2, 0)
                        $iCopy = 0
                    Else
                        $iCopy = DirCopy($sSrc, $sDest, $FC_OVERWRITE)
                    EndIf
                Case 5 ; Échec au 1er essai, succès au 2ème (test retry)
                    $TEST_ERREUR_COPIE_COMPTEUR += 1
                    If $iTry = 1 And $TEST_ERREUR_COPIE_COMPTEUR <= 2 Then
                        _Logging("[TEST] Simulation échec essai 1 (retry attendu) : " & $sDest, 2, 0)
                        $iCopy = 0
                    Else
                        $iCopy = DirCopy($sSrc, $sDest, $FC_OVERWRITE)
                    EndIf
            EndSwitch
        Else
            ; Mode normal
            $iCopy = DirCopy($sSrc, $sDest, $FC_OVERWRITE)
        EndIf
        ; ===== FIN SIMULATION =====

        If $iCopy = 1 Then
            Local $bIntegrite = $bForcerEchecIntegrite ? 0 : _VerifierIntegriteCopieDossier($sSrc, $sDest)
            If $bIntegrite = 1 Then
                If $iTry > 1 Then _Logging("Copie réussie au " & $iTry & "ème essai : " & $sDest, 2, 0)
                Return 1
            EndIf
            _Logging("Copie OK mais intégrité ÉCHEC (essai " & $iTry & "/" & $iMaxTry & ") : " & $sDest, 5, 0)
        Else
            _Logging("Copie ÉCHEC (essai " & $iTry & "/" & $iMaxTry & ") : " & $sDest, 5, 0)
        EndIf
        If $iTry < $iMaxTry Then Sleep(300)
    Next
    Return 0
EndFunc   ;==>_CopierDossierFiable

; ============================================================================
; Wrapper sécurisé autour de DirGetSize($sPath, 1).
; Retourne TOUJOURS un tableau valide [taille, nbFichiers, nbDossiers].
; En cas d'erreur (dossier inexistant, accès refusé, résultat non-array),
; retourne [0, 0, 0] au lieu de planter avec
; "Subscript used on non-accessible variable".
; ============================================================================
Func _DirGetSizeSafe($sPath)
    Local $aDefault[3] = [0, 0, 0]
    If Not FileExists($sPath) Then Return $aDefault
    Local $aInfo = DirGetSize($sPath, 1)
    If @error Or Not IsArray($aInfo) Or UBound($aInfo) < 3 Then Return $aDefault
    Return $aInfo
EndFunc   ;==>_DirGetSizeSafe

; ============================================================================
; Signale les erreurs de copie non vérifiées :
;   1. Renomme le dossier candidat en NNNNNN_ (ajout d'un tiret bas)
;      sur chaque destination concernée — le dossier n'est plus détecté
;      dans la liste des récupérations normales, forçant une vérification manuelle.
;   2. Crée _ERREURS_COPIE.txt dans le dossier renommé avec la liste
;      des sources non supprimées et l'action requise.
; $aDest    : tableau ["USB_path_candidat", "Local_path_candidat"]
;             (chemins complets incluant déjà le numéro du candidat)
; $aErreurs : tableau [0]=nb, [1..n]= "SOURCE|RAISON"
; $sCandidat: numéro du candidat (pour l'en-tête du fichier)
; ============================================================================
; ============================================================================
; Crée ou met à jour _ERREURS_COPIE.txt dans les dossiers destination
; (USB et local) pour signaler les fichiers/dossiers dont la copie a échoué.
; Le renommage NNNNNN → NNNNNN_ est fait séparément par _FinaliserErreursCopie,
; appelé UNE SEULE FOIS à la fin, après que TOUTES les copies soient terminées.
; $aDest    : tableau [chemin_USB_candidat, chemin_Local_candidat]
; $aErreurs : tableau [0]=nb, [1..n]= "SOURCE|RAISON"
; $sCandidat: numéro du candidat (pour l'en-tête du fichier)
; ============================================================================
Func _EcrireFichierErreursCopie($aDest, $aErreurs, $sCandidat)
    If Not IsArray($aErreurs) Or $aErreurs[0] = 0 Then Return

    Local $sNomFichier = "\_ERREURS_COPIE.txt"
    Local $sContenu = $PROG_TITLE & $PROG_VERSION & " — Rapport d'erreurs de copie" & @CRLF
    $sContenu &= "Candidat : " & $sCandidat & @CRLF
    $sContenu &= "Date     : " & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & @CRLF
    $sContenu &= @CRLF
    $sContenu &= "ATTENTION — " & $aErreurs[0] & " source(s) NON supprimée(s) car la copie n'a pas pu être vérifiée :" & @CRLF
    $sContenu &= @CRLF

    For $i = 1 To $aErreurs[0]
        Local $aParts = StringSplit($aErreurs[$i], "|", 2)
        If UBound($aParts) >= 2 Then
            $sContenu &= "  * " & $aParts[0] & @CRLF
            $sContenu &= "    Raison : " & $aParts[1] & @CRLF & @CRLF
        Else
            $sContenu &= "  * " & $aErreurs[$i] & @CRLF & @CRLF
        EndIf
    Next

    $sContenu &= "ACTION REQUISE :" & @CRLF
    $sContenu &= "  Les fichiers/dossiers listés ci-dessus sont CONSERVES sur le poste candidat." & @CRLF
    $sContenu &= "  Copier manuellement le contenu manquant dans ce dossier." & @CRLF
    $sContenu &= "  Renommer ensuite ce dossier de """ & $sCandidat & "_"" vers """ & $sCandidat & """." & @CRLF

    ; Écrire dans chaque destination (USB et locale) — sans renommer
    ; FO_APPEND : cumule les erreurs si le fichier existe déjà (plusieurs appels)
    For $k = 0 To UBound($aDest) - 1
        Local $sChemin = $aDest[$k]
        If $sChemin = "" Then ContinueLoop
        If Not FileExists($sChemin) Then DirCreate($sChemin)
        Local $bFichierExiste = FileExists($sChemin & $sNomFichier)
        Local $hFile = FileOpen($sChemin & $sNomFichier, $FO_APPEND + $FO_UTF8)
        If $hFile <> -1 Then
            ; En-tête uniquement si nouveau fichier
            If Not $bFichierExiste Then
                FileWrite($hFile, $sContenu)
            Else
                ; Ajouter seulement les nouvelles lignes d'erreurs
                FileWrite($hFile, @CRLF & "--- Erreurs supplémentaires ---" & @CRLF)
                For $i = 1 To $aErreurs[0]
                    Local $aParts2 = StringSplit($aErreurs[$i], "|", 2)
                    If UBound($aParts2) >= 2 Then
                        FileWrite($hFile, "  * " & $aParts2[0] & @CRLF)
                        FileWrite($hFile, "    Raison : " & $aParts2[1] & @CRLF & @CRLF)
                    Else
                        FileWrite($hFile, "  * " & $aErreurs[$i] & @CRLF & @CRLF)
                    EndIf
                Next
            EndIf
            FileClose($hFile)
            _Logging("Fichier erreurs mis à jour : """ & $sChemin & $sNomFichier & """", 5, 0)
        EndIf
    Next
EndFunc   ;==>_EcrireFichierErreursCopie

; ============================================================================
; Renomme les dossiers candidat NNNNNN → NNNNNN_ sur USB et local,
; UNIQUEMENT si _ERREURS_COPIE.txt y existe.
; À appeler UNE SEULE FOIS, après que TOUTES les copies soient terminées,
; juste avant le message final.
; Retour : True si au moins un dossier renommé (= erreurs de copie détectées).
; ============================================================================
Func _FinaliserErreursCopie($sDestUSB, $sDestLocal)
    Local $bErreurs = False
    Local $aDests[2] = [$sDestUSB, $sDestLocal]

    For $k = 0 To 1
        Local $sChemin = $aDests[$k]
        If $sChemin = "" Or Not FileExists($sChemin & "\_ERREURS_COPIE.txt") Then ContinueLoop

        Local $sCheminRenomme = $sChemin & "_"
        If FileExists($sCheminRenomme) Then DirRemove($sCheminRenomme, 1)
        If DirMove($sChemin, $sCheminRenomme) Then
            _Logging("Dossier renommé : """ & $sChemin & """ → """ & $sCheminRenomme & """", 5, 0)
        Else
            _Logging("Renommage ÉCHEC : """ & $sChemin & """ (dossier non renommé)", 5, 0)
        EndIf
        $bErreurs = True
    Next

    Return $bErreurs
EndFunc   ;==>_FinaliserErreursCopie

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

; ============================================================================
; Ferme les processus qui pourraient verrouiller les fichiers du dossier
; www/htdocs avant suppression.
; ============================================================================
Func _LibererVerrousDossierWeb()
    ; Serveurs web + Navigateurs (cache localhost / onglet ouvert)
    Local $aProcs = [ _
        "httpd.exe", _
        "chrome.exe", "firefox.exe", "msedge.exe", "iexplore.exe", _
        "opera.exe", "brave.exe", "vivaldi.exe" _
    ]

    Local $iKilled = 0
    For $p In $aProcs
        Local $iTries = 0
        While ProcessExists($p) And $iTries < 5
            ProcessClose($p)
            Sleep(150)
            $iTries += 1
        WEnd
        If Not ProcessExists($p) And $iTries > 0 Then $iKilled += 1
    Next

    If $iKilled > 0 Then
        _Logging("Libération des verrous web : " & $iKilled & " processus fermé(s).", 2, 0)
        Sleep(400)
    EndIf
    Return $iKilled
EndFunc   ;==>_LibererVerrousDossierWeb

; ============================================================================
; Suppression robuste d'un dossier web (www/htdocs).
; - Retire l'attribut Lecture-seule (et autres) sur tous les fichiers.
; - Utilise DirRemove récursif.
; - En cas d'échec partiel, libère à nouveau les verrous et réessaie.
; - Vérifie post-suppression : log la liste des fichiers résiduels si présent.
; Retour : 1 si tout est supprimé, 0 si des fichiers résistent.
; ============================================================================
Func _SuppressionRobusteDossierWeb($sFolder)
    If Not FileExists($sFolder) Then Return 1

    ; Étape 1 : retirer les attributs R/A/S/H récursivement
    ; (-RASH supprime Read-only, Archive, System, Hidden)
    FileSetAttrib($sFolder & "\*.*", "-RASH", 1) ; 1 = récursif

    ; Étape 2 : 1ère tentative de suppression
    Local $iDel = DirRemove($sFolder, 1)

    ; Étape 3 : si le dossier existe encore, retenter après une courte pause
    If FileExists($sFolder) Then
        _LibererVerrousDossierWeb()
        Sleep(500)
        FileSetAttrib($sFolder & "\*.*", "-RASH", 1)
        $iDel = DirRemove($sFolder, 1)
    EndIf

    ; Étape 4 : vérification finale et log des fichiers résiduels
    If FileExists($sFolder) Then
        Local $aRest = _FileListToArrayRec($sFolder, "*", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
        If IsArray($aRest) And $aRest[0] > 0 Then
            _Logging("Fichiers résiduels non supprimés (" & $aRest[0] & ") :", 5, 0)
            For $k = 1 To ($aRest[0] > 10 ? 10 : $aRest[0])
                _Logging("    » " & $aRest[$k], 5, 0)
            Next
            If $aRest[0] > 10 Then _Logging("    » ... (+ " & ($aRest[0] - 10) & " autres)", 5, 0)
        EndIf
        Return 0
    EndIf

    Return 1
EndFunc   ;==>_SuppressionRobusteDossierWeb

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

; Décompresse un ZIP via Shell.Application.
Func _UnZip($sZipFile, $sDestFolder)
    If Not FileExists($sDestFolder) Then DirCreate($sDestFolder)
    Local $oShell = ObjCreate("Shell.Application")
    If Not IsObj($oShell) Then
        Return False
    EndIf
    Local $oDest = $oShell.NameSpace($sDestFolder)
    Local $oZip  = $oShell.NameSpace($sZipFile)
    If IsObj($oZip) And IsObj($oDest) Then
        $oDest.CopyHere($oZip.Items, 20)
        Return True
    EndIf
    Return False
EndFunc
