#include-once
#include <Crypt.au3>
#include <APIDiagConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <File.au3>
;#NoTrayIcon

;~ Global $DossierDest = IniRead(@ScriptDir & "\BacBackup.ini", "Params", "DossierDest", "")
;~ Global $Lecteur = IniRead(@ScriptDir & "\BacBackup.ini", "Params", "Lecteur", "")

;Nombre d'occurrences de $substring dans $string
Func _NbOccurrences($substring, $string)
	StringReplace($string, $substring, $substring)
	Return @extended
EndFunc

Func _FineSize($iTaille) ;reçoit une taille en Octet >> retourne la taille en "multiple" approprié Ko ou Mo...
	if $iTaille<1024 Then
		return $iTaille & " o"
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " Ko" ;Kilo
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " Mo" ;Mega
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " Go" ;Gigaoctet
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " To" ;Tera
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " Po" ;Peta
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " Eo" ;Exaoctet
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " Zo" ;Zetaoctet
	EndIf

	$iTaille = Round($iTaille / 1024, 1)
	if $iTaille<1024 Then
		return $iTaille & " Yo" ;Yotaoctet
	EndIf

EndFunc

Func DossiersBac($Path = 1) ; 1:Chemins complets, 0:Chemins relatifs
;~ 	dans certain cas @HomeDrive retoune une chaîne vide --> remplacée par : StringLeft(@WindowsDir,2)
	Local $Bac = _FileListToArray(StringLeft(@WindowsDir,2), "bac*2*", 2, $Path)
	Local $Liste[1] = [0] ;

	If IsArray($Bac) Then
		$Liste[0] += $Bac[0] ;
		_ArrayDelete($Bac, 0)
		_ArrayAdd($Liste, $Bac) ;
	EndIf
	Return $Liste
EndFunc   ;==>DossiersBac
;#########################################################################################
;#########################################################################################
;#########################################################################################
Func DossiersEasyPHPwww($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
	Local $aEasyPHP[1] = [0]

	Local $aTmpEasyPHP = _FileListToArrayRec(StringLeft(@WindowsDir,2), "EasyPHP*;wamp*;xampp*;apachefriends*", 30, 0, 2, $FullPath + 1)

	If IsArray($aTmpEasyPHP) Then
				$aEasyPHP[0] += $aTmpEasyPHP[0]
				_ArrayDelete($aTmpEasyPHP, 0)
				_ArrayAdd($aEasyPHP, $aTmpEasyPHP)
	EndIf

	Local $aTmpEasyPHP = _FileListToArray(@ProgramFilesDir, "EasyPHP*", 2, $FullPath)

	If IsArray($aTmpEasyPHP) Then
				$aEasyPHP[0] += $aTmpEasyPHP[0]
				_ArrayDelete($aTmpEasyPHP, 0)
				_ArrayAdd($aEasyPHP, $aTmpEasyPHP)
	EndIf

	Local $Liste[1] = [0]
	Local $Liste[1] = [0], $www


	If IsArray($aEasyPHP) Then
		For $i = 1 To $aEasyPHP[0]
			If (FileExists($aEasyPHP[$i] & '\www')) Then
				$Liste[0] += 1     ;
				_ArrayAdd($Liste, $aEasyPHP[$i] & '\www')     ;
			ElseIf (FileExists($aEasyPHP[$i] & '\eds-www')) Then
				$Liste[0] += 1         ;
				_ArrayAdd($Liste, $aEasyPHP[$i] & '\eds-www')         ;
			ElseIf (FileExists($aEasyPHP[$i] & '\data\localweb')) Then
				$Liste[0] += 1             ;
				_ArrayAdd($Liste, $aEasyPHP[$i] & '\data\localweb')
			ElseIf (FileExists($aEasyPHP[$i] & '\htdocs')) Then
				$Liste[0] += 1                 ;
				_ArrayAdd($Liste, $aEasyPHP[$i] & '\htdocs')
			ElseIf (FileExists($aEasyPHP[$i] & '\xampp\htdocs')) Then
				$Liste[0] += 1                 ;
				_ArrayAdd($Liste, $aEasyPHP[$i] & '\xampp\htdocs')
			EndIf
		Next
	EndIf

	Return $Liste
EndFunc   ;==>DossiersEasyPHPwww
;#########################################################################################
;#########################################################################################
;#########################################################################################
Func DossiersEasyPHPdata($FullPath = 1) ; 1:Chemins complets, 0:Chemins relatifs
	Local $aEasyPHP[1] = [0]

	Local $aTmpEasyPHP = _FileListToArrayRec(StringLeft(@WindowsDir,2), "EasyPHP*;wamp*;xampp*;apachefriends*", 30, 0, 2, $FullPath + 1)

	If IsArray($aTmpEasyPHP) Then
				$aEasyPHP[0] += $aTmpEasyPHP[0]
				_ArrayDelete($aTmpEasyPHP, 0)
				_ArrayAdd($aEasyPHP, $aTmpEasyPHP)
	EndIf

	Local $aTmpEasyPHP = _FileListToArray(@ProgramFilesDir, "EasyPHP*", 2, $FullPath)

	If IsArray($aTmpEasyPHP) Then
				$aEasyPHP[0] += $aTmpEasyPHP[0]
				_ArrayDelete($aTmpEasyPHP, 0)
				_ArrayAdd($aEasyPHP, $aTmpEasyPHP)
	EndIf

	Local $Liste[1] = [0], $data

	If IsArray($aEasyPHP) Then
		For $i = 1 To $aEasyPHP[0]
			$data = $aEasyPHP[$i] & '\apps\mysql\data'  ;xampp-lite
			If FileExists($data) Then
				$Liste[0] += 1 ;
				_ArrayAdd($Liste, $data) ;
			ElseIf FileExists($aEasyPHP[$i] & '\binaries\mysql\data') Then ;EasyPHP 13.x & 14.x
				$Liste[0] += 1 ;
				_ArrayAdd($Liste, $aEasyPHP[$i] & '\binaries\mysql\data') ;
			ElseIf FileExists($aEasyPHP[$i] & '\mysql\data') Then ;easyphp < 7 & xampp
				$Liste[0] += 1 ;
				_ArrayAdd($Liste, $aEasyPHP[$i] & '\mysql\data') ;
			ElseIf FileExists($aEasyPHP[$i] & '\eds-binaries\dbserver')  Then ;EasyPHP 15.x & 16.x & 17.x
				$data = _FindDataFldr($aEasyPHP[$i] & '\eds-binaries\dbserver')  ;EasyPHP 15.x & 16.x & 17.x
				If FileExists($data) Then
					$Liste[0] += 1 ;
					_ArrayAdd($Liste, $data) ;
				EndIf
			ElseIf (FileExists($aEasyPHP[$i] & '\bin\mysql')) Then ;Wamp
				$data = _FindDataFldr($aEasyPHP[$i] & '\bin\mysql')
				If FileExists($data) Then
					$Liste[0] += 1 ;
					_ArrayAdd($Liste, $data) ;
				EndIf
			ElseIf (FileExists($aEasyPHP[$i] & '\bin\mariadb')) Then ;Wamp
				$data = _FindDataFldr($aEasyPHP[$i] & '\bin\mariadb')
				If FileExists($data) Then
					$Liste[0] += 1 ;
					_ArrayAdd($Liste, $data) ;
				EndIf
			EndIf
		Next
	EndIf
	Return $Liste
EndFunc   ;==>DossiersEasyPHPdata
;#########################################################################################
;#########################################################################################
;#########################################################################################
Func _FindDataFldr($PathEasy)
	Local $aSearch = _FileListToArrayRec($PathEasy,"data",2,1,0,2)
	If IsArray($aSearch) Then Return $aSearch[1]

	Return 0
EndFunc


;#########################################################################################
;#########################################################################################
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
	if $iMinutes < 0 Then $iMinutes = 9999
	Return $iMinutes
EndFunc   ;==>AgeDuFichierEnMinutesModification

Func AgeDuFichierEnMinutesCreation($cFichier)
	Local $cFileDate = FileGetTime($cFichier, 1) ;Get Creation date and time
	If @error Then Return 9999

	Local $iMinutes = _DateDiff("n", $cFileDate[0] & "/" & $cFileDate[1] & "/" & $cFileDate[2] & " " & $cFileDate[3] & ":" & $cFileDate[4] & ":" & $cFileDate[5], _NowCalc())
	if $iMinutes < 0 Then $iMinutes = 9999
	Return $iMinutes
EndFunc   ;==>AgeDuFichierEnMinutesCreation

Func _LockFolder($sObj, $sUserName = @UserName)
	If FileExists($sObj) = 0 Then Return SetError(1, 0, -1)
	RunWait('"' & @ComSpec & '" /c cacls.exe "' & $sObj & '" /E /P "' & $sUserName & '":N', '', @SW_HIDE)
	If $sUserName <> @UserName Then
		RunWait('"' & @ComSpec & '" /c cacls.exe "' & $sObj & '" /E /P "' & @UserName & '":N', '', @SW_HIDE)
	EndIf
EndFunc   ;==>_LockFolder

Func _UnlockFolder($sObj, $sUserName = @UserName)
	If FileExists($sObj) = 0 Then Return SetError(1, 0, -1)
	RunWait('"' & @ComSpec & '" /c cacls.exe "' & $sObj & '" /E /P "' & $sUserName & '":F', '', @SW_HIDE)
	If $sUserName <> @UserName Then
		RunWait('"' & @ComSpec & '" /c cacls.exe "' & $sObj & '" /E /P "' & @UserName & '":F', '', @SW_HIDE)
	EndIf
EndFunc   ;==>_UnlockFolder


;~ Find out current username when executed as admin
;~ https://www.autoitscript.com/forum/topic/183689-find-out-current-username-when-executed-as-admin/#comment-1319682
Func _GetUsername()
    Local $aResult = DllCall("Wtsapi32.dll", "int", "WTSQuerySessionInformationW", "int", 0, "dword", -1, "int", 5, "dword*", 0, "dword*", 0)
    If @error Or $aResult[0] = 0 Then Return SetError(1, 0, @UserName)
    Local $sUsername = BinaryToString(DllStructGetData(DllStructCreate("byte[" & $aResult[5] & "]", $aResult[4]), 1), 2)
    DllCall("Wtsapi32.dll", "int", "WTSFreeMemory", "ptr", $aResult[4])
    Return $sUsername
EndFunc


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
	if IsObj($objWMIService) Then
		$colCompSysPro = $objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct", "WQL", 0x10 + 0x20)

		;If variable is acceptable
		If IsObj($colCompSysPro) Then
			For $objCompSysPro In $colCompSysPro
				;The Hardware Info
				;Manufacturer/Marque/Vendor
				Local $TmpOrdinateur = $objCompSysPro.Vendor
				Local $BuiltInTunisia = 0
				If $TmpOrdinateur = "ECS" Then
					$TmpOrdinateur = "PC assemblé, probablement ""Versus"""
					$BuiltInTunisia = 1
				ElseIf $TmpOrdinateur = "To be filled by O.E.M." Then
					$TmpOrdinateur = "PC assemblé, probablement ""Microlux"" ou ""Discovery"""
					$BuiltInTunisia = 1
				ElseIf $TmpOrdinateur = "System manufacturer" Then
					$TmpOrdinateur = "PC assemblé (Microlux, Versus, Discovery...)"
					$BuiltInTunisia = 1
				EndIf

				$Output &= "Ordinateur  : " & $TmpOrdinateur & @CRLF

				;Model
				If Not $BuiltInTunisia Then
					If $objCompSysPro.Vendor = "Lenovo" Then
						$objModel = $objCompSysPro.Version
					Else
						$objModel = $objCompSysPro.Name
					EndIf
					$Output &= "Modèle      : " & $objModel & @CRLF

				EndIf
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
	$Output &= "Computer ID : "	& _GetUUID()

	Return $Output


EndFunc   ;==>_Hardware

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


Func _MonthName($MonthNum)

    Switch $MonthNum
        Case 01
            $MON = "Janv";
        Case 02
            $MON = "Févr";
        Case 03
            $MON = "Mars";
        Case 04
            $MON = "Avr.";
        Case 05
            $MON = "Mai";
        Case 06
            $MON = "Juin";
        Case 07
            $MON = "Juil";
        Case 08
            $MON = "Août";
        Case 09
            $MON = "Sept";
        Case 10
            $MON = "Oct.";
        Case 11
            $MON = "Nov.";
        Case 12
            $MON = "Déc.";
	EndSwitch

	Return $MON
EndFunc

Func _MonthFullName($MonthNum)

    Switch $MonthNum
        Case 01
            $MON = "Janvier";
        Case 02
            $MON = "Février";
        Case 03
            $MON = "Mars";
        Case 04
            $MON = "Avril";
        Case 05
            $MON = "Mai";
        Case 06
            $MON = "Juin";
        Case 07
            $MON = "Juillet";
        Case 08
            $MON = "Août";
        Case 09
            $MON = "Septembre";
        Case 10
            $MON = "Octobre";
        Case 11
            $MON = "Novembre";
        Case 12
            $MON = "Décembre";
	EndSwitch

	Return $MON
EndFunc

Func _GetUUID($annee=2025)
	$Uuid = RegRead("HKCU\SOFTWARE\BacBackup", "UUID")
	If @error = 0 And StringRegExp($Uuid, "^((\d{4}-)(\w{4} ){2}\w{4})$") = 1 And StringLeft($Uuid, 4)>=2025 Then Return $Uuid
    $Uuid = $annee & "-" & StringFormat('%04X %04X %04X', _
            Random(0, 0xffff), _
            BitOR(Random(0, 0x0fff), 0x4000), _
            BitOR(Random(0, 0x3fff), 0x8000) _
        )
	RegWrite("HKCU\SOFTWARE\BacBackup", "UUID", "REG_SZ", $Uuid)
	Return $Uuid
EndFunc

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
