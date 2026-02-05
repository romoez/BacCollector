
#Region "Générer Grille d'évaluation" --------------------
Func _GenererGrilleEvaluation($aNumCandiats)
	Local $chunkSize = 15
	Local $sSeance = GUICtrlRead($cSeance)
	Local $sLabo = GUICtrlRead($cLabo)
	Local $total = UBound($aNumCandiats)
	Local $fileCount = Ceiling($total / $chunkSize)

	Local $sCheminScripDir = @ScriptDir
	If StringRight($sCheminScripDir, 1) <> "\" Then
		$sCheminScripDir = $sCheminScripDir & "\"
	EndIf

	Local $KesNomGrille = "GrilleEval_" & $sSeance & "_" &  $sLabo & ".xlsx"
	If $fileCount = 1 Then
		Local $fileNumber = 1
		While FileExists($sCheminScripDir & $KesNomGrille)
			$KesNomGrille = "GrilleEval_" & $sSeance & "_" &  $sLabo & "_(" & $fileNumber & ")" & ".xlsx"
			$fileNumber += 1
		WEnd
		return _GenererFichiersExcel($aNumCandiats, $sSeance, $sLabo, $KesNomGrille)
	EndIf

	For $i = 0 To $fileCount - 1
		Local $start = $i * $chunkSize
		Local $end = $start + $chunkSize - 1
		If $end >= $total Then $end = $total - 1

		Local $chunk = _ArrayExtract($aNumCandiats, $start, $end)
		Local $KesNomGrille = "GrilleEval_" & $sSeance & "_" &  $sLabo & "_Part" & ($i + 1) & ".xlsx"
		Local $fileNumber = 1
		While FileExists($sCheminScripDir & $KesNomGrille)
			$KesNomGrille = "GrilleEval_" & $sSeance & "_" &  $sLabo & "_Part" & ($i + 1) & "_(" & $fileNumber & ")"  & ".xlsx"
			$fileNumber += 1
		WEnd
		Local $result = _GenererFichiersExcel($chunk, $sSeance, $sLabo, $KesNomGrille)
		If Not $result Then Return 0
	Next
	Return $result
EndFunc

Func _GenererFichiersExcel($aNumCandiats, $sSeance, $sLabo, $NomGrille)
	If Not IsArray($aNumCandiats) Then Return 0
	Local $sSheet = ''
	Local $sFldr = _TmpGrilleVide()


	$sSheet = $sSheet & '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	$sSheet = $sSheet &  @CRLF & '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{482217F5-CA01-4A07-A820-71D429B33AFE}">'
	$sSheet = $sSheet &  @CRLF & '<dimension ref="A2:N24"/>'
	$sSheet = $sSheet &  @CRLF & '<sheetViews><sheetView showGridLines="0" tabSelected="1" zoomScale="85" zoomScaleNormal="85" workbookViewId="0"><selection activeCell="F2" sqref="F2:I4"/></sheetView></sheetViews>'
	$sSheet = $sSheet &  @CRLF & '<sheetFormatPr baseColWidth="10" defaultRowHeight="12.75" x14ac:dyDescent="0.2"/><cols><col min="1" max="1" width="3.85546875" customWidth="1"/><col min="2" max="2" width="11.7109375" customWidth="1"/><col min="3" max="12" width="11.28515625" customWidth="1"/><col min="13" max="13" width="7.42578125" customWidth="1"/><col min="14" max="14" width="11.28515625" customWidth="1"/></cols>'
	$sSheet = $sSheet &  @CRLF & '<sheetData>'
	$sSheet = $sSheet &  @CRLF & '<row r="2" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="B2" s="26" t="s"><v>5</v></c><c r="C2" s="16"><v>' & StringRight($sSeance,1)& '</v></c><c r="D2" s="16"/><c r="E2" s="2"/><c r="F2" s="17"/><c r="G2" s="18"/><c r="H2" s="18"/><c r="I2" s="19"/><c r="J2" s="10"/><c r="K2" s="17"/><c r="L2" s="18"/><c r="M2" s="19"/><c r="N2" s="3"/></row>'
	$sSheet = $sSheet &  @CRLF & '<row r="3" spans="1:14" ht="14.25" x14ac:dyDescent="0.2"><c r="B3" s="26" t="s"><v>7</v></c><c r="C3" s="16"><v>' & StringRight($sLabo,1)& '</v></c><c r="D3" s="16"/><c r="E3" s="3"/><c r="F3" s="20"/><c r="G3" s="21"/><c r="H3" s="21"/><c r="I3" s="22"/><c r="J3" s="9"/><c r="K3" s="20"/><c r="L3" s="21"/><c r="M3" s="22"/><c r="N3" s="3"/></row>'
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
			$Note = '<f>IF(SUM(C%ExcelLineNumber:L%ExcelLineNumber)=0,"",SUM(C%ExcelLineNumber:L%ExcelLineNumber))</f><v/>'
			$CellsErrorXL_Range = $CellsErrorXL_Range & "M" & $ExcelLineNumber & " "
			$StyleCells = "7"
		EndIf

		$Ligne ='<row r="%ExcelLineNumber" spans="1:14" ht="23.1" customHeight="1" x14ac:dyDescent="0.2"><c r="A%ExcelLineNumber" s="31"><v>' & $NumAutoCandid & '</v></c><c r="B%ExcelLineNumber" s="14"><v>' & $CandidNum & '</v></c><c r="C%ExcelLineNumber" s="' & $StyleCells & '"/><c r="D%ExcelLineNumber" s="' & $StyleCells & '"/><c r="E%ExcelLineNumber" s="' & $StyleCells & '"/><c r="F%ExcelLineNumber" s="' & $StyleCells & '"/><c r="G%ExcelLineNumber" s="' & $StyleCells & '"/><c r="H%ExcelLineNumber" s="' & $StyleCells & '"/><c r="I%ExcelLineNumber" s="' & $StyleCells & '"/><c r="J%ExcelLineNumber" s="' & $StyleCells & '"/><c r="K%ExcelLineNumber" s="' & $StyleCells & '"/><c r="L%ExcelLineNumber" s="' & $StyleCells & '"/><c r="M%ExcelLineNumber" s="8" t="str">' & $Note & '</c><c r="N%ExcelLineNumber" s="14"><f t="shared" si="0"/><v>' & $CandidNum & '</v></c></row>'
		$sSheet = $sSheet & @CRLF & StringReplace($Ligne, "%ExcelLineNumber", $ExcelLineNumber)
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

	$sSheetName = $sSeance & ' ' & $sLabo

	Local $sWorkbook = ''
	$sWorkbook = $sWorkbook &  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	$sWorkbook = $sWorkbook &  @CRLF & '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"><fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="21001"/><workbookPr/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="x15"><x15ac:absPath url="" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><xr:revisionPtr revIDLastSave="0" documentId="13_ncr:1_{33728A07-704C-4A9F-86E5-EBF9C54C3A43}" xr6:coauthVersionLast="38" xr6:coauthVersionMax="38" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/><bookViews><workbookView xWindow="120" yWindow="75" windowWidth="13995" windowHeight="10230" xr2:uid="{00000000-000D-0000-FFFF-FFFF00000000}"/></bookViews><sheets><sheet name="' & $sSheetName & '" sheetId="12" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">''' & $sSheetName & '''!$A$1:$N$24</definedName></definedNames><calcPr calcId="191029"/></workbook>'
	Local $FileWorkbook = FileOpen(@TempDir & "\BacCollector\" & $sFldr & "\xl\workbook.xml", 256 + 2) ;$FO_UTF8_NOBOM + $FO_OVERWRITE
	FileWrite($FileWorkbook, $sWorkbook)
	FileClose($FileWorkbook) ; Close the open file after writing

	Local $sApp = ''
	$sApp = $sApp &  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	$sApp = $sApp &  @CRLF & '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Feuilles de calcul</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Plages nommées</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>' & $sSheetName & '</vt:lpstr><vt:lpstr>''' & $sSheetName & '''!Zone_d_impression</vt:lpstr></vt:vector></TitlesOfParts><Manager>Moez Romdhane</Manager><Company>La Communauté des Enseignants d''Informatique en Tunisie</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinkBase>https://github.com/romoez/BacCollector</HyperlinkBase><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>'
	Local $FileApp = FileOpen(@TempDir & "\BacCollector\" & $sFldr & "\docProps\app.xml", 256 + 2) ;$FO_UTF8_NOBOM + $FO_OVERWRITE
	FileWrite($FileApp, $sApp)
	FileClose($FileApp) ; Close the open file after writing

;~ 	Run("explorer.exe /n, /e, " & '"' & @TempDir & "\BacCollector\" & $sFldr & '"')
;~ ************** Zip avec zipfldr.dll
;~ 	Local $Zip = _Zip_Create(@TempDir & "\BacCollector\Grille.zip") ;Create The Zip File. Returns a Handle to the zip File
;~ 	_Zip_AddFolderContents($Zip, @TempDir & "\BacCollector\" & $sFldr) ;Add a folder's content in the zip file


;~ ************** Zip avec external 7-zip32.dll/7-zip64.dll
	Local $ArcFile = @TempDir & "\BacCollector\Grille.zip"
	Local $FolderName = @TempDir & "\BacCollector\" & $sFldr & "\*"

	If _7ZipStartup() Then ; To test dll opening (can be omitted for some functions)
		$retResult = _7ZipAdd(0, $ArcFile, $FolderName, 0, 1)  ; $hWnd, $sArcName, $sFileName, $sFileType = 0, $sHide = 0, $sFilesOpen = 0,...
		If @error Then
			MsgBox(16, "Erreur Grille d'évaluation", "Erreur lors de la préparation de la Grille d'évaluation")
		EndIf
	Else
		MsgBox(16, "Erreur Grille d'évaluation", "Erreur lors de la préparation de la Grille d'évaluation")
	EndIf
	_7ZipShutdown()



	DirRemove(@TempDir & "\BacCollector\" & $sFldr, 1)
	If @error Then
		FileDelete(@TempDir & "\BacCollector\Grille.zip")
		Return 0
	EndIf

	Local $sCheminScripDir = @ScriptDir
	If StringRight($sCheminScripDir, 1) <> "\" Then
		$sCheminScripDir = $sCheminScripDir & "\"
	EndIf

	If Not FileMove( @TempDir & "\BacCollector\Grille.zip", $sCheminScripDir & $NomGrille , 1 ) Then
		MsgBox(16, "Erreur Grille d'évaluation", "Erreur lors de la création de la Grille d'évaluation")
		Return 0
	Else
		Return $sCheminScripDir & $NomGrille
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


Func _FormatCol($sString, $iReservedSpace)
	$sString &= "                                    "
	Return StringLeft($sString, $iReservedSpace)
EndFunc   ;==>_FormatCol

;=========================================================

Func FilesByExtension($sSearchDir)
    Local Const $MAX_EXTENSIONS = 6
    Local Const $MAX_FILES_PER_EXTENSION = 9

;~     If Not FileExists($sSearchDir) Or Not StringInStr(FileGetAttrib($sSearchDir), "D") Then
;~         Return ""
;~     EndIf

    Local $sExtensions = "*.sql;*.htm;*.html;*.css;*.js;*.php;*.accdb;*.csv;*.py;*.ipynb;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.ui;*.dat;*.txt"

    Local $aArray = _FileListToArrayRec($sSearchDir, $sExtensions, $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_NOSORT, $FLTAR_FULLPATH)
    If @error Or Not IsArray($aArray) Or $aArray[0] = 0 Then
        Return ""
    EndIf

    Local $sDrive = "", $sDir = "", $sFileName = "", $sExtension = ""
    For $i = 1 To $aArray[0]
        Local $aPathSplit = _PathSplit($aArray[$i], $sDrive, $sDir, $sFileName, $sExtension)
        $aArray[$i] = StringLower($sExtension)
    Next

    Local $sd = ObjCreate("Scripting.Dictionary")
    For $i = 1 To $aArray[0]
        Local $sExt = $aArray[$i]
        $sd.Item($sExt) = ($sd.Exists($sExt) ? $sd.Item($sExt) + 1 : 1)
    Next

    ; Convertir en tableau et trier
    $aArray = _SDtoArray($sd)
    Local $N = $aArray[0][0]
    If $N > $MAX_EXTENSIONS Then
        $N = $MAX_EXTENSIONS
    EndIf

    ; Construire la chaîne de résultat
    Local $aResults[$N + 1] ; +1 pour l'index 0
    For $i = 1 To $N
        Local $iCount = ($aArray[$i][1] > $MAX_FILES_PER_EXTENSION) ? $MAX_FILES_PER_EXTENSION : $aArray[$i][1]
        $aResults[$i] = $iCount & " " & $aArray[$i][0]
    Next

    Local $sNbFilesByExtension = _ArrayToString($aResults, " • ", 1) ; Concatène du 1er au dernier élément
    Return $sNbFilesByExtension
EndFunc
;=========================================================

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

;=========================================================

Func _GenererRapportPdf($sRapportFileName, $Liste, $DestLocalFldr, $sExamMatiere)
	_SetTitle($sRapportFileName)
	_SetSubject("Bac Pratique " & GUICtrlRead($cBac) & " - " & JourDeLaSemaine() & " " & @MDAY & "-" & @MON & "-" & @YEAR)
	_SetAuthor("Communauté Tunisienne des Enseignants d'Informatique")
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

	Local $chunkSize = 15
	_ArrayDelete($Liste, 0)
	Local $total = UBound($Liste)
	Local $pageCount = Ceiling($total / $chunkSize)

	For $pageNumber = 0 To $pageCount - 1
		Local $start = $pageNumber * $chunkSize
		Local $end = $start + $chunkSize - 1
		If $end >= $total Then $end = $total - 1
		Local $chunk = _ArrayExtract($Liste, $start, $end)

        _BeginPage() ; begin page
            _SetColourStroke($_iColorBorder)
            _Draw_Rectangle(1.5, $pHight - 4.2, 18, 2.8, $PDF_STYLE_STROKED, 0.5, 0xFFFFFF, 0.05)
            _SetColourStroke(0)
            _InsertRenderedText(10.5, $pHight - 2.4, "Bac Pratique " & GUICtrlRead($cBac), "_Calibri", 20, 100, $PDF_ALIGN_CENTER, $_iColorText1, $_iColorText1)
            _InsertRenderedText(10.5, $pHight - 3.1, $sExamMatiere, "_Calibri", 13, 100, $PDF_ALIGN_CENTER, $_iColorText1, $_iColorText1)
            _InsertRenderedText(10.5, $pHight - 3.8, JourDeLaSemaine() & "  " & @MDAY & " " & _MonthFullName(@MON) & " " & @YEAR, "_CalibriB", 13, 100, $PDF_ALIGN_CENTER, $_iColorText2, $_iColorText2)

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
            Local $aCellsContent[6] = ['#', "N° Ins.", "Files", "Nombre de Fichiers / Type (top 6)", "Taille", "Remaques"]
            _InsertPdfTableRow(1.5, $iY, $aCellsWidth,  $aCellsContent, "_CalibriB" , $iRowH, $iStepUp, 0xefefef, $CellBorderColor, $fThickness, $_FontSize  + 1.5, 100, $PDF_ALIGN_CENTER , $_iColorText1, $_iColorText1 )
            Local $N = UBound($chunk, 1)
;~             $N = $N <= 15 ? $N : 15
            $_FontSize = 11.5
            $iY += $iRowH
            For $i = 0 To $N - 1
                Local $array1D[UBound($chunk, 2) + 1]

                ; Copie de la ligne N°1 dans le tableau 1D
                $array1D[0] = $i + 1 + ($pageNumber * 15)
                $array1D[1] = StringLeft($chunk[$i][0], 3) & "-" & StringRight($chunk[$i][0], 3)
                For $j = 1 To UBound($chunk, 2) - 1
                    $array1D[$j + 1] = $chunk[$i][$j]
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
            _Paragraph( $_sConsigne , 1.5 + 1.2, $pHight - $iY, 16.8, "_Calibri", 11)

    ;~ 		===== La liste des extensions
            If StringInStr($sExamMatiere, "STI") Then
                $iY = 23
                Local $iHeaderWidth = 9
                Local $aCellsWidth[1] = [$iHeaderWidth]
                Local $aCellsContent[1] = ['Types de fichiers attendus pour l’examen de STI']
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
                Local $aCellsContent[1] = ['Types de fichiers attendus par section']
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
            Local $_sConsigne = "Veuillez noter que cette liste est indicative et sujette à modification."
            _Paragraph( $_sConsigne , 1.5, $pHight - $iY, $iHeaderWidth, "_Calibri", 9)
        _EndPage()
	Next
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

