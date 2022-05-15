#include-once

; #INDEX# ============================================================================================================
; Title .........: GUIExtender
; AutoIt Version : v3.3 or higher
; Language ......: English
; Description ...: Extends and retracts user-defined sections of a GUI either vertically or horizontally
; Remarks .......:
; Note ..........:
; Author(s) .....: Melba23
; ====================================================================================================================

;#AutoIt3Wrapper_au3check_parameters=-d -w 1 -w 2 -w 3 -w 4 -w 5 -w 6 -w- 7

; #GLOBAL VARIABLES# =================================================================================================
; Declare array to hold data for each intialised GUI
Global $aGUIExt_GUI_Data[1][2] = [[0, 9999]]
; [0][0] = GUI count       [n][0] = GUI handle
;                          [n][1] = Section array
; See _GUIExtender_Init for details of section array

; Declare array to hold data about embedded objects
Global $aGUIExt_Obj_Data[1][2] = [[0]]
; [0][0] = Object count    [n][0] = Embedded object ControlID
;                          [n][1] = Original object handle

; Declare an array to hold data on UDF controls and child GUIs within the sections
Global $aGUIExt_UDF_Data[1][5] = [[0]]
; [0][0] = Handle count    [n][0] = Main GUI handle
;                          [n][1] = Control/child GUI handle
;                          [n][2] = Section in which control/child is located
;                          [n][3] = X-coord
;                          [n][4] = Y-coord

; #CURRENT# ==========================================================================================================
; _GUIExtender_Init:              Initialises a GUI to contain sections and sets orientation and fixed point
; _GUIExtender_Clear:             Called on GUI deletion to clear the relevant section from data array
; _GUIExtender_Section_Create:    Defines a GUI section
; _GUIExtender_Section_Activate:  Creates buttons for extension/retraction of GUI sections
; _GUIExtender_Section_BaseCoord: Returns the base coordinate of a section
; _GUIExtender_Section_State:     Returns state of section
; _GUIExtender_Handle_Data:       Stores data on UDF-created controls and child windows within main GUI sections
; _GUIExtender_Obj_Data:          Stores data on embedded objects
; _GUIExtender_Section_Action:    Actions section extension or retraction programatically or via section action buttons
; _GUIExtender_EventMonitor:      Used in GUIGetMsg loops to detect events requiring UDF action
; ====================================================================================================================

; #INTERNAL_USE_ONLY#=================================================================================================
; __GUIExtender_Section_Event:  Used in OnEvent mode to detect section actions and $GUI_EVENT_RESTORE events
; __GUIExtender_Action_Section: Actions extension/retraction for specified sections
; __GUIExtender_Restore:        Used to reset GUI after a MINIMIZE and RESTORE cycle
; __GUIExtender_Handle_Check:   Hides/Shows and moves UDF controls and child GUIs when sections are extended/retracted
; ====================================================================================================================

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Init
; Description ...: Initialises a GUI to contain sections and sets orientation and fixed point
; Syntax.........: _GUIExtender_Init($hWnd[, $iOrient = 0[, $iFixed = 0[, $bComplex = False]]])
; Parameters ....: $hWnd - Handle of GUI containing the sections
;                  $iOrient  - 0 = Vert - GUI extends and retracts in vertical sense
;                              1 = Horz - GUI extends and retracts in horizontal sense
;                  $iFixed   - 0 = GUI Left/Top fixed (Default)
;                              1 = GUI centre fixed
;                              2 = GUI Right/Bottom fixed
;                  $bComplex - True = Complex controls in GUI so use different code
;                              False (default) = No complex controls in GUI
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns 1
;                  Failure:  Returns 0 and sets @error as follows:
;                  1 = Invalid handle
;                  2 = GUI already initialised
;                  3 = Invalid parameter with @extended set: 1 = $iOrient, 2 = $iFixed
; Author ........: Melba23
; Remarks .......: This function should be called before any controls are created within the GUI
;                  Complex controls include: UpDown
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Init($hWnd, $iOrient = 0, $iFixed = 0, $bComplex = False)

	; Check valid handle
	If Not IsHWnd($hWnd) Then Return SetError(1, 0, 0)

	; Check for valid parameters
	Switch $iOrient
		Case 0, 1
			; Valid
		Case Else
			Return SetError(3, 1, 0)
	EndSwitch
	Switch $iFixed
		Case 0, 1, 2
			; Valid
		Case Else
			Return SetError(3, 2, 0)
	EndSwitch

	; See if GUI has already been initialised
	For $i = 1 To $aGUIExt_GUI_Data[0][0]
		If $hWnd = $aGUIExt_GUI_Data[$i][0] Then
			; GUI already in array
			Return SetError(2, 0, 0)
		EndIf
	Next

	; Check for empty element in data array
	Local $iGUI_Index = -1
	For $i = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$i][0] = 0 Then
			$iGUI_Index = $i
			ExitLoop
		EndIf
	Next
	; If none then increase array size
	If $iGUI_Index = -1 Then
		$aGUIExt_GUI_Data[0][0] += 1
		$iGUI_Index = $aGUIExt_GUI_Data[0][0]
		ReDim $aGUIExt_GUI_Data[$iGUI_Index + 1][2]
	EndIf

	; Store GUI handle
	$aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd

	; Declare array to hold GUI section data
	Local $aActive_Section_Data[1][11] = [[0, 0, 1, 0, "", 9999]]
	; [0][0]  = Section count                                            [n][0] = Initial section X/Y-coord
	; [0][1]  = Orientation - 0 = Vertical, 1 = Horizontal               [n][1] = Section height/width (inc border)
	; [0][2]  = All state - 0 = all retracted, 1 = at least 1 extended   [n][2] = Section state - 0 = Retracted, 1 = extended, 2 = static
	; [0][3]  = GUI handle                                               [n][3] = Section anchor ControlID (invisible disabled label)
	; [0][4]  = Fixed point - 0 = Left/Top, 1 = Centre, 2 = Right/Bottom [n][4] = Section final ControlID
	; [0][5]  = Action all button ControlID                              [n][5] = Section action button ControlID
	; [0][6]  = Action all button extended text                          [n][6] = Section action button extended text
	; [0][7]  = Action all button retracted text                         [n][7] = Section action button retracted text
	; [0][8]  = Action all button type                                   [n][8] = Section action button type
	; [0][9]  = Complex controls flag
	; [0][10] = RTL flag

	; Store GUI handle
	$aActive_Section_Data[0][3] = $hWnd
	; Store orientation
	$aActive_Section_Data[0][1] = $iOrient
	; Store fixed point
	$aActive_Section_Data[0][4] = $iFixed
	; Store complex control mode
	$aActive_Section_Data[0][9] = $bComplex
	; Set RTL flag if required
	Local $sFuncName = "GetWindowLongW"
	If @AutoItX64 Then $sFuncName = "GetWindowLongPtrW"
	Local $aResult = DllCall("user32.dll", "long_ptr", $sFuncName, "hwnd", $hWnd, "int", 0xFFFFFFEC) ; $GWL_EXSTYLE
	If BitAND($aResult[0], 0x400000) = 0x0400000 Then ; $WS_EX_LAYOUTRTL
		$aActive_Section_Data[0][10] = 1
	EndIf

	; Store array
	$aGUIExt_GUI_Data[$iGUI_Index][1] = $aActive_Section_Data

	; Force resizing mode to prevent resizing of controls
	Opt("GUIResizeMode", 0x0322) ; $GUI_DOCKALL

	; Set automatic restore function for OnEvent mode scripts
	GUISetOnEvent(-5, "__GUIExtender_Restore") ; $GUI_EVENT_RESTORE

	Return 1

EndFunc   ;==>_GUIExtender_Init

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Clear
; Description ...: Called on GUI deletion to clear the relevant section from data array
; Syntax.........: _GUIExtender_Clear($hWnd)
; Parameters ....: $hWnd - Handle of GUI to delete from data array
; Return values .: Success:  Returns 1
;                  Failure:  Returns 0 and sets @error as follows:
;                  1 = GUI not initialised
; Requirement(s).: v3.3 +
; Author ........: Melba23
; Remarks .......: This function is only to save memory during execution and need not be called on exit
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Clear($hWnd)

	; Get index of GUI in array
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, 0)
	EndIf

	; Clear GUI data from array
	$aGUIExt_GUI_Data[$iGUI_Index][0] = 0
	$aGUIExt_GUI_Data[$iGUI_Index][1] = 0

	; Clear any imbedded objects
	If $aGUIExt_Obj_Data[0][0] Then
		; Extract GUI data
		Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
		; Get CID range
		Local $cStart = $aActive_Section_Data[1][3]
		Local $cEnd = $aActive_Section_Data[$aActive_Section_Data[0][0]][4]
		; Clear any of these CIDs from object array
		For $i = 1 To $aGUIExt_Obj_Data[0][0]
			Switch $aGUIExt_Obj_Data[$i][0]
				Case $cStart To $cEnd
					; Clear CID
					$aGUIExt_Obj_Data[$i][0] = 0
			EndSwitch
		Next
	EndIf

	; Clear any UDF controls/child GUIs
	For $i = 1 To $aGUIExt_UDF_Data[0][0]
		If $aGUIExt_UDF_Data[$i][0] = $hWnd Then
			$aGUIExt_UDF_Data[$i][0] = 0
		EndIf
	Next

	Return 1

EndFunc   ;==>_GUIExtender_Clear

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Section_Create
; Description ...: Defines a GUI section
; Syntax.........: _GUIExtender_Section_Create($hWnd[, $iSection_Coord = -1[, $iSection_Size = -1]])
; Parameters ....: $hWnd - Handle of GUI containing the section
;                  $iSection_Coord - Coordinates of left/top edge of section depending on orientation
;                                    -1 (default) = continue from previous section, or at top of GUI if first section
;                                    -99 = End of section creation
;                  $iSection_Size  - Width/Height of section
;                                    -1 (default) = Fill remaining GUI
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns section ID
;                  Failure:  Returns 0 and sets @error as follows:
;                  1 = GUI not initialised
;                  2 = Section does not follow previous section (either overlap or gap)
;                  3 = Section too large for remaining GUI
; Author ........: Melba23
; Remarks .......: The function creates a disabled label to act as an anchor for the section position
;                  The function must be called BEFORE any native controls in the section have been created
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Section_Create($hWnd, $iSection_Coord = -1, $iSection_Size = -1)

	; Get copy of active GUI section data array
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, 0)
	EndIf

	; Get GUI size
	Local $aWinSize = WinGetClientSize($hWnd)

	; Terminate previous section if not first
	If $aActive_Section_Data[0][0] Then
		$aActive_Section_Data[$aActive_Section_Data[0][0]][4] = GUICtrlCreateDummy() - 1
	EndIf

	; Force parameters to integers
	If $iSection_Coord = Default Then
		$iSection_Coord = -1
	Else
		$iSection_Coord = Int($iSection_Coord)
	EndIf
	If $iSection_Size = Default Then
		$iSection_Size = -1
	Else
		$iSection_Size = Int($iSection_Size)
	EndIf

	; Check if final termination
	If $iSection_Coord = -99 Then
		; Resave array
		$aGUIExt_GUI_Data[$iGUI_Index][1] = $aActive_Section_Data
		Return 99
	EndIf

	; Check section coordinate
	If $aActive_Section_Data[0][0] Then ; Existing sections
		If $iSection_Coord = -1 Then
			; Determine section start coordinate
			$iSection_Coord = $aActive_Section_Data[$aActive_Section_Data[0][0]][0] + $aActive_Section_Data[$aActive_Section_Data[0][0]][1]
		Else
			; Check required section start matches previous section end
			If $aActive_Section_Data[$aActive_Section_Data[0][0]][0] + $aActive_Section_Data[$aActive_Section_Data[0][0]][1] <> $iSection_Coord Then
				Return SetError(2, 0, 0)
			EndIf
		EndIf
	Else ; First section, so default to 0
		If $iSection_Coord = -1 Then
			$iSection_Coord = 0
		EndIf
	EndIf

	; Get max available section size
	Local $iMax_Section_Size
	If $aActive_Section_Data[0][1] = 0 Then
		$iMax_Section_Size = $aWinSize[1] - $iSection_Coord
	Else
		$iMax_Section_Size = $aWinSize[0] - $iSection_Coord
	EndIf
	; Check section size will fit in GUI
	If $iSection_Size = -1 Then
		; Fill remaining GUI
		$iSection_Size = $iMax_Section_Size
	Else
		; Confirm required size will fit
		If $iSection_Size > $iMax_Section_Size Then
			Return SetError(3, 0, 0)
		EndIf
	EndIf

	; Add new section
	$aActive_Section_Data[0][0] += 1
	; ReDim array if required
	If UBound($aActive_Section_Data) < $aActive_Section_Data[0][0] + 1 Then
		ReDim $aActive_Section_Data[$aActive_Section_Data[0][0] + 1][UBound($aActive_Section_Data, 2)]
	EndIf
	; Store passed position and size parameters
	$aActive_Section_Data[$aActive_Section_Data[0][0]][0] = $iSection_Coord
	$aActive_Section_Data[$aActive_Section_Data[0][0]][1] = $iSection_Size
	; Set state to static if not already set
	If $aActive_Section_Data[$aActive_Section_Data[0][0]][2] = "" Then $aActive_Section_Data[$aActive_Section_Data[0][0]][2] = 2
	; Create a zero size disabled label to act as an anchor for the section
	If $aActive_Section_Data[0][1] Then ; Depending on orientation
		$aActive_Section_Data[$aActive_Section_Data[0][0]][3] = GUICtrlCreateLabel("", $iSection_Coord, 0, 1, 1)
	Else
		$aActive_Section_Data[$aActive_Section_Data[0][0]][3] = GUICtrlCreateLabel("", 0, $iSection_Coord, 1, 1)
	EndIf
	GUICtrlSetBkColor(-1, -2) ; $GUI_BKCOLOR_TRANSPARENT
	GUICtrlSetState(-1, 128) ; $GUI_DISABLE
	; Set dummy action ControlID if needed
	If $aActive_Section_Data[$aActive_Section_Data[0][0]][5] = "" Then $aActive_Section_Data[$aActive_Section_Data[0][0]][5] = 9999

	; Resave array
	$aGUIExt_GUI_Data[$iGUI_Index][1] = $aActive_Section_Data

	; Return section ID
	Return $aActive_Section_Data[0][0]

EndFunc   ;==>_GUIExtender_Section_Create

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Section_Activate
; Description ...: Creates buttons for extension/retraction of GUI sections
; Syntax.........: _GUIExtender_Section_Activate($hWnd, $iSection[, $sText_Extended = ""[, $sText_Retracted = ""[, $iX = 0[, $iY = 0[, $iW = 0[, $iH = 0[, $iType = 0[, $iEventMode = 0]]]]]]]]])
; Parameters ....: $hWnd            - Handle of GUI containing the section
;                  $iSection        - Section to action
;                                     0 = all extendable sections
;                  $sText_Extended  - Text on button when section is extended - default: small up/left arrow
;                  $sText_Retracted - Text on button when section is retracted - default: small down/right arrow
;                  $iX              - Left side of the button
;                  $iY              - Top of the button
;                  $iW              - Width of the button
;                  $iH              - Height of the button
;                  $iType           - Type of button:
;                                     0 = pushbutton (default)
;                                     1 = normal button
;                  $iEventMode      - 0 = (default) MessageLoop mode
;                                     1 = OnEvent mode - control automatically linked to __GUIExtender_Section_Event function
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns 1
;                  Failure:  Returns 0 and sets @error as follows:
;                  1 = GUI not initialised
;                  2 = Control not created
; Author ........: Melba23
; Remarks .......: Sections are static unless an action control has been set
;                  Omitting all optional parameters creates a section which can only be actioned programmatically
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Section_Activate($hWnd, $iSection, $sText_Extended = "", $sText_Retracted = "", $iX = 0, $iY = 0, $iW = 1, $iH = 1, $iType = 0, $iEventMode = 0)

	; Get copy of active GUI section data array
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, 0)
	EndIf
	; ReDim array if required
	If $iSection > 1 And UBound($aActive_Section_Data) < $iSection + 1 Then
		ReDim $aActive_Section_Data[$iSection + 1][UBound($aActive_Section_Data, 2)]
	EndIf

	; Set state to extended indicating a extendable section
	$aActive_Section_Data[$iSection][2] = 1
	; If default arrows required
	Local $iDef_Arrows = 0
	If $sText_Extended = "" And $sText_Retracted = "" Then
		$iDef_Arrows = 1
		If $aActive_Section_Data[0][1] Then ; Depending on orientation
			$sText_Extended = 3
			$sText_Retracted = 4
		Else
			$sText_Extended = 5
			$sText_Retracted = 6
		EndIf
	EndIf
	; What type of button?
	Switch $iType
		; Pushbutton
		Case 0
			; Create normal button
			$aActive_Section_Data[$iSection][5] = GUICtrlCreateButton($sText_Extended, $iX, $iY, $iW, $iH)
		Case Else
			; Create pushbutton
			$aActive_Section_Data[$iSection][5] = GUICtrlCreateCheckbox($sText_Extended, $iX, $iY, $iW, $iH, 0x1000) ; $BS_PUSHLIKE
			; Set button state
			GUICtrlSetState(-1, 1) ; $GUI_CHECKED
	EndSwitch
	; Check for error
	If $aActive_Section_Data[$iSection][5] = 0 Then Return SetError(2, 0, 0)
	; Change font if default arrows required
	If $iDef_Arrows Then GUICtrlSetFont($aActive_Section_Data[$iSection][5], 10, 400, 0, "Webdings")
	; Set event function if required
	If $iEventMode Then GUICtrlSetOnEvent($aActive_Section_Data[$iSection][5], "__GUIExtender_Section_Event")
	; Store required text
	$aActive_Section_Data[$iSection][6] = $sText_Extended
	$aActive_Section_Data[$iSection][7] = $sText_Retracted
	; Store button type
	$aActive_Section_Data[$iSection][8] = $iType

	; Resave array
	$aGUIExt_GUI_Data[$iGUI_Index][1] = $aActive_Section_Data

	Return 1

EndFunc   ;==>_GUIExtender_Section_Activate

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Section_BaseCoord
; Description ...: Returns the base coordinate of a section
; Syntax.........: _GUIExtender_Section_BaseCoord($hWnd, $iSection)
; Parameters ....: $hWnd     - Handle of GUI containing the section
;                  $iSection - Specified section
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns section base coordinate
;                  Failure:  Returns -1 and sets @error as follows:
;                  1 = GUI not initialised
;                  2 = Invalid section number
; Author ........: Melba23
; Remarks .......: When default $iSection_Coord parameter used to create a section this function returns the actual section base coordinate.
;                  This can be used to position controls relative to the section rather than using absolute GUI coordinates
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Section_BaseCoord($hWnd, $iSection)

	; Get copy of active GUI section data array
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, -1)
	EndIf

	; Check for valid section
	If $iSection > UBound($aActive_Section_Data) - 1 Then Return SetError(2, 0, -1)

	; Return base coordinate
	Return $aActive_Section_Data[$iSection][0]

EndFunc

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Section_State
; Description ...: Returns current state of section
; Syntax.........: _GUIExtender_Section_State($hWnd, $iSection)
; Parameters ....: $hWnd - Handle of GUI containing the section
;                  $iSection - Section to check
; Requirement(s).: v3.3 +
; Return values .: Success: State of section: 0 = Retracted, 1 = Extended, 2 = Static
;                  Failure:  Returns -1 and sets @error as follows:
;                  1 = GUI not initialised
;                  2 = Invalid section number
; Author ........: Melba23
; Remarks .......: This allows additional GUI controls to reflect the section state
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Section_State($hWnd, $iSection)

	; Get copy of active GUI section data array
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, -1)
	EndIf

	; Check for valid section
	If $iSection > UBound($aActive_Section_Data) - 1 Then Return SetError(2, 0, -1)

	; Return state
	Return $aActive_Section_Data[$iSection][2]

EndFunc   ;==>_GUIExtender_Section_State

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Handle_Data
; Description ...: Stores data on UDF-created controls and child windows within main GUI sections
; Syntax.........: _GUIExtender_Handle_Data($hWnd, $hHandle, $iSection, $iX, $iY)
; Parameters ....: $hWnd     - Handle of GUI containing the section
;                  $hhandle  - Handle of UDF-created control or child GUI
;                  $iSection - Section within which control/child is situated
;                  $iX, $iY  - Coords of control/child relative to main GUI
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns 1
;                  Failure:  Returns 0 and sets @error as follows:
;                  1 = GUI not initialised
;                  2 = Invalid handle
;                  3 = Invalid section
;                  4 = Invalid coordinate value
; Author ........: Melba23
; Remarks .......: This allows UDF-created controls and child GUIs ot be used with the UDF
;                  For child GUIs: Use $WS_POPUP style and _WinAPI_SetParent
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Handle_Data($hWnd, $hHandle, $iSection, $iX, $iY)

	; Check GUI initiated
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, 0)
	EndIf

	; Check parameters
	If Not IsHWnd($hHandle) Or Not WinExists($hHandle) Then Return SetError(2, 0, 0)
	If $iSection > UBound($aActive_Section_Data) - 1 Then Return SetError(3, 0, 0)
	If Not IsInt($iX) Or Not IsInt($iY) Then Return SetError(4, 0, 0)

	; Determine coords in relation to section
	If $aActive_Section_Data[0][1] = 1 Then
		; Horizontal expansion
		$iX = $iX - $aActive_Section_Data[$iSection][0]
	Else
		; Vertical expansion
		$iY = $iY - $aActive_Section_Data[$iSection][0]
	EndIf

	; See if there is an available empty element
	Local $iIndex = -1
	For $i = 1 To $aGUIExt_UDF_Data[0][0]
		If $aGUIExt_UDF_Data[$i][0] = 0 Then
			$iIndex = $i
			ExitLoop
		EndIf
	Next
	; If not then increase array size
	If $iIndex = -1 Then
		$aGUIExt_UDF_Data[0][0] += 1
		$iIndex = $aGUIExt_UDF_Data[0][0]
		ReDim $aGUIExt_UDF_Data[$iIndex + 1][5]
	EndIf
	$aGUIExt_UDF_Data[$iIndex][0] = $hWnd
	$aGUIExt_UDF_Data[$iIndex][1] = $hHandle
	$aGUIExt_UDF_Data[$iIndex][2] = $iSection
	$aGUIExt_UDF_Data[$iIndex][3] = $iX
	$aGUIExt_UDF_Data[$iIndex][4] = $iY

	Return 1

EndFunc   ;==>_GUIExtender_Handle_Data

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Obj_Data
; Description ...: Store data on embedded objects
; Syntax.........: _GUIExtender_Obj_Data($iCID, $oObj)
; Parameters ....: $iCID - Returned ControlID from GUICtrlCreateObj
;                  $iObj - Object reference number from initial object creation
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns 1
;                  Failure:  Returns 0 and sets @error to 1
; Author ........: Melba23, DllCall from trancexx
; Remarks .......: This allows embedded objects to be used in the UDF
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Obj_Data($iCID, $iObj)

	; See if there is an available empty element
	Local $iIndex = -1
	For $i = 1 To $aGUIExt_Obj_Data[0][0]
		If $aGUIExt_Obj_Data[$i][0] = 0 Then
			$iIndex = $i
			ExitLoop
		EndIf
	Next
	; If not then increase array size
	If $iIndex = -1 Then
		; Increase array size
		$aGUIExt_Obj_Data[0][0] += 1
		$iIndex = $aGUIExt_Obj_Data[0][0]
		ReDim $aGUIExt_Obj_Data[$iIndex + 1][2]
	EndIf

	; Store ControlID
	$aGUIExt_Obj_Data[$iIndex][0] = $iCID
	; Determine and store object handle
	Local $aRet = DllCall("oleacc.dll", "int", "WindowFromAccessibleObject", "idispatch", $iObj, "hwnd*", 0)
	If @error Or $aRet[0] Then Return SetError(1, 0, 0)
	$aGUIExt_Obj_Data[$iIndex][1] = $aRet[2]

	Return 1

EndFunc   ;==>_GUIExtender_Obj_Data

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_Section_Action
; Description ...: Actions section extension or retraction programatically or via section action buttons
; Syntax.........: _GUIExtender_Section_Action($hWnd, $iSection[, $iMode = 1[, $iFixed]])
; Parameters ....: $hWnd     - Handle of GUI containing the section
;                  $iSection - Section to action
;                              0 = all moveable sections
;                  $iMode    - 1 (default) = extend section
;                              0 = retract section
;                              9 = toggle section state
;                  $iFixed   - 0 = Left/Top fixed
;                              1 = Expand/contract centred
;                              2 = Right/bottom fixed
;                              Any other value = As set by _GUIExtender_Init (default)
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns 1
;                  Failure:  Returns 0 and sets @error as follows:
;                  1 = GUI not initialised
;                  2 = Invalid section ID (@extended: 1 = not in array, 2 = no _Start function used)
;                  3 = Invalid mode
;                  4 = GUI minimized
; Author ........: Melba23
; Remarks .......: This function is called by the UDF when the section action buttons are clicked,
;                  but can also be called programatically
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_Section_Action($hWnd, $iSection, $iMode = 1, $iFixed = 9)

	; Get copy of active GUI section data array
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, 0)
	EndIf

	; Check mode
	Switch $iMode
		Case 0, 1, 9
			;
		Case Else
			Return SetError(3, 0, 0)
	EndSwitch

	; Check GUI is not minimized
	If BitAND(WinGetState($aActive_Section_Data[0][3]), 16) Then Return SetError(4, 0, 0)

	Local $iStart = $iSection, $iEnd = $iSection

	; If all sections to be toggled
	If $iSection = 0 Then
		; Set bounds to all sections
		$iStart = 1
		$iEnd = $aActive_Section_Data[0][0]
		; If not toggling
		If $iMode <> 9 Then
			; Set mode to correct "all" value
			$iMode = Not($aActive_Section_Data[0][2])
		EndIf
	Else
		; Check for invalid section (either outside array or no _Start function called)
		If $iSection > UBound($aActive_Section_Data) - 1 Then Return SetError(2, 1, 0)
		If $aActive_Section_Data[$iSection][1] = "" Then Return SetError(2, 2, 0)
	EndIf

	; Lock the GUI to prevent excess flicker
	GUISetState(@SW_LOCK, $aActive_Section_Data[0][3])

	; Run action function on selected section(s)
	__GUIExtender_Action_Section($aActive_Section_Data, $iStart, $iEnd, $iMode, $iFixed)

	; Unlock GUI again
	GUISetState(@SW_UNLOCK, $aActive_Section_Data[0][3])

	; Set correct "all" state if there is a "all" control
	If $aActive_Section_Data[0][5] <> 9999 Then
		Local $iAll_State = 0
		; Check if any sections extended
		For $i = 1 To $aActive_Section_Data[0][0]
			If $aActive_Section_Data[$i][2] = 1 Then
				$iAll_State = 1
				ExitLoop
			EndIf
		Next
		; Sync "all" sections control
		Switch $iAll_State
			; None extended
			Case 0
				; Clear flag
				$aActive_Section_Data[0][2] = 0
				; Set text
				GUICtrlSetData($aActive_Section_Data[0][5], $aActive_Section_Data[0][7])
				; Set state if required
				If $aActive_Section_Data[0][8] = 1 Then GUICtrlSetState($aActive_Section_Data[0][5], 4) ; $GUI_UNCHECKED
				; Some extended
			Case Else
				; Set flag
				$aActive_Section_Data[0][2] = 1
				; Set text
				GUICtrlSetData($aActive_Section_Data[0][5], $aActive_Section_Data[0][6])
				; Set state if required
				If $aActive_Section_Data[0][8] = 1 Then GUICtrlSetState($aActive_Section_Data[0][5], 1) ; $GUI_CHECKED
		EndSwitch
	EndIf

	; Resave array
	$aGUIExt_GUI_Data[$iGUI_Index][1] = $aActive_Section_Data

	; Check if any UDF controls or child windows are used
	If $aGUIExt_UDF_Data[0][0] Then
		; Set correct visibility for these items
		__GUIExtender_Handle_Check($hWnd, $aActive_Section_Data)
	EndIf

	Return 1

EndFunc   ;==>_GUIExtender_Section_Action

; #FUNCTION# =========================================================================================================
; Name...........: _GUIExtender_EventMonitor
; Description ...: Used in GUIGetMsg loops to detect events requiring UDF action
; Syntax.........: _GUIExtender_EventMonitor($hWnd, $iMsg)
; Parameters ....: $hWnd - Handle of GUI containing the section
;                  $iMsg - Return from GUIGetMsg
; Requirement(s).: v3.3 +
; Return values .: Success:  Returns 1
;                  Failure:  Returns 0 and sets @error as follows:
;                  1 = GUI not initialised
; Author ........: Melba23
; Remarks .......: The function uses the return from GUIGetMsg to determine if a section action control was clicked
;                  When using multiple GUIs, use "advanced" parameter with GUIGetMsg and pass the returned values to this function
;                  Also checks for $GUI_EVENT_RESTORE and calls internal function to correctly display GUI
; Example........: Yes
;=====================================================================================================================
Func _GUIExtender_EventMonitor($hWnd, $iMsg)

	Local $aActive_Section_Data

	; Get copy of active GUI section data array
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			$aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If Not IsArray($aActive_Section_Data) Then
		Return SetError(1, 0, 0)
	EndIf

	; If GUI is restored, run the restore function
	If $iMsg = -5 Then
		__GUIExtender_Restore() ; $GUI_EVENT_RESTORE
	Else
		; Check if "all" action button clicked
		If $iMsg = $aActive_Section_Data[0][5] Then
			; Action all sections - required state determined later
			_GUIExtender_Section_Action($hWnd, 0)
		Else
			; Check if a section action control has been clicked
			For $i = 1 To $aActive_Section_Data[0][0]
				; If action button clicked
				If $iMsg = $aActive_Section_Data[$i][5] Then
					; Toggle section state
					_GUIExtender_Section_Action($hWnd, $i, 9)
					ExitLoop
				EndIf
			Next
		EndIf
	EndIf

	Return 1

EndFunc   ;==>_GUIExtender_EventMonitor

; #INTERNAL_USE_ONLY#============================================================= ===============================================
; Name...........: __GUIExtender_Section_Event
; Description ...: Used to detect clicks on section action buttons in OnEvent mode
; Author ........: Melba23
; Modified.......:
; Remarks .......: This function is called when section action button are clicked in OnEvent mode
; ===============================================================================================================================
Func __GUIExtender_Section_Event()

	_GUIExtender_EventMonitor(@GUI_WinHandle, @GUI_CtrlId)

EndFunc   ;==>__GUIExtender_Section_Event

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __GUIExtender_Action_Section
; Description ...: Actions extension/retraction for specified sections
; Author ........: Melba23
; Modified.......:
; Remarks .......: This function is called to get the specified sections into the correct state
; ===============================================================================================================================
Func __GUIExtender_Action_Section(ByRef $aActive_Section_Data, $iStart, $iEnd, $iMode, $iFixed = 9)

	Local $iState, $aPos, $iCID, $iDelta_GUI

	For $iSection = $iStart To $iEnd

		; Get current section state
		$iState = $aActive_Section_Data[$iSection][2]

		; Check if state already matches demand or is fixed section
		Switch $iState
			Case $iMode, 2
				; Skip section
				ContinueLoop
		EndSwitch

		; Section needs to toggle state - so set required state
		$iState = Not ($iState)

		; Check Move state
		Switch $iFixed
			Case 0, 1, 2
				; Do nothing
			Case Else
				; Set default value
				$iFixed = $aActive_Section_Data[0][4]
		EndSwitch

		; Get current GUI size and set function calculation values
		Local $aGUIPos = WinGetPos($aActive_Section_Data[0][3])
		Local $iGUI_Fixed = $aGUIPos[2]
		Local $iGUI_Adjust = $aGUIPos[3]
		If $aActive_Section_Data[0][1] Then ; If Horz orientation
			$iGUI_Fixed = $aGUIPos[3]
			$iGUI_Adjust = $aGUIPos[2]
		EndIf

		; Check for RTL GUI
		Local $fRTL = (($aActive_Section_Data[0][10] = 1) ? (True) : (False))

		; Determine whether action required is extension or retraction
		If $iState Then
			; Change button text
			GUICtrlSetData($aActive_Section_Data[$iSection][5], $aActive_Section_Data[$iSection][6])
			; Force action control state if needed
			If $aActive_Section_Data[$iSection][8] = 1 Then GUICtrlSetState($aActive_Section_Data[$iSection][5], 1) ; $GUI_CHECKED
			; Set section state
			$aActive_Section_Data[$iSection][2] = 1
			; Add size of section being extended
			$iGUI_Adjust += $aActive_Section_Data[$iSection][1]
		Else
			; Change button text
			GUICtrlSetData($aActive_Section_Data[$iSection][5], $aActive_Section_Data[$iSection][7])
			; Force action control state if needed
			If $aActive_Section_Data[$iSection][8] = 1 Then GUICtrlSetState($aActive_Section_Data[$iSection][5], 4) ; $GUI_UNCHECKED
			; Set section state
			$aActive_Section_Data[$iSection][2] = 0
			; Subtract size of section being hidden
			$iGUI_Adjust -= $aActive_Section_Data[$iSection][1]
		EndIf

		; Hide controls to prevent ghosting when changing GUI size
		Local $bComplex = $aActive_Section_Data[0][9] ; Check for complex flag
		For $i = $aActive_Section_Data[1][3] To $aActive_Section_Data[$aActive_Section_Data[0][0]][4]
			If $bComplex Then
				; If complex flag set then check control type
				Local $aResult = DllCall("user32.dll", "int", "GetClassNameW", "hwnd", GUICtrlGetHandle($i), "wstr", "", "int", 4096)
				If Not @error Then
					Switch $aResult[2]
						Case "msctls_updown32"
							; Do not hide or ControlGetPos fails
						Case Else
							; Hide all other control types
							GUICtrlSetState($i, 32) ; $GUI_HIDE
					EndSwitch
				EndIf
			Else ; Hide all controls
				GUICtrlSetState($i, 32) ; $GUI_HIDE
			EndIf
		Next

		; Resize and possibly move GUI
		If $aActive_Section_Data[0][1] Then ; Depending on orientation
			; Calculate change in GUI size
			$iDelta_GUI = $aGUIPos[2] - $iGUI_Adjust
			; Check GUI fixed point
			Switch $iFixed
				Case 0
					WinMove($aActive_Section_Data[0][3], "", Default, Default, $iGUI_Adjust, $iGUI_Fixed)
				Case 1
					WinMove($aActive_Section_Data[0][3], "", $aGUIPos[0] + Int($iDelta_GUI / 2), Default, $iGUI_Adjust, $iGUI_Fixed)
				Case 2
					WinMove($aActive_Section_Data[0][3], "", $aGUIPos[0] + $iDelta_GUI, Default, $iGUI_Adjust, $iGUI_Fixed)
			EndSwitch
		Else
			$iDelta_GUI = $aGUIPos[3] - $iGUI_Adjust
			Switch $iFixed
				Case 0
					WinMove($aActive_Section_Data[0][3], "", Default, Default, $iGUI_Fixed, $iGUI_Adjust)
				Case 1
					WinMove($aActive_Section_Data[0][3], "", Default, $aGUIPos[1] + Int($iDelta_GUI / 2), $iGUI_Fixed, $iGUI_Adjust)
				Case 2
					WinMove($aActive_Section_Data[0][3], "", Default, $aGUIPos[1] + $iDelta_GUI, $iGUI_Fixed, $iGUI_Adjust)
			EndSwitch
		EndIf

		; Initial section position = section 1 start
		Local $iNext_Coord = $aActive_Section_Data[1][0]
		; Move sections to required position
		Local $iAdjust_X = 0, $iAdjust_Y = 0
		For $i = 1 To $aActive_Section_Data[0][0]
			; Is this section visible?
			If $aActive_Section_Data[$i][2] > 0 Then
				; Get current position of section anchor
				$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $aActive_Section_Data[$i][3])
				; Adjust for RTL if required
				If $fRTL Then $aPos[0] -= 1

				If $aActive_Section_Data[0][1] Then ; Depending on orientation
					; Determine required change in X position for section controls
					$iAdjust_X = $aPos[0] - $iNext_Coord
					; Determine if controls need to be moved back into the GUI
					If $aPos[1] > $iGUI_Fixed Then $iAdjust_Y = $iGUI_Fixed
				Else
					; Determine required change in Y position for section controls
					$iAdjust_Y = $aPos[1] - $iNext_Coord
					; Determine if controls need to be moved back into the GUI
					If $aPos[0] > $iGUI_Fixed Then $iAdjust_X = $iGUI_Fixed
				EndIf

				; For all controls in this section
				For $j = $aActive_Section_Data[$i][3] To $aActive_Section_Data[$i][4]
					$iCID = $j
					; Adjust the position
					$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $iCID)

					If @error Then
						For $k = 1 To $aGUIExt_Obj_Data[0][0]
							If $aGUIExt_Obj_Data[$k][0] = $j Then
								$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $aGUIExt_Obj_Data[$k][1])
								$iCID = $aGUIExt_Obj_Data[$k][1]
								ExitLoop
							EndIf
						Next
						; If not an object  see if the ControlID returns a handle (an internal tab will not)
						If $iCID = $j And GUICtrlGetHandle($j) = 0 Then $iCID = "Ignore"
					EndIf
					If $iCID = "Ignore" Then ContinueLoop

					; Adjust for RTL if required
					If $fRTL Then $aPos[0] -= $aPos[2]

					; Move control
					ControlMove($aActive_Section_Data[0][3], "", $iCID, $aPos[0] - $iAdjust_X, $aPos[1] - $iAdjust_Y)
					; And show the control
					GUICtrlSetState($j, 16) ; $GUI_SHOW
				Next
				; Determine start position for next visible section
				$iNext_Coord += $aActive_Section_Data[$i][1]
			Else
				; Get current position of hidden section anchor
				$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $aActive_Section_Data[$i][3])
				; Determine if controls in this section need to be moved outside GUI to prevent possible overlap
				If $aActive_Section_Data[0][1] Then ; Depending on orientation
					If $aPos[1] < $iGUI_Fixed Then
						For $j = $aActive_Section_Data[$i][3] To $aActive_Section_Data[$i][4]
							$iCID = $j
							; Adjust the position
							$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $j)
							If @error Then
								For $k = 1 To $aGUIExt_Obj_Data[0][0]
									If $aGUIExt_Obj_Data[$k][0] = $j Then
										$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $aGUIExt_Obj_Data[$k][1])
										$iCID = $aGUIExt_Obj_Data[$k][1]
										ExitLoop
									EndIf
								Next
								; If not an object  see if the ControlID returns a handle (an internal tab will not)
								If $iCID = $j And GUICtrlGetHandle($j) = 0 Then $iCID = "Ignore"
							EndIf
							If $iCID = "Ignore" Then ContinueLoop
							; Move control
							ControlMove($aActive_Section_Data[0][3], "", $iCID, $aPos[0], $aPos[1] + $iGUI_Fixed)
						Next
					EndIf
				Else
					If $aPos[0] < $iGUI_Fixed Then
						For $j = $aActive_Section_Data[$i][3] To $aActive_Section_Data[$i][4]
							$iCID = $j
							; Adjust the position
							$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $j)
							If @error Then
								For $k = 1 To $aGUIExt_Obj_Data[0][0]
									If $aGUIExt_Obj_Data[$k][0] = $j Then
										$aPos = ControlGetPos($aActive_Section_Data[0][3], "", $aGUIExt_Obj_Data[$k][1])
										$iCID = $aGUIExt_Obj_Data[$k][1]
										ExitLoop
									EndIf
								Next
								; If not an object  see if the ControlID returns a handle (an internal tab will not)
								If $iCID = $j And GUICtrlGetHandle($j) = 0 Then $iCID = "Ignore"
							EndIf
							If $iCID = "Ignore" Then ContinueLoop
							; Move control
							ControlMove($aActive_Section_Data[0][3], "", $iCID, $aPos[0] + $iGUI_Fixed, $aPos[1])
						Next
					EndIf
				EndIf
			EndIf
		Next

	Next

EndFunc   ;==>__GUIExtender_Action_Section

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __GUIExtender_Restore
; Description ...: Used to reset GUI after a MINIMIZE and RESTORE cycle
; Author ........: Melba23
; Modified.......:
; Remarks .......: The GUI is not correctly restored after a MINIMIZE if a section is retracted.  This function cycles
;                  the highest extendable section to reset the correct position of the controls
;=====================================================================================================================
Func __GUIExtender_Restore()

	; Get copy of active GUI section data array
	Local $hWnd = WinGetHandle("[ACTIVE]")
	For $iGUI_Index = 1 To $aGUIExt_GUI_Data[0][0]
		If $aGUIExt_GUI_Data[$iGUI_Index][0] = $hWnd Then
			Local $aActive_Section_Data = $aGUIExt_GUI_Data[$iGUI_Index][1]
			ExitLoop
		EndIf
	Next
	If $iGUI_Index > $aGUIExt_GUI_Data[0][0] Then
		Return SetError(1, 0, 0)
	EndIf

	; Lock the GUI to prevent excess flicker
	GUISetState(@SW_LOCK, $aActive_Section_Data[0][3])
	; Look for the highest extendable function
	For $i = UBound($aActive_Section_Data) - 1 To 1 Step -1
		If $aActive_Section_Data[$i][2] <> 2 Then
			; Cycle section
			__GUIExtender_Action_Section($aActive_Section_Data, $i, $i, 9)
			__GUIExtender_Action_Section($aActive_Section_Data, $i, $i, 9)
			ExitLoop
		EndIf
	Next
	; Unlock GUI again
	GUISetState(@SW_UNLOCK, $aActive_Section_Data[0][3])

	Return 1

EndFunc   ;==>__GUIExtender_Restore

; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __GUIExtender_Handle_Check
; Description ...: Hides/Shows and moves UDF controls and child GUIs when sections are extended/retracted
; Author ........: Melba23
; Modified.......:
; Remarks .......: This function is called when section action button are clicked in OnEvent mode
; Example........: Yes
;=====================================================================================================================
Func __GUIExtender_Handle_Check($hWnd, ByRef $aActive_Section_Data)

	; Loop through array and check all handles
	For $i = 1 To $aGUIExt_UDF_Data[0][0]
		; If control on actioned GUI
		If $aGUIExt_UDF_Data[$i][0] = $hWnd Then
			; Check GUI is visible
			If Not BitAND(WinGetState($hWnd), 16) Then
				Local $hHandle = $aGUIExt_UDF_Data[$i][1]
				Local $iSection = $aGUIExt_UDF_Data[$i][2]
				Local $iX = $aGUIExt_UDF_Data[$i][3]
				Local $iY = $aGUIExt_UDF_Data[$i][4]
				; Check section state
				Switch _GUIExtender_Section_State($hWnd, $iSection)
					Case 0
						; Hide $hHandle
						ControlHide($aActive_Section_Data[0][3], "", $hHandle)
					Case 1
						; Locate section
						Local $aPos = ControlGetPos($hWnd, "", $aActive_Section_Data[$iSection][3])
						; Move $hHandle
						WinMove($hHandle, "", $aPos[0] + $iX, $aPos[1] + $iY)
						; Show $hHandle
						ControlShow($aActive_Section_Data[0][3], "", $hHandle)
						; Redraw $hHandle
						DllCall("user32.dll", "bool", "RedrawWindow", "hwnd", $hHandle, "struct*", 0, "handle", 0, "uint", 5)
				EndSwitch
			EndIf
		EndIf
	Next

	Return 1

EndFunc   ;==>__GUIExtender_Handle_Check
