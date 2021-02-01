
#include-once

;#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w- 7

; #INDEX# =======================================================================================================================
; Title .........: GUITreeViewEx
; AutoIt Version : 3.3.12 +
; Language ......: English
; Description ...: Functions that assist with TreeView loading/saving and checkbox management.
; Remarks .......: - It is good practice to use _GUITreeViewEx_Close when an initiated TreeView is deleted to free the memory used
;                    by the $g_GTVEx_aTVData array which shadows the TreeView contents.
;                  - If the script already has a WM_NOTIFY handler then call the _GUITreeViewEx_WM_NOTIFY_Handler from within the
;                    existing handler and do not use _GUITreeViewEx_RegMsg
; Author(s) .....: Melba23
; ===============================================================================================================================

; #INCLUDES# =========================================================================================================
#include <WindowsConstants.au3>
#include <GuiTreeView.au3>
#include <Array.au3>

; #GLOBAL VARIABLES# =================================================================================================
; TV item selection change flag
Global $g_GTVEx_hItemSelected = 0
; TV check data array
Global $g_GTVEx_aTVData[1][2] = [[0, 0]]
; [0][0] = TreeView count      [n][0] = Handle of intitiated TreeView
; [0][1] = TreeView activated  [n][1] = Array holding TreeView checkbox data

; #CURRENT# ==========================================================================================================
; _GUITreeViewEx_LoadTV            : Fills the TreeView from a delimited string of item titles, level and check state
; _GUITreeViewEx_SaveTV            : Saves the TreeView into a delimited string of item titles, level and check state
; _GUITreeViewEx_InitTV            : Parses the TreeView to create an array of current checkbox states
; _GUITreeViewEx_CloseTV           : Removes the TreeView from the initialised list
; _GUITreeViewEx_RegMsg            : Registers the _GUITreeViewEx_WM_NOTIFY_Handler handler
; _GUITreeViewEx_WM_NOTIFY_Handler : _WM_NOTIFY_Handler handler for the UDF
; _GUITreeViewEx_AutoCheck         : Checks if a checkbox has been altered and adjust parent and children accordingly
; _GUITreeViewEx_Check_All         : Check or clear all checkboxes in an initiated TreeView
; ====================================================================================================================

; #INTERNAL_USE_ONLY#=================================================================================================
; __GTVEx_Adjust_Parents  : Adjusts checkboxes above the one changed
; __GTVEx_Adjust_Children : Adjusts checkboxes below the one changed
; ====================================================================================================================

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_LoadTV
; Description ...: Loads the TreeView from a delimited string of item titles, level and check state
; Syntax.........: _GUITreeViewEx_LoadTV($hTV, $sString [, $sDelimiter = "|" [, $sLevel = "~" [, $sChecked = "#"]]])
; Parameters ....: $hTV        - Handle or ControlID of TreeView
;                  $sString    - Delimited string holding item titles and level information
;                  $sDelimiter - Character delimiting item titles (default = |)
;                  $sLevel     - Character indicating level status (default - ~)
;                  $sChecked   - Character indicating item checked
; Requirement(s).: v3.3.12 +
; Return values .: None
; Author ........: Melba23
; Modified ......:
; Remarks .......: This function works independently of _GUITreeViewEx_RegMsg
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_LoadTV($hTV, $sString, $sDelimiter = "|", $sLevel = "~", $sChecked = "#")

	Local $sTVItem, $iLevel, $bChecked

	; Ensure handle
	If Not IsHWnd($hTV) Then $hTV = GUICtrlGetHandle($hTV)
	; Array to hold current parent handles - base TreeView set to 0
	Local $aLevelParent[100] = [0]
	; Split string
	Local $aTVItems = StringSplit($sString, $sDelimiter)
	; Loop through items
	For $i = 1 To $aTVItems[0]
		; Get required level
		$sTVItem = StringReplace($aTVItems[$i], $sLevel, "")
		$iLevel = @extended
		; Check for checked flag
		$sTVItem = StringReplace($sTVItem, $sChecked, "")
		$bChecked = ( (@extended) ? (True) : (False) )
		; Create item using parent handle
		$aLevelParent[$iLevel + 1] = _GUICtrlTreeView_AddChild($hTV, $aLevelParent[$iLevel], $sTVItem)
		; Check if required
		If $bChecked Then _GUICtrlTreeView_SetChecked($hTV, $aLevelParent[$iLevel + 1], True)
	Next

EndFunc   ;==>_GUITreeViewEx_LoadTV

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_SaveTV
; Description ...: Saves the TreeView into a delimited string of item titles, level and check state
; Syntax.........: _GUITreeViewEx_LoadTV($hTV [, $sDelimiter = "|" [, $sLevel = "~"]])
; Parameters ....: $hTV        - Handle or ControlID of TreeView
;                  $sDelimiter - Character delimiting item titles (default = |)
;                  $sLevel     - Character indicating level status (default - ~)
;                  $sChecked   - Character indicating item checked
; Requirement(s).: v3.3.12 +
; Return values .: String containing TreeView content - format as _GUITreeViewEx_LoadTV
; Author ........: Melba23
; Modified ......:
; Remarks .......: This function works independently of _GUITreeViewEx_RegMsg
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_SaveTV($hTV, $sDelimiter = "|", $sLevel = "~", $sChecked = "#")

	Local $sString = "", $sText, $sLevelCount, $sCheck

	; Ensure handle
	If Not IsHWnd($hTV) Then $hTV = GUICtrlGetHandle($hTV)
	; Work through TreeView items
	Local $hHandle = _GUICtrlTreeView_GetFirstItem($hTV)
	While 1
		; Set level
		$sLevelCount = ""
		For $i = 1 To _GUICtrlTreeView_Level($hTV, $hHandle)
			$sLevelCount &= $sLevel
		Next
		; Get checked state
		$sCheck = ""
		If _GUICtrlTreeView_GetChecked($hTV, $hHandle) Then $sCheck = $sChecked
		; Get text
		$sText = _GUICtrlTreeView_GetText($hTV, $hHandle)
		; Add to string
		$sString &= $sLevelCount & $sCheck & $sText & $sDelimiter
		; Move to next item
		$hHandle = _GUICtrlTreeView_GetNext($hTV, $hHandle)
		; Exit if at end
		If $hHandle = 0 Then ExitLoop
	WEnd
	; Remove final delimiter
	Return StringTrimRight($sString, 1)

EndFunc

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_InitTV
; Description ...: Parses the TreeView to create an array of current checkbox states
; Syntax.........: _GUITreeViewEx_InitTV($hTV)
; Parameters ....: $hTV - Handle or ControlID of TreeView
; Requirement(s).: v3.3 12 +
; Return values .: None
; Author ........: Melba23
; Modified ......:
; Remarks .......:
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_InitTV($hTV)

	; Ensure handle
	If Not IsHWnd($hTV) Then $hTV = GUICtrlGetHandle($hTV)

	; Basic check data array and item count
	Local $aParseTV[10][2], $iParseCount = 0

	; Work through TreeView items
	Local $hHandle = _GUICtrlTreeView_GetFirstItem($hTV)
	While 1
		; Add item to array
		$aParseTV[$iParseCount][0] = $hHandle
		$aParseTV[$iParseCount][1] = _GUICtrlTreeView_GetChecked($hTV, $hHandle)
		; increase count
		$iParseCount += 1
		; Enlarge array if required (minimizes ReDim usage)
		If $iParseCount > UBound($aParseTV) - 1 Then
			ReDim $aParseTV[$iParseCount * 2][2]
		EndIf
		; Move to next item
		$hHandle = _GUICtrlTreeView_GetNext($hTV, $hHandle)
		; Exit if at end
		If $hHandle = 0 Then ExitLoop
	WEnd
	; Remove any empty array elements
	ReDim $aParseTV[$iParseCount][2]

	; Resize main data array
	$g_GTVEx_aTVData[0][0] += 1
	ReDim $g_GTVEx_aTVData[$g_GTVEx_aTVData[0][0] + 1][2]
	; Store TreeView handle and check data
	$g_GTVEx_aTVData[$g_GTVEx_aTVData[0][0]][0] = $hTV
	$g_GTVEx_aTVData[$g_GTVEx_aTVData[0][0]][1] = $aParseTV

EndFunc   ;==>_GUITreeViewEx_InitTV

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_CloseTV
; Description ...: Removes the TreeView from the initialised list
; Syntax.........: _GUITreeViewEx_CloseTV($hTV)
; Parameters ....: $hTV - Handle or ControlID of TreeView
; Requirement(s).: v3.3 +
; Return values .: None
; Author ........: Melba23
; Modified ......:
; Remarks .......:
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_CloseTV($hTV)

	; Ensure handle
	If Not IsHWnd($hTV) Then $hTV = GUICtrlGetHandle($hTV)
	; Search array
	For $i = 1 To $g_GTVEx_aTVData[0][0]
		If $hTV = $g_GTVEx_aTVData[$i][0] Then
			_ArrayDelete($g_GTVEx_aTVData, $i)
			$g_GTVEx_aTVData[0][0] -= 1
		EndIf
		ExitLoop
	Next

EndFunc   ;==>_GUITreeViewEx_CloseTV

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_RegMsg
; Description ...: Registers the _GUITreeViewEx_WM_NOTIFY_Handler handler
; Syntax.........: _GUITreeViewEx_RegMsg()
; Parameters ....: None
; Requirement(s).: v3.3.12 +
; Return values .: None
; Author ........: Melba23
; Modified ......:
; Remarks .......: If the script already has a WM_NOTIFY handler then call the _GUITreeViewEx_WM_NOTIFY_Handler from within the
;                  existing handler and do not use _GUITreeViewEx_RegMsg
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_RegMsg()

	; Register handler
	GUIRegisterMsg($WM_NOTIFY, "_GUITreeViewEx_WM_NOTIFY_Handler")

EndFunc   ;==>_GUITreeViewEx_RegMsg

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_WM_NOTIFY_Handler
; Description ...: _WM_NOTIFY_Handler handler for the UDF
; Syntax.........: _GUITreeViewEx_WM_NOTIFY_Handler($hWnd, $iMsg, $wParam, $lParam)
; Parameters ....: None
; Requirement(s).: v3.3.12 +
; Return values .: None
; Author ........: Melba23
; Modified ......:
; Remarks .......: If the script already has a WM_NOTIFY handler then call the _GUITreeViewEx_WM_NOTIFY_Handler from within the
;                  existing handler and do not use _GUITreeViewEx_RegMsg
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_WM_NOTIFY_Handler($hWnd, $iMsg, $wParam, $lParam)

	#forceref $hWnd, $iMsg, $wParam
	; Create NMTREEVIEW structure
	Local $tStruct = DllStructCreate("struct;hwnd hWndFrom;uint_ptr IDFrom;INT Code;endstruct;" & _
			"uint Action;struct;uint OldMask;handle OldhItem;uint OldState;uint OldStateMask;" & _
			"ptr OldText;int OldTextMax;int OldImage;int OldSelectedImage;int OldChildren;lparam OldParam;endstruct;" & _
			"struct;uint NewMask;handle NewhItem;uint NewState;uint NewStateMask;" & _
			"ptr NewText;int NewTextMax;int NewImage;int NewSelectedImage;int NewChildren;lparam NewParam;endstruct;" & _
			"struct;long PointX;long PointY;endstruct", $lParam)
	Local $hWndFrom = DllStructGetData($tStruct, "hWndFrom")
	; Check TreeView initiated
	For $i = 1 To $g_GTVEx_aTVData[0][0]
		If $hWndFrom = $g_GTVEx_aTVData[$i][0] Then
			Switch DllStructGetData($tStruct, "Code")
				; If item selection changed
				Case $TVN_SELCHANGEDA, $TVN_SELCHANGEDW
					Local $hItem = DllStructGetData($tStruct, "NewhItem")
					; Set flag to selected item handle
					If $hItem Then $g_GTVEx_hItemSelected = $hItem
					; Store TreeView handle
					$g_GTVEx_aTVData[0][1] = $hWndFrom
			EndSwitch
		EndIf
	Next
EndFunc   ;==>_GUITreeViewEx_WM_NOTIFY_Handler

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_AutoCheck
; Description ...: Checks if an item chaeckbox has been altered and adjust parent and children accordingly
; Syntax.........: _GUITreeViewEx_AutoCheck()
; Parameters ....: None
; Requirement(s).: v3.3.12 +
; Return values .: None
; Author ........: Melba23
; Modified ......:
; Remarks .......: This function must be placed in the script idle loop
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_AutoCheck()

	Local $aTVCheckData, $bState, $iItemIndex

	; If an item has been selected
	If $g_GTVEx_hItemSelected Then
		; Read TreeView handle and confirm initiated
		Local $hTV = $g_GTVEx_aTVData[0][1]
		For $iTVIndex = 1 To $g_GTVEx_aTVData[0][0]
			If $hTV = $g_GTVEx_aTVData[$iTVIndex][0] Then
				; Extract check data array
				$aTVCheckData = $g_GTVEx_aTVData[$iTVIndex][1]
				; Determine checked state
				$bState = _GUICtrlTreeView_GetChecked($hTV, $g_GTVEx_hItemSelected)
				; Find item in array
				$iItemIndex = _ArraySearch($aTVCheckData, $g_GTVEx_hItemSelected)
				; If checked state has altered
				If $aTVCheckData[$iItemIndex][1] <> $bState Then
					; Store new state
					$aTVCheckData[$iItemIndex][1] = $bState
					; Adjust parents and children as required
					__GTVEx_Adjust_Parents($hTV, $g_GTVEx_hItemSelected, $aTVCheckData, $bState)
					__GTVEx_Adjust_Children($hTV, $g_GTVEx_hItemSelected, $aTVCheckData, $bState)
				EndIf
				; Store amended array
				$g_GTVEx_aTVData[$iTVIndex][1] = $aTVCheckData
				; No point in looping further
				ExitLoop
			EndIf
		Next
		; Clear selected flag
		$g_GTVEx_hItemSelected = 0
	EndIf

EndFunc   ;==>_GUITreeViewEx_AutoCheck

; #FUNCTION# =========================================================================================================
; Name...........: _GUITreeViewEx_Check_All
; Description ...: Check or clear all checkboxes in an initiated TreeView
; Syntax.........: _GUITreeViewEx_Check_All($hTV [, $bState = True])
; Parameters ....: $hTV    - Handle or ControlID of TreeView
;                  $bState - True (default) = set all checkboxes
;                            False          = clear all checkboxes
; Requirement(s).: v3.3.12 +
; Return values .: None
; Author ........: Melba23
; Modified ......:
; Remarks .......:
; Example........: Yes
;=====================================================================================================================
Func _GUITreeViewEx_Check_All($hTV, $bState = True)

	; Ensure handle
	If Not IsHWnd($hTV) Then $hTV = GUICtrlGetHandle($hTV)

	; Confirm TreeView is initiated
	For $iIndex = 1 To $g_GTVEx_aTVData[0][0]
		If $hTV = $g_GTVEx_aTVData[$iIndex][0] Then
			; Extract check data
			Local $aTVData = $g_GTVEx_aTVData[$iIndex][1]
			; Loop through items
			For $i = 0 To UBound($aTVData) - 1
				; Adjust item
				_GUICtrlTreeView_SetChecked($hTV, $aTVData[$i][0], $bState)
				; Adjust array
				$aTVData[$i][1] = $bState
			Next
			; Store amended array
			$g_GTVEx_aTVData[$iIndex][1] = $aTVData
			; No point in looping further
			ExitLoop
		EndIf
	Next

EndFunc   ;==>_GUITreeViewEx_Check_All

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __GTVEx_Adjust_Parents
; Description ...: Adjusts checkboxes above the one changed
; Author ........: Melba23
; ===============================================================================================================================
Func __GTVEx_Adjust_Parents($hTV, $hPassedItem, ByRef $aTVCheckData, $bState = True)

	; Get handle of parent
	Local $hParent = _GUICtrlTreeView_GetParentHandle($hTV, $hPassedItem)
	If $hParent = 0 Then Return
	; Assume parent is to be adjusted
	Local $bAdjustParent = True
	; Find parent in array
	Local $iItemIndex = _ArraySearch($aTVCheckData, $hParent)
	; Need to confirm all siblings clear before clearing parent
	If $bState = False Then
		; Check on number of siblings
		Local $iCount = _GUICtrlTreeView_GetChildCount($hTV, $hParent)
		; If only 1 sibling then parent can be cleared - if more then need to look at them all
		If $iCount <> 1 Then
			; Number of siblings checked
			Local $iCheckCount = 0
			; Move through previous siblings
			Local $hSibling = $hPassedItem
			While 1
				$hSibling = _GUICtrlTreeView_GetPrevSibling($hTV, $hSibling)
				; If found
				If $hSibling Then
					; Is sibling checked)
					If _GUICtrlTreeView_GetChecked($hTV, $hSibling) Then
						; Increase count if so
						$iCheckCount += 1
					EndIf
				Else
					; No point in continuing
					ExitLoop
				EndIf
			WEnd
			; Move through later siblings
			$hSibling = $hPassedItem
			While 1
				$hSibling = _GUICtrlTreeView_GetNextSibling($hTV, $hSibling)
				If $hSibling Then
					If _GUICtrlTreeView_GetChecked($hTV, $hSibling) Then
						$iCheckCount += 1
					EndIf
				Else
					ExitLoop
				EndIf
			WEnd
			; If at least one sibling checked then do not clear parent
			If $iCheckCount Then $bAdjustParent = False
		EndIf
	EndIf
	; If parent is to be adjusted
	If $bAdjustParent Then
		; Adjust the array
		$aTVCheckData[$iItemIndex][1] = $bState
		; Adjust the parent
		_GUICtrlTreeView_SetChecked($hTV, $hParent, $bState)
		; And now do the same for the generation above
		__GTVEx_Adjust_Parents($hTV, $hParent, $aTVCheckData, $bState)
	EndIf

EndFunc   ;==>__GTVEx_Adjust_Parents

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __GTVEx_Adjust_Children
; Description ...: Adjusts checkboxes below the one changed
; Author ........: Melba23
; ===============================================================================================================================
Func __GTVEx_Adjust_Children($hTV, $hPassedItem, ByRef $aTVCheckData, $bState = True)

	Local $iItemIndex

	; Get the handle of the first child
	Local $hChild = _GUICtrlTreeView_GetFirstChild($hTV, $hPassedItem)
	If $hChild = 0 Then Return
	While 1
		; Find child index
		$iItemIndex = _ArraySearch($aTVCheckData, $hChild)
		; Adjust the array
		$aTVCheckData[$iItemIndex][1] = $bState
		; Adjust the child
		_GUICtrlTreeView_SetChecked($hTV, $hChild, $bState)
		; And now do the same for the generation beow
		__GTVEx_Adjust_Children($hTV, $hChild, $aTVCheckData, $bState)
		; Now get next child
		$hChild = _GUICtrlTreeView_GetNextChild($hTV, $hChild)
		; Exit the loop if no more found
		If $hChild = 0 Then ExitLoop
	WEnd

EndFunc   ;==>__GTVEx_Adjust_Children
