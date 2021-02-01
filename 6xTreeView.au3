#include-once
#include <6xUtil.au3>
#include <GuiTreeView.au3>
#include <GUITreeViewEx.au3>

; #FUNCTION# ====================================================================================================================
; Name ..........: _GUICtrlTreeView_CreateArray
; Description ...: Creates a 1-dimensional array from a TreeView.
; Syntax ........: _GUICtrlTreeView_CreateArray($hTreeView)
; Parameters ....: $hTreeView            - Control ID/Handle to the control
; Return values .: Success - The array returned is one-dimensional and is made up of the following:
;                                $aArray[0] = Number of rows
;                                $aArray[1] = 1st row
;                                $aArray[2] = 2nd row
;                                $aArray[n] = nth row
;                  Failure - Sets @error to non-zero and returns an empty an array.
; Author ........: guinness
; Remarks .......: GUICtrlTreeView.au3 should be included.
; Example .......: Yes
; ===============================================================================================================================
Func _GUICtrlTreeView_CreateArray($hTreeView)
    Local $iItemCount = _GUICtrlTreeView_GetCount($hTreeView)
    Local $aReturn[$iItemCount + 1] = [$iItemCount]
    Local $hItem = _GUICtrlTreeView_GetFirstItem($hTreeView)
    If $iItemCount And $hItem Then
        $aReturn[1] = _GUICtrlTreeView_GetText($hTreeView, $hItem)
        For $i = 2 To $iItemCount
            $hItem = _GUICtrlTreeView_GetNext($hTreeView, $hItem)
            $aReturn[$i] = _GUICtrlTreeView_GetText($hTreeView, $hItem)
        Next
    EndIf
    Return SetError(Int($aReturn[0] = 0), 0, $aReturn)
EndFunc   ;==>_GUICtrlTreeView_CreateArray

Func TreeViewAdd($control)
	_GUICtrlTreeView_BeginUpdate( $control )
	_GUICtrlTreeView_Add($control, 0, "Set -> dosi-imp , %DOSLIIM.DOSI-IMP%" )
	_GUICtrlTreeView_EndUpdate($control)
EndFunc

Func TreeViewCountRamaCompleta($control, $item)
	Local $hijos = _GUICtrlTreeView_GetChildCount($control, $item)

	If($hijos < 1) Then
		$hijos = 0
	Else
		;Vamos al primer hijo
		$item = _GUICtrlTreeView_GetFirstChild($control,$item)
		For $i = 1 To $hijos
			$hijos += TreeViewCountRamaCompleta($control, $item) ;Contabilizamos el total
			$item = _GUICtrlTreeView_GetNextChild($control, $item) ;Nos movemos al siguiente
		Next
	EndIf
	return $hijos
EndFunc