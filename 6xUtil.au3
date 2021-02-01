#include-once

#include <Misc.au3>
#include <MsgBoxConstants.au3>
#include <Array.au3>

#include <GuiListView.au3>

Func Debug($msg = "")
	If (Not @Compiled) Then
		ConsoleWrite($msg & @CRLF)
	EndIf
EndFunc

Func TrueFalse($valor)
	If($valor == true) Then
		Return 1
	Else
		Return 0
	EndIf
EndFunc

Func Consola($msg = "")
	ConsoleWrite($msg & @CRLF)
EndFunc

Func _SendEx($ss, $warn = "")
	Local $iT = TimerInit()

	While _IsPressed("10") Or _IsPressed("11") Or _IsPressed("12")
		If $warn <> "" And TimerDiff($iT) > 1000 Then
			MsgBox($MB_TOPMOST, "Warning", $warn)
		EndIf
		Sleep(50)
	WEnd
	Send($ss)
EndFunc

;~ Func EsperarVentana($ventana, $timeout = 1)
;~ 	Local $hwd = WinWaitActive($ventana, "", $timeout)

;~ 	If $hwd Then
;~ 		Return $ventana
;~ 	Else
;~ 		Debug("No se ha encontrado: " & $ventana)
;~ 		Return 0
;~ 	EndIf
;~ EndFunc

Func EsperarVentana($ventana, $timeout = 0)
    Local $hTimer = TimerInit()
	Local $sText = ""

    While 1
        If $timeout <> 0 And TimerDiff($hTimer) >= $timeout Then ExitLoop
        If WinExists($ventana, $sText) And WinActive($ventana, $sText) Then
			Return WinGetHandle($ventana, $sText)
		EndIf
    WEnd
    Return 0
EndFunc

Func ExtraerNombreMapa($senda)
	Return StringRegExp($senda, "(?i)(?:.+\\)*(.+)(\..+)", $STR_REGEXPARRAYGLOBALMATCH)[0]
EndFunc

Func ExtraerNombreMapaDevelop($senda)
	Return StringRegExp($senda, "(?i)(?:.+\\)*(.+)(\..+)\]", $STR_REGEXPARRAYGLOBALMATCH)[0]
EndFunc

; El modo extendido recorre todos los objetos donde se usa y devuelve el nº verdadero de usos.
Func DondeSeUsa($control = "", $extendido = false)
	Local $count = 0
	Local $indice = false
	Local $campo = false

	ControlSend("","", $control, "{F9}")

	Local $hwd = WinWaitActive("[REGEXPTITLE:^(INSPECTOR: Objetos donde se usa el objeto|Velneo vDevelop)$]","", 10)

	Local $titulo = WinGetTitle("[ACTIVE]")

	If $titulo = "INSPECTOR: Objetos donde se usa el objeto" Then
		$control = ControlGetHandle("","","SysListView323")
		$count = _GUICtrlListView_GetItemCount($control)

		If($extendido == true) Then
			; Recorremos todos los objetos donde se usa
			For $i = 0 To ($count - 1)
				$indice = false
				$campo = false

				_GUICtrlListView_SetItemSelected($control,$i,true,true);
				;Leemos el edit donde aparecen la información
				Local $infoEx = ControlGetText("","","Edit1")

				;Desglosamos la información extendida
				;Debug($infoEx)
				Local $vUsoCampos = StringSplit($infoEx,@CR)
				;_ArrayDisplay($vUsoCampos)
				if(StringInStr($infoEx, "indice")) Then $indice = true
				if(StringInStr($infoEx, "campo")) Then $campo = true
			Next
		EndIf
		ControlSend("", "", $hwd, "{ESC}")
	Else
		ControlSend("", "", $hwd, "{ESC}")
	EndIf

	;Si está en modo extendido devolvemos si el campo se usa en índices o campos.
	;Si un campo tiene usos y lo demás es false, el campo se usa solo en objetos no en otras tablas.
	If($extendido) Then
		Return "" & $count & "|" & TrueFalse($indice) & "|" & TrueFalse($campo)
	Else
		Return $count
	EndIf
EndFunc