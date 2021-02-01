#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=flammable-64.ico
#AutoIt3Wrapper_Outfile=6x.exe
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=Extras 6x
#AutoIt3Wrapper_Res_Fileversion=1.0.0.22
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_ProductName=6x Hack
#AutoIt3Wrapper_Res_CompanyName=Visual MS
#AutoIt3Wrapper_Res_LegalCopyright=Visual MS
#AutoIt3Wrapper_Res_LegalTradeMarks=Visual MS
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(inputboxres, true)

#include <SQLite.au3>
#include <SQLite.dll.au3>

#include <GuiListBox.au3>
#include <GuiListView.au3>
#include <GuiTreeView.au3>
#include <GuiComboBox.au3>
#include <GuiToolbar.au3>

#include <MsgBoxConstants.au3>
#include <AutoItConstants.au3>
#include <StringConstants.au3>
;~ #include <WindowsConstants.au3>
;~ #include <GUIConstantsEx.au3>

#include <Array.au3>
#include <Clipboard.au3>
#include <String.au3>
#include <Misc.au3>
#include <TrayConstants.au3>

#include <HotKey_21b.au3>
#include <WinAPIvkeysConstants.au3>
;~ #include <vkConstants.au3>

#include <6xUtil.au3>
#include <6xListView.au3>
#include <6xTreeView.au3>
#include <6xFormularioControles.au3>

Global Const $fileConfig = @ScriptDir & "\config.ini"
Global Const $vDevelop = "[REGEXPTITLE:(Velneo vDevelop.*)]"
Global Const $ventanaProcesos = "[REGEXPTITLE:(^Proceso.*|^Función.*)]"
Global Const $ventanaInspectorPorColor = "INSPECTOR: Objetos por color"
Local $ventanas = "[REGEXPTITLE:(INSPECTOR: Objetos por tipo.*|Lista de Variables:.*|^Proceso.*|^Función.*|Asistente de funciones proceso|Línea de Proceso)|(Seleccione campo de:.*)|(Asistente para edición de fórmulas)]"

Global $controlFocus, $textoUltimo, $itemUltimo, $AuxNum
Global $ventanaActiva;
Global $db

ConsoleWrite("KeyDelay: " & IniRead($fileConfig, "config", "KeyDelay", 20) & @CRLF)

Opt("WinTitleMatchMode", 2)
Opt("SendKeyDelay", IniRead($fileConfig, "config", "KeyDelay", 20))
Opt("SendKeyDownDelay", 10)

;^ = ctrl, ! = alt
_HotKey_Assign(BitOR($CK_CONTROL, $VK_R), "Reemplazar", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_NOBLOCKHOTKEY, $HK_FLAG_EXTENDEDCALL), $ventanaProcesos)

_HotKey_Assign(BitOR($CK_CONTROL, $VK_F2), "Refactorizar", 0, $ventanaProcesos)
_HotKey_Assign(BitOR($CK_CONTROL, $CK_SHIFT, $VK_F3 ), "Quitarcolores", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_EXTENDEDCALL), $vDevelop)
_HotKey_Assign(BitOR($CK_CONTROL, $VK_F3 ), "If1", 0, $ventanaProcesos)
_HotKey_Assign(BitOR($CK_CONTROL, $VK_F4 ), "ErroresDelProyecto", $HK_FLAG_DEFAULT, $vDevelop)

_HotKey_Assign(BitOR($CK_CONTROL, $VK_F5 ), "ObtenerPorColor", 0, $vDevelop)
HotKeySet("!{F5}", "ListaDeObjetos")

_HotKey_Assign(BitOR($CK_CONTROL, $VK_K), "Comentarios", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_EXTENDEDCALL), $ventanaProcesos)
_HotKey_Assign(BitOR($CK_CONTROL, $VK_F6 ), "QuitarColor", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_NOBLOCKHOTKEY, $HK_FLAG_EXTENDEDCALL), $vDevelop)
_HotKey_Assign($VK_F7, "ObjetosPorColor", $HK_FLAG_DEFAULT, $vDevelop)


_HotKey_Assign(BitOR($CK_CONTROL, $VK_F9 ), "ProcesoTXT", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_EXTENDEDCALL), $ventanaProcesos)
_HotKey_Assign(BitOR($CK_CONTROL, $VK_F10 ), "ProcesoVariablesNoUsadas", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_EXTENDEDCALL), $ventanaProcesos)

;Diseño
_HotKey_Assign(BitOR($CK_CONTROL, $VK_F11 ), "frmAdjuntarVertical", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_EXTENDEDCALL), $vDevelop)

HotKeySet("^!{+}", ExitProg)
Func ExitProg()
	_SQLite_Close($db)
	_SQLite_Shutdown()
    Exit 0
EndFunc
OnAutoItExitRegister("ExitProg")

; Buscar texto
_HotKey_Assign(BitOR($CK_CONTROL, $VK_F ), "BuscarTexto", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_EXTENDEDCALL), $ventanas)
_HotKey_Assign($VK_F3, "BuscarTexto", BitOR($HK_FLAG_DEFAULT, $HK_FLAG_EXTENDEDCALL), $ventanas)

;IniciarBaseDatos()
While 1
	Sleep(200)
	;if($ventanaActiva <> WinGetTitle("[ACTIVE]")) Then
	;	$ventanaActiva = WinGetTitle("[ACTIVE]")
	;	GuardarProcesoFuncionPosicion($ventanaActiva)
	;EndIf
WEnd

Func Quitarcolores()
	Opt("SendKeyDelay", 5)

	Send("!{i}{o}")
	$hWnd = WinWait("INSPECTOR: Objetos por color")
	$control = ControlGetHandle("", "", "SysListView321")
	Local $objetos = ListViewDevolverListaObjetos($control)
	$count = $objetos[0]
	Send("{ENTER}")

	;Pulsamos shit + f3 por cada uno
	Do
		WinWait(("[REGEXPTITLE:(Velneo vDevelop.*)]"))

		;Abrimos la selección de color
		Send("+{F3}")

		$count -= 1
		Send("!{i}{o}")

		WinWait("INSPECTOR: Objetos por color")
		Send("{ENTER}")
	Until $count = 0

	Opt("SendKeyDelay", IniRead($fileConfig, "config", "KeyDelay", 5))
EndFunc

Func Reemplazar()
	Local $cadenaBuscar = InputBox("Texto a buscar", "Introduce el texto que quieres sustituir")
	If($cadenaBuscar == "") Then
		Return
	EndIf

	Local $cadenaReemplazar = InputBox("Reemplazar por", "Introduce el texto por el que será sustituido")
	Local $controlFocus

	Local $control = ControlGetHandle("", "", "SysTreeView321")
	Local $tipoSeleccion = BitOR($TVGN_CARET, $TVGN_FIRSTVISIBLE)

	; Número de hijos
	Local $lineas = _GUICtrlTreeView_GetCount($control) ; Devuelve el total de líneas

	; Vamos al principio
	$item = _GUICtrlTreeView_GetFirstItem($control)
	Local $itemUltimo ; Guardamos el último procesado
	_GUICtrlTreeView_SelectItem($control, $item, $tipoSeleccion)
	_GUICtrlTreeView_SetSelected( $control, $item )

	Local $i
	For $i=0 to $lineas
		Debug($i)

		; Buscamos el item que contenga la cadena
		;$item = _GUICtrlTreeView_FindItem($control, $cadenaBuscar, true, $itemUltimo)
		_GUICtrlTreeView_SelectItemByIndex($control, $item, $lineas )
		$texto = _GUICtrlTreeView_GetText($control, $item)

		; Obtenemos la instrucción
		$instruccion = ObtenerInstruccion($texto)
		Debug($item)
		Debug($texto)

		Switch $instruccion
			Case "Rem"
				$controlFocus = "Edit1"
			Case "Set"
				; Además el edit2
				$controlFocus &= "Edit1:Edit2"
			Case Else
				Debug("Ignoramos " & $instruccion)
		EndSwitch

		If(StringLen($controlFocus)) Then
			If(StringInStr($texto, $cadenaBuscar)) Then
				_GUICtrlTreeView_SelectItem($control, $item, $tipoSeleccion)
				_GUICtrlTreeView_SetSelected( $control, $item )

				Debug($controlFocus)
				Local $controles = StringSplit($controlFocus, ":", $STR_NOCOUNT)

				; Reemplazamos en los controles indicados
				For $controlSeleccionar In $controles
					Debug("reemplazamos")
					ReemplazarTextoInstruccion($control, $controlFocus, $cadenaBuscar, $cadenaReemplazar)
				Next

				$controlFocus = ""
			EndIf
		EndIf

		; Vamos al siguiente
		Debug($item)
		$item = _GUICtrlTreeView_GetNextVisible($control, $item)
		Debug($item)
		#cs
		if($item = 0) Then
			$item = _GUICtrlTreeView_GetFirstItem($control)
			_GUICtrlTreeView_SelectItem($control, $item)
			_GUICtrlTreeView_SetSelected( $control, $item )
		EndIf
		#ce

		$itemUltimo = $item
		;$item = ""
	Next

	Debug("adiós")
	Exit

EndFunc

Func ObtenerInstruccion($texto)
	return StringLeft( $texto, StringInStr( $texto, "->" ) - 2)
EndFunc

Func ReemplazarTextoInstruccion($hwd, $control, $textoBuscar, $textoReemplazar)
	ControlSend("","", $hwd, "{ENTER}")
	$ventana = WinWaitActive("Línea de Proceso")
	ControlFocus($ventana, "", $control)
	$texto = ControlGetText($ventana, "", $control)
	ControlSetText( $ventana, "", $control, StringReplace( $texto, $textoBuscar, $textoReemplazar))
	ControlSend($ventana,"","", "{ENTER}")
	WinWaitActive($ventanaProcesos)
EndFunc

Func ProcesoTXT()
	Local $control, $item, $items, $lineas, $text, $nivel

	$control = ControlGetHandle("", "", "SysTreeView321")

	; Número de hijos
	$items = ControlTreeView("", "", $control, "GetItemCount") ; Devuelve las líneas principales
	$lineas = _GUICtrlTreeView_GetCount($control) ; Devuelve el total de líneas

	$item = _GUICtrlTreeView_GetFirstItem($control)

	For $i = 0 To $lineas - 1
		$nivel = _GUICtrlTreeView_Level($control, $item)
		$text &= _StringRepeat(@TAB, $nivel) & _GUICtrlTreeView_GetText($control, $item) & @CRLF

		; Vamos al siguiente
		$item = _GUICtrlTreeView_GetNext($control, $item)
	Next

	ClipPut($text)
	return $text
EndFunc   ;==>ProcesoTXT

Func ProcesoVariablesNoUsadas()
	Local $proceso = ProcesoTXT()
	Local $variableUsos, $item
	Local $control = ControlGetHandle("", "", "SysTreeView321")

	;Captura todas las declaraciones de variables
	Local $variables = StringRegExp($proceso, "(?im)^.*Set -> (.*) ,.*$", 3)

	if @error Then
		Debug(@extended)
	EndIf
	; Quitamos las repeticiones
	$variables = _ArrayUnique($variables)
	_ArrayDelete($variables, 0)

	Local $variableNoUsadas[1]

	; Por cada variable buscamo si se usa
	Local $i = 0
	For $variable In $variables
		$variableUsos = StringRegExp($proceso, "'" & $variable & "'", 3)
		If UBound($variableUsos) = 0 Then
			_ArrayAdd($variableNoUsadas, $variable, 0, "|", @CRLF, 6)
			$i += 1

			$item = _GUICtrlTreeView_FindItem($control, $variable, True)
			_GUICtrlTreeView_SetBold($control, $item)
		EndIf
	Next
	$variableNoUsadas[0] = $i

	MsgBox($MB_OK, "Proceso finalizado", "Proceso finalizado")

	_ArrayDisplay($variables, "Variables del proceso" )
	_ArrayDisplay($variableNoUsadas, "Variables no usadas" )
EndFunc

Func ObjetosPorColor()
	Send("!{i}{o}")
EndFunc   ;==>ObjetosPorColor

Func ObtenerPorColor()
	Send("!{i}{o}")

	Local $hWnd
	Local $control

	Local $ventana = "INSPECTOR: Objetos por color"

	$hWnd = WinWait("INSPECTOR: Objetos por color")
	$control = ControlGetHandle("", "", "SysListView321")

	Local $objetos = ListViewDevolverListaObjetos($control)

	Send("{ESC}")

	MsgBox($MB_OK, "Objetos por color: " & $objetos[0], "Ya los tienes en el portapapeles")
EndFunc   ;==>ObtenerPorColor

Func QuitarColor()
	Local $arriba = InputBox("Cuantos arriba?", "Indica cuantos cuadros arriba")
	Local $izquierda = InputBox("Cuantos izquierda?", "Indica cuantos cuadros a la izquierda")

	If ($arriba + $izquierda = 0) Then Return

	;Obtenemos el total de objetos
	Send("!{i}{o}")
	$hWnd = WinWait("INSPECTOR: Objetos por color")
	$control = ControlGetHandle("", "", "SysListView321")
	Local $objetos = ListViewDevolverListaObjetos($control)
	$count = $objetos[0]
	Send("{ENTER}")

	;Abrimos el cuadro de objetos por color
	Do
		WinWait(("[REGEXPTITLE:(Velneo vDevelop.*)]"))

		;Abrimos la selección de color
		Send("{F3}")

		;Nos movemos para quitarlo
		Send("{UP " & $arriba & "}")
		Send("{LEFT " & $izquierda & "}")

		;Acepta
		Send("{ENTER}")
		$count -= 1

		Send("!{i}{o}")

		WinWait("INSPECTOR: Objetos por color")
		Send("{ENTER}")
	Until $count = 0

	MsgBox($MB_OK, "Objetos modificados: " & $objetos[0], _ArrayToString($objetos, @CRLF, 1) & @CRLF & "Además los tienes en el portapales ;)")
EndFunc   ;==>QuitarColor

Func Refactorizar()
	Local $control, $item, $cuantos, $multirama, $instruccion, $baja, $hijos

	$control = ControlGetHandle("","", "SysTreeView321")
	$item = _GUICtrlTreeView_GetSelection($control)
	$instruccion = _GUICtrlTreeView_GetText($control, $item)

	If(($instruccion = "if -> 1") Or ($instruccion = "if -> (1=1)") Or ($instruccion = "if -> 1=1")) Then
		$cuantos = _GUICtrlTreeView_GetChildCount($control, $item)
		Debug($cuantos)
		ControlSend("","",$control, "{INS}")
		ControlSend("","",$control, "{DOWN 2}")
	Else
		$cuantos = InputBox("Cuantos??", "Valor: ", 1)
	EndIf

	$AuxNum = $cuantos

	For $i = 1 To $cuantos
		$item = _GUICtrlTreeView_GetSelection($control)
		$instruccion = _GUICtrlTreeView_GetText($control, $item)

		Switch $instruccion
			Case "Libre"
				Debug("Borrando libre")
				ControlSend("","",$control, "{DEL}")
			Case Else
				$hijos = TreeViewCountRamaCompleta($control, $item)

				if($hijos > 0) Then
					$baja = $hijos + 2
				Else
					$baja = 2
				EndIf

				ControlSend("","",$control, "^x")
				ControlSend("","",$control, "{UP 2}")
				ControlSend("","",$control, "^v")
				ControlSend("","",$control, "{DOWN " & ($baja - 1) & "}")

				if($hijos > 0) Then
					ControlSend("","",$control, "{INS}{DOWN}")
				EndIf
				ControlSend("","",$control, "{DOWN}")
		EndSwitch
		Debug($i)
	Next
EndFunc   ;==>Refactorizar

Func RemplazarCampo()
	$casilla = ""
	$casilla = InputBox("Sustituir campo por....", "Valor:")

	If $casilla <> "" Then
		Send("{ENTER}")
		Send("{TAB 2}")
		Send("^c")
		Sleep(200)
		$clipboard = ClipGet()
		$clipboard = StringRegExpReplace($clipboard, "(\%)(.*)(\%)", $casilla)
	EndIf
	Sleep(200)
	ClipPut($clipboard)
	Send("^v")
	Send("{ENTER}")
EndFunc   ;==>RemplazarCampo

Func RemplazarContenedor()
	$casilla = ""
	$casilla = InputBox("Sustituir contenedor por....", "Valor:")

	If $casilla <> "" Then
		Send("{ENTER}")
		Send("{TAB 2}")
		Send("^c")
		Sleep(200)
		$clipboard = ClipGet()
		$clipboard = StringRegExpReplace($clipboard, "(\#|\[)(.*)(\]|\#)", "$1" & $casilla & "$3")
	EndIf
	Sleep(200)
	ClipPut($clipboard)
	Send("^v")
	Send("{ENTER}")
EndFunc   ;==>RemplazarContenedor

Func BuscarTextoObjetosPorTipo($tecla)
	Local $control
	Debug("BuscarTextoObjetosPorTipo")
	$control = ControlGetHandle("", "", "SysListView322")
	Debug($control)

	if( $tecla = $VK_F3 ) Then
		$texto = $textoUltimo
		$item = $itemUltimo
	Else
		$texto = InputBox( "Buscar en proceso", "Texto a buscar: ", "", "", -1, -1, Default, Default, 0, WinGetHandle("[ACTIVE]"))
		$textoUltimo = $texto
		$item = Null
	EndIf

	$itemUltimo = ListViewBuscarTexto($control, $texto, $itemUltimo)
EndFunc

Func BuscarTextoProceso($texto, $itemUltimo)
	Local $control, $item

	Debug("BuscarTextoProceso")

	Debug("buscando: " & $texto)
	Debug("último item: " & $itemUltimo)

	$control = ControlGetHandle("", "", "SysTreeView321")
	$itemUltimo = _GUICtrlTreeView_GetNextVisible($control, $itemUltimo)

	$item = _GUICtrlTreeView_FindItem($control, $texto, true, $itemUltimo)
	Debug($item)
	if($item) Then
		$itemUltimo = $item
		_GUICtrlTreeView_SelectItem($control, $item)
		_GUICtrlTreeView_SetSelected($control, $item)
	EndIf
	return $item
EndFunc

Func BuscarVariableGlobal($texto, $itemUltimo)
	Local $control, $item

	Debug("Buscar combo instrucción")

	Local $instruccion = ControlGetText( "", "", ControlGetHandle("", "", "ComboBox1"))
	Debug($instruccion)

	$control = ControlGetHandle("", "", "ComboBox2")
	$item = _GUICtrlComboBox_SelectString ( $control, $texto, $item )
	return $item
EndFunc

Func BuscarTexto($tecla)
	Local $ventana = WinGetTitle("")
	Debug("Ventana: " & $ventana)

	if( $tecla = $VK_F3 ) Then
		$texto = $textoUltimo
	Else
		$texto = InputBox( "Buscar:", "Texto a buscar: ", $textoUltimo, "", -1, -1, Default, Default, 0, WinGetHandle("[ACTIVE]"))
		$textoUltimo = $texto
		__HK_KeyUp($VK_F)
	EndIf
	Debug("Buscando: " & $texto)

	Select
		Case ($ventana = "Asistente de funciones proceso")
			ContinueCase
		Case ($ventana = "Lista de Variables: GLOBALES")
			ContinueCase
		Case ($ventana = "Lista de Variables: LOCALES")
			$control = ControlGetHandle("", "", "SysListView321")
			$itemUltimo = ListViewBuscarTexto($control, $texto, $itemUltimo)
		Case ($ventana = "INSPECTOR: Objetos por tipo")
			$control = ControlGetHandle("", "", "SysListView322")
			$itemUltimo = ListViewBuscarTexto($control, $texto, $itemUltimo)
		Case ($ventana = "Línea de Proceso")
			Debug("2")
			$itemUltimo = BuscarVariableGlobal($texto, $itemUltimo)
		Case StringRegExp($ventana, "(^Proceso.*|^Función.*|^Seleccione campo de.*)") = 1
			Debug("3")
			$itemUltimo = BuscarTextoProceso($texto, $itemUltimo)
		Case Else
			Debug("Nadaaaa")
	EndSelect
EndFunc

Func If1()
 	Opt("SendKeyDelay", 20)

	Local $cuantos, $sube = 1, $hijos = 0, $instruccion, $control, $item, $itemAnterior
	$cuantos = InputBox("Cuantos??", "Valor: ", 0)

	if($cuantos > 0) Then
		$control = ControlGetHandle("","", "SysTreeView321")
		Send("{INS}")
		Send("{ENTER}")
		Send("if")
		Send("{TAB}")
		Send("1")
		Send("{ENTER}")
		Send("{DOWN 1}")
		$posicionPegar = _GUICtrlTreeView_GetSelection($control)
		Send("{DOWN 1}")

		For $i = 1 To $cuantos
			$item = _GUICtrlTreeView_GetSelection($control)

			If((_GUICtrlTreeView_GetNext($control, $item) = 0)) Then
				Debug("Salimos por que no hay más")
				ExitLoop
			EndIf

			$instruccion = StringUpper(_GUICtrlTreeView_GetText($control, $item))

			Select
				Case $instruccion = "LIBRE"
					Send("{DEL}")
				Case Else
				Send("^x")
				Send("{UP " & $sube & "}")
				Send("^v")

				if($instruccion = "REM -> STOP") Then
					ExitLoop
				EndIf

				$item = _GUICtrlTreeView_GetSelection($control)
				$hijos = TreeViewCountRamaCompleta($control, $item)
				Debug("Hijos: " & $hijos)
				If($hijos <> 0) Then
					Send("{LEFT}")
				EndIf
				Send("{DOWN}")

				$sube = $hijos + 1
			EndSelect
		Next
	EndIf
	Opt("SendKeyDelay", 5)
EndFunc

Func ListaDeObjetos()
	$control = ControlGetHandle("", "", "SysListView323")

	$item = _GUICtrlListView_GetItem($control, 2)
	Local $objetos = ListViewDevolverListaObjetos($control)
	Send("{ESC}")
	MsgBox($MB_OK, "Objetos por color: " & $objetos[0], "Ya los tienes en el portapapeles")
EndFunc

; Ctrl + r
Func Comentarios()
	$control = ControlGetHandle("","", "SysTreeView321")
	ControlSend("","",$control,"{ENTER}")
	ControlSend("","","","rem")
	ControlSend("","","","{ENTER}")
EndFunc

Func ErroresDelProyecto()
	Send("!i")
	Send("{UP}")
	Send("{ENTER}")
	$ventana = "INSPECTOR: Objetos con errores"
	WinWaitActive($ventana)

	$ancho = (@DesktopWidth / 2)
	$alto = (@DesktopHeight / 2)
	$x = (@DesktopWidth / 2) - ($ancho/2)
	$y = (@DesktopHeight / 2) - ($alto/2)

	WinMove($ventana,"", $x, $y, $ancho, $alto)
EndFunc

; Crear base de datos sqlite
Func IniciarBaseDatos()
	_SQLite_Startup()
	ConsoleWrite("_SQLite_LibVersion=" & _SQLite_LibVersion() & @CRLF)

	ConsoleWrite(@ScriptDir & "\database.db" & @CRLF)
	;$db = _SQLite_Open() ; Creates a :memory: database
	$db = _SQLite_Open(@ScriptDir & "\database.db")
	If @error Then
		MsgBox($MB_SYSTEMMODAL, "SQLite Error", "Can't create a memory Database!")
	Else
		If Not _SQLite_Exec(-1, "CREATE TABLE IF NOT EXISTS ProcesosFunciones (Nombre TEXT NOT NULL PRIMARY KEY, Posicion INTEGER);") Then _
			MsgBox($MB_SYSTEMMODAL, "SQLite Error", _SQLite_ErrMsg())
	EndIf
EndFunc

Func GuardarProcesoFuncionPosicion($ventana)
	Local $query
	Local $aRow
	Local $sMsg

	If(StringRegExp($ventana, "^Proceso.*|^Función.*")) Then
		$ventana = StringRegExpReplace($ventana, "Proceso: |Función: ", "")
		ConsoleWrite($ventana & @CRLF)
		_SQLite_Query(-1, "SELECT COUNT(*) FROM ProcesosFunciones WHERE Nombre='" & $ventana & "'", $query)
		While _SQLite_FetchData($query, $aRow) = $SQLITE_OK
			$sMsg &= $aRow[0]
			If($aRow[0] == 0) Then
				$control = ControlGetHandle("", "", "SysListView321")
				_SQLite_Exec(-1, "INSERT INTO ProcesosFunciones VALUES (" & $ventana & "," & 0 & ")")
			EndIf
		WEnd
		ConsoleWrite($sMsg & @CRLF)
	EndIf
EndFunc