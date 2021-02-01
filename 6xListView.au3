#include-once
#include <6xUtil.au3>
#include <GuiListView.au3>
#include <Array.au3>

Global Const $LV_STRING = 0x00
Global Const $LV_CLIP = 0x10
Global Const $LV_ARRAY = 0x20
Global Const $LV_NONE = 0x99

Func ListViewBuscarTexto($control, $texto, $itemAnterior = -1)
	Local $item

	ControlFocus("","",$control)
	$item = _GUICtrlListView_FindInText( $control, $texto, $itemAnterior )
	_GUICtrlListView_EnsureVisible ( $control, $item )
	_GUICtrlListView_SetItemSelected ( $control, $item, True, True)
	return $item
EndFunc

Func ListViewDevolverListaObjetos($control = "SysListView323", $TipoRetorno = 0x00)
	Local $items = "", $count, $item

	$count = _GUICtrlListView_GetItemCount($control)

	For $i = 0 To $count - 1
		$items &= _GUICtrlListView_GetItemText($control, $i) & (($i < $count - 1) ? "|" : "")
	Next

	Local $objetos = StringSplit($items, "|")

	Debug(BitOR($TipoRetorno, $LV_CLIP))

	Select
		Case ($TipoRetorno = 0x00)
			;Enviamos al portapaples los objetos de la lista
			_ArrayToClip($objetos, @CRLF, 1)
	EndSelect

	Return $objetos
EndFunc