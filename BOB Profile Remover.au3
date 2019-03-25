#include <IE.au3>
#include <Excel.au3>
#include <MsgBoxConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include "MetroGUI-UDF\MetroGUI_UDF.au3"
#include "MetroGUI-UDF\_GUIDisable.au3"

;Variables
Global $oIE, $oExcel, $sCompanyId, $oWorkbook, $lLastRow

;ESC key will stop this bot
HotKeySet("{ESC}", "_Exit")
Func _Exit()
	Exit
EndFunc   ;==>_Exit

#Region GUI
_Metro_EnableHighDPIScaling()
_SetTheme("LightGray")
$GUIThemeColor = 0xeff4ff
$Form1 = _Metro_CreateGUI("BOB Profile Remover", 310, 175, -1, -1, True)
$ButtonBKColor = 0x603cba
$Button1 = _Metro_CreateButtonEx2("Start", 210, 50, 80, 30)
$Button2 = _Metro_CreateButtonEx2("Stop", 210, 100, 80, 30)
$Label1 = GUICtrlCreateLabel("BOB Profile Remover", 50, 5, 97, 30)
GUICtrlSetFont(-1, 9, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0x603cba)
$Label2 = GUICtrlCreateLabel("", 50, 70, 100, 40)
GUICtrlSetFont(-1, 11, 400, 0, "Segoe UI")
GUICtrlSetColor(-1, 0xee1111)
GUICtrlSetState(-1, $GUI_HIDE)
$Control_Buttons = _Metro_AddControlButtons(True, False, True, False, True) ;CloseBtn = True, MaximizeBtn = True, MinimizeBtn = True, FullscreenBtn = True, MenuBtn = True
$GUI_CLOSE_BUTTON = $Control_Buttons[0]
$GUI_MINIMIZE_BUTTON = $Control_Buttons[3]
$GUI_MENU_BUTTON = $Control_Buttons[6]
GUISetState(@SW_SHOW)
#EndRegion GUI

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE, $GUI_CLOSE_BUTTON
			ExitLoop
			Exit
		Case $Form1
		Case $GUI_MINIMIZE_BUTTON
			GUISetState(@SW_MINIMIZE, $Form1)
		Case $GUI_MENU_BUTTON
			Local $MenuButtonsArray[2] = ["About", "Exit"]
			Local $MenuSelect = _Metro_MenuStart($Form1, 50, $MenuButtonsArray)
			Switch $MenuSelect ;Above function returns the index number of the selected button from the provided buttons array.
				Case "0"
				Case "1"
					_Metro_GUIDelete($Form1)
					Exit
			EndSwitch
		Case $Button1
			OpenExcelInput()
			RemoveLoop()
			_Excel_Close($oExcel)
		Case $Button2
			Exit
		Case $Label1
	EndSwitch
WEnd

Func OpenExcelInput()
	$oExcel = _Excel_Open(False, False, False, False, True)
	$sWorkbook = @ScriptDir & "\input.xlsx"
	$oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook)
	$lLastRow = $oWorkbook.ActiveSheet.UsedRange.Rows.Count
EndFunc   ;==>OpenExcelInput

Func RemoveLoop()
	WinActivate("Company Lookup")
	WinWaitActive("Company Lookup")
	$oIE = _IEAttach("Company Lookup")
	For $i = 2 To $lLastRow Step 1

		GUICtrlSetData($Label2, "Working on cell       " & $i & " out of " & $lLastRow)
		GUICtrlSetState($Label2, $GUI_SHOW)

		$sCompanyId = _Excel_RangeRead($oWorkbook, "Sheet1", "A" & $i, 1)

		_Send($oIE, "ctl00_m_phContent_vtbCompanyID", $sCompanyId)
		_Click($oIE, "ctl00_m_phContent_m_btnSearch")
		_Click($oIE, "ctl00_m_phContent_gv_CompanySearchResults_ctl02_lnk_Company")
		While 1
			Sleep(100)
			$tid = _IEGetObjById($oIE, "ctl00_ctl10_m_lblClientId")
			$ttext = _IEPropertyGet($tid, "innertext")
			If $ttext = $sCompanyId Then ExitLoop
		WEnd
		_IELinkClickByText($oIE, "Maintain Company Setup")
		_Click($oIE, "ctl00_m_phContent_m_LnkDelete")
		WinWait("Delete Company")
		_Click($oIE, "ctl00_m_phContent_m_Delete")
		_IELoadWait($oIE)
		Print()
		_Click($oIE, "ctl00_ctl10_CompanyLookupHyperlink")
	Next
EndFunc   ;==>RemoveLoop

Func Print()
	_IEAction($oIE, "print")
	WinWait("Print")
	WinActivate("Print")
	Sleep(500)
	ControlClick("Print", "", "Button10")
	Sleep(500)
	ControlClick("Print", "", "Button13")
	Sleep(500)
EndFunc   ;==>Print

#Region MyFunctions ===================================================================
Func _Click($Tab, $ObjIdOrName)
	While 1
		Sleep(100)
		$Obj = _IEGetObjByName($Tab, $ObjIdOrName)
			If Not @error Then
				$Obj = _IEGetObjByName($Tab, $ObjIdOrName)
				ExitLoop
			EndIf
		$Obj = _IEGetObjById($Tab, $ObjIdOrName)
			If Not @error Then
				$Obj = _IEGetObjById($Tab, $ObjIdOrName)
				ExitLoop
			EndIf
	WEnd
	_IEAction($Obj, "click")
EndFunc   ;==>_Click


Func _Send($Tab, $ObjIdOrName, $Text)
	While 1
		Sleep(100)
		$Obj = _IEGetObjByName($Tab, $ObjIdOrName)
			If Not @error Then
				_IEFormElementSetValue($Obj, $Text)
				ExitLoop
			EndIf
		$Obj = _IEGetObjById($Tab, $ObjIdOrName)
			If Not @error Then
				_IEFormElementSetValue($Obj, $Text)
				ExitLoop
			EndIf
	WEnd
EndFunc   ;==>_Send
#EndRegion MyFunctions ===================================================================

Exit