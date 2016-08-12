#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=E:\Soft\Download\Untitled.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
; coded by VinhPham
; opdo.vn
; Chỉ sử dụng cho mục đích cá nhân.
#include <ListViewConstants.au3>
#include <GuiListView.au3>
#include <StaticConstants.au3>
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <GuiButton.au3>
#include <Array.au3>
#include <Icons.au3>
#include <Sound.au3>
#include <APIGdiConstants.au3>
#include <WinAPIGdi.au3>
#include <Excel.au3>

_WinAPI_AddFontResourceEx(@ScriptDir & '\fontawesome.ttf', $FR_PRIVATE) ; load font

Global $_MAIN_Pic_Control[1] = [0], $_LAST_HOVER = -1, $_ANSWER = '', $_NUMBER, $_Question, $_SOUND, $_COUNT = 1, $_LISTEN_FLAG = False, $myNewList[0]

SplashTextOn("Load data...", "Just a moment", 100, 70)
_Get_Excel_List()
SplashOff()
$_NUMBER = _Choose_Level() ; chọn lv, trả về number là số pic
Global $_List = _Get_List() ; lấy danh sách các từ vựng
#Region Tạo gui
$w = Int((@DesktopWidth - 100) / 5) + 50
Do
	$w -= 50
	$w2 = 160 + Int($_NUMBER / 5) * ($w + 13)
Until $w2 < @DesktopHeight - 100
$w3 = ($w + 13) * 5 + 47
Global $GUIMAIN = GUICreate("Study English", $w3, $w2, -1, -1, BitOR($WS_POPUP, $WS_BORDER, $WS_SIZEBOX), -1)
GUISetBkColor(0xFFFFFF, $GUIMAIN)
GUICtrlCreateLabel("", 0, 0, $w3, 70, -1, $GUI_WS_EX_PARENTDRAG)
GUICtrlSetBkColor(-1, "0x0066cc")
$Question_Txt = GUICtrlCreateLabel("Question " & $_COUNT, 0, 14, $w3, 49, $SS_CENTER, -1)
GUICtrlSetFont(-1, 18, 350, 0, "Segoe UI Semilight")
GUICtrlSetColor(-1, "0xFFFFFF")
GUICtrlSetBkColor(-1, "-2")
$Listen_Again = GUICtrlCreateLabel("Show word's mean", 0, 90, $w3 - 50, 30, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
GUICtrlSetState(-1, $GUI_SHOW)
GUICtrlSetFont(-1, 16, 400, 0, "FontAwesome")
GUICtrlSetBkColor(-1, "-2")
GUICtrlSetCursor(-1, 0)
$Favorite = GUICtrlCreateLabel("", $w3 - 50, 90, 50, 30, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
GUICtrlSetState(-1, $GUI_SHOW)
GUICtrlSetFont(-1, 16, 400, 0, "FontAwesome")
GUICtrlSetBkColor(-1, "-2")
GUICtrlSetCursor(-1, 0)

$Exit = GUICtrlCreateLabel("", $w3 - 23, 0, 23, 21, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
GUICtrlSetFont(-1, 14, 400, 0, "FontAwesome")
GUICtrlSetColor(-1, "0xFFFFFF")
GUICtrlSetBkColor(-1, "-2")
$Mini = GUICtrlCreateLabel("", $w3 - 23 * 2, 0, 23, 21, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
GUICtrlSetFont(-1, 14, 400, 0, "FontAwesome")
GUICtrlSetColor(-1, "0xFFFFFF")
GUICtrlSetBkColor(-1, "-2")
_Create_Picture($_NUMBER, $w) ; tạo các pic hiển thị
_Set_Question() ; set câu hỏi
GUISetState(@SW_SHOW, $GUIMAIN) ; hiển thị GUI
#EndRegion Tạo gui
;WinMove($GUIMAIN,'',0,0,@DesktopWidth,@DesktopHeight-30)
While 1
	If $_LISTEN_FLAG Then ; thay đổi chữ Playing khi sound đã play xong
		If _SoundStatus($_SOUND) == 'stopped' Then
			_SoundPlay($_SOUND)
			;GUICtrlSetData($Listen_Again," Listen again")
			;$_LISTEN_FLAG = False
		EndIf
	EndIf

	$info = GUIGetCursorInfo($GUIMAIN) ; lấy thông tin event click và hover control
	If $info[2] Then ; click control
		Do ; đợi buông click
			$info = GUIGetCursorInfo($GUIMAIN)
		Until $info[2] = 0
		_ControlClick($info[4])
	EndIf
	If $info[4] Then ; hover control
		_ControlHover($info[4])
	Else
		_ControlNormal($_LAST_HOVER)
	EndIf

WEnd
#Region Các hàm control click, hover
Func _ControlClick($control)
	Switch $control
		Case $Favorite
			If IniRead(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN', $_ANSWER, '') = $_ANSWER Then
				IniDelete(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN', $_ANSWER)
				GUICtrlSetColor($Favorite, 0x808080)
			Else
				IniWrite(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN', $_ANSWER, $_ANSWER)
				GUICtrlSetColor($Favorite, 0xFF0000)
			EndIf
			
		Case $Listen_Again ; bắt sự kiện khi ấn nghe lại
			;_SoundPlay($_SOUND)
			;$_LISTEN_FLAG = True
			;GUICtrlSetData($Listen_Again," Playing...")
			If GUICtrlRead($Listen_Again) == "Hide word's mean" Then
				GUICtrlSetData($Listen_Again, "Show word's mean")
				GUICtrlSetData($Question_Txt, 'Question ' & $_COUNT)
			Else
				GUICtrlSetData($Question_Txt, $_ANSWER & ' \' & IniRead(@ScriptDir & '\Data\word.ini', $_ANSWER, 'b', '') & '\ (' & IniRead(@ScriptDir & '\Data\word.ini', $_ANSWER, 'c', '') & ') : ' & IniRead(@ScriptDir & '\Data\word.ini', $_ANSWER, 'd', ''))
				GUICtrlSetData($Listen_Again, "Hide word's mean")
			EndIf
		Case $Mini
			GUISetState(@SW_MINIMIZE, $GUIMAIN)
		Case $GUI_EVENT_CLOSE, $Exit
			Exit
		Case Else
			For $i = 1 To $_NUMBER ; bắt sự kiện khi click vào các hình ảnh trả lời
				$myControl = $_MAIN_Pic_Control[$i]
				If $control = $myControl[4] Or $control = $myControl[0] Or $control = $myControl[1] Or $control = $myControl[2] Or $control = $myControl[3] Then
					If $_ANSWER == $_Question[$i - 1] Then ; nếu câu trả lời đúng
						_SetPic($myControl[1], @ScriptDir & '\Data\true.png')
						GUICtrlSetData($myControl[3], '')
						Sleep(1000)
						$_COUNT += 1
						;GUICtrlSetData($Question_Txt,'Question '&$_COUNT)
						_Set_Question() ; chuyển câu hỏi
					Else
						_SetPic($myControl[1], @ScriptDir & '\Data\false.png')
						GUICtrlSetData($myControl[3], '')
					EndIf
				EndIf
			Next
	EndSwitch
EndFunc   ;==>_ControlClick
Func _ControlHover($control)
	If $control <> $_LAST_HOVER Then
		If $_LAST_HOVER <> -1 Then _ControlNormal($_LAST_HOVER)
		Switch $control
			Case $Listen_Again
				GUICtrlSetColor($control, 0x0066cc)
			Case $Exit
				GUICtrlSetColor($control, 0xe50000)
			Case $Mini
				GUICtrlSetColor($control, 0xff7f00)
			Case Else
				For $i = 1 To $_NUMBER
					$myControl = $_MAIN_Pic_Control[$i]
					If $control = $myControl[4] Or $control = $myControl[0] Or $control = $myControl[1] Or $control = $myControl[2] Or $control = $myControl[3] Then
						GUICtrlSetBkColor($myControl[2], 0x004c99)
					EndIf
				Next
		EndSwitch
	EndIf
	$_LAST_HOVER = $control
EndFunc   ;==>_ControlHover
Func _ControlNormal($control)
	If $_LAST_HOVER <> -1 Then
		Switch $control
			Case $Listen_Again
				GUICtrlSetColor($control, 0x000000)
			Case $Exit
				GUICtrlSetColor($control, 0xFFFFFF)
			Case $Mini
				GUICtrlSetColor($control, 0xFFFFFF)
			Case Else
				For $i = 1 To $_NUMBER
					$myControl = $_MAIN_Pic_Control[$i]
					If $control = $myControl[4] Or $control = $myControl[0] Or $control = $myControl[1] Or $control = $myControl[2] Or $control = $myControl[3] Then
						GUICtrlSetBkColor($myControl[2], 0x198cff)
					EndIf
				Next
		EndSwitch
	EndIf
	$_LAST_HOVER = -1
EndFunc   ;==>_ControlNormal
#EndRegion Các hàm control click, hover
#Region Các hàm xử lý câu hỏi, câu trả lời
Func _Word_Check()
	$Form1 = GUICreate("Word check", 347, 343, 192, 124)
	$ListView1 = GUICtrlCreateListView("Word|Mean", 0, 0, 346, 342)
	GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 140)
	GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 200)
	_GUICtrlListView_SetExtendedListViewStyle($ListView1, BitOR($LVS_EX_GRIDLINES, $LVS_EX_CHECKBOXES))
	$read = IniReadSectionNames(@ScriptDir & '\Data\word.ini')
	For $i = 1 To $read[0]
		_GUICtrlListView_AddItem($ListView1, $read[$i], 0)
		_GUICtrlListView_AddSubItem($ListView1, $i - 1, IniRead(@ScriptDir & '\Data\word.ini', $read[$i], 'd', ''), 1)
		If IniRead(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN', $read[$i], '') = $read[$i] Then _GUICtrlListView_SetItemChecked($ListView1, $i - 1)
	Next

	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				IniDelete(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN')
				For $i = 1 To $read[0]
					If _GUICtrlListView_GetItemChecked($ListView1, $i - 1) Then IniWrite(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN', $read[$i], $read[$i])
				Next
				GUIDelete($Form1)
				ExitLoop

		EndSwitch
	WEnd
EndFunc   ;==>_Word_Check
Func _Get_Excel_List()
	Local $oExcel = _Excel_Open(False)
	Local $oWorkbook = _Excel_BookOpen($oExcel, @ScriptDir & '\Data\word.xls')
	$i = 0
	If Not FileExists(@ScriptDir & '\Data\word.ini') Then
		FileDelete(@ScriptDir & '\Data\word.ini')
		$file = FileOpen(@ScriptDir & '\Data\word.ini', 34)
		FileWrite($file, '// coded by opdo.vn')
		FileClose($file)
	EndIf
	$aResult = _Excel_RangeRead($oWorkbook)
	For $i = 0 To UBound($aResult) - 1
		IniWrite(@ScriptDir & '\Data\word.ini', $aResult[$i][0], 'b', $aResult[$i][1])
		IniWrite(@ScriptDir & '\Data\word.ini', $aResult[$i][0], 'c', $aResult[$i][2])
		IniWrite(@ScriptDir & '\Data\word.ini', $aResult[$i][0], 'd', $aResult[$i][3])
	Next
	$read = IniReadSectionNames(@ScriptDir & '\Data\word.ini')
	For $i = 1 To $read[0]
		If $read[$i] == 'PROGRAMWORDMEAN' Then ContinueLoop
		If Not FileExists(@ScriptDir & '\Data\Sounds\' & $read[$i] & '.mp3') Or Not FileExists(@ScriptDir & '\Data\Images\' & $read[$i] & '.jpg') Then IniDelete(@ScriptDir & '\Data\word.ini', $read[$i])
	Next
	_Excel_Close($oExcel)
EndFunc   ;==>_Get_Excel_List
Func _Random_List(ByRef $List, $number) ; random câu hỏi + câu trả lời
	If UBound($List) < $number Then ; nếu ko đủ từ vựng để rando,
		MsgBox(16, UBound($List), "Error: Less than " & $number & " items.")
		Exit
	EndIf
	Local $List_temp = $List ; tạo mảng khác để lưu trữ list từ vựng
	Local $Return[$number] ; mảng trả về các từ được random
	For $i = 1 To $number
		Local $random = Random(0, UBound($List_temp) - 1, 1)
		$Return[$i - 1] = $List_temp[$random]
		_ArrayDelete($List_temp, $random) ; xóa bỏ từ đã được đưa vào mảng trả về
	Next
	Return $Return
EndFunc   ;==>_Random_List
Func _Get_List() ; lấy danh sách từ vựng
	#
	$hSearch = FileFindFirstFile(@ScriptDir & "\Data\Sounds\*.mp3") ; lấy danh sách từ các file sound mp3
	If $hSearch = -1 Then
		MsgBox(16, "ERROR", "Error: No sound files exist in Data\Sounds.")
		Exit
	EndIf
	While 1
		$sFileName = FileFindNextFile($hSearch)
		If @error Then ExitLoop
		If @extended Then ContinueLoop
		_ArrayAdd($myNewList, StringTrimRight($sFileName, 4))
	WEnd
	Local $List[0]
	Local $read = IniReadSection(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN')
	If Not @error Then
		For $i = 1 To $read[0][0]
			_ArrayAdd($List, $read[$i][0])
		Next
	EndIf
	Return $List
EndFunc   ;==>_Get_List
Func _Set_Question() ; set câu hỏi, set các picture hiển thị
	Global $_Question = _Random_List($myNewList, $_NUMBER) ; lấy danh sách các từ random
	If Not IsArray($_List) Then
		$_ANSWER = $myNewList[Random(0, UBound($myNewList) - 1, 1)]
	Else
		If UBound($_List) < 1 Then
			$_ANSWER = $myNewList[Random(0, UBound($myNewList) - 1, 1)]
		Else
			If Random(1, 100, 1) > 75 Then
				$_ANSWER = $myNewList[Random(0, UBound($myNewList) - 1, 1)]
			Else
				$_ANSWER = $_List[Random(0, UBound($_List) - 1, 1)]
			EndIf
		EndIf
	EndIf
	
	$_Question[Random(0, UBound($_Question) - 1, 1)] = $_ANSWER
	If IniRead(@ScriptDir & '\Data\word.ini', 'PROGRAMWORDMEAN', $_ANSWER, '') = $_ANSWER Then
		GUICtrlSetColor($Favorite, 0xFF0000)
	Else
		GUICtrlSetColor($Favorite, 0x808080)
	EndIf
	;$Favorite

	
	_SoundClose($_SOUND)
	$_SOUND = _SoundOpen(@ScriptDir & "\Data\Sounds\" & $_ANSWER & '.mp3') ; mở âm thanh
	If GUICtrlRead($Listen_Again) == "Show word's mean" Then
		GUICtrlSetData($Question_Txt, 'Question ' & $_COUNT)
	Else
		GUICtrlSetData($Question_Txt, $_ANSWER & ' \' & IniRead(@ScriptDir & '\Data\word.ini', $_ANSWER, 'b', '') & '\ (' & IniRead(@ScriptDir & '\Data\word.ini', $_ANSWER, 'c', '') & ') : ' & IniRead(@ScriptDir & '\Data\word.ini', $_ANSWER, 'd', ''))
	EndIf

	For $i = 0 To UBound($_Question) - 1 ; set hình ảnh và một số control khác
		Local $control = $_MAIN_Pic_Control[$i + 1]
		GUICtrlSetData($control[3], '')
		GUICtrlSetBkColor($control[2], 0x198cff)
		Local $mypos = ControlGetPos($GUIMAIN, '', $control[4])
		GUICtrlSetPos($control[0], $mypos[0], $mypos[1], $mypos[2], $mypos[3])
		_SetPic($control[1], '')
		_SetPic($control[0], @ScriptDir & '\Data\Images\' & $_Question[$i] & '.jpg', 1)
	Next
	_SoundPlay($_SOUND) ; chạy âm thanh
	$_LISTEN_FLAG = True
EndFunc   ;==>_Set_Question
Func _Create_Picture($number, $w) ; tạo các control để hiển thị các đáp án trả lời
	GUISwitch($GUIMAIN)
	Local $y = 145
	For $i = 0 To $number - 1
		Local $control[5]
		If $i <> 0 And Mod($i, 5) = 0 Then $y += ($w + 13)
		$control[2] = GUICtrlCreateLabel("", 23 + Mod($i, 5) * ($w + 13), $y - 2, $w + 4, $w + 4) ; label khung nền
		GUICtrlSetBkColor(-1, 0x0066cc)
		GUICtrlSetCursor(-1, 0)
		$control[4] = GUICtrlCreateLabel("", 25 + Mod($i, 5) * ($w + 13), $y, $w, $w) ; label khung nền
		GUICtrlSetBkColor(-1, 0xFFFFFF)
		GUICtrlSetCursor(-1, 0)

		$control[0] = GUICtrlCreatePic("", 25 + Mod($i, 5) * ($w + 13), $y, $w, $w) ; pic chính hiển thị
		GUICtrlSetResizing(-1, 1)
		GUICtrlSetCursor(-1, 0)
		$control[1] = GUICtrlCreatePic("", 23 + Mod($i, 5) * ($w + 13), $y - 2, $w + 4, $w + 4) ; pic hiển thị màu đè lên khi trả lời
		GUICtrlSetResizing(-1, 1)
		GUICtrlSetCursor(-1, 0)
		$control[3] = GUICtrlCreateLabel("", 25 + Mod($i, 5) * ($w + 13), $y, $w, $w, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1) ; hiển thị icon đúng sai
		GUICtrlSetFont(-1, 40, 400, 0, "FontAwesome")
		GUICtrlSetColor(-1, "0xFFFFFF")
		GUICtrlSetBkColor(-1, "-2")
		GUICtrlSetCursor(-1, 0)

		_ArrayAdd($_MAIN_Pic_Control, '') ; add control vào mảng pic_control
		$_MAIN_Pic_Control[0] += 1
		$_MAIN_Pic_Control[$_MAIN_Pic_Control[0]] = $control
	Next
EndFunc   ;==>_Create_Picture
#EndRegion Các hàm xử lý câu hỏi, câu trả lời
#Region Các hàm xử lý hình ảnh
Func _SetPic($control, $pic, $flag = 0) ; hàm set hình ảnh, thay thế _setimage đang lỗi
	Local $pos = ControlGetPos($GUIMAIN, '', $control)
	_SetImage($control, $pic)
	If $pic <> '' And $flag = 1 Then
		$size = _GetSize($pic)
		_Scale_Pic($size, $pos[2])
		GUICtrlSetPos($control, Default, Default, $pos[2], $pos[3] + 1)
		If $size[0] == $pos[2] Then
			GUICtrlSetPos($control, Default, $pos[1] + Int(($pos[2] - $size[1]) / 2), $size[0], $size[1])
		Else
			GUICtrlSetPos($control, $pos[0] + Int(($pos[2] - $size[0]) / 2), Default, $size[0], $size[1])
		EndIf
	Else
		GUICtrlSetPos($control, Default, Default, $pos[2], $pos[3] + 1)
		GUICtrlSetPos($control, Default, Default, $pos[2], $pos[3])
	EndIf
EndFunc   ;==>_SetPic
Func _GetSize($file) ; lấy kích thước của ảnh
	_GDIPlus_Startup()

	Local $hImage = _GDIPlus_ImageLoadFromFile($file)
	If @error Then
		MsgBox(16, "Error", "Does the file " & $file & " exist?")
		Exit 1
	EndIf
	Local $size[2] = [_GDIPlus_ImageGetWidth($hImage), _GDIPlus_ImageGetHeight($hImage)]
	_GDIPlus_ImageDispose($hImage)
	_GDIPlus_Shutdown()
	Return $size
EndFunc   ;==>_GetSize
Func _Scale_Pic(ByRef $size, $scale) ; scale ảnh
	If $size[0] > $size[1] Then
		$size[1] = Round(($size[1] * $scale) / $size[0])
		$size[0] = $scale
	Else
		$size[0] = Round(($size[0] * $scale) / $size[1])
		$size[1] = $scale
	EndIf
EndFunc   ;==>_Scale_Pic
#EndRegion Các hàm xử lý hình ảnh
Func _Choose_Level() ; chọn level
	Global $LEVELGUI = GUICreate("GUIMAIN", 285, 251, -1, -1, BitOR($WS_POPUP, $WS_BORDER), -1)
	GUISetBkColor(0x0066cc, $LEVELGUI)
	$Level_1 = GUICtrlCreateLabel(" Easy (10 Pictures)", 43, 57, 197, 34, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
	GUICtrlSetState(-1, $GUI_SHOW)
	GUICtrlSetFont(-1, 15, 400, 0, "FontAwesome")
	GUICtrlSetColor(-1, "0xFFFFFF")
	GUICtrlSetBkColor(-1, "-2")
	GUICtrlSetCursor(-1, 0)
	GUICtrlCreateLabel("Choose level", 43, 7, 197, 43, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
	GUICtrlSetState(-1, $GUI_SHOW)
	GUICtrlSetFont(-1, 16, 350, 0, "Segoe UI Semilight")
	GUICtrlSetColor(-1, "0xFFFFFF")
	GUICtrlSetBkColor(-1, "-2")
	$Level_2 = GUICtrlCreateLabel(" Normal (15 Pictures)", 43, 97, 197, 34, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
	GUICtrlSetState(-1, $GUI_SHOW)
	GUICtrlSetFont(-1, 15, 400, 0, "FontAwesome")
	GUICtrlSetColor(-1, "0xFFFFFF")
	GUICtrlSetBkColor(-1, "-2")
	GUICtrlSetCursor(-1, 0)
	$Level_3 = GUICtrlCreateLabel(" Hard (20 Pictures)", 43, 134, 197, 34, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
	GUICtrlSetState(-1, $GUI_SHOW)
	GUICtrlSetFont(-1, 15, 400, 0, "FontAwesome")
	GUICtrlSetColor(-1, "0xFFFFFF")
	GUICtrlSetBkColor(-1, "-2")
	GUICtrlSetCursor(-1, 0)
	$Level_4 = GUICtrlCreateLabel(" Favorite", 43, 168, 197, 34, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
	GUICtrlSetState(-1, $GUI_SHOW)
	GUICtrlSetFont(-1, 15, 400, 0, "FontAwesome")
	GUICtrlSetColor(-1, "0xFFFFFF")
	GUICtrlSetBkColor(-1, "-2")
	GUICtrlSetCursor(-1, 0)
	$Level_Exit = GUICtrlCreateLabel(" Exit", 43, 202, 197, 34, BitOR($SS_CENTER, $SS_CENTERIMAGE), -1)
	GUICtrlSetState(-1, $GUI_SHOW)
	GUICtrlSetFont(-1, 15, 400, 0, "FontAwesome")
	GUICtrlSetColor(-1, "0xFFFFFF")
	GUICtrlSetBkColor(-1, "-2")
	GUICtrlSetCursor(-1, 0)
	GUISetState(@SW_SHOW, $LEVELGUI)


	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $Level_1, $Level_2, $Level_3
				$text = GUICtrlRead($nMsg)
				$s1 = StringRegExp($text, "\d+", 3)
				GUIDelete($LEVELGUI)
				Return Number($s1[0])
			Case $Level_4
				_Word_Check()
			Case $GUI_EVENT_CLOSE, $Level_Exit
				Exit

		EndSwitch
	WEnd
EndFunc   ;==>_Choose_Level
