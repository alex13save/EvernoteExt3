#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=AutoIt_Main_v11_256x256_RGB-A.ico
#AutoIt3Wrapper_Res_Comment=Вставка даты и форматирование в Evernote
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Perfiliev Alex

 Script Function:
	Extension Evernote functions v.3.1

#ce ----------------------------------------------------------------------------

#include <Misc.au3>
#Include <Constants.au3>
#include <Date.au3>

#include <DateTimeConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>

Opt("SendKeyDelay",50)
$TL = "[CLASS:ENMainFrame]"

HotKeySet("^+л","RedKursiv")	;Ctrl+Shift+K - красный, подчеркнутфй курсив (выделенный текст)
HotKeySet("^+в","RedDateTime")	;Ctrl+Shift+D - текущая дата и время красным цветом (в поз. курсора)
HotKeySet("^+ь","GreyDone")		;Ctrl+Shift+M - серый, перечеркнутый (выделенный текст)
HotKeySet("^!в","RedDateWeek")	;Ctrl+Alt+D - выбирает и выводит в позицию курсора дату и день недели

Opt("TrayMenuMode",1)
TraySetState()
$trayhint = "EvernoteExt3 - доп. инструменты для Evernote"
TraySetToolTip($trayhint)

$helpitem = TrayCreateItem("Помощь")
TrayCreateItem("")
$exititem = TrayCreateItem("Выход")

while 1
   Sleep(10)

    $msg = TrayGetMsg()
    Select
        Case $msg = 0
            ContinueLoop
        Case $msg = $helpitem

			ShowHelpWindows()

        Case $msg = $exititem
            ExitLoop
    EndSelect

WEnd

Func ShowHelpWindows()
#Region ### START Koda GUI section ### Form=form1.kxf
$Form1_1 = GUICreate("Дополнительные инструменты для Evernote", 650, 130, -1, -1)
$Edit1 = GUICtrlCreateEdit("", 8, 8, 633, 113, BitOR($ES_AUTOVSCROLL,$ES_AUTOHSCROLL,$ES_WANTRETURN))
GUICtrlSetData(-1, StringFormat("Горячие клавиши:\r\n\r\n• Ctrl+Alt+D   - выбрать и вывести дату и день недели\r\n• Ctrl+Shift+D - вывести текущую дату и время\r\n• Ctrl+Shift+K - применить к выделенному тексту стиль "&Chr(34)&"Красный, подчеркнутый, кусив"&Chr(34)&"\r\n• Ctrl+Shift+M - применить к выделенному тексту стиль "&Chr(34)&"Серый, перечеркнутый"&Chr(34)&""))
GUICtrlSetFont(-1, 9, 400, 0, "Courier New")
GUICtrlSetState(-1, $GUI_DISABLE)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			GUIDelete();
			ExitLoop

	EndSwitch
WEnd


EndFunc

Func RedDateWeek()
	if WinExists($TL) Then
		$rez = WinWaitActive($TL,"",3)
		if $rez=0 then
			Return
		EndIf

		if CtrlMButton() Then
			Return
		EndIf

		Send("+{HOME}")
		Send("^ш")	;Курсив
		Send("^г")	;Подчеркнутый
		Sleep(50)
		SetColorFont(9)
		Send("{RIGHT}")

	EndIf
EndFunc

Func CtrlMButton()

	Local $aMyDate, $aMyTime

	$hWnd = WinGetHandle("[ACTIVE]")

	$PreselectDate = @YEAR & "/" & @MON & "/" & @MDAY

	$pos = MouseGetPos()

	#Region ### START Koda GUI section ### Form=D:\AutoIt\InsDate\FormCalendar.kxf

	$Form1_1 = GUICreate("Выбери", 182, 179, $pos[0], $pos[1], BitOR($GUI_SS_DEFAULT_GUI,$DS_MODALFRAME), BitOR($WS_EX_TOOLWINDOW,$WS_EX_TOPMOST,$WS_EX_WINDOWEDGE))
	$MonthCal1 = GUICtrlCreateMonthCal($PreselectDate, 8, 8, 166, 164)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				$CloseCal = True
				ExitLoop

			Case $MonthCal1
				$CloseCal = False
				ExitLoop
		EndSwitch
	WEnd

	$SelDate = GUICtrlRead($MonthCal1)
	$rez = _DateTimeFormat($SelDate, 2)

	GUIDelete();

	If $CloseCal Then
		Return $CloseCal
	EndIf

	_DateTimeSplit($SelDate, $aMyDate, $aMyTime)

	$DOW = _DateToDayOfWeek($aMyDate[1], $aMyDate[2], $aMyDate[3])

	Select
		Case $DOW = 1
			$sShortDayName = "Вс"
		Case $DOW = 2
			$sShortDayName = "Пн"
		Case $DOW = 3
			$sShortDayName = "Вт"
		Case $DOW = 4
			$sShortDayName = "Ср"
		Case $DOW = 5
			$sShortDayName = "Чт"
		Case $DOW = 6
			$sShortDayName = "Пт"
		Case $DOW = 7
			$sShortDayName = "Сб"
	EndSelect

	WinActivate($hWnd)

	Opt("SendKeyDelay", 1)
	_SendEx($rez & ", " & $sShortDayName)
	Opt("SendKeyDelay", 50)

	Return $CloseCal

EndFunc


Func GreyDone()
   if WinExists($TL) Then
		$rez = WinWaitActive($TL,"",3)
		if $rez=0 then
			Return
		EndIf
		Send("^е")
		Sleep(500)
		SetColorFont(8)
		Send("{RIGHT}")
   EndIf
EndFunc

Func RedDateTime()
   if WinExists($TL) Then
		$rez = WinWaitActive($TL,"",3)
		if $rez=0 then
			Return
		EndIf
	  Send("^ж")
	  Send("+{HOME}")
	  SetColorFont(9)
	  Send("{RIGHT}{DOWN}")
   EndIf
EndFunc

Func RedKursiv()
   if WinExists($TL) Then
		$rez = WinWaitActive($TL,"",3)
		if $rez=0 then
			Return
		EndIf
	  Send("^ш")	;Курсив
	  Send("^г")	;Подчеркнутый
	  Sleep(50)
	  SetColorFont(9)
 	  Send("{RIGHT}")
   EndIf
EndFunc

;Меняет цвет выделенного текста через диалог "Шрифт"
;$n - число смещений (9-красный, 8-серебристый, 0-черный)
Func SetColorFont($n)
  _SendExEx("^в")	;Ctrl+D - вызов диалога
  $hwnd = WinWaitActive("[CLASS:#32770]","",5) ;Диалог Шрифт
   if $hWnd=0 Then
	  Return
   EndIf
	Sleep(50)

	Send("{TAB 5}")
	Send("{HOME}")
	Send("{DOWN " & $n & "}")

   Send("{ENTER}")
EndFunc

;Автор: CreatoR
;Интерпритация на функцию Send(), только с использованием б.обмена - обход проблемы с кодировками
Func _SendEx($sString)
Local $sOld_Clip = ClipGet()

ClipPut($sString)
Sleep(10)
Send("+{INSERT}")

ClipPut($sOld_Clip)
EndFunc

;Автор: CreatoR
;Обход проблемы с отправкой нажатии клавиш в русской раскладке клавиатуры
Func _SendExEx($sKeys, $iFlag=0)
If @KBLayout = 0419 Then
Local $sANSI_Chars = "ёйцукенгшщзхъфывапролджэячсмитьбю.?"
Local $sASCII_Chars = "`qwertyuiop[]asdfghjkl;'zxcvbnm,./&"

Local $aSplit_Keys = StringSplit($sKeys, "")
Local $sKey
$sKeys = ""

For $i = 1 To $aSplit_Keys[0]
$sKey = StringMid($sANSI_Chars, StringInStr($sASCII_Chars, $aSplit_Keys[$i]), 1)

If $sKey <> "" Then
$sKeys &= $sKey
Else
$sKeys &= $aSplit_Keys[$i]
EndIf
Next
EndIf

Return Send($sKeys, $iFlag)
EndFunc
