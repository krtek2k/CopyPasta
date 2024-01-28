; Script:    CopyPasta.ahk
; License:   The Unlicense
; Author:    krtek2k
; Github:    https://github.com/krtek2k/CopyPasta
; Date:      2024-01-28
; Version:   1.0

/*
 * Informs about clipboard events using tooltips at caret or at mouse
 * Prevents blank Copy&Pasta and repeating the process after nothing was copied and pasted
 * - informs about sucessful Copy&Pasta
 * - informs about wrong CTRL+C combinations such as LAlt+C or LWin+C
 * - informs about CRLF (Carriage Return/Line Feed) as an empty Copy&Pasta
 * - informs about empty Copy&Pasta or A_Tab or up to 3 Space 
 */

#Requires AutoHotkey v2.0
#SingleInstance Force

CoordMode("ToolTip", "Screen")

; Ctrl=^
$^c::{
	A_Clipboard := ""
	Send "^c"
	HandleCopyPasta()
}

; Ctrl=^
$^x::{
	A_Clipboard := ""
	Send "^x"
	HandleCopyPasta()
}

HandleCopyPasta() {
	if ClipWait(1, 1) {
		CopyPasta(CopyPastaEvent()).ShowTooltip()
	}
	else {
		CopyPasta(CopyPastaLongerThanExpectedEvent()).ShowTooltip()
		
		if ClipWait((CopyPasta.Settings.ttip_cooldown_ms+250)/1000, 1) {
			CopyPasta(CopyPastaEvent()).ShowTooltip()
		}
		else {
			CopyPasta(CopyPastaWaitingEvent()).ShowTooltip()
			
			if ClipWait(10, 1) {
				CopyPasta(CopyPastaEvent()).ShowTooltip()
			}
			else {
				CopyPasta(CopyPastaTimeoutEvent()).ShowTooltip()
			}
		}
	}
}

; FN key - vkFF sc163 not found 
$vkFF::{ 	
	HotIf (*) => GetKeyState("vkFF", "P")
	Hotkey "*c",(*) => CopyPasta(CopyPastaFnEvent()).ShowTooltip()
	KeyWait "vkFF"
}

; LWin=#
$#c::{
	CopyPasta(CopyPastaLWinEvent()).ShowTooltip()
	Send "#c"
}

; Alt=! 
$!c::{
	CopyPasta(CopyPastaLAltEvent()).ShowTooltip()
	Send "!c"
}

class CopyPasta {
		
	class Settings {
	
		static cb_buffer_obj_size_as_empty := 32 ; ClipboardAll().Size - 3 spaces take 32 as well as one A_Tab, anything bigger in buffer consider as data if not text
		
		static ttip_cooldown_ms := 1300
		static ttip_cooldown_error_multiplier := 2
		static ttip_offset := 10
		
		static ttip_msg_clip_lenght_limit := 200
		
		static ttip_msg_app := "COPYPASTA!"
		static ttip_msg_line_icon := "✔"
		static ttip_msg_error_line_icon := "❌"
		static ttip_msg_error_line_full := "{1} {1} {1} {1} {1} {1} {1} {1} {1}"
		static ttip_msg_error_line_side := "{1}                                    {1}"
		static ttip_msg_error_line_app := "{1}       {2}       {1}"
		static ttip_msg_error_line_emoji_shrug := "{1}         ¯\_( •︠ ͜ʖ ︡•)_/¯       {1}"
		static ttip_msg_error_line_err_empty := "{1}            EMPTY            {1}"
		static ttip_msg_error_line_err_fn := "{1}             Fn+C             {1}"
		static ttip_msg_error_line_err_win := "{1}             Win+C           {1}"
		static ttip_msg_error_line_err_alt := "{1}              Alt+C            {1}"
		static ttip_msg_error_line_err_longer_than_expected := "{1}      HMMMMM...      {1}"
		static ttip_msg_error_line_err_waiting := "{1}⏳ WAITING...(10s) ⏳ {1}"
		static ttip_msg_error_line_err_timeout := "{1}     🏁 TIMEOUT 🏁     {1}"
	}
	
	__Init(){
		ToolTip() ; off all tooltips
	}
	__New(copyPastaEvent) {
		this.Event := copyPastaEvent
	}
	__Delete() {
		this.QueueTooltipDissmiss(this.Event.IsError ? CopyPasta.Settings.ttip_cooldown_ms * CopyPasta.Settings.ttip_cooldown_error_multiplier : CopyPasta.Settings.ttip_cooldown_ms)
	}
	
	ShowTooltip() {
		if hwnd := GetCaretPosEx(&x, &y, &w, &h){
			ToolTip(this.Event.Message, x + CopyPasta.Settings.ttip_offset, y + h + CopyPasta.Settings.ttip_offset)
		}
		else {
			ToolTip(this.Event.Message)
		}
	}
	
	QueueTooltipDissmiss(periodMs) {
		SetTimer () => (ToolTip()), -periodMs  
	}
}

class CopyPastaEventBase {

	IsError => true
	
	BuildMessage(msg) {
		if (this.IsError) {
			return (
				Format(CopyPasta.Settings.ttip_msg_error_line_full, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_side, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_app, CopyPasta.Settings.ttip_msg_error_line_icon, CopyPasta.Settings.ttip_msg_app)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_side, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				msg
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_emoji_shrug, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_side, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_full, CopyPasta.Settings.ttip_msg_error_line_icon)
			)
		}
		else {
			return (
				CopyPasta.Settings.ttip_msg_line_icon
				" "
				CopyPasta.Settings.ttip_msg_app
				"`n" msg
			)
		}
	}
}

class CopyPastaFnEvent extends CopyPastaEventBase {
	Message => this.BuildMessage(Format(CopyPasta.Settings.ttip_msg_error_line_err_fn, CopyPasta.Settings.ttip_msg_error_line_icon))
}
class CopyPastaLWinEvent extends CopyPastaEventBase {
	Message => this.BuildMessage(Format(CopyPasta.Settings.ttip_msg_error_line_err_win, CopyPasta.Settings.ttip_msg_error_line_icon))
}
class CopyPastaLAltEvent extends CopyPastaEventBase  {
	Message => this.BuildMessage(Format(CopyPasta.Settings.ttip_msg_error_line_err_alt, CopyPasta.Settings.ttip_msg_error_line_icon))
}
class CopyPastaLongerThanExpectedEvent extends CopyPastaEventBase  {
	Message => this.BuildMessage(Format(CopyPasta.Settings.ttip_msg_error_line_err_longer_than_expected, CopyPasta.Settings.ttip_msg_error_line_icon))
}
class CopyPastaWaitingEvent extends CopyPastaEventBase  {
	Message => this.BuildMessage(Format(CopyPasta.Settings.ttip_msg_error_line_err_waiting, CopyPasta.Settings.ttip_msg_error_line_icon))
}
class CopyPastaTimeoutEvent extends CopyPastaEventBase  {
	Message => this.BuildMessage(Format(CopyPasta.Settings.ttip_msg_error_line_err_timeout, CopyPasta.Settings.ttip_msg_error_line_icon))
}

class CopyPastaEvent extends CopyPastaEventBase  {

	IsError
	{
		get => this._isError
	}
	
	Message
	{
		get {
			if (this._isError) {
				return this.BuildMessage(Format(CopyPasta.Settings.ttip_msg_error_line_err_empty, CopyPasta.Settings.ttip_msg_error_line_icon))
			}
			return this.BuildMessage(this._message)
		}
		set => this._message := value
	}
	
	__New() {
		; trim and shorten for tooltip message
		this._message  := SubStr(
			StrReplace(Trim(A_Clipboard), A_Tab), 
			1, 
			CopyPasta.Settings.ttip_msg_clip_lenght_limit
		)
		
		this._isError := (ClipboardAll().Size <= CopyPasta.Settings.cb_buffer_obj_size_as_empty && (StrLen(A_Clipboard) <= 3 && (this._message = "" || this._message = "`n" || this._message = "`r" || this._message = "`r`n"))) ? true : false 
	}
}

GetCaretPosEx(&x?, &y?, &w?, &h?) {
    #Requires AutoHotkey v2.0-rc.1 64-bit
    x := h := w := h := 0
    static iUIAutomation := 0, hOleacc := 0, IID_IAccessible, guiThreadInfo, _ := init()
    if !iUIAutomation || ComCall(8, iUIAutomation, "ptr*", eleFocus := ComValue(13, 0), "int") || !eleFocus.Ptr
        goto useAccLocation
    if !ComCall(16, eleFocus, "int", 10002, "ptr*", valuePattern := ComValue(13, 0), "int") && valuePattern.Ptr
        if !ComCall(5, valuePattern, "int*", &isReadOnly := 0) && isReadOnly
            return 0
    useAccLocation:
    ; use IAccessible::accLocation
    hwndFocus := DllCall("GetGUIThreadInfo", "uint", DllCall("GetWindowThreadProcessId", "ptr", WinExist("A"), "ptr", 0, "uint"), "ptr", guiThreadInfo) && NumGet(guiThreadInfo, 16, "ptr") || WinExist()
    if hOleacc && !DllCall("Oleacc\AccessibleObjectFromWindow", "ptr", hwndFocus, "uint", 0xFFFFFFF8, "ptr", IID_IAccessible, "ptr*", accCaret := ComValue(13, 0), "int") && accCaret.Ptr {
        NumPut("ushort", 3, varChild := Buffer(24, 0))
        if !ComCall(22, accCaret, "int*", &x := 0, "int*", &y := 0, "int*", &w := 0, "int*", &h := 0, "ptr", varChild, "int")
            return hwndFocus
    }
    if iUIAutomation && eleFocus {
        ; use IUIAutomationTextPattern2::GetCaretRange
        if ComCall(16, eleFocus, "int", 10024, "ptr*", textPattern2 := ComValue(13, 0), "int") || !textPattern2.Ptr
            goto useGetSelection
        if ComCall(10, textPattern2, "int*", &isActive := 0, "ptr*", caretTextRange := ComValue(13, 0), "int") || !caretTextRange.Ptr || !isActive
            goto useGetSelection
        if !ComCall(10, caretTextRange, "ptr*", &rects := 0, "int") && rects && (rects := ComValue(0x2005, rects, 1)).MaxIndex() >= 3 {
            x := rects[0], y := rects[1], w := rects[2], h := rects[3]
            return hwndFocus
        }
        useGetSelection:
        ; use IUIAutomationTextPattern::GetSelection
        if textPattern2.Ptr
            textPattern := textPattern2
        else if ComCall(16, eleFocus, "int", 10014, "ptr*", textPattern := ComValue(13, 0), "int") || !textPattern.Ptr
            goto useGUITHREADINFO
        if ComCall(5, textPattern, "ptr*", selectionRangeArray := ComValue(13, 0), "int") || !selectionRangeArray.Ptr
            goto useGUITHREADINFO
        if ComCall(3, selectionRangeArray, "int*", &length := 0, "int") || length <= 0
            goto useGUITHREADINFO
        if ComCall(4, selectionRangeArray, "int", 0, "ptr*", selectionRange := ComValue(13, 0), "int") || !selectionRange.Ptr
            goto useGUITHREADINFO
        if ComCall(10, selectionRange, "ptr*", &rects := 0, "int") || !rects
            goto useGUITHREADINFO
        rects := ComValue(0x2005, rects, 1)
        if rects.MaxIndex() < 3 {
            if ComCall(6, selectionRange, "int", 0, "int") || ComCall(10, selectionRange, "ptr*", &rects := 0, "int") || !rects
                goto useGUITHREADINFO
            rects := ComValue(0x2005, rects, 1)
            if rects.MaxIndex() < 3
                goto useGUITHREADINFO
        }
        x := rects[0], y := rects[1], w := rects[2], h := rects[3]
        return hwndFocus
    }
    useGUITHREADINFO:
    if hwndCaret := NumGet(guiThreadInfo, 48, "ptr") {
        if DllCall("GetWindowRect", "ptr", hwndCaret, "ptr", clientRect := Buffer(16)) {
            w := NumGet(guiThreadInfo, 64, "int") - NumGet(guiThreadInfo, 56, "int")
            h := NumGet(guiThreadInfo, 68, "int") - NumGet(guiThreadInfo, 60, "int")
            DllCall("ClientToScreen", "ptr", hwndCaret, "ptr", guiThreadInfo.Ptr + 56)
            x := NumGet(guiThreadInfo, 56, "int")
            y := NumGet(guiThreadInfo, 60, "int")
            return hwndCaret
        }
    }
    return 0
    static init() {
        try
            iUIAutomation := ComObject("{E22AD333-B25F-460C-83D0-0581107395C9}", "{30CBE57D-D9D0-452A-AB13-7AC5AC4825EE}")
        hOleacc := DllCall("LoadLibraryW", "str", "Oleacc.dll", "ptr")
        NumPut("int64", 0x11CF3C3D618736E0, "int64", 0x719B3800AA000C81, IID_IAccessible := Buffer(16))
        guiThreadInfo := Buffer(72), NumPut("uint", guiThreadInfo.Size, guiThreadInfo)
    }
}