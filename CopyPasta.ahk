; Script:    CopyPasta.ahk
; License:   The Unlicense
; Author:    krtek2k
; Github:    https://github.com/krtek2k/CopyPasta
; Date:      2024-01-28
; Version:   1.2.2

/*
 * Informs about blank Copy&Pasta CTRL+C and prevents, by this, repeating the process of copying nothing in the clipboard
 * Basically just shows tooltip at caret or at mouse. Does not interrupt or change any thing.
 * - informs about sucessful Copy&Pasta CTRL+C
 * - informs about wrong CTRL+C combinations such as LAlt+C or LWin+C
 * - informs about empty Copy&Pasta CTRL+C or A_Tab or up to 3 Space
 * - informs about CRLF (Carriage Return/Line Feed) as an empty Copy&Pasta
 * - When Copy&Pasta CTRL+C process take too long, there is a little information flow with timeout
 */

#Requires AutoHotkey v2.0-rc.1 64-bit
#SingleInstance Force
if not (A_IsAdmin or RegExMatch(DllCall("GetCommandLine", "str"), " /restart(?!\S)")) { ; needs to be run as admin, copypaste not working in apps that are in administrator mode 
  try { ; try to run as an administrator, if could not be run as an administrator, it will not work with apps run as an administrator.
    if A_IsCompiled
      Run '*RunAs "' A_ScriptFullPath '" /restart'
    else
      Run '*RunAs "' A_AhkPath '" /restart "' A_ScriptFullPath '"'
  }
}
CoordMode("ToolTip", "Screen")
DEBUG_MODE := false

class CopyPasta {
		
	class Settings {
	
		static ttip_cooldown_ms := 1500
		static ttip_cooldown_error_multiplier := 1.5
		static ttip_offset := 10
		
		static ttip_msg_clip_lenght_limit := 200
		
		static ttip_msg_line_icon := "⎀✔"
		static ttip_msg_line_success := "{1}"
		static ttip_msg_error_line_icon := "❌"
		static ttip_msg_error_line_full := "{1} {1} {1} {1} {1} {1} {1} {1} {1}"
		static ttip_msg_error_line_side := "{1}                                    {1}"
		static ttip_msg_error_line_app := "{1}       {2}       {1}"
		static ttip_msg_error_line_emoji_shrug := "{1}             ( •︠ ͜ʖ ︡•)             {1}"
		static ttip_msg_error_line_err_empty := "{1}             EMPTY           {1}"
		static ttip_msg_error_line_err_fn := "{1}             Fn+C             {1}"
		static ttip_msg_error_line_err_win := "{1}             Win+C           {1}"
		static ttip_msg_error_line_err_alt := "{1}              Alt+C            {1}"
		static ttip_msg_error_line_err_longer_than_expected := "{1}      HMMMMM...      {1}"
		static ttip_msg_error_line_err_waiting := "{1} ⏳ WAITING....(5s) ⏳ {1}"
		static ttip_msg_error_line_err_timeout := "{1}     🏁 TIMEOUT 🏁     {1}"
		static ttip_msg_debug_line_full := "🛠 DEBUG_MODE 🛠"
		static ttip_msg_debug_line_duplicate := "{1}   ALREADY COPIED  {1}"
	}
	
	__Init(){
		this.QueueTooltipDissmiss(0) ; off all tooltips
	}
	__New(copyPastaEvent) {
		this.Event := copyPastaEvent
	}
	__Delete() {
		calcDissmissTimer := CopyPasta.Settings.ttip_cooldown_ms
		if (this.Event.IsError)
			calcDissmissTimer *= CopyPasta.Settings.ttip_cooldown_error_multiplier 
		if (DEBUG_MODE)
			calcDissmissTimer += 2000
		this.QueueTooltipDissmiss(calcDissmissTimer)
		global CbOrigCb := ""
		global CbOrigBufferSize := 0
	}
	
	ShowTooltip() {
		if (DEBUG_MODE)
			return this
		if hwnd := GetCaretPosEx(&x, &y, &w, &h){
			ToolTip(this.Event.Message, x + CopyPasta.Settings.ttip_offset, y + h + CopyPasta.Settings.ttip_offset)
		}
		else {
			ToolTip(this.Event.Message)
		}
		return this
	}
	
	ShowDebug(callStack) {
		if (!DEBUG_MODE)
			return this
			
		this.QueueTooltipDissmiss(0)
		if hwnd := GetCaretPosEx(&x, &y, &w, &h){
			ToolTip(
				this.BuildDebugMessage(callStack)
				, x + CopyPasta.Settings.ttip_offset, y + h + CopyPasta.Settings.ttip_offset
			)
		} 
		else {
			ToolTip(this.BuildDebugMessage(callStack))
		}
		return this
	}
	QueueTooltipDissmiss(periodMs) {
		SetTimer () => (ToolTip()), -periodMs 
	}	
	BuildDebugMessage(callStack) {
		eventMessage := ""
		eventType := StrReplace(Type(this.Event), "CopyPasta", "", 1) ;eventType := StrReplace(StrReplace(Type(this.Event), "CopyPasta", "", 1), "Event", "")
		Loop Parse, this.Event.Message, "`n"
		{
			eventMessage := eventMessage A_Tab A_Tab A_Space A_LoopField "`n"
		}
		return (
			CopyPasta.Settings.ttip_msg_debug_line_full
			"`n`n"
			"CallStack:"
			A_Tab callStack A_Tab
			"`n"
			A_Tab A_Tab "BufferSize: " ClipboardAll().Size " DataType: " (DllCall("IsClipboardFormatAvailable", "Uint", 1) ? 1 : 2)
			"`n"
			"Event:"
			A_Tab A_Tab (eventType!="" ? eventType: (Format(CopyPasta.Settings.ttip_msg_line_success, CopyPasta.Settings.ttip_msg_line_icon, CopyPasta.Settings.ttip_msg_app)))
			"`n"
			A_Tab A_Tab "{"
			"`n"
			eventMessage A_Tab
			A_Tab "}"
			"`n "
		)
	}
}

class CopyPastaEventBase {

	IsError => true ; most of events are errors
	Message => this.BuildMessage("")
	
	BuildMessage(msg) {
		if (this.IsError) {
			return (
				Format(CopyPasta.Settings.ttip_msg_error_line_full, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_side, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_emoji_shrug, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_side, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				msg
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_side, CopyPasta.Settings.ttip_msg_error_line_icon)
				"`n"
				Format(CopyPasta.Settings.ttip_msg_error_line_full, CopyPasta.Settings.ttip_msg_error_line_icon)
			)
		}
		else {
			return (
				Format(CopyPasta.Settings.ttip_msg_line_success, CopyPasta.Settings.ttip_msg_line_icon)
				"`n" 
				msg
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
		get => this.BuildMessage(this._message)
	}
	
	__New() {
		this._isError := false
		; trim and shorten the tooltip message
		this._message := SubStr(
			Trim(A_Clipboard), 
			1, 
			CopyPasta.Settings.ttip_msg_clip_lenght_limit
		)
		
		; determine dataType - 0 empty 1 text+copy file >1 not text
		dataType := (DllCall("IsClipboardFormatAvailable", "Uint", 1) ? 1 : 2)
		; validate
		Switch dataType
		{
			Case 1:
				if (this._message = "" || this._message = "`n" || this._message = "`r" || this._message = "`r`n") {
						this._isError := true
						this._message := Format(CopyPasta.Settings.ttip_msg_error_line_err_empty, CopyPasta.Settings.ttip_msg_error_line_icon)
				}
				else if (CbOrigCb = A_Clipboard) {
					this._isError := true
					this._message := Format(CopyPasta.Settings.ttip_msg_debug_line_duplicate, CopyPasta.Settings.ttip_msg_error_line_icon)
				}
			Case 2:
				if (CbOrigBufferSize = ClipboardAll().Size) {
					this._isError := true
					this._message := Format(CopyPasta.Settings.ttip_msg_debug_line_duplicate, CopyPasta.Settings.ttip_msg_error_line_icon)
				}
			Case 0:
				this._isError := true
				this._message := Format(CopyPasta.Settings.ttip_msg_error_line_err_empty, CopyPasta.Settings.ttip_msg_error_line_icon)
		}
	}
}

; Ctrl=^
#HotIf !WinActive("ahk_class QPasteClass") ;skip on Ditto the clipboard manager, before paste its using CTRL+C
	$^c::{
		global CbOrigBufferSize := ClipboardAll().Size
		global CbOrigCb := A_Clipboard
		Send "^c"
		Sleep 20
		HandleOnClipboardNotChangedYet("CTRL+C")
	}
#HotIf 

; Ctrl=^
$^x::{ 
	global CbOrigBufferSize := ClipboardAll().Size
	global CbOrigCb := A_Clipboard
	Send "^x"
	Sleep 20
	HandleOnClipboardNotChangedYet("CTRL+X")
}

; FN key - vkFF sc163 not found 
$vkFF::{ 	
	HotIf (*) => GetKeyState("vkFF", "P")
	Hotkey "*c",(*) => CopyPasta(CopyPastaFnEvent()).ShowTooltip().ShowDebug("Fn+C")
	KeyWait "vkFF"
}

; LWin=#
$#c::{
	CopyPasta(CopyPastaLWinEvent()).ShowTooltip().ShowDebug("LWin+C")
	Send "#c"
}

; Alt=! 
$!c::{
	CopyPasta(CopyPastaLAltEvent()).ShowTooltip().ShowDebug("Alt+C")
	Send "!c"
}

HandleOnClipboardNotChangedYet(stackTrace) {
	if !ClipWait(1, 1) {
			CopyPasta(CopyPastaLongerThanExpectedEvent()).ShowTooltip().ShowDebug(stackTrace "/NotChangedYet/Long1")
			if (ClipWait((CopyPasta.Settings.ttip_cooldown_ms+250)/1000, 1)) {
				CopyPasta(CopyPastaEvent()).ShowTooltip().ShowDebug(stackTrace "/NotChangedYet/CLIPWAIT{ANY-OK}")
			}
			else {
				CopyPasta(CopyPastaWaitingEvent()).ShowTooltip().ShowDebug(stackTrace "/NotChangedYet/Long2")
				if ClipWait(5, 1)
					CopyPasta(CopyPastaEvent()).ShowTooltip().ShowDebug(stackTrace "/NotChangedYet/CLIPWAIT{ANY-OK}")
				else 
					CopyPasta(CopyPastaTimeoutEvent()).ShowTooltip().ShowDebug(stackTrace "/NotChangedYet/Long3")
			}
	}
	else {
		CopyPasta(CopyPastaEvent()).ShowTooltip().ShowDebug(stackTrace "/NotChangedYet/CLIPWAIT{ANY-OK}")
	}
}

; ############ TOOLS ############
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
; script auto reload on save in debug mode
#HotIf WinActive("ahk_class Notepad++") ; Reload ahk on CTRL+S when debugging
    ~^s:: {
		if (!DEBUG_MODE) {
			Send "^s"
			return
		}
		winTitle := WinGetTitle("A")  ; "A" matches "Active" window
		if (InStr(winTitle, A_Scriptdir) or InStr(winTitle, A_ScriptName)) { ; Only when the script dir/filename is in the titlebar
		  Reload
		  return
		}
		
		SplitPath(A_Scriptdir, &topDir) ; Only when the top dir name is in the titlebar
		if (InStr(winTitle, topDir)) {
		  Reload
		  return
		}
    }
#HotIf 
