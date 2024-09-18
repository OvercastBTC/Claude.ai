#Requires AutoHotkey v2+
AE_Claude()
AE_Claude.SM_BISL(&sm)

; VSCode Claude Integration Extension
#Include <..\Personal\Common_Personal>

; Global variables
; Global API_KEY := '' ;log.Claude.API_Key
Global API_KEY := log.Claude.API_Key
global API_URL := "https://api.anthropic.com/v1/messages"

; Main hotkey to trigger Claude interaction
^!c::Copy_Code_Claude() ; Ctrl+Alt+C
Copy_Code_Claude(){
	response := selected_text := cBak := ''
	if (WinActive("ahk_exe Code.exe"))
	{
		selected_text := GetSelectedText(&cBak)
		if (selected_text != '') {
			response := QueryClaude(selected_text)
			InsertTextInVSCode(response, cBak)
		}
		else {
			MsgBox("Please select some text in VSCode before triggering Claude.")
		}
	}
}
GetSelectedText(&cBak?){
	cBak := cliptext := ''
	AE_Claude.cBakClr(&cBak)
	Send('{sc1D down}{sc2E}{sc1D up}')
	AE_Claude.cSleep(100)
	cliptext := A_Clipboard
	AE_Claude.cSleep(100)
	return cliptext
}

QueryClaude(prompt){
	parsed := response := payload := headers := ''
	headers := "Content-Type: application/json`nX-API-Key: " . API_KEY

	payload := '
	(
	{
		"messages": [
			{
				"role": "user",
				"content": "' . prompt . '"
			}
		],
		"model": "claude-3-opus-20240229",
		"max_tokens": 1000
	}
	)'

	try	{
		whr := ComObject("WinHttp.WinHttpRequest.5.1")
		whr.Open("POST", API_URL, true)
		whr.SetRequestHeader("Content-Type", "application/json")
		whr.SetRequestHeader("X-API-Key", API_KEY)
		whr.Send(payload)
		whr.WaitForResponse()
		response := whr.ResponseText
	}
	catch as err {
		MsgBox("Error querying Claude API: " . err.Message)
		return ""
	}

	; Parse JSON response and extract content
	parsed := jsongo_claude.Parse(response)
	return parsed.content[1].text
}

InsertTextInVSCode(text, cBak?) {
	A_Clipboard := text
	AE_Claude.cSleep(100)
	Send('{sc1D down}{sc2F}{sc1D up}')
	AE_Claude.Timer(-500)
	Sleep(100)
	AE_Claude.EmptyClipboard()
	if (cBak != ''){
		AE_Claude.cRestore(cBak)
	}
	else {
		return
	}
}


/************************************************************************
 * function ......: Auto Execution (AE)
 * @description ..: A work in progress (WIP) of standard AE setup(s)
 * @file AE_Claude.v2.ahk
 * @author OvercastBTC
 * @date 2024.06.13
 * @version 2.0.0
 * @ahkversion v2+
 ***********************************************************************/
; @revision(2.0.0)...: Converted all functions to a Class
; --------------------------------------------------------------------------------
/************************************************************************
 * function ...........: Resource includes for .exe standalone
 * @author OvercastBTC
 * @date 2023.08.15
 * @version 3.0.2
 ***********************************************************************/
;@Ahk2Exe-IgnoreBegin
; #Include <CheckUpdate\ScriptVersionMap>
; version :=  ScriptVersion.ScriptVersionMap['main'] 
;@Ahk2Exe-IgnoreEnd
SetVersion := "3.0.0" ; If quoted literal not empty, do 'SetVersion'
;@Ahk2Exe-Nop
;@Ahk2Exe-Obey U_V, = "%A_PriorLine~U)^(.+")(.*)".*$~$2%" ? "SetVersion" : "Nop"
;@Ahk2Exe-%U_V% %A_PriorLine~U)^(.+")(.*)".*$~$2%
; --------------------------------------------------------------------------------
#Requires AutoHotkey v2+
#Warn All, OutputDebug
#SingleInstance Force
#WinActivateForce
; --------------------------------------------------------------------------------
;! Need to investigate this stuff
; #MaxThreads 255 ; Allows a maximum of 255 instead of default threads.
; #MaxThreadsBuffer true
; A_MaxHotkeysPerInterval := 1000
; --------------------------------------------------------------------------------
SetWorkingDir(A_ScriptDir)
SetTitleMatchMode(2)
; --------------------------------------------------------------------------------
; AE_Claude.DH(true)
; AE_Claude.SetDelays(-1)
; AE_Claude.DPIAwareness()
; --------------------------------------------------------------------------------
/**
 * Function: Includes
 */
; #Include <Paths>
; #Include <Includes\Base>
; #Include <Includes\ObjectTypeExtensions>
; #Include <Includes\Apps>
; #Include <Includes\Includes_DPI>
; #Include <Utils\ClipSend>
; #Include <Tools\explorerGetPath.v2>
; #Include <Chrome>
; #Include <App\WebInspector>
; #Include <Tools\InternetSearch>
; #Include <_GuiReSizer>
; #Include <Tools\WAPI\WAPI>
; #Include <System\UIA>
; #Include <Tools\Info> ;! Moved to Gui.ahk
; #Include <App\Autohotkey> ;! Included in Includes\Apps
; ---------------------------------------------------------------------------
Hotstring(':?*:/new', '{#}Requires AutoHotkey v2.0{+}`n*Esc::ExitApp()')
; Typing '!quit!' exits the script
Hotstring(':?*X:!quit!', (*) => ExitApp())
; HotString(':?*XC1:bonnell.i', (*) => key.SendVK(Tredgar.i))
; HotString(':?*XC1:bonnell.c', (*) => key.SendVK(Tredgar.c))
; HotString(':?*XC1:bonnell.n', (*) => key.SendVK(Tredgar.ln))
; HotString(':?*XC1:bonnell.l', (*) => key.SendVK(Tredgar.loc))

class AE_Claude {
	/************************************************************************
	* @description Initialize the class with default settings
	* @example AE class is instantiated
	***********************************************************************/
	static __New() {
		this.DH(1)
		this.SetDelays(-1)
	}

	/************************************************************************
	* @description Toggle the CapsLock state
	* @example AE_Claude.toggleCapsLock()
	***********************************************************************/
	static toggleCapsLock() => SetCapsLockState(!GetKeyState('CapsLock', 'T'))

	/************************************************************************
	* @description Set detection for hidden windows and text
	* @example AE_Claude.DH(1)
	***********************************************************************/
	static _DetectHidden_Text_Windows(n := 1) {
		DetectHiddenText(n)
		DetectHiddenWindows(n)
	}
	static DH(n) => this._DetectHidden_Text_Windows(n)
	static DetectHidden(n) => this._DetectHidden_Text_Windows(n)

	/************************************************************************
	* @description Set various delay settings
	* @example AE_Claude.SetDelays(-1)
	* @var {Integer} : delay_key := d := n := -1
	* @var {Integer} : hold_time := delay_press := p := -1
	***********************************************************************/
	static _SetDelays(n := -1, p:=-1) {
		delay_key := d := n
		hold_time := delay_press := p
		SetControlDelay(n)
		SetMouseDelay(n)
		SetWinDelay(n)
		SetKeyDelay(delay_key, delay_press)
	}
	static SetDelays(n) => this._SetDelays(n)

	/************************************************************************
	* @description Set BlockInput and SendLevel
	* @example AE_Claude.BISL(1)
	* @var {Integer} : Send_Level := A_SendLevel
	* @var {Integer} : Block_Input := bi := 0
	* @var {Integer} : n = send level increase number
	* @returns {Integer}
	***********************************************************************/
	static _BlockInputSendLevel(n := 1, bi := 0, &Send_Level?) {
		SendLevel(0)
		Send_Level := sl := A_SendLevel
		(sl < 100) ? SendLevel(sl + n) : SendLevel(n + n)
		(n >= 1) ? bi := 1 : bi := 0 
		BlockInput(bi)
		return Send_Level
	}
	static BISL(n := 1, bi := 0, &sl?) => this._BlockInputSendLevel(n, bi, &sl?)
	static BI(bi := 1) => BlockInput(bi)
	static rBI(bi := 0) => BlockInput(bi)
	static slBISL(&sl) => this.BISL(n := 1, bi := 0, &sl?)
	static rBISL(sl) => this._restoreBlockInputSendLevel(sl)
	static _restoreBlockInputSendLevel(sl) {
		this.rBI(0)
		SendLevel(sl)
	}

	/************************************************************************
	* @description Change SendMode and SetKeyDelay
	* @example AE_Claude.SM(&SendModeObj)
	* @var {Object} : SendModeObject,
	* @var {Integer} : s: A_SendMode,
			d: A_KeyDelay,
			p: A_KeyDuration
	***********************************************************************/
	static _SendMode(&SendModeObj := {}) {
		SendModeObj := {
			s: A_SendMode,
			d: A_KeyDelay,
			p: A_KeyDuration
		}
		SendMode('Event')
		SetKeyDelay(-1, -1)
		return SendModeObj
	}
	static SM(&SendModeObj?) => this._SendMode(&SendModeObj?)

	/************************************************************************
	* @description Restore SendMode and SetKeyDelay
	* @example AE_Claude.rSM(RestoreObject)
	***********************************************************************/
	static _RestoreSendMode(RestoreObject) {
		SetKeyDelay(RestoreObject.d, RestoreObject.p)
		SendMode(RestoreObject.s)
	}
	static rSM(RestoreObject) => this._RestoreSendMode(RestoreObject)

	/************************************************************************
	* @description Set SendMode, SendLevel, and BlockInput
	* @example AE_Claude.SM_BISL(&SendModeObj, 1)
	***********************************************************************/
	static _SendMode_SendLevel_BlockInput(&SendModeObj?, n := 1) {
		this.SM(&SendModeObj)
		this.BISL(1)
		return SendModeObj
	}
	static SM_BISL(&SendModeObj?, n := 1) => this._SendMode_SendLevel_BlockInput(&SendModeObj?, n:=1)

	/************************************************************************
	* @description Restore SendMode, SendLevel, and BlockInput
	* @example AE_Claude.rSM_BISL(RestoreObj)
	***********************************************************************/
	static _restore_SendMode_SendLevel_BlockInput(RestoreObj) {
		this.BISL(0)
		this.rSM(RestoreObj)
	}
	static rSM_BISL(RestoreObj) => this._restore_SendMode_SendLevel_BlockInput(RestoreObj)

	/************************************************************************
	* @description Sleep while clipboard is in use
	* @example AE_Claude._Clipboard_Sleep(10)
	***********************************************************************/
	; static _Clipboard_Sleep(n := 10) {
	;     loop n {
	;         Sleep(n)
	;     } Until (!this.GetOpenClipboardWindow() || (A_Index = 50))
	; }

	/************************************************************************
	* @description Wait for the clipboard to be available
	* @example AE_Claude.WaitForClipboard()
	***********************************************************************/
	; static WaitForClipboard(timeout := 1000) {
	; 	startTime := A_TickCount
	; 	while (this.IsClipboardBusy()) {
	; 		if (A_TickCount - startTime > timeout) {
	; 			throw Error("Clipboard timeout")
	; 		}
	; 		Sleep(10)
	; 	}
	; }
	static WaitForClipboard(timeout := 1000) {
		clipboardReady := false
		startTime := A_TickCount

		checkClipboard := (*) => _checkClipboard()
		_checkClipboard() {
			if (!this.IsClipboardBusy()) {
				clipboardReady := true
				SetTimer(checkClipboard, 0)  ; Turn off the timer
			} else if (A_TickCount - startTime > timeout) {
				SetTimer(checkClipboard, 0)  ; Turn off the timer
				; throw Error("Clipboard timeout")
			}
		}
		
		SetTimer(checkClipboard, 10)  ; Check every 10ms

		; Wait for the clipboard to be ready or for a timeout
		while (!clipboardReady) {
			Sleep(10)
		}
	}
	
	; static cSleep(n := 10) => this._Clipboard_Sleep(n)
	static cSleep(n := 10) => this.WaitForClipboard(n)
	/************************************************************************
		* @description Safely copy content to clipboard with verification
		* @context_sensitive Yes
		* @example result := AE_Claude.SafeCopyToClipboard()
		***********************************************************************/
	static SafeCopyToClipboard() {
		; cBak := this.BackupAndClearClipboard()
		cBak := ''
		this.BackupAndClearClipboard(&cBak)
		this.WaitForClipboard()
		this.SelectAllText()
		this.CopyToClipboard()
		this.WaitForClipboard()
		clipContent := this.GetClipboardText()
		return clipContent
	}

	/************************************************************************
	* @description Backup current clipboard content and clear it
	* @example cBak := AE_Claude.BackupAndClearClipboard()
	***********************************************************************/
	static BackupAndClearClipboard(&backup?) {
		backup := DllCall("OleGetClipboard", "Ptr", 0, "Ptr")
		DllCall("EmptyClipboard")
		return backup
	}

	/************************************************************************
	* @description Select all text in the focused control
	* @context_sensitive Yes
	* @example AE_Claude.SelectAllText()
	***********************************************************************/
	static SelectAllText() {
		static EM_SETSEL := 0x00B1
		hCtl := this.hfCtl()
		DllCall("SendMessage", "Ptr", hCtl, "UInt", EM_SETSEL, "Ptr", 0, "Ptr", -1)
	}

	/************************************************************************
	* @description Copy selected text to clipboard
	* @context_sensitive Yes
	* @example AE_Claude.CopyToClipboard()
	***********************************************************************/
	static CopyToClipboard() {
		static WM_COPY := 0x0301
		hCtl := this.hfCtl()
		DllCall("SendMessage", "Ptr", hCtl, "UInt", WM_COPY, "Ptr", 0, "Ptr", 0)
	}

	/************************************************************************
	* @description Check if the clipboard is currently busy
	* @example if AE_Claude.IsClipboardBusy()
	***********************************************************************/
	static IsClipboardBusy() {
		return DllCall("GetOpenClipboardWindow", "Ptr") ;!= 0
	}

	/************************************************************************
	* @description Get text from clipboard using DllCalls
	* @example clipText := AE_Claude.GetClipboardText()
	***********************************************************************/
	static GetClipboardText() {
		if (!DllCall("OpenClipboard", "Ptr", 0)) {
			return ""
		}

		hData := DllCall("GetClipboardData", "UInt", 13, "Ptr") ; CF_UNICODETEXT
		if (hData == 0) {
			DllCall("CloseClipboard")
			return ""
		}

		pData := DllCall("GlobalLock", "Ptr", hData, "Ptr")
		if (pData == 0) {
			DllCall("CloseClipboard")
			return ""
		}

		text := StrGet(pData, "UTF-16")

		DllCall("GlobalUnlock", "Ptr", hData)
		DllCall("CloseClipboard")

		return text
	}
	/************************************************************************
	* @description Empty the clipboard
	* @example AE_Claude.EmptyClipboard()
	***********************************************************************/
	static EmptyClipboard() => DllCall("User32.dll\EmptyClipboard", "Int")

	/************************************************************************
	* @description Close the clipboard
	* @example AE_Claude.CloseClipboard()
	***********************************************************************/
	static CloseClipboard() => DllCall("User32.dll\CloseClipboard", "Int")

	/************************************************************************
	* @description Get the handle of the window with an open clipboard
	* @example AE_Claude.GetOpenClipboardWindow()
	***********************************************************************/
	static GetOpenClipboardWindow() => DllCall("User32.dll\GetOpenClipboardWindow", "Ptr")
	static GetOpenClipWin() => this.GetOpenClipboardWindow()

	/************************************************************************
	* @description Backup and clear clipboard
	* @example AE_Claude._Clipboard_Backup_Clear(&cBak)
	***********************************************************************/
	/**
	 * @description Backup ClipboardAll() and clear clipboard
	 * @param cBak 
	 * @returns {ClipboardAll} 
	 */
	static cBakClr(&cBak?) => this._Clipboard_Backup_Clear(&cBak?)
	static _Clipboard_Backup_Clear(&cBak?) {
		cBak := ClipboardAll()
		this.EmptyClipboard()
		this.cSleep(100)
		this.CloseClipboard()
		return cBak
	}

	/************************************************************************
	* @description Restore clipboard from backup
	* @example AE_Claude._Clipboard_Restore(cBak)
	***********************************************************************/
	static _Clipboard_Restore(cBak) {
		SetTimer(() => this.cSleep(50), -500)
		A_Clipboard := cBak
		this.CloseClipboard()
	}
	static cRestore(cBak) => this._Clipboard_Restore(cBak)

	static SelectAll() {
		static EM_SETSEL := 0x00B1
		focusedControl := this.hfCtl()
		if (focusedControl) {
			SendMessage(EM_SETSEL, 0, -1, focusedControl)
		}
	}

	static Copy() {
		static WM_COPY := 0x0301
		focusedControl := this.hfCtl()
		if (focusedControl) {
			SendMessage(WM_COPY, 0, 0, focusedControl)
		}
	}

	static Paste() {
		static WM_PASTE := 0x0302
		focusedControl := this.hfCtl()
		if (focusedControl) {
			SendMessage(WM_PASTE, 0, 0, focusedControl)
		}
	}

	/************************************************************************
	* @description Set a timer for clipboard operations
	* @example AE_Claude.Timer(-100)
	***********************************************************************/
	static Timer(time?) => SetTimer((*) => AE_Claude.cSleep(0), time := -100)

	/************************************************************************
	* @description Get the handle of the focused control
	* @context_sensitive Yes
	* @example AE_Claude.hfCtl(&fCtl)
	***********************************************************************/
	static hfCtl(&fCtl?) {
		return fCtl := ControlGetFocus('A')
	}

	/************************************************************************
	* @description Select text in the focused control
	* @context_sensitive Yes
	* @example AE_Claude._Select(0, -1)
	***********************************************************************/
	static EM_SETSEL := 177
	static _Select(wParam, lParam) {
		return DllCall('SendMessage', 'UInt', this.hfCtl(), 'UInt', this.EM_SETSEL, 'UInt', wParam, (wParam = 0 && lParam = -1) ? 'UIntP' : 'UInt', lParam)
	}
	static _Select_All() => this._Select(0, -1)
	static SelectAll => (*) => this._Select(0, -1)
	static _Select_Beginning() => this._Select(0, 0)
	static SelectHome => (*) => this._Select(0, 0)
	static _Select_End() => this._Select(-1, -1)
	static SelectEnd => (*) => this._Select(-1, -1)

	/************************************************************************
	* @description Convert scan code to hexadecimal format
	* @example AE_Claude.SC_Convert("a")
	***********************************************************************/
	static SC_Convert(key) {
		key_SC := GetKeySC(key)
		return Format("sc{:X}", key_SC)
	}

	/************************************************************************
	* @description Draw a border around the control
	* @example AE_Claude.DrawBorder(ControlGetFocus('A'))
	***********************************************************************/
	static DrawBorder(hWnd:=0) => this.DrawBorder1(hWnd:= 0)
	static DrawBorder1(WA := WinActive("A")) {
		Static OS:=3
		Static BG:="FF0000"
		Static myGui := Gui("+AlwaysOnTop +ToolWindow -Caption","GUI4Border")
		myGui.BackColor := BG
		; WA:=WinActive("A")
		If WA && !WinGetMinMax(WA) && !WinActive("GUI4Border ahk_class AutoHotkeyGUI"){
			WinGetPos(&wX,&wY,&wW,&wH,WA)
			myGui.Show("x" wX " y" wY " w" wW " h" wH " NA")
			Try WinSetRegion("0-0 " wW "-0 " wW "-" wH " 0-" wH " 0-0 " OS "-" OS " " wW-OS
			. "-" OS " " wW-OS "-" wH-OS " " OS "-" wH-OS " " OS "-" OS,"GUI4Border")
		}
		Else {
			myGui.Hide()
		}
	}
	static DrawBorder2(hwnd, color:=0x04a230, enable:=1) {
		static DWMWA_BORDER_COLOR := 34
		static DWMWA_COLOR_DEFAULT	:= 0xFFFFFFFF
		R := (color & 0xFF0000) >> 16
		G := (color & 0xFF00) >> 8
		B := (color & 0xFF)
		color := (B << 16) | (G << 8) | R
		DllCall("dwmapi\DwmSetWindowAttribute", "ptr", hwnd, "int", DWMWA_BORDER_COLOR, "int*", enable ? color : DWMWA_COLOR_DEFAULT, "int", 4)
	}
	static GetAncestorTitles(winHandle, txtormap := true) {
		static arrayGA := [GA_PARENT := 1, GA_ROOT := 2, GA_ROOTOWNER := 3, GA_NOTOPDESCENT := 4]
		static GA_PARENT := 1, GA_ROOT := 2, GA_ROOTOWNER := 3, GA_NOTOPDESCENT := 4
		titles := Map()
		WinTitle_NoTopDescent := WinTitle_Parent := WinTitle_RootOwner := WinTitle_Root := ''
		for each, flag in arrayGA {
			; infos(each ' ' flag)
			try {
				try WinTitle_Root := WinGetTitle(DllCall('GetAncestor', 'Ptr', winHandle, 'Int', GA_ROOT))
				titles.set('WinTitle_Root', WinTitle_Root)
				try WinTitle_RootOwner := WinGetTitle(DllCall('GetAncestor', 'Ptr', winHandle, 'Int', GA_ROOTOWNER))
				titles.set('WinTitle_RootOwner', WinTitle_RootOwner)
				try WinTitle_Parent := WinGetTitle(DllCall('GetAncestor', 'Ptr', winHandle, 'Int', GA_PARENT))
				titles.set('WinTitle_Parent', WinTitle_Parent)
				try WinTitle_NoTopDescent := WinGetTitle(DllCall('GetAncestor', 'Ptr', winHandle, 'Int', GA_NOTOPDESCENT))
				titles.set('WinTitle_NoTopDescent', WinTitle_NoTopDescent)
				; try ancestor := DllCall('GetAncestor', 'Ptr', winHandle, 'Int', flag, 'Ptr')
				; try ancestor := DllCall( "GetAncestor", 'uint', winHandle, 'uint', flag)
				; try ancestor := DllCall( "GetAncestor", 'uint', winHandle, 'uint', flag, 'Ptr')
				; titles[flag] := A_Index ': ' WinGetTitle('ahk_id ' . ancestor)
				; if (ancestor) {
				; 	titles[flag] := A_Index ': ' WinGetTitle('ahk_id ' . ancestor) '`n'
				; }
			}
		}
		if txtormap == true {
			return titles
		}
		else {
			return titles.ToString('`n')
		}
	
		; return titles
	}
}

class jsongo_claude {
    ; Creator: GroggyOtter
    ; Created: 20230622
    ; Updated: 20230630
    ; Website: https://github.com/GroggyOtter/jsongo_AHKv2
    ; License: GNU (Free to use but please keep these top comment lines with the code)
    #Requires AutoHotkey 2.0+
    static version := 'BETA'
    
    ; User Options:
    static escape_slash     := 1    ; true => Adds the optional escape character to forward slashes
        ,  escape_backslash := 1    ; true => Uses \\ for backslash escaping instead of \u005C
        ,  inline_arrays    := 0    ; true => Arrays containing only strings/numbers are kept on 1 line
        ,  extract_objects  := 1    ; true => Attempts to extract literal objects in map format instead of erroring
        ,  extract_all      := 1    ; true => Attempts to extract any object to map format instead of erroring
        ,  silent_error     := 1    ; true => No more error popups and error_log property receives error message
        ,  error_log        := ''   ; Used to store error message when there's an error and silent_error is true
    
    ; User methods:
    ; Parse(jtxt [,reviver])
    ;   jtxt [Str]      |  JSON string to convert into an AHK object
    ;   reviver [Func]  |  [Optional] Reference to a reviver(key, value, remove) function.
    ;                   |  The function must have at least 3 parameters to accept each key, value, and a special remove variable.
    ;                   |  A reviver allows you to interact with each key:value pair before being added to the object.
    ;   RETURN          |  Success => Map, Array, String, or Number [Depending on JSON text]
    ;                   |  Failure => an error is displayed
    ;                   |  Failure + .silent_error is true => An empty string is returned and .error_log is set to error message
    static Parse(jtxt, reviver:='') => this._Parse(jtxt, reviver)

    ; Stringify(obj [,replacer ,spacer ,extract_all])
    ;   obj [Map|Arr]        |  Map, array, string, and number are accepted JSON values.
    ;                        |  Literal objects are accepted if the .extract_objects property is set to true.
    ;                        |  All objects are accepted if the .extract_all property or extract_all parameter are true.
    ;                        |  A replacer allows you to interact with each key:value pair before being added to the JSON string.
    ;   replacer [Func|Arr]  |  [Optional] A replacer allows you to interact with each key:value pair before it's addedto the object.
    ;                        |  If replacer is function, it must have at least 3 parameters and works similarly to a Parse() reviver.
    ;                        |  If replacer is array, each key is compared to each element and if a match is made, that key:value pair is discarded.
    ;   spacer [Str|Num]     |  [Optional] Defines the character set used to space each level of the JSON tree.
    ;                        |  Use a number indicates the number of spaces to use. 4 => 4 spaces => '    '
    ;                        |  Str indiciates the string to use. '`t' => 1 tab for each indent level
    ;                        |  If omitted or an empty string is passed in, the JSON string will export as a single line of text
    ;   extract_all [Bool]   |  [Optional] true => all object types will be processed and exported in map format
    ;                        |  If omitted, the extract_all property value is used instead.
    ;   RETURN               |  Success => Formatted JSON string
    ;                        |  Failure => an error is displayed
    ;                        |  Failure + .silent_error is true => An empty string is returned and .error_log is set to error message
    static Stringify(base_item, replacer:='', spacer:='', extract_all:=0) => this._Stringify(base_item, replacer, spacer, extract_all)
    
    static _Parse(jtxt, reviver:='') {
        this.error_log := '', if_rev := (reviver is Func && reviver.MaxParams > 2) ? 1 : 0, xval := 1, xobj := 2, xarr := 3, xkey := 4, xstr := 5, xend := 6, xcln := 7, xeof := 8, xerr := 9, null := '', str_flag := Chr(5), tmp_q := Chr(6), tmp_bs:= Chr(7), expect := xval, json := [], path := [json], key := '', is_key:= 0, remove := jsongo_claude.JSON_Remove(), fn := A_ThisFunc
        loop 31
            (A_Index > 13 || A_Index < 9 || A_Index = 11 || A_Index = 12) && (i := InStr(jtxt, Chr(A_Index), 1)) ? err(21, i, 'Character number: 9, 10, 13 or anything higher than 31.', A_Index) : 0
        for k, esc in [['\u005C', tmp_bs], ['\\', tmp_bs], ['\"',tmp_q], ['"',str_flag], [tmp_q,'"'], ['\/','/'], ['\b','`b'], ['\f','`f'], ['\n','`n'], ['\r','`r'], ['\t','`t']]
            this.replace_if_exist(&jtxt, esc[1], esc[2])
        i := 0
        while (i := InStr(jtxt, '\u', 1, ++i))
            IsNumber('0x' (hex := SubStr(jtxt, i+2, 4))) ? jtxt := StrReplace(jtxt, '\u' hex, Chr(('0x' hex)), 1) : err(22, i+2, '\u0000 to \uFFFF', '\u' hex)
        (i := InStr(jtxt, '\', 1)) ? err(23, i+1, '\b \f \n \r \t \" \\ \/ \u', '\' SubStr(jtxt, i+1, 1)) : jtxt := StrReplace(jtxt, tmp_bs, '\', 1)
        jlength := StrLen(jtxt) + 1, ji := 1
        
        while (ji < jlength) {
            if InStr(' `t`n`r', (char := SubStr(jtxt, ji, 1)), 1)
                ji++
            else switch expect {
                case xval:
                    v:
                    (char == '{') ? (o := Map(), (path[path.Length] is Array) ? path[path.Length].Push(o) : path[path.Length][key] := o, path.Push(o), expect := xobj, ji++)
                    : (char == '[') ? (a := [], (path[path.Length] is Array) ? path[path.Length].Push(a) : path[path.Length][key] := a, path.Push(a), expect := xarr, ji++)
                    : (char == str_flag) ? (end := InStr(jtxt, str_flag, 1, ji+1)) ? is_key ? (is_key := 0, key := SubStr(jtxt, ji+1, end-ji-1), expect := xcln, ji := end+1) : (rev(SubStr(jtxt, ji+1, end-ji-1)), expect := xend, ji := end+1) : err(24, ji, '"', SubStr(jtxt, ji))
                    : InStr('-0123456789', char, 1) ? RegExMatch(jtxt, '(-?(?:0|[123456789]\d*)(?:\.\d+)?(?:[eE][-+]?\d+)?)', &match, ji) ? (rev(Number(match[])), expect := xend, ji := match.Pos + match.Len ) : err(25, ji, , SubStr(jtxt, ji))
                    : (char == 't') ? (SubStr(jtxt, ji, 4) == 'true')  ? (rev(true) , ji+=4, expect := xend) : err(26, ji + tfn_idx('true', SubStr(jtxt, ji, 4)), 'true' , SubStr(jtxt, ji, 4))
                    : (char == 'f') ? (SubStr(jtxt, ji, 5) == 'false') ? (rev(false), ji+=5, expect := xend) : err(27, ji + tfn_idx('false', SubStr(jtxt, ji, 5)), 'false', SubStr(jtxt, ji, 5))
                    : (char == 'n') ? (SubStr(jtxt, ji, 4) == 'null')  ? (rev(null) , ji+=4, expect := xend) : err(28, ji + tfn_idx('null', SubStr(jtxt, ji, 4)), 'null' , SubStr(jtxt, ji, 4))
                    : err(29, ji, '`n`tArray: [ `n`tObject: { `n`tString: " `n`tNumber: -0123456789 `n`ttrue/false/null: tfn ', char)
                case xarr: if (char == ']')
                        path_pop(&char), expect := (path.Length = 1) ? xeof : xend, ji++
                    else goto('v')
                case xobj: 
                    switch char {
                        case str_flag: goto((is_key := 1) ? 'v' : 'v')
                        case '}': path_pop(&char), expect := (path.Length = 1) ? xeof : xend, ji++
                        default: err(31, ji, '"}', char)
                    }
                case xkey: if (char == str_flag)
                        goto((is_key := 1) ? 'v' : 'v')
                    else err(32, ji, '"', char)
                case xcln: (char == ':') ? (expect := xval, ji++) : err(33, ji, ':', char)
                case xend: (char == ',') ? (ji++, expect := (path[path.Length] is Array) ? xval : xkey)
                    : (char == '}') ? (ji++, (path[path.Length] is Map)   ? path_pop(&char) : err(34, ji, ']', char), (path.Length = 1) ? expect := xeof : 0)
                    : (char == ']') ? (ji++, (path[path.Length] is Array) ? path_pop(&char) : err(35, ji, '}', char), (path.Length = 1) ? expect := xeof : 0)
                    : err(36, ji, '`nEnd of array: ]`nEnd of object: }`nNext value: ,`nWhitespace: [Space] [Tab] [Linefeed] [Carriage Return]', char)
                case xeof: err(40, ji, 'End of JSON', char)
                case xerr: return ''
            }
        }
        
        return (path.Length != 1) ? err(37, ji, 'Size: 1', 'Actual size: ' path.Length) : json[1]
        
        path_pop(&char) => (path.Length > 1) ? path.Pop() : err(38, ji, 'Size > 0', 'Actual size: ' path.Length-1)
        rev(value) => (path[path.Length] is Array) ? (if_rev ? value := reviver((path[path.Length].Length), value, remove) : 0, (value == remove) ? '' : path[path.Length].Push(value) ) : (if_rev ? value := reviver(key, value, remove) : 0, (value == remove) ? '' : path[path.Length][key] := value )
        err(msg_num, idx, ex:='', rcv:='') => (clip := '`n',  offset := 50,  clip := 'Error Location:`n', clip .= (idx > 1) ? SubStr(jtxt, 1, idx-1) : '',  (StrLen(clip) > offset) ? clip := SubStr(clip, (offset * -1)) : 0,  clip .= '>>>' SubStr(jtxt, idx, 1) '<<<',  post_clip := (idx < StrLen(jtxt)) ? SubStr(jtxt, ji+1) : '',  clip .= (StrLen(post_clip) > offset) ? SubStr(post_clip, 1, offset) : post_clip,  clip := StrReplace(clip, str_flag, '"'),  this.error(msg_num, fn, ex, rcv, clip), expect := xerr)
        tfn_idx(a, b) {
            loop StrLen(a)
                if SubStr(a, A_Index, 1) !== SubStr(b, A_Index, 1)
                    Return A_Index-1
        }
    }
    
    static _Stringify(base_item, replacer, spacer, extract_all) {
        switch Type(replacer) {
            case 'Func': if_rep := (replacer.MaxParams > 2) ? 1 : 0
            case 'Array':
                if_rep := 2, omit := Map(), omit.Default := 0
                for i, v in replacer
                    omit[v] := 1
            default: if_rep := 0
        }
        
        switch Type(spacer) {
            case 'String': _ind := spacer, lf := (spacer == '') ? '' : '`n'
                if (spacer == '')
                    _ind := lf := '', cln := ':'
                else _ind := spacer, lf := '`n', cln := ': '
            case 'Integer','Float','Number':
                lf := '`n', cln := ': ', _ind := ''
                loop Floor(spacer)
                    _ind .= ' '
            default: _ind := lf := '', cln := ':'
        }
        
        this.error_log := '', extract_all := (extract_all) ?  1 : this.extract_all ? 1 : 0, remove := jsongo_claude.JSON_Remove(), value_types := 'String Number Array Map', value_types .= extract_all ? ' AnyObject' : this.extract_objects ? ' LiteralObject' : '', fn := A_ThisFunc
        
        (if_rep = 1) ? base_item := replacer('', base_item, remove) : 0
        if (base_item = remove)
            return ''
        else jtxt := extract_data(base_item)
        
        loop 33
            switch A_Index {
                case 9,10,13: continue
                case  8: this.replace_if_exist(&jtxt, Chr(A_Index), '\b')
                case 12: this.replace_if_exist(&jtxt, Chr(A_Index), '\f')
                case 32: (this.escape_slash) ? this.replace_if_exist(&jtxt, '/', '\/') : 0
                case 33: (this.escape_backslash) ? this.replace_if_exist(&jtxt, '\u005C', '\\') : 0 
                default: this.replace_if_exist(&jtxt, Chr(A_Index), Format('\u{:04X}', A_Index))
            }
        
        return jtxt
        
        extract_data(item, ind:='') {
            switch Type(item) {
                case 'String': return '"' encode(&item) '"'
                case 'Integer','Float': return item
                case 'Array':
                    str := '['
                    if (ila := this.inline_arrays ?  1 : 0)
                        for i, v in item
                            InStr('String|Float|Integer', Type(v), 1) ? 1 : ila := ''
                        until (!ila)
                    for i, v in item
                        (if_rep = 2 && omit[i]) ? '' : (if_rep = 1 && (v := replacer(i, v, remove)) = remove) ? '' : str .= (ila ? extract_data(v, ind _ind) ', ' : lf ind _ind extract_data(v, ind _ind) ',')
                    return ((str := RTrim(str, ', ')) == '[') ? '[]' : str (ila ? '' : lf ind) ']'
                case 'Map':
                    str := '{'
                    for k, v in item
                        (if_rep = 2 && omit[k]) ? '' : (if_rep = 1 && (v := replacer(k, v, remove)) = remove) ? '' : str .= lf ind _ind (k is String ? '"' encode(&k) '"' cln : err(11, 'String', Type(k))) extract_data(v, ind _ind) ','
                    return ((str := RTrim(str, ',')) == '{') ? '{}' : str lf ind '}'
                case 'Object':
                    (this.extract_objects) ? 1 : err(12, value_types, Type(item))
                    Object:
                    str := '{'
                    for k, v in item.OwnProps()
                        (if_rep = 2 && omit[k]) ? '' : (if_rep = 1 && (v := replacer(k, v, remove)) = remove) ? '' : str .= lf ind _ind (k is String ? '"' encode(&k) '"' cln : err(11, 'String', Type(k))) extract_data(v, ind _ind) ','
                    return ((str := RTrim(str, ',')) == '{') ? '{}' : str lf ind '}'
                case 'VarRef','ComValue','ComObjArray','ComObject','ComValueRef': return err(15, 'These are not of type "Object":`nVarRef ComValue ComObjArray ComObject and ComValueRef', Type(item))
                default:
                    !extract_all ? err(13, value_types, Type(item)) : 0
                    goto('Object')
            }
        }
        
        encode(&str) => (this.replace_if_exist(&str ,  '\', '\u005C'), this.replace_if_exist(&str,  '"', '\"'), this.replace_if_exist(&str, '`t', '\t'), this.replace_if_exist(&str, '`n', '\n'), this.replace_if_exist(&str, '`r', '\r')) ? str : str
        err(msg_num, ex:='', rcv:='') => this.error(msg_num, fn, ex, rcv)
    }

    class JSON_Remove {
    }
    static replace_if_exist(&txt, find, replace) => (InStr(txt, find, 1) ? txt := StrReplace(txt, find, replace, 1) : 0)
    static error(msg_num, fn, ex:='', rcv:='', extra:='') {
        err_map := Map(11,'Stringify error: Object keys must be strings.'  ,12,'Stringify error: Literal objects are not extracted unless:`n-The extract_objects property is set to true`n-The extract_all property is set to true`n-The extract_all parameter is set to true.'  ,13,'Stringify error: Invalid object found.`nTo extract all objects:`n-Set the extract_all property to true`n-Set the extract_all parameter to true.'  ,14,'Stringify error: Invalid value was returned from Replacer() function.`nReplacer functions should always return a string or the "remove" value passed into the 3rd parameter.'  ,15,'Stringify error: Invalid object encountered.'  ,21,'Parse error: Forbidden character found.`nThe first 32 ASCII chars are forbidden in JSON text`nTab, linefeed, and carriage return may appear as whitespace.'  ,22,'Parse error: Invalid hex found in unicode escape.`nUnicode escapes must be in the format \u#### where #### is a hex value between 0000 and FFFF.`nHex values are not case sensitive.'  ,23,'Parse error: Invalid escape character found.'  ,24,'Parse error: Could not find end of string'  ,25,'Parse error: Invalid number found.'  ,26,'Parse error: Invalid `'true`' value.'  ,27,'Parse error: Invalid `'false`' value.'  ,28,'Parse error: Invalid `'null`' value.'  ,29,'Parse error: Invalid value encountered.'  ,31,'Parse error: Invalid object item.'  ,32,'Parse error: Invalid object key.`nObject values must have a string for a key name.'  ,33,'Parse error: Invalid key:value separator.`nAll keys must be separated from their values with a colon.'  ,34,'Parse error: Invalid end of array.'  ,35,'Parse error: Invalid end of object.'  ,36,'Parse error: Invalid end of value.'  ,37,'Parse error: JSON has objects/arrays that have not been terminated.'  ,38,'Parse error: Cannot remove an object/array that does not exist.`nThis error is usually thrown when there are extra closing brackets (array)/curly braces (object) in the JSON string.'  ,39,'Parse error: Invalid whitespace character found in string.`nTabs, linefeeds, and carriage returns must be escaped as \t \n \r (respectively).'  ,40,'Characters appears after JSON has ended.' )
        msg := err_map[msg_num], (ex != '') ? msg .= '`nEXPECTED: ' ex : 0, (rcv != '') ? msg .= '`nRECEIVED: ' rcv : 0
        if !this.silent_error
            throw Error(msg, fn, extra)
        this.error_log := 'JSON ERROR`n`nTimestamp:`n' A_Now '`n`nMessage:`n' msg '`n`nFunction:`n' fn '()' (extra = '' ? '' : '`n`nExtra:`n') extra '`n'
        return ''
    }
}
