#Requires AutoHotkey v2+
#Include <Directives\__AE.v2>
AE()
AE.SM_BISL(&sm)

; VSCode Claude Integration Extension
#Include <..\Personal\Common_Personal>

; Global variables
global API_KEY := log.Claude.API_Key
global API_URL := "https://api.anthropic.com/v1/messages"

; Main hotkey to trigger Claude interaction
^!c::Copy_Code_Claude() ; Ctrl+Alt+C
Copy_Code_Claude(){
	cBak := ''
	if (WinActive("ahk_exe Code.exe"))
	{
		selected_text := GetSelectedText(&cBak)
		if (selected_text != "") {
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
	AE.cBakClr(&cBak)
	key.sSend(key.copy)
	AE.cSleep(100)
	cliptext := A_Clipboard
	AE.cSleep(100)
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
	parsed := jsongo.Parse(response)
	return parsed.content[1].text
}

InsertTextInVSCode(text, cBak?) {
	key.cSend(text)
	AE.Timer(-500)
	Sleep(100)
	AE.EmptyClipboard()
	AE.cSleep(100)
	if (cBak != ''){
		AE.cRestore(cBak)
	} else {
		return
	}
}
