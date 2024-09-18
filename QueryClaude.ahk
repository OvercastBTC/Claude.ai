#Requires AutoHotkey v2.0+
#SingleInstance Force
#Include <JSON> ;JSON must have JSON.Load and JSON.Dump methods.

GetClaudeData(inputText, API_KEY := "Your API key here") {
    headers := "Content-Type: application/json`nX-API-Key: " . API_KEY

    payload := ''
    . '    {'
    . '    "messages": ['
    . '        {'
    . '            "role": "user",'
    . '            "content": "' . inputText . '"'
    . '        }'
    . '    ],'
    . '    "model": "claude-3-opus-20240229",'
    . '    "max_tokens": 1000'
    . '    }'

    response := ""
    try
    {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("POST", 'https://api.anthropic.com/v1/messages', true)
        whr.SetRequestHeader("Content-Type", "application/json")
        whr.SetRequestHeader("anthropic-version", "2023-06-01")
        whr.SetRequestHeader("X-API-Key", API_KEY)
        whr.Send(payload)
        whr.WaitForResponse()
        response := whr.ResponseText
    }
    catch as err
    {
        MsgBox("Error querying Claude API: " . err.Message)
        return ""
    }

    ; Parse JSON response and extract content
    parsed := JSON.Load(response)
    return parsed["content"][1].get('text')
    ;You could compress those last four lines. I spread them out so it'll be easier to understand

}

;Example of how to use the function.
QueryClaudeAI() {
    input := InputBox("Ask me anything...", "Claude AI API")
    if (input.result = "cancel")
        ExitApp()
    MsgBox(GetClaudeData(input.Value))
}
Loop {
    QueryClaudeAI()
}

Esc:: ExitApp()
