; Function to escape special characters for JSON
EscapeJSON(str) {
    str := StrReplace(str, "\", "\\")
    str := StrReplace(str, "`"", "\`"")
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`t", "\t")
    return str
}
#Requires AutoHotkey v2.0
#SingleInstance Force
#NoTrayIcon

; Add UTF-8 file encoding to ensure proper handling of non-ASCII characters
FileEncoding "UTF-8"

; Variable declarations
ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
ModernbrowsersProcesses := "msedge.exe,iexplore.exe,chrome.exe,opera.exe,brave.exe,vivaldi.exe,MicrosoftEdge.exe"

; Main loop instead of timer
Loop {
    try {
        SaveURL()
    } catch Error as e {
        ; If there's an error, output a JSON error object and continue
        jsonOutput := "{`"name`":null,`"title`":null,`"url`":null,`"error`":`"" . EscapeJSON(e.Message) . "`"}"
        FileAppend(jsonOutput "`n", "*", "UTF-8")
    }
    Sleep 1800  ; 1.8 seconds between checks
}

SaveURl()
{
    ; Get active window
    activeWin := WinExist("A")
    if (!activeWin) {
        ; No active window, return null JSON
        jsonOutput := "{`"name`":null,`"title`":null,`"url`":null}"
        FileAppend(jsonOutput "`n", "*", "UTF-8")
        Return
    }

    ; Try to get window class, title and process name with error handling
    try {
        sClass := WinGetClass("ahk_id " activeWin)
    } catch {
        sClass := ""
    }
    
    try {
        title := WinGetTitle("ahk_id " activeWin)
    } catch {
        title := ""
    }
    
    try {
        name := WinGetProcessName("ahk_id " activeWin)
    } catch {
        name := ""
    }
    
    if InStr(ModernbrowsersProcesses, name)
    {
        If InStr(ModernBrowsers, sClass)
        {
            accData := GetAccData()
            if !accData {
                Return
            }
            _2Data := accData[2]
            if !_2Data {
                _2Data := "new tab"
            }
            ; Format output as JSON
            jsonOutput := "{`"name`":`"" . EscapeJSON(name) . "`",`"title`":`"" . EscapeJSON(title) . "`",`"url`":`"" . EscapeJSON(_2Data) . "`"}"
            FileAppend(jsonOutput "`n", "*", "UTF-8")
            _2Data := ""
            Return
        } else {
            ddeData := GetBrowserURL_DDE(sClass)
            If !ddeData {
                ddeData := "new tab"
            }
            ; Format output as JSON
            jsonOutput := "{`"name`":`"" . EscapeJSON(name) . "`",`"title`":`"" . EscapeJSON(title) . "`",`"url`":`"" . EscapeJSON(ddeData) . "`"}"
            FileAppend(jsonOutput "`n", "*", "UTF-8")
            ddeData := ""
            Return
        }
    }
    ; Format output as JSON with null url
    jsonOutput := "{`"name`":`"" . EscapeJSON(name) . "`",`"title`":`"" . EscapeJSON(title) . "`",`"url`":null}"
    FileAppend(jsonOutput "`n", "*", "UTF-8")
    Return
}

;-------Function-------
GetTitle() 
{
    Title := WinGetTitle("A")
    Return Title
}

GetText() 
{
    Text := WinGetText("A")
    Return Text
}

GetName() 
{
    Active_ID := WinGetID("A")
    Active_Process := WinGetProcessName("ahk_id " Active_ID)
    return Active_Process
}

GetAccData(WinId := "A") 
{
    static w := Map(), n := 0
    th := WinExist(WinId)
    if GetKeyState("Ctrl", "P")
    {
        w := Map()
        n := 0
    }
    
    for i, v in w
    {
        if (th = v[1])
            Return [GetAccObjectFromWindow(v[1]).accName(0), ParseAccData(v[4])[2]]
    }
    
    tr := ParseAccData(GetAccObjectFromWindow(th))
    
    ; Make sure tr is properly initialized as an array with at least 2 elements
    if !IsObject(tr) {
        tr := [0, 0]
    } else if tr.Length < 2 {
        tr.Push(0)
    }
    
    if tr[2]
    {
        n++
        w[n] := [th, tr[1], tr[2], tr[3]]
    }
    
    Return [tr[1], tr[2]]
}

ParseAccData(accObj, accData := "") 
{
    ; Initialize accData as an array if not provided
    if (accData = "") {
        accData := [0, 0, 0]  ; Pre-initialize with 3 elements
    }
    
    ; Safety check for accObj
    if (!IsObject(accObj)) {
        return accData
    }
    
    ; Try to get accName
    try {
        if (accData[1] = 0 || accData[1] = "") {
            accData[1] := accObj.accName(0)
        }
    } catch {
        accData[1] := ""
    }
    
    ; Try to get URL from accValue if the role is correct
    try {
        if (accObj.accRole(0) = 42 && accObj.accName(0) && accObj.accValue(0)) {
            u := accObj.accValue(0)
            accData[2] := SubStr(u, 1, 4) = "http" ? u : "https://" u
            accData[3] := accObj
        }
    } catch {
        ; Do nothing if this fails
    }
    
    ; Try to process children if we don't have a URL yet
    try {
        if (!accData[2]) {  ; Check if element 2 exists AND is empty
            children := GetAccChildren(accObj)
            if (IsObject(children)) {
                for _, accChild in children {
                    if (IsObject(accChild)) {
                        ParseAccData(accChild, accData)
                        if (accData[2]) {  ; If we found a URL, stop processing
                            break
                        }
                    }
                }
            }
        }
    } catch {
        ; Do nothing if child processing fails
    }
    
    Return accData
}

GetAccInit() 
{
    static hw := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    return hw
}

GetAccObjectFromWindow(hWnd, idObject := 0) 
{
    static IID_IAccessible := "{618736E0-3C3D-11CF-810C-00AA00389B71}"
    
    ; Load oleacc.dll if needed
    if !DllCall("GetModuleHandle", "Str", "oleacc", "Ptr")
        DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    
    ; Create a GUID from the IID string
    GUID := Buffer(16, 0)
    DllCall("ole32\CLSIDFromString", "WStr", IID_IAccessible, "Ptr", GUID)
    
    ; Send WM_GETOBJECT message
    SendMessage 0x003D, 0, 1, "Chrome_RenderWidgetHostHWND1", "ahk_id " WinExist("A")  ; WM_GETOBJECT
    
    ; Try to get the accessibility object
    pacc := 0
    loop 60 {
        ; Try to get the accessible object
        hr := DllCall("oleacc\AccessibleObjectFromWindow", 
                    "Ptr", hWnd, 
                    "UInt", idObject & 0xFFFFFFFF, 
                    "Ptr", GUID, 
                    "Ptr*", &pacc)
                    
        if (hr = 0 && pacc != 0)  ; S_OK and got an object
            break
            
        if (A_Index >= 60)
            return 0
            
        Sleep 30
    }
    
    if (pacc = 0)
        return 0
        
    return ComObjFromPtr(pacc)
}

GetAccQuery(objAcc) 
{
    try {
        if ComObjType(objAcc, "Name") != "IAccessible"
            return 0
        return ComObjQuery(objAcc, "{618736e0-3c3d-11cf-810c-00aa00389b71}")
    }
}

GetAccChildren(objAcc) 
{
    ; Safety check
    if (!IsObject(objAcc)) {
        return []
    }
    
    try {
        if (ComObjType(objAcc, "Name") != "IAccessible") {
            return []
        }
        
        cChildren := objAcc.accChildCount
        Children := []
        
        if (cChildren <= 0) {
            return Children
        }
        
        varChildren := Buffer(cChildren * (8 + 2 * A_PtrSize), 0)
        
        if (!DllCall("oleacc\AccessibleChildren", 
                  "Ptr", ComObjValue(objAcc), 
                  "Int", 0, 
                  "Int", cChildren, 
                  "Ptr", varChildren, 
                  "Int*", &cChildren)) {
            
            Loop cChildren {
                i := (A_Index - 1) * (A_PtrSize * 2 + 8) + 8
                child := NumGet(varChildren, i, "Ptr")
                vt := NumGet(varChildren, i - 8, "UChar")
                
                if (vt = 9 && child) {  ; VT_DISPATCH and valid pointer
                    try {
                        childObj := ComObjFromPtr(child)
                        Children.Push(childObj)
                    } catch {
                        ; Skip this child if there's an error
                    }
                    
                    ; Release the object even if we had an error
                    if (child) {
                        DllCall("OleAut32\VariantClear", "Ptr", child)
                    }
                } else if (child) {
                    Children.Push(child)
                }
            }
        }
        
        return Children
    } catch {
        return []
    }
}

GetBrowserURL_DDE(sClass) 
{
    sServer := WinGetProcessName("ahk_class " sClass)
    sServer := SubStr(sServer, 1, StrLen(sServer) - 4)
    
    iCodePage := 0x04B0  ; Always use Unicode in v2
    
    idInst := 0
    DllCall("DdeInitialize", "UInt*", &idInst, "UInt", 0, "UInt", 0, "UInt", 0)
    
    hServer := DllCall("DdeCreateStringHandle", "UInt", idInst, "Str", sServer, "Int", iCodePage)
    hTopic := DllCall("DdeCreateStringHandle", "UInt", idInst, "Str", "WWW_GetWindowInfo", "Int", iCodePage)
    hItem := DllCall("DdeCreateStringHandle", "UInt", idInst, "Str", "0xFFFFFFFF", "Int", iCodePage)
    
    hConv := DllCall("DdeConnect", "UInt", idInst, "UInt", hServer, "UInt", hTopic, "UInt", 0)
    
    sData := ""
    if (hConv) {
        nResult := 0
        hData := DllCall("DdeClientTransaction", "Ptr", 0, "UInt", 0, "UInt", hConv, "UInt", hItem, "UInt", 1, "UInt", 0x20B0, "UInt", 10000, "UInt*", &nResult)
        
        if (hData) {
            ; Get data size first
            cbData := 0
            pData := DllCall("DdeAccessData", "UInt", hData, "UInt*", &cbData, "Ptr")
            
            if (pData && cbData) {
                ; Create a buffer and copy the data
                sData := StrGet(pData, cbData, "CP0")
                DllCall("DdeUnaccessData", "UInt", hData)
            }
            DllCall("DdeFreeDataHandle", "UInt", hData)
        }
        DllCall("DdeDisconnect", "UInt", hConv)
    }
    
    ; Clean up
    DllCall("DdeFreeStringHandle", "UInt", idInst, "UInt", hServer)
    DllCall("DdeFreeStringHandle", "UInt", idInst, "UInt", hTopic)
    DllCall("DdeFreeStringHandle", "UInt", idInst, "UInt", hItem)
    DllCall("DdeUninitialize", "UInt", idInst)
    
    ; Parse the result if we got any
    if (sData) {
        sWindowInfo := StrSplit(sData, "`"")
        if (sWindowInfo.Length >= 3)
            return sWindowInfo[2]
    }
    
    return ""
}