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
ModernBrowsers :=
    "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
ModernbrowsersProcesses := "msedge.exe,iexplore.exe,chrome.exe,opera.exe,brave.exe,vivaldi.exe,MicrosoftEdge.exe"

; Last known good output — used as fallback when a transient window (notification, toast, etc.) steals focus
lastValidOutput := ""

; Window classes and processes that are transient/overlay and should be ignored
TransientClasses := "Windows.UI.Core.CoreWindow,NativeHWNDHost,Shell_TrayWnd,Shell_SecondaryTrayWnd,NotifyIconOverflowWindow,TopLevelWindowForOverflowXamlIsland,XamlExplorerHostIslandWindow,InputApp"
TransientProcesses := "ShellExperienceHost.exe,SearchHost.exe,StartMenuExperienceHost.exe,TextInputHost.exe"

; Main loop instead of timer
loop {
    try {
        SaveURL()
    } catch Error as e {
        ; If there's an error, output a JSON error object and continue
        jsonOutput := "{`"name`":null,`"displayName`":null,`"title`":null,`"url`":null,`"error`":`"" . EscapeJSON(e.Message
        ) . "`"}"
        ; MsgBox jsonOutput
        FileAppend(jsonOutput "`n", "*", "UTF-8")
    }
    Sleep 1500 ; 1.5 seconds between checks
}

SaveURl() {
    global lastValidOutput, TransientClasses, TransientProcesses

    ; Get active window
    activeWin := WinExist("A")

    ; Validate the active window — check if it's real, visible, and not transient
    if (activeWin) {
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
            name := GetRealProcessName(activeWin)
        } catch {
            name := ""
        }

        isTransient := InStr(TransientClasses, sClass) || InStr(TransientProcesses, name) || (!name && !title)
    } else {
        isTransient := true
    }

    ; If no active window or it's transient — find the topmost real window instead
    if (!activeWin || isTransient) {
        activeWin := GetTopVisibleWindow()
        if (!activeWin) {
            ; Nothing found — fall back to last valid output
            if (lastValidOutput) {
                ; MsgBox lastValidOutput
                FileAppend(lastValidOutput "`n", "*", "UTF-8")
            }
            return
        }
        ; Re-read properties from the fallback window
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
            name := GetRealProcessName(activeWin)
        } catch {
            name := ""
        }
    }

    try {
        displayName := GetAppDisplayName(name)
    } catch {
        displayName := ""
    }

    if InStr(ModernbrowsersProcesses, name) {
        if InStr(ModernBrowsers, sClass) {
            accData := GetAccData("ahk_id " activeWin)
            if !accData {
                return
            }
            _2Data := accData[2]
            if !_2Data {
                _2Data := "new tab"
            }
            ; Format output as JSON
            jsonOutput := "{`"name`":`"" . EscapeJSON(name) . "`",`"displayName`":`"" . EscapeJSON(displayName) .
            "`",`"title`":`"" . EscapeJSON(title) . "`",`"url`":`"" . EscapeJSON(_2Data) . "`"}"
            lastValidOutput := jsonOutput
            ; MsgBox jsonOutput
            FileAppend(jsonOutput "`n", "*", "UTF-8")
            _2Data := ""
            return
        } else {
            ddeData := GetBrowserURL_DDE(sClass)
            if !ddeData {
                ddeData := "new tab"
            }
            ; Format output as JSON
            jsonOutput := "{`"name`":`"" . EscapeJSON(name) . "`",`"displayName`":`"" . EscapeJSON(displayName) .
            "`",`"title`":`"" . EscapeJSON(title) . "`",`"url`":`"" . EscapeJSON(ddeData) . "`"}"
            lastValidOutput := jsonOutput
            ; MsgBox jsonOutput
            FileAppend(jsonOutput "`n", "*", "UTF-8")
            ddeData := ""
            return
        }
    }
    ; Format output as JSON with null url
    jsonOutput := "{`"name`":`"" . EscapeJSON(name) . "`",`"displayName`":`"" . EscapeJSON(displayName) .
    "`",`"title`":`"" . EscapeJSON(title) . "`",`"url`":null}"
    lastValidOutput := jsonOutput
    ; MsgBox jsonOutput
    FileAppend(jsonOutput "`n", "*", "UTF-8")
    return
}

; Walk windows in Z-order (top to bottom) and return the first visible, non-minimized, non-transient window
GetTopVisibleWindow() {
    global TransientClasses, TransientProcesses

    ; Start from the topmost window
    hw := DllCall("GetTopWindow", "Ptr", 0, "Ptr")  ; desktop's first child = topmost Z-order

    loop {
        if (!hw)
            break

        ; Check: visible and not minimized
        isVisible := DllCall("IsWindowVisible", "Ptr", hw, "Int")
        isMinimized := DllCall("IsIconic", "Ptr", hw, "Int")

        if (isVisible && !isMinimized) {
            try {
                wClass := WinGetClass("ahk_id " hw)
            } catch {
                wClass := ""
            }
            try {
                wTitle := WinGetTitle("ahk_id " hw)
            } catch {
                wTitle := ""
            }
            try {
                wName := WinGetProcessName("ahk_id " hw)
            } catch {
                wName := ""
            }

            ; Skip transient windows, empty titles, and the desktop shell window
            if (wTitle && wName
                && !InStr(TransientClasses, wClass)
                && !InStr(TransientProcesses, wName)
                && wClass != "Progman"         ; desktop
                && wClass != "WorkerW"         ; desktop wallpaper worker
                ) {
                return hw
            }
        }

        ; Next window in Z-order
        hw := DllCall("GetWindow", "Ptr", hw, "UInt", 2, "Ptr")  ; GW_HWNDNEXT = 2
    }

    return 0
}

;-------Function-------
GetTitle() {
    Title := WinGetTitle("A")
    return Title
}

GetText() {
    Text := WinGetText("A")
    return Text
}

GetName() {
    Active_ID := WinGetID("A")
    Active_Process := WinGetProcessName("ahk_id " Active_ID)
    return Active_Process
}

GetAccData(WinId := "A") {
    static w := Map(), wKey := 0, callCount := 0
    th := WinExist(WinId)

    ; Periodically clear stale cache entries (every 50 calls ~75 seconds)
    callCount++
    if (Mod(callCount, 50) = 0) {
        w := Map()
        wKey := 0
        callCount := 0
    }

    for i, v in w {
        if (th = v[1])
            return [GetAccObjectFromWindow(v[1]).accName(0), ParseAccData(v[4])[2]]
    }

    tr := ParseAccData(GetAccObjectFromWindow(th))

    ; Make sure tr is properly initialized as an array with at least 2 elements
    if !IsObject(tr) {
        tr := [0, 0]
    } else if tr.Length < 2 {
        tr.Push(0)
    }

    if tr[2] {
        wKey++
        w[wKey] := [th, tr[1], tr[2], tr[3]]
    }

    return [tr[1], tr[2]]
}

ParseAccData(accObj, accData := "") {
    ; Initialize accData as an array if not provided
    if (accData = "") {
        accData := [0, 0, 0] ; Pre-initialize with 3 elements
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
        if (!accData[2]) { ; Check if element 2 exists AND is empty
            children := GetAccChildren(accObj)
            if (IsObject(children)) {
                for _, accChild in children {
                    if (IsObject(accChild)) {
                        ParseAccData(accChild, accData)
                        if (accData[2]) { ; If we found a URL, stop processing
                            break
                        }
                    }
                }
            }
        }
    } catch {
        ; Do nothing if child processing fails
    }

    return accData
}

GetAccInit() {
    static hw := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
    return hw
}

GetAccObjectFromWindow(hWnd, idObject := 0) {
    static IID_IAccessible := "{618736E0-3C3D-11CF-810C-00AA00389B71}"

    ; Load oleacc.dll if needed
    if !DllCall("GetModuleHandle", "Str", "oleacc", "Ptr")
        DllCall("LoadLibrary", "Str", "oleacc", "Ptr")

    ; Create a GUID from the IID string
    GUID := Buffer(16, 0)
    DllCall("ole32\CLSIDFromString", "WStr", IID_IAccessible, "Ptr", GUID)

    ; Send WM_GETOBJECT message to the specific window, not the current active window
    SendMessage 0x003D, 0, 1, "Chrome_RenderWidgetHostHWND1", "ahk_id " hWnd ; WM_GETOBJECT

    ; Try to get the accessibility object
    pacc := 0
    loop 60 {
        ; Try to get the accessible object
        hr := DllCall("oleacc\AccessibleObjectFromWindow",
            "Ptr", hWnd,
            "UInt", idObject & 0xFFFFFFFF,
            "Ptr", GUID,
            "Ptr*", &pacc)

        if (hr = 0 && pacc != 0) ; S_OK and got an object
            break

        if (A_Index >= 60)
            return 0

        Sleep 30
    }

    if (pacc = 0)
        return 0

    return ComObjFromPtr(pacc)
}

GetAccQuery(objAcc) {
    try {
        if ComObjType(objAcc, "Name") != "IAccessible"
            return 0
        return ComObjQuery(objAcc, "{618736e0-3c3d-11cf-810c-00aa00389b71}")
    }
}

GetAccChildren(objAcc) {
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

            loop cChildren {
                i := (A_Index - 1) * (A_PtrSize * 2 + 8) + 8
                child := NumGet(varChildren, i, "Ptr")
                vt := NumGet(varChildren, i - 8, "UChar")

                ; Calculate pointer to the VARIANT structure in the buffer
                variantPtr := varChildren.Ptr + (A_Index - 1) * (A_PtrSize * 2 + 8)

                if (vt = 9 && child) { ; VT_DISPATCH and valid pointer
                    try {
                        childObj := ComObjFromPtr(child)
                        Children.Push(childObj)
                    } catch {
                        ; Skip this child if there's an error
                        ; Clear the VARIANT to release the COM object
                        DllCall("OleAut32\VariantClear", "Ptr", variantPtr)
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

; Get the real process name, handling UWP apps in ApplicationFrameHost
GetRealProcessName(hWnd) {
    processName := WinGetProcessName("ahk_id " hWnd)

    ; If it's ApplicationFrameHost, dig deeper to find the real UWP app
    if (processName = "ApplicationFrameHost.exe") {
        try {
            ; Method 1: Try to find CoreWindow child
            childHwnd := DllCall("FindWindowEx", "Ptr", hWnd, "Ptr", 0, "Str", "Windows.UI.Core.CoreWindow", "Ptr", 0,
                "Ptr")

            if (childHwnd) {
                childPid := 0
                DllCall("GetWindowThreadProcessId", "Ptr", childHwnd, "UInt*", &childPid)

                if (childPid) {
                    realName := GetProcessNameByPID(childPid)
                    if (realName && realName != "ApplicationFrameHost.exe")
                        return realName
                }
            }

            ; Method 2: Try ApplicationFrameWindow child
            childHwnd := DllCall("FindWindowEx", "Ptr", hWnd, "Ptr", 0, "Str", "ApplicationFrameWindow", "Ptr", 0,
                "Ptr")

            if (childHwnd) {
                childPid := 0
                DllCall("GetWindowThreadProcessId", "Ptr", childHwnd, "UInt*", &childPid)

                if (childPid) {
                    realName := GetProcessNameByPID(childPid)
                    if (realName && realName != "ApplicationFrameHost.exe")
                        return realName
                }
            }

            ; Method 3: Manually walk through child windows
            WinGetPID(&parentPid, "ahk_id " hWnd)

            childWnd := DllCall("GetWindow", "Ptr", hWnd, "UInt", 5, "Ptr") ; GW_CHILD = 5
            loop {
                if (!childWnd)
                    break

                childPid := 0
                DllCall("GetWindowThreadProcessId", "Ptr", childWnd, "UInt*", &childPid)

                if (childPid && childPid != parentPid) {
                    realName := GetProcessNameByPID(childPid)
                    if (realName && realName != "ApplicationFrameHost.exe")
                        return realName
                }

                ; Get next sibling
                childWnd := DllCall("GetWindow", "Ptr", childWnd, "UInt", 2, "Ptr") ; GW_HWNDNEXT = 2
            }

        } catch as e {
            ; If all methods fail, return ApplicationFrameHost
        }
    }

    return processName
}

; Helper function to get process name by PID
GetProcessNameByPID(pid) {
    try {
        hProcess := DllCall("OpenProcess", "UInt", 0x1000, "Int", 0, "UInt", pid, "Ptr")
        if (hProcess) {
            size := 260
            exeName := Buffer(size * 2, 0)
            if (DllCall("QueryFullProcessImageName", "Ptr", hProcess, "UInt", 0, "Ptr", exeName, "UInt*", &size)) {
                fullPath := StrGet(exeName, "UTF-16")
                DllCall("CloseHandle", "Ptr", hProcess)
                ; Extract just the filename
                SplitPath fullPath, &fileName
                return fileName
            }
            DllCall("CloseHandle", "Ptr", hProcess)
        }
    }
    return ""
}

GetAppDisplayName(processName) {
    try {
        ; Get the full path of the process
        for proc in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process where Name='" processName "'") {
            processPath := proc.ExecutablePath
            if (processPath) {
                ; Get the FileDescription from version info
                fileDesc := ""

                ; Get version info size
                size := DllCall("version\GetFileVersionInfoSize", "Str", processPath, "UInt*", 0, "UInt")
                if (size > 0) {
                    verInfo := Buffer(size)
                    if (DllCall("version\GetFileVersionInfo", "Str", processPath, "UInt", 0, "UInt", size, "Ptr",
                        verInfo)) {

                        ; Try to get the translation table to find available languages
                        pTranslate := 0
                        lenTranslate := 0

                        if (DllCall("version\VerQueryValue", "Ptr", verInfo, "Str", "\VarFileInfo\Translation", "Ptr*", &
                            pTranslate, "UInt*", &lenTranslate)) {
                            ; We have translations, try each one
                            numTranslations := lenTranslate // 4 ; Each translation is 4 bytes (WORD + WORD)

                            loop numTranslations {
                                offset := (A_Index - 1) * 4
                                lang := Format("{:04x}", NumGet(pTranslate + offset, "UShort"))
                                codepage := Format("{:04x}", NumGet(pTranslate + offset + 2, "UShort"))

                                ; Try to get FileDescription for this language/codepage
                                queryStr := "\StringFileInfo\" . lang . codepage . "\FileDescription"
                                pValue := 0
                                len := 0

                                if (DllCall("version\VerQueryValue", "Ptr", verInfo, "Str", queryStr, "Ptr*", &pValue,
                                    "UInt*", &len)) {
                                    fileDesc := StrGet(pValue, "UTF-16")
                                    if (fileDesc) {
                                        return fileDesc ; Return first valid description found
                                    }
                                }
                            }
                        }

                        ; Fallback: Try common language codes if translation table didn't work
                        commonLangCodes := ["040904b0", "040904e4", "000004b0", "000004e4"]
                        for code in commonLangCodes {
                            queryStr := "\StringFileInfo\" . code . "\FileDescription"
                            pValue := 0
                            len := 0

                            if (DllCall("version\VerQueryValue", "Ptr", verInfo, "Str", queryStr, "Ptr*", &pValue,
                                "UInt*", &len)) {
                                fileDesc := StrGet(pValue, "UTF-16")
                                if (fileDesc) {
                                    return fileDesc
                                }
                            }
                        }
                    }
                }
            }
            break ; Only check first instance
        }
    }
    ; Fallback: remove .exe extension
    return RegExReplace(processName, "\.exe$", "")
}

GetBrowserURL_DDE(sClass) {
    sServer := WinGetProcessName("ahk_class " sClass)
    sServer := SubStr(sServer, 1, StrLen(sServer) - 4)

    iCodePage := 0x04B0 ; Always use Unicode in v2

    idInst := 0
    DllCall("DdeInitialize", "UInt*", &idInst, "UInt", 0, "UInt", 0, "UInt", 0)

    hServer := DllCall("DdeCreateStringHandle", "UInt", idInst, "Str", sServer, "Int", iCodePage)
    hTopic := DllCall("DdeCreateStringHandle", "UInt", idInst, "Str", "WWW_GetWindowInfo", "Int", iCodePage)
    hItem := DllCall("DdeCreateStringHandle", "UInt", idInst, "Str", "0xFFFFFFFF", "Int", iCodePage)

    hConv := DllCall("DdeConnect", "UInt", idInst, "UInt", hServer, "UInt", hTopic, "UInt", 0)

    sData := ""
    if (hConv) {
        nResult := 0
        hData := DllCall("DdeClientTransaction", "Ptr", 0, "UInt", 0, "UInt", hConv, "UInt", hItem, "UInt", 1, "UInt",
            0x20B0, "UInt", 10000, "UInt*", &nResult)

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
