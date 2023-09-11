FileEncoding, UTF-8
#NoEnv
#SingleInstance, force
#NoTrayIcon
ModernBrowsers := "ApplicationFrameWindow,Chrome_WidgetWin_0,Chrome_WidgetWin_1,Maxthon3Cls_MainFrm,MozillaWindowClass,Slimjet_WidgetWin_1"
ModernbrowsersProcesses := "msedge.exe,iexplore.exe,chrome.exe,opera.exe,brave.exe,vivaldi.exe,MicrosoftEdge.exe"

SetTimer, SaveURL, 1800
return

SaveURl:
    WinGetClass, sClass, A

    title:=GetTitle()
    name:=GetName()
    If (name == "ApplicationFrameHost.exe") {
        name:=GetText()
    }

    if name In % ModernbrowsersProcesses
    {
        If sClass In % ModernBrowsers
        {
            accData := GetAccData()
            2Data := accData.2
            If !accData {
                Return
            }
            if !2Data {
                2Data := "new tab"
            }
            FileAppend,%name%`n%title%`n%2Data%`n,*
            ; MsgBox, %name%`n%title%`n%2Data%
            2Data :=""
            Return
        } else {
            ddeData := GetBrowserURL_DDE(sClass)
            If !ddeData {
                ddeData := "new tab"
            }
            FileAppend,%name%`n%title%`n%ddeData%`n,*
            ; MsgBox, %name%`n%title%`n%ddeData%
            ddeData :=""
            Return
        }
    }
    2Data := ""
    FileAppend,%name%`n%title%`n,*
    ; MsgBox, %name%`n%title%
Return

!F12::
    SetTimer, SaveURL, Off
Return

StopTimer:
    SetTimer, SaveURL, Off
Return

;-------Function-------
GetTitle() {
    WinGetTitle, Title, A
    Return Title
}

GetText() {
    WinGetText, Text, A
    Return Text
}

GetName() {
    WinGet, Active_ID, ID, A
    WinGet, Active_Process, ProcessName, ahk_id %Active_ID%
    return Active_Process
}
GetAccData(WinId:="A") { ;by RRR based on atnbueno's https://www.autohotkey.com/boards/viewtopic.php?f=6&t=3702
    Static w:= [], n:=0
    Global vtr1, vtr2
    th:= WinExist(WinId), GetKeyState("Ctrl", "P")? (w:= [], n:= 0): ""
    WinGet, bp, ProcessName, ahk_id %th%
    for i, v in w
        if (th=v.1)
            Return [ GetAccObjectFromWindow(v.1).accName(0), ParseAccData(v.4).2 ]
    Return ( [(tr:= ParseAccData(GetAccObjectFromWindow(th))).1, tr.2]
             , tr.2? ( w[++n]:= [], w[n].1:=th, w[n].2:=tr.1, w[n].3:=tr.2, w[n].4:=tr.3): "" ) ; save AccObj history
}

ParseAccData(accObj, accData:="") {
    try   accData? "": accData:= [accObj.accName(0)]
    try   if accObj.accRole(0) = 42 && accObj.accName(0) && accObj.accValue(0)
              accData.2:= SubStr(u:=accObj.accValue(0), 1, 4)="http"? u: "http://" u, accData.3:= accObj
          For nChild, accChild in GetAccChildren(accObj)
              accData.2? "": ParseAccData(accChild, accData)
          Return accData
}

GetAccInit() {
    Static hw:= DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
}

GetAccObjectFromWindow(hWnd, idObject = 0) {
    SendMessage, WM_GETOBJECT := 0x003D, 0, 1, Chrome_RenderWidgetHostHWND1, % "ahk_id " WinExist("A") ; by malcev
    While DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", idObject&=0xFFFFFFFF, "Ptr"
        , -VarSetCapacity(IID, 16)+NumPut(idObject==0xFFFFFFF0? 0x46000000000000C0: 0x719B3800AA000C81
        , NumPut(idObject==0xFFFFFFF0? 0x0000000000020400: 0x11CF3C3D618736E0, IID, "Int64"), "Int64"), "Ptr*", pacc)!=0
        && A_Index < 60
        sleep, 30
    Return ComObjEnwrap(9, pacc, 1)
}

GetAccQuery(objAcc) {
    Try Return ComObj(9, ComObjQuery(objAcc, "{618736e0-3c3d-11cf-810c-00aa00389b71}"), 1)
}

GetAccChildren(objAcc) {
    try If ComObjType(objAcc,"Name") != "IAccessible"
        ErrorLevel := "Invalid IAccessible Object"
    Else {
        cChildren:= objAcc.accChildCount, Children:= []
        If DllCall("oleacc\AccessibleChildren", "Ptr", ComObjValue(objAcc), "Int", 0, "Int", cChildren, "Ptr"
          , VarSetCapacity(varChildren, cChildren * (8 + 2 * A_PtrSize), 0) * 0 + &varChildren, "Int*", cChildren) = 0 {
            Loop, % cChildren {
                i:= (A_Index - 1) * (A_PtrSize * 2 + 8) + 8, child:= NumGet(varChildren, i)
                Children.Insert(NumGet(varChildren, i - 8) = 9? GetAccQuery(child): child)
                NumGet(varChildren, i - 8) != 9? "": ObjRelease(child)
            }   Return (Children.MaxIndex()? Children: "")
        }   Else ErrorLevel := "AccessibleChildren DllCall Failed"
    }
}

GetBrowserURL_DDE(sClass) {
    WinGet, sServer, ProcessName, % "ahk_class " sClass
    StringTrimRight, sServer, sServer, 4
    iCodePage := A_IsUnicode ? 0x04B0 : 0x03EC ; 0x04B0 = CP_WINUNICODE, 0x03EC = CP_WINANSI
    DllCall("DdeInitialize", "UPtrP", idInst, "Uint", 0, "Uint", 0, "Uint", 0)
    hServer := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", sServer, "int", iCodePage)
    hTopic := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "WWW_GetWindowInfo", "int", iCodePage)
    hItem := DllCall("DdeCreateStringHandle", "UPtr", idInst, "Str", "0xFFFFFFFF", "int", iCodePage)
    hConv := DllCall("DdeConnect", "UPtr", idInst, "UPtr", hServer, "UPtr", hTopic, "Uint", 0)
    hData := DllCall("DdeClientTransaction", "Uint", 0, "Uint", 0, "UPtr", hConv, "UPtr", hItem, "UInt", 1, "Uint", 0x20B0, "Uint", 10000, "UPtrP", nResult) ; 0x20B0 = XTYP_REQUEST, 10000 = 10s timeout
    sData := DllCall("DdeAccessData", "Uint", hData, "Uint", 0, "Str")
    DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hServer)
    DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hTopic)
    DllCall("DdeFreeStringHandle", "UPtr", idInst, "UPtr", hItem)
    DllCall("DdeUnaccessData", "UPtr", hData)
    DllCall("DdeFreeDataHandle", "UPtr", hData)
    DllCall("DdeDisconnect", "UPtr", hConv)
    DllCall("DdeUninitialize", "UPtr", idInst)
    csvWindowInfo := StrGet(&sData, "CP0")
    StringSplit, sWindowInfo, csvWindowInfo, `" ;"; comment to avoid a syntax highlighting issue in autohotkey.com/boards
    Return sWindowInfo2
}
