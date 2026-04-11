; ============================================
; Macro Recorder Pro - Multi-Macro Support
; AutoHotkey v1.1
; ============================================

#SingleInstance Force
#NoTrayIcon
SetBatchLines, -1
SetKeyDelay, -1, -1
SetMouseDelay, -1
SetControlDelay, -1
CoordMode, Mouse, Screen

OnMessage(0x4A, "Receive_WM_COPYDATA")  ; Hook for Python IPC

; ============================================
; CONFIGURATION
; ============================================

Config_ProfileDir := A_ScriptDir "\Profiles"
Config_LastProfileFile := A_ScriptDir "\lastprofile.txt"
ifNotExist, %Config_ProfileDir%
    FileCreateDir, %Config_ProfileDir%

; ============================================
; GLOBAL STATE
; ============================================

Macros := []               ; Array of macro objects
HotkeyMap := {}            ; hotkey string -> macro index lookup
SelectedIdx := 1           ; Currently selected macro index
IsRecording := 0
MacrosEnabled := 1         ; Master enable/disable toggle
StopKeyName := "``"        ; Default stop-recording key (backtick)
BoundStopKey := ""
PanicKeyName := "End"      ; Default panic exit key
BoundPanicKey := ""
MasterHotkeyName := ""     ; Hotkey to toggle enable/disable
BoundMasterHotkey := ""
BlockMouseState := 0       ; Block mouse during playback
LoopDelay := 500           ; Default loop delay between repetitions
CurrentProfile := "Default"
overlayHwnd := 0
playHwnd := 0
NextMacroId := 2
BlockMouseActive := 0      ; Global flag for hotkey click blocking

; Seed first macro
m := {}
m.Name := "Macro 1"
m.Hotkey := "F1"
m.Actions := ""
m.IsPlaying := false
m.LimitLoops := false      
m.LoopLimit := 5           
m.CurrentLoop := 0         
Macros.Push(m)

; ============================================
; GUI
; ============================================

Gui, Main:New,, Config Manager
Gui, Main:Margin, 10, 10

; --- Profile ---
Gui, Main:Add, GroupBox, x10 y10 w460 h75, Profile
Gui, Main:Add, Text, x20 y33, Profile:
Gui, Main:Add, DropDownList, x65 y30 w135 vProfileDD gOnProfileChange, Default||
Gui, Main:Add, Button, x255 y30 w50 h23 gNewProfileBtn, New
Gui, Main:Add, Button, x310 y30 w50 h23 gSaveProfileBtn, Save

Gui, Main:Add, Text, x20 y60, Toggle Key:
Gui, Main:Add, Edit, x80 y57 w50 vMasterHotkeyInput, 
Gui, Main:Add, Button, x135 y57 w30 h23 gApplyMasterHotkey, Set
Gui, Main:Add, Checkbox, x250 y59 w110 vOverlayToggle gOnOverlayToggle Checked, Overlay Enabled?
Gui, Main:Add, Checkbox, x375 y59 w70 vEnableToggle gOnEnableToggle Checked, Enable

; --- Macro Settings ---
Gui, Main:Add, GroupBox, x10 y95 w460 h220, Macro Settings
Gui, Main:Add, Text, x20 y118, Macro:
Gui, Main:Add, DropDownList, x65 y115 w145 vMacroDD gOnMacroSelect, Macro 1||
Gui, Main:Add, Button, x215 y115 w40 h23 gAddMacroBtn, +Add
Gui, Main:Add, Button, x259 y115 w50 h23 gDeleteMacroBtn, Delete
Gui, Main:Add, Button, x313 y115 w55 h23 gRenameMacroBtn, Rename

Gui, Main:Add, Text, x20 y150, Play Key:
Gui, Main:Add, Edit, x75 y147 w50 vPlayKeyInput, F1
Gui, Main:Add, Button, x130 y147 w30 h23 gApplyPlayKey, Set

Gui, Main:Add, Text, x170 y150, Stop Rec Key:
Gui, Main:Add, Edit, x245 y147 w40 vStopKeyInput, ``
Gui, Main:Add, Button, x290 y147 w30 h23 gOnStopKeyChange, Set

Gui, Main:Add, Text, x325 y150, Panic Exit:
Gui, Main:Add, Edit, x385 y147 w40 vPanicKeyInput, End
Gui, Main:Add, Button, x430 y147 w30 h23 gOnPanicKeyChange, Set

Gui, Main:Add, Checkbox, x20 y180 vBlockMouseCB gOnBlockMouseCB, Block Mouse Input
Gui, Main:Add, Text, x310 y180, Loop Delay:
Gui, Main:Add, Edit, x370 y177 w35 vLoopDelayMin, 500
Gui, Main:Add, Text, x408 y180, -
Gui, Main:Add, Edit, x418 y177 w35 vLoopDelayMax, 
Gui, Main:Add, Text, x310 y198 cGray, (Leave 2nd box blank for static)

; New Loop Limit Controls
Gui, Main:Add, Checkbox, x20 y210 vLimitLoopsCB gOnLimitLoopsCB, Limit Loops to:
Gui, Main:Add, Edit, x125 y207 w40 vLoopLimitInput gOnLoopLimitChange, 5
Gui, Main:Add, Text, x170 y210 cGray, times

Gui, Main:Add, Button, x20 y245 w80 h30 gStartRecording, Record
Gui, Main:Add, Button, x105 y245 w80 h30 gStopRecording, Stop Rec
Gui, Main:Add, Button, x190 y245 w80 h30 gPlayMacroBtn, Play
Gui, Main:Add, Button, x275 y245 w80 h30 gStopAllBtn, Stop All

; --- Editor ---
Gui, Main:Add, GroupBox, x10 y325 w460 h195, Editor (Auto-applied)  (Use format: delay: 10ms  OR  delay: random(10,50)ms)
Gui, Main:Add, Edit, x20 y345 w440 h140 vEditBox Multi gApplyEditorBtn
; Button removed

; --- Status ---
Gui, Main:Add, Text, x10 y520 w460 cGreen vStatusText, Ready

Gui, Main:Show, w480 h545 Center, Config Manager

; --- Overlay ---
Gui, Overlay:New, +AlwaysOnTop -Caption +ToolWindow +LastFound +E0x20
overlayHwnd := WinExist()
Gui, Overlay:Color, Red
Gui, Overlay:Font, s24 bold
Gui, Overlay:Add, Text, x20 y20 cWhite BackgroundTrans, RECORDING
Gui, Overlay:Add, Text, x20 y60 cYellow BackgroundTrans vOverlayStatus, Press stop key...
Gui, Overlay:Show, Hide x0 y0 w%A_ScreenWidth% h%A_ScreenHeight% NoActivate
WinSet, Transparent, 40, ahk_id %overlayHwnd%
Gui, Overlay:Hide

; --- PlayOverlay ---
Gui, PlayOverlay:New, +AlwaysOnTop -Caption +ToolWindow +LastFound +E0x20
playHwnd := WinExist()
Gui, PlayOverlay:Color, Red
Gui, PlayOverlay:Show, Hide x0 y0 w%A_ScreenWidth% h%A_ScreenHeight% NoActivate
WinSet, Transparent, 30, ahk_id %playHwnd%
Gui, PlayOverlay:Hide

; --- ProfileOSD ---
Gui, ProfileOSD:New, +AlwaysOnTop -Caption +ToolWindow +LastFound +E0x20
osdHwnd := WinExist()
Gui, ProfileOSD:Color, Black
Gui, ProfileOSD:Font, s48 bold, Arial
Gui, ProfileOSD:Add, Text, w%A_ScreenWidth% Center cLime vOSDText BackgroundTrans, Profile: Default
Gui, ProfileOSD:Show, Hide x0 y100 w%A_ScreenWidth% NoActivate
WinSet, TransColor, Black 200, ahk_id %osdHwnd%
Gui, ProfileOSD:Hide

; --- EnabledOSD --- (Top Left)
Gui, EnabledOSD:New, +AlwaysOnTop -Caption +ToolWindow +LastFound +E0x20
enabledOSDHwnd := WinExist()
Gui, EnabledOSD:Color, Black
Gui, EnabledOSD:Font, s18 bold, Arial
Gui, EnabledOSD:Add, Text, x10 y10 w500 c0088CC vEnabledOSDText BackgroundTrans, %CurrentProfile%
Gui, EnabledOSD:Show, x0 y0 Hide NoActivate
WinSet, TransColor, Black, ahk_id %enabledOSDHwnd%
Gui, EnabledOSD:Hide

; --- Initial bindings ---
GoSub, BindAllHotkeys
GoSub, BindStopKey
GoSub, LoadLastProfile
GoSub, OnEnableToggle
return

; ============================================
; WINDOW EVENTS
; ============================================

MainGuiClose:
MainGuiEscape:
ExitApp
return

LoadLastProfile:
if FileExist(Config_LastProfileFile)
{
    FileRead, lastP, %Config_LastProfileFile%
    lastP := Trim(lastP)
    if (lastP != "")
        CurrentProfile := lastP
}

GoSub, RefreshProfileList
GoSub, LoadProfileBtn
return

RefreshProfileList:
profileList := ""
Loop, Files, %Config_ProfileDir%\*.ini
{
    SplitPath, A_LoopFileName,,,, noExt
    if (noExt = CurrentProfile)
        profileList .= noExt "||"
    else
        profileList .= noExt "|"
}
if (profileList = "")
    profileList := CurrentProfile "||"
GuiControl, Main:, ProfileDD, |%profileList%
return

; ============================================
; ENABLE / DISABLE TOGGLE
; ============================================

OnEnableToggle:
Gui, Main:Submit, NoHide
MacrosEnabled := EnableToggle
if (MacrosEnabled)
{
    GoSub, BindAllHotkeys
    GoSub, BindStopKey
    GoSub, BindPanicKey
    GuiControl, Main:, StatusText, Macros enabled. All hotkeys rebound.
    if (OverlayToggle)
        Gui, EnabledOSD:Show, NoActivate
}
else
{
    GoSub, UnbindAllHotkeys
    GoSub, UnbindStopKey
    GoSub, UnbindPanicKey
    GuiControl, Main:, StatusText, Macros disabled. All keys restored to normal.
    Gui, EnabledOSD:Hide
}
return

OnOverlayToggle:
Gui, Main:Submit, NoHide
if (MacrosEnabled && OverlayToggle)
    Gui, EnabledOSD:Show, NoActivate
else
    Gui, EnabledOSD:Hide
return

; ============================================
; PROFILE MANAGEMENT
; ============================================

OnProfileChange:
Gui, Main:Submit, NoHide
CurrentProfile := ProfileDD
GoSub, LoadProfileBtn
return

NewProfileBtn:
InputBox, newProfName, New Profile, Enter profile name:,, 280, 130
if (ErrorLevel || newProfName = "")
    return
; Stop everything
GoSub, StopAllBtn
GoSub, UnbindAllHotkeys

CurrentProfile := newProfName
; Reset macros to a single empty one
Macros := []
m := {}
m.Name := "Macro 1"
m.Hotkey := "F1"
m.Actions := ""
m.IsPlaying := false
m.LimitLoops := false
m.LoopLimit := 5
m.CurrentLoop := 0
Macros.Push(m)
SelectedIdx := 1
NextMacroId := 2

GoSub, RebuildMacroDD
GoSub, LoadMacroToGUI
GoSub, BindAllHotkeys
; Add to and select in profile dropdown
GuiControl, Main:, ProfileDD, %newProfName%||
GuiControl, Main:, StatusText, Profile "%newProfName%" created.
GuiControl, EnabledOSD:, EnabledOSDText, %CurrentProfile%
return

SaveProfileBtn:
GoSub, SaveCurrentToMacro
profilePath := Config_ProfileDir "\" CurrentProfile ".ini"
FileDelete, %profilePath%

; General settings
IniWrite, % Macros.Length(), %profilePath%, General, MacroCount
IniWrite, %StopKeyName%, %profilePath%, General, StopKey
IniWrite, %PanicKeyName%, %profilePath%, General, PanicKey
IniWrite, %MasterHotkeyName%, %profilePath%, General, MasterHotkey
IniWrite, %OverlayToggle%, %profilePath%, General, OverlayEnabled

Gui, Main:Submit, NoHide
IniWrite, %BlockMouseCB%, %profilePath%, General, BlockMouse
IniWrite, %LoopDelayMin%, %profilePath%, General, LoopDelayMin
IniWrite, %LoopDelayMax%, %profilePath%, General, LoopDelayMax

FileDelete, %Config_LastProfileFile%
FileAppend, %CurrentProfile%, %Config_LastProfileFile%

; Each macro
Loop, % Macros.Length()
{
    section := "Macro" A_Index
    mc := Macros[A_Index]
    IniWrite, % mc.Name, %profilePath%, %section%, Name
    IniWrite, % mc.Hotkey, %profilePath%, %section%, Hotkey
    IniWrite, % mc.Actions, %profilePath%, %section%, Actions
    IniWrite, % mc.LimitLoops, %profilePath%, %section%, LimitLoops
    IniWrite, % mc.LoopLimit, %profilePath%, %section%, LoopLimit
}
GoSub, RefreshProfileList
GuiControl, Main:, StatusText, Profile "%CurrentProfile%" saved.
return

LoadProfileBtn:
Gui, Main:Submit, NoHide
profilePath := Config_ProfileDir "\" CurrentProfile ".ini"
if (!FileExist(profilePath))
{
    GuiControl, Main:, StatusText, No saved file for "%CurrentProfile%".
    return
}
GoSub, StopAllBtn
GoSub, UnbindAllHotkeys

IniRead, macroCount, %profilePath%, General, MacroCount, 0
IniRead, sk, %profilePath%, General, StopKey, ``
IniRead, pk, %profilePath%, General, PanicKey, End
IniRead, bm, %profilePath%, General, BlockMouse, 0
IniRead, ldMin, %profilePath%, General, LoopDelayMin, 500
IniRead, ldMax, %profilePath%, General, LoopDelayMax, % ""
IniRead, mhk, %profilePath%, General, MasterHotkey, % ""
IniRead, ovl, %profilePath%, General, OverlayEnabled, 1

StopKeyName := sk
PanicKeyName := pk
BlockMouseState := (bm = 1) ? 1 : 0

GuiControl, Main:, StopKeyInput, %StopKeyName%
GuiControl, Main:, PanicKeyInput, %PanicKeyName%
GuiControl, Main:, BlockMouseCB, %BlockMouseState%
GuiControl, Main:, LoopDelayMin, %ldMin%
GuiControl, Main:, LoopDelayMax, %ldMax%
GuiControl, Main:, MasterHotkeyInput, %mhk%
GuiControl, Main:, OverlayToggle, %ovl%

FileDelete, %Config_LastProfileFile%
FileAppend, %CurrentProfile%, %Config_LastProfileFile%

Macros := []
Loop, %macroCount%
{
    section := "Macro" A_Index
    mc := {}
    IniRead, val, %profilePath%, %section%, Name, Macro %A_Index%
    mc.Name := val
    IniRead, val, %profilePath%, %section%, Hotkey, % ""
    mc.Hotkey := val
    IniRead, val, %profilePath%, %section%, Actions, % ""
    mc.Actions := val
    IniRead, val, %profilePath%, %section%, LimitLoops, 0
    mc.LimitLoops := (val = 1)
    IniRead, val, %profilePath%, %section%, LoopLimit, 5
    mc.LoopLimit := val
    mc.IsPlaying := false
    mc.CurrentLoop := 0
    Macros.Push(mc)
}

if (Macros.Length() = 0)
{
    mc := {}
    mc.Name := "Macro 1"
    mc.Hotkey := "F1"
    mc.Actions := ""
    mc.IsPlaying := false
    mc.LimitLoops := false
    mc.LoopLimit := 5
    mc.CurrentLoop := 0
    Macros.Push(mc)
}

SelectedIdx := 1
NextMacroId := Macros.Length() + 1
GoSub, RebuildMacroDD
GoSub, LoadMacroToGUI
GoSub, BindStopKey
GoSub, BindPanicKey
GoSub, ApplyMasterHotkey
GoSub, BindAllHotkeys
GuiControl, Main:, StatusText, Profile "%CurrentProfile%" loaded (%macroCount% macros).

GuiControl, EnabledOSD:, EnabledOSDText, %CurrentProfile%
GuiControl, ProfileOSD:, OSDText, Profile: %CurrentProfile%
Gui, ProfileOSD:Show, NoActivate x0 y100
SetTimer, HideProfileOSD, -3000
return

HideProfileOSD:
Gui, ProfileOSD:Hide
return

; ============================================
; MACRO ADD / DELETE / RENAME
; ============================================

AddMacroBtn:
GoSub, SaveCurrentToMacro
newName := "Macro " NextMacroId
NextMacroId++
mc := {}
mc.Name := newName
mc.Hotkey := ""
mc.Actions := ""
mc.IsPlaying := false
mc.LimitLoops := false
mc.LoopLimit := 5
mc.CurrentLoop := 0
Macros.Push(mc)

SelectedIdx := Macros.Length()
GoSub, RebuildMacroDD
GoSub, LoadMacroToGUI
GoSub, BindAllHotkeys
GuiControl, Main:, StatusText, Added "%newName%". Set a play key for it.
return

DeleteMacroBtn:
if (Macros.Length() <= 1)
{
    GuiControl, Main:, StatusText, Can't delete the last macro.
    return
}
delName := Macros[SelectedIdx].Name
; Stop it if playing
Macros[SelectedIdx].IsPlaying := false

Macros.RemoveAt(SelectedIdx)
if (SelectedIdx > Macros.Length())
    SelectedIdx := Macros.Length()

GoSub, RebuildMacroDD
GoSub, LoadMacroToGUI
GoSub, BindAllHotkeys
GuiControl, Main:, StatusText, Deleted "%delName%".
return

RenameMacroBtn:
oldName := Macros[SelectedIdx].Name
InputBox, newName, Rename Macro, New name for "%oldName%":,, 280, 130,,,,, %oldName%
if (ErrorLevel || newName = "")
    return
Macros[SelectedIdx].Name := newName
GoSub, RebuildMacroDD
GuiControl, Main:, StatusText, Renamed to "%newName%".
return

; ============================================
; MACRO SELECTION / GUI SYNC
; ============================================

OnMacroSelect:
Gui, Main:Submit, NoHide
; Save outgoing macro's settings
GoSub, SaveCurrentToMacro
; Find index by name
Loop, % Macros.Length()
{
    if (Macros[A_Index].Name = MacroDD)
    {
        SelectedIdx := A_Index
        break
    }
}
GoSub, LoadMacroToGUI
return

OnBlockMouseCB:
Gui, Main:Submit, NoHide
BlockMouseState := BlockMouseCB
return

OnLimitLoopsCB:
Gui, Main:Submit, NoHide
if (SelectedIdx >= 1 && SelectedIdx <= Macros.Length())
    Macros[SelectedIdx].LimitLoops := LimitLoopsCB
return

OnLoopLimitChange:
Gui, Main:Submit, NoHide
if (SelectedIdx >= 1 && SelectedIdx <= Macros.Length())
    Macros[SelectedIdx].LoopLimit := LoopLimitInput
return

RebuildMacroDD:
list := ""
Loop, % Macros.Length()
    list .= "|" Macros[A_Index].Name
GuiControl, Main:, MacroDD, %list%
GuiControl, Main:Choose, MacroDD, %SelectedIdx%
return

LoadMacroToGUI:
if (SelectedIdx < 1 || SelectedIdx > Macros.Length())
    return
mc := Macros[SelectedIdx]
GuiControl, Main:, PlayKeyInput, % mc.Hotkey
GuiControl, Main:, LimitLoopsCB, % mc.LimitLoops
GuiControl, Main:, LoopLimitInput, % mc.LoopLimit
readable := FormatForDisplay(mc.Actions)
GuiControl, Main:, EditBox, %readable%
return

SaveCurrentToMacro:
if (SelectedIdx < 1 || SelectedIdx > Macros.Length())
    return
Gui, Main:Submit, NoHide
Macros[SelectedIdx].LimitLoops := LimitLoopsCB
Macros[SelectedIdx].LoopLimit := LoopLimitInput
return

; ============================================
; MASTER HOTKEY BINDING
; ============================================

ApplyMasterHotkey:
Gui, Main:Submit, NoHide
newKey := MasterHotkeyInput
newKey := Trim(newKey)
if (newKey = A_Space)
    newKey := "Space"
GuiControl, Main:, MasterHotkeyInput, %newKey%

; Unbind old
if (BoundMasterHotkey != "")
{
    try
    {
        ; We use the bound key string to unbind
        if (InStr(BoundMasterHotkey, " "))
            oldHk := "~*" BoundMasterHotkey
        else
            oldHk := "*" BoundMasterHotkey
        Hotkey, %oldHk%, MasterHotkeyFired, Off
    }
    BoundMasterHotkey := ""
}

MasterHotkeyName := newKey
if (MasterHotkeyName = "")
{
    GuiControl, Main:, StatusText, Master hotkey cleared.
    return
}

try
{
    if (InStr(MasterHotkeyName, " "))
        newHk := "~*" MasterHotkeyName
    else
        newHk := "*" MasterHotkeyName
        
    Hotkey, %newHk%, MasterHotkeyFired, On
    BoundMasterHotkey := MasterHotkeyName
    GuiControl, Main:, StatusText, Master hotkey "%MasterHotkeyName%" bound.
}
catch e
{
    GuiControl, Main:, StatusText, Failed to bind master hotkey "%MasterHotkeyName%".
}
return

MasterHotkeyFired:
Gui, Main:Default
GuiControlGet, currentEnable,, EnableToggle
newEnable := !currentEnable
GuiControl, Main:, EnableToggle, %newEnable%
GoSub, OnEnableToggle
return

; ============================================
; PLAY KEY BINDING
; ============================================

ApplyPlayKey:
Gui, Main:Submit, NoHide
newKey := PlayKeyInput
if (newKey = A_Space)
    newKey := "Space"
else
    newKey := Trim(newKey)
GuiControl, Main:, PlayKeyInput, %newKey%
if (newKey = "")
{
    Macros[SelectedIdx].Hotkey := ""
    GoSub, BindAllHotkeys
    GuiControl, Main:, StatusText, Play key cleared.
    return
}
if (SelectedIdx < 1 || SelectedIdx > Macros.Length())
    return
; Check for duplicates
Loop, % Macros.Length()
{
    if (A_Index != SelectedIdx && Macros[A_Index].Hotkey = newKey && newKey != "")
    {
        dupe := Macros[A_Index].Name
        GuiControl, Main:, StatusText, Key "%newKey%" already used by "%dupe%"!
        return
    }
}
Macros[SelectedIdx].Hotkey := newKey
GoSub, BindAllHotkeys
; Verify it actually bound
if (HotkeyMap.HasKey(newKey))
{
    tempName := Macros[SelectedIdx].Name
    GuiControl, Main:, StatusText, Play key "%newKey%" bound to "%tempName%".
}
else
{
    GuiControl, Main:, StatusText, Failed to bind "%newKey%". Try a key name like F1, F2, z, Space, etc.
}
return

; ============================================
; HOTKEY BINDING (ALL MACROS)
; ============================================

BindAllHotkeys:
; Unbind everything first
GoSub, UnbindAllHotkeys

if (!MacrosEnabled)
    return

; Ensure we create hotkeys in the global context (no #If condition)
; This prevents conflict with the #If IsRecording script hotkeys
Hotkey, If

; Bind each macro's unique hotkey
boundCount := 0
Loop, % Macros.Length()
{
    hk := Macros[A_Index].Hotkey
    hk := Trim(hk)
    if (hk = "")
        continue
    if (HotkeyMap.HasKey(hk))
        continue  ; skip duplicates
    try
    {
        if (InStr(hk, " "))
            boundHk := "~*" hk
        else
            boundHk := "*" hk
            
        Hotkey, %boundHk%, MacroHotkeyFired, On
        HotkeyMap[hk] := boundHk
        boundCount++
    }
    catch e
    {
        tempName := Macros[A_Index].Name
        GuiControl, Main:, StatusText, Failed to bind "%hk%" for "%tempName%".
    }
}
return

UnbindAllHotkeys:
for hk, boundHk in HotkeyMap
{
    try
        Hotkey, %boundHk%, MacroHotkeyFired, Off
}
HotkeyMap := {}
Gui, PlayOverlay:Hide
; Stop all playing macros
Loop, % Macros.Length()
    Macros[A_Index].IsPlaying := false
SetTimer, PlayMacroLoop, Off
return

#MaxThreadsPerHotkey 3
MacroHotkeyFired:
pressedKey := A_ThisHotkey
; Strip AHK modifiers (~, *, $, +, ^, #, !) that might be attached 
cleanKey := RegExReplace(pressedKey, "^[~*$+^#!]+")
StringLower, pressedKeyLower, cleanKey
idx := 0
Loop, % Macros.Length()
{
    hk := Macros[A_Index].Hotkey
    StringLower, hkLower, hk
    if (hkLower == pressedKeyLower)
    {
        idx := A_Index
        break
    }
}

if (!idx || idx < 1 || idx > Macros.Length())
{
    ; Build debug string
    debugStr := ""
    Loop, % Macros.Length()
        debugStr .= "M" A_Index ":" Macros[A_Index].Hotkey " "
    
    ToolTip, Hotkey "%pressedKey%" fired. No mapped macro.`nCurrently mapped: %debugStr%
    SetTimer, ClearToolTip, -4000
    return
}
toggleMacro(idx)

if (cleanKey != "")
    KeyWait, %cleanKey%, T0.2
return

ClearToolTip:
ToolTip
return

; ============================================
; PLAYBACK CONTROL
; ============================================

toggleMacro(idx) {
    global Macros, BlockMouseState, BlockMouseActive
    mc := Macros[idx]

    if (mc.Actions = "")
    {
        tempName := mc.Name
        GuiControl, Main:, StatusText, "%tempName%" has no recorded actions.
        return
    }

    if (mc.IsPlaying)
    {
        Macros[idx].IsPlaying := false
        tempName := mc.Name
        GuiControl, Main:, StatusText, Stopped "%tempName%".
        
        ; If no macros are playing anymore, hide the overlay and block
        anyLeft := false
        Loop, % Macros.Length()
            if (Macros[A_Index].IsPlaying)
                anyLeft := true
        if (!anyLeft)
        {
            Gui, PlayOverlay:Hide
            BlockInput, MouseMoveOff
            BlockMouseActive := 0
            ReleaseStuckInputs()
        }
    }
    else
    {
        Macros[idx].IsPlaying := true
        Macros[idx].CurrentLoop := 0 ; Reset counter on start
        SetTimer, PlayMacroLoop, -1
        tempName := mc.Name
        GuiControl, Main:, StatusText, Playing "%tempName%"...
        
        ; Show red play overlay (very transparent, click-through)
        Gui, PlayOverlay:Show, NoActivate
        
        ; Enable mouse blocking immediately if requested
        if (BlockMouseState)
        {
            BlockInput, MouseMove
            BlockMouseActive := 1
        }
    }
}

PlayMacroBtn:
if (SelectedIdx < 1 || SelectedIdx > Macros.Length())
    return
idx := SelectedIdx
toggleMacro(idx)
return

StopAllBtn:
Loop, % Macros.Length()
    Macros[A_Index].IsPlaying := false
GuiControl, Main:, StatusText, All macros stopped.
Gui, PlayOverlay:Hide
BlockInput, MouseMoveOff
BlockMouseActive := 0
ReleaseStuckInputs()
return

; ============================================
; PLAYBACK LOOP (services all playing macros)
; ============================================

PlayMacroLoop:
Gui, Main:Default
fastLoopCounter := 0
Loop
{
    fastLoopCounter++
    anyPlaying := false
    
    ; Update block state at start of loop
    if (BlockMouseState)
    {
        BlockInput, MouseMove
        BlockMouseActive := 1
    }
    else
    {
        BlockInput, MouseMoveOff
        BlockMouseActive := 0
    }
    
    GuiControlGet, ldMin,, LoopDelayMin
    GuiControlGet, ldMax,, LoopDelayMax
    
    ldm1 := Trim(ldMin)
    ldm2 := Trim(ldMax)
    
    Loop, % Macros.Length()
    {
        if (Macros[A_Index].IsPlaying)
        {
            ExecuteMacro(A_Index)
            
            ; Increment loop count if limit enabled
            if (Macros[A_Index].LimitLoops)
            {
                Macros[A_Index].CurrentLoop++
                if (Macros[A_Index].CurrentLoop >= Macros[A_Index].LoopLimit)
                {
                    Macros[A_Index].IsPlaying := false
                }
            }

            ; Only count as still playing if it wasn't just stopped
            if (Macros[A_Index].IsPlaying)
            {
                anyPlaying := true
                if (ldm1 != "")
                {
                    if (ldm2 != "")
                        Random, rndWait, %ldm1%, %ldm2%
                    else
                        rndWait := ldm1
                        
                    ResponsiveSleep(rndWait, A_Index)
                }
            }
        }
    }
    
    
    if (!anyPlaying)
    {
        BlockInput, MouseMoveOff
        BlockMouseActive := 0
        Gui, PlayOverlay:Hide
        ReleaseStuckInputs()
        break
    }
    
    ; Yield to message pump every 5 loops to prevent hotkey starvation
    if (Mod(fastLoopCounter, 5) = 0)
        Sleep, -1
}
return

ResponsiveSleep(ms, idx) {
    global Macros
    if (ms < 100) { ; Use native sleep for high-speed delays
        if (ms > 0)
            Sleep, %ms%
        return
    }
    
    stopAt := A_TickCount + ms
    while (A_TickCount < stopAt) {
        if (!Macros[idx].IsPlaying)
            return
            
        remaining := stopAt - A_TickCount
        sleepTime := (remaining > 50) ? 50 : remaining
        if (sleepTime > 0)
            Sleep, %sleepTime%
    }
}

ExecuteMacro(idx) {
    global Macros
    actions := Macros[idx].Actions
    actLoopCnt := 0
    Loop, Parse, actions, |
    {
        actLoopCnt++
        if (Mod(actLoopCnt, 20) = 0)
            Sleep, -1
            
        ; Bail if stopped mid-execution
        if (!Macros[idx].IsPlaying)
            return

        parts := StrSplit(A_LoopField, ":")
        delay := 0
        type := parts[1]
        
        if (InStr(type, "KEY") || InStr(type, "MOUSE"))
        {
            delay := InStr(type, "KEY") ? parts[3] : parts[5]
            
            if InStr(delay, "-")
            {
                dp := StrSplit(delay, "-")
                d1 := Trim(dp[1])
                d2 := Trim(dp[2])
                Random, rDelay, %d1%, %d2%
                delay := rDelay
            }
            
            if (delay > 0)
            {
                ResponsiveSleep(delay, idx)
                if (!Macros[idx].IsPlaying)
                    return
            }
        }

        if (type = "KEY")
        {
            key := parts[2]
            Send, {%key%}
        }
        else if (type = "KEY_DOWN")
        {
            key := parts[2]
            Send, {%key% down}
        }
        else if (type = "KEY_UP")
        {
            key := parts[2]
            Send, {%key% up}
        }
        else if (type = "MOUSE")
        {
            btn := parts[2]
            x := parts[3]
            y := parts[4]
            MouseMove, %x%, %y%, 0
            if (btn = "L")
                Click, %x%, %y%
            else
                Click, right, %x%, %y%
        }
        else if (type = "MOUSE_DOWN")
        {
            btn := parts[2]
            x := parts[3]
            y := parts[4]
            MouseMove, %x%, %y%, 0
            if (btn = "L")
                Click, down, %x%, %y%
            else
                Click, right down, %x%, %y%
        }
        else if (type = "MOUSE_UP")
        {
            btn := parts[2]
            x := parts[3]
            y := parts[4]
            MouseMove, %x%, %y%, 0
            if (btn = "L")
                Click, up, %x%, %y%
            else
                Click, right up, %x%, %y%
        }
    }
}

; ============================================
; STOP-RECORDING KEY
; ============================================

OnStopKeyChange:
Gui, Main:Submit, NoHide
StopKeyName := StopKeyInput
GoSub, BindStopKey
return

BindStopKey:
GoSub, UnbindStopKey
if (!MacrosEnabled)
    return
if (StopKeyName != "")
{
    try
    {
        bound := "*" StopKeyName
        Hotkey, %bound%, StopRecordingHotkey, On
        BoundStopKey := bound
    }
    catch e
    {
        GuiControl, Main:, StatusText, Bad stop key: %StopKeyName%
    }
}
return

UnbindStopKey:
if (BoundStopKey != "")
{
    try
        Hotkey, %BoundStopKey%, StopRecordingHotkey, Off
    BoundStopKey := ""
}
return

StopRecordingHotkey:
if (IsRecording)
    GoSub, StopRecording
    
prKey := RegExReplace(A_ThisHotkey, "^[~*$+^#!]+")
if (prKey != "")
    KeyWait, %prKey%, T0.2
return

; ============================================
; PANIC EXIT KEY
; ============================================

OnPanicKeyChange:
Gui, Main:Submit, NoHide
PanicKeyName := PanicKeyInput
GoSub, BindPanicKey
return

BindPanicKey:
GoSub, UnbindPanicKey
if (!MacrosEnabled)
    return
if (PanicKeyName != "")
{
    try
    {
        bound := "*" PanicKeyName
        Hotkey, %bound%, PanicExitAction, On
        BoundPanicKey := bound
    }
    catch e
    {
        GuiControl, Main:, StatusText, Bad panic key: %PanicKeyName%
    }
}
return

UnbindPanicKey:
if (BoundPanicKey != "")
{
    try
        Hotkey, %BoundPanicKey%, PanicExitAction, Off
    BoundPanicKey := ""
}
return

PanicExitAction:
BlockInput, MouseMoveOff
ExitApp
return

; ============================================
; RECORDING
; ============================================

StartRecording:
if (IsRecording)
    return
Gui, Main:Submit, NoHide
if (MacroDD = "")
    return
    
if (!MacrosEnabled)
{
    GuiControl, Main:, StatusText, Enable macros first before recording.
    return
}

; Find SelectedIdx
SelectedIdx := 0
Loop, % Macros.Length()
{
    if (Macros[A_Index].Name = MacroDD)
    {
        SelectedIdx := A_Index
        break
    }
}
if (!SelectedIdx)
    return

IsRecording := 1
RecordingBuffer := ""
LastEventTime := 0

GoSub, UnbindAllHotkeys

; Hook all keys for down and up natively
Loop, 255
{
    key := GetKeyName(Format("vk{:x}", A_Index))
    if (key != "" && key != StopKeyName && key != PanicKeyName)
    {
        try
        {
            Hotkey, ~*%key%, RecordKeyDown, On
            Hotkey, ~*%key% Up, RecordKeyUp, On
        }
    }
}

; We bind mouse explicitly
Hotkey, ~*LButton, RecordMouseLDown, On
Hotkey, ~*LButton Up, RecordMouseLUp, On
Hotkey, ~*RButton, RecordMouseRDown, On
Hotkey, ~*RButton Up, RecordMouseRUp, On

tempName := Macros[SelectedIdx].Name
GuiControl, Main:, StatusText, Recording "%tempName%"... Press "%StopKeyName%" to stop.

Gui, Overlay:Show, NoActivate
return

StopRecording:
IsRecording := 0
Gui, Overlay:Hide

; unhook all
Loop, 255
{
    key := GetKeyName(Format("vk{:x}", A_Index))
    if (key != "" && key != StopKeyName && key != PanicKeyName)
    {
        try
        {
            Hotkey, ~*%key%, RecordKeyDown, Off
            Hotkey, ~*%key% Up, RecordKeyUp, Off
        }
    }
}
try Hotkey, ~*LButton, RecordMouseLDown, Off
try Hotkey, ~*LButton Up, RecordMouseLUp, Off
try Hotkey, ~*RButton, RecordMouseRDown, Off
try Hotkey, ~*RButton Up, RecordMouseRUp, Off

if (SelectedIdx >= 1 && SelectedIdx <= Macros.Length())
{
    Macros[SelectedIdx].Actions := RecordingBuffer
    readable := FormatForDisplay(RecordingBuffer)
    GuiControl, Main:, EditBox, %readable%
}

; Restore play keys and stop key standard behavior
GoSub, BindAllHotkeys
GoSub, BindStopKey
GoSub, BindPanicKey

GuiControl, Main:, StatusText, Recording saved. Play keys active.
return

; ============================================
; DISPLAY FORMAT FUNCTIONS
; ============================================

FormatForDisplay(raw) {
    if (raw = "")
        return ""
    result := ""
    Loop, Parse, raw, |
    {
        parts := StrSplit(A_LoopField, ":")
        type := parts[1]

        if (InStr(type, "KEY"))
        {
            key := parts[2]
            delay := parts[3]
            if InStr(delay, "-")
            {
                dp := StrSplit(delay, "-")
                dispDelay := "random(" dp[1] ", " dp[2] ")"
            }
            else
                dispDelay := delay
                
            line := type " " key
            pad := 30 - StrLen(line)
            if (pad < 1)
                pad := 1
            Loop, %pad%
                line .= " "
            line .= "delay: " dispDelay "ms"
        }
        else if (InStr(type, "MOUSE"))
        {
            btn := parts[2]
            x := parts[3]
            y := parts[4]
            delay := parts[5]
            if InStr(delay, "-")
            {
                dp := StrSplit(delay, "-")
                dispDelay := "random(" dp[1] ", " dp[2] ")"
            }
            else
                dispDelay := delay
            
            btnName := (btn = "L") ? "L click" : "R click"
            line := type " " btnName " @ (" x ", " y ")"
            pad := 30 - StrLen(line)
            if (pad < 1)
                pad := 1
            Loop, %pad%
                line .= " "
            line .= "delay: " dispDelay "ms"
        }
        else
            line := A_LoopField

        result .= (result ? "`n" : "") . line
    }
    return result
}

ParseFromDisplay(text) {
    if (text = "")
        return ""
    result := ""
    Loop, Parse, text, `n, `r
    {
        line := Trim(A_LoopField)
        if (line = "")
            continue
        if (InStr(line, "KEY"))
        {
            RegExMatch(line, "(KEY|KEY_DOWN|KEY_UP)\s+(\S+)\s+delay:\s*(.+?)ms", m)
            if (m1 != "" && m2 != "" && m3 != "")
            {
                if InStr(m3, "random(")
                {
                    RegExMatch(m3, "random\((\d+),\s*(\d+)\)", rm)
                    val := rm1 "-" rm2
                }
                else
                    val := m3
                result .= (result ? "|" : "") m1 ":" m2 ":" val
            }
        }
        else if (InStr(line, "MOUSE"))
        {
            RegExMatch(line, "(MOUSE|MOUSE_DOWN|MOUSE_UP)\s+([LR])\s+click\s+@\s+\((\d+),\s*(\d+)\)\s+delay:\s*(.+?)ms", m)
            if (m1 != "" && m2 != "" && m3 != "" && m4 != "" && m5 != "")
            {
                if InStr(m5, "random(")
                {
                    RegExMatch(m5, "random\((\d+),\s*(\d+)\)", rm)
                    val := rm1 "-" rm2
                }
                else
                    val := m5
                result .= (result ? "|" : "") m1 ":" m2 ":" m3 ":" m4 ":" val
            }
        }
    }
    return result
}

; ============================================
; APPLY EDITOR CHANGES
; ============================================

ApplyEditorBtn:
GuiControlGet, editContent, Main:, EditBox
raw := ParseFromDisplay(editContent)
if (SelectedIdx >= 1 && SelectedIdx <= Macros.Length())
{
    Macros[SelectedIdx].Actions := raw
}
return

; ============================================
; RECORD INPUT (context-sensitive)
; ============================================

#If IsRecording

CalcDelay:
if (LastEventTime = 0)
{
    r_delay := 0
    LastEventTime := A_TickCount
}
else
{
    r_delay := A_TickCount - LastEventTime
    LastEventTime := A_TickCount
}
return

RecordKeyDown:
key := RegExReplace(A_ThisHotkey, "^[~*]+")
GoSub, CalcDelay
RecordingBuffer .= (RecordingBuffer ? "|" : "") "KEY_DOWN:" key ":" r_delay
return

RecordKeyUp:
key := RegExReplace(A_ThisHotkey, "^[~*]+")
key := RegExReplace(key, "i)\s+Up$", "")
GoSub, CalcDelay
RecordingBuffer .= (RecordingBuffer ? "|" : "") "KEY_UP:" key ":" r_delay
return

RecordMouseLDown:
MouseGetPos, mx, my
GoSub, CalcDelay
RecordingBuffer .= (RecordingBuffer ? "|" : "") "MOUSE_DOWN:L:" mx ":" my ":" r_delay
return

RecordMouseLUp:
MouseGetPos, mx, my
GoSub, CalcDelay
RecordingBuffer .= (RecordingBuffer ? "|" : "") "MOUSE_UP:L:" mx ":" my ":" r_delay
return

RecordMouseRDown:
MouseGetPos, mx, my
GoSub, CalcDelay
RecordingBuffer .= (RecordingBuffer ? "|" : "") "MOUSE_DOWN:R:" mx ":" my ":" r_delay
return

RecordMouseRUp:
MouseGetPos, mx, my
GoSub, CalcDelay
RecordingBuffer .= (RecordingBuffer ? "|" : "") "MOUSE_UP:R:" mx ":" my ":" r_delay
return

#If

; ============================================
; GLOBAL HOTKEYS
; ============================================

#Esc::ExitApp
#r::Reload

; ============================================
; IPC (INTER-PROCESS COMMUNICATION)
; ============================================

Receive_WM_COPYDATA(wParam, lParam)
{
    global CurrentProfile
    StringAddress := NumGet(lParam + 2*A_PtrSize)
    CopyOfData := StrGet(StringAddress, "UTF-16")
    
    parts := StrSplit(CopyOfData, ":")
    if (parts[1] = "LOAD_PROFILE")
    {
        targetProf := Trim(parts[2])
        if (targetProf != "")
        {
            CurrentProfile := targetProf
            GoSub, RefreshProfileList
            GoSub, LoadProfileBtn
        }
    }
    return 1
}

; ============================================
; RELEASE LOGICAL STUCK INPUTS
; ============================================

ReleaseStuckInputs() {
    if GetKeyState("LButton")
        Send, {Blind}{LButton Up}
    if GetKeyState("RButton")
        Send, {Blind}{RButton Up}
    if GetKeyState("MButton")
        Send, {Blind}{MButton Up}
    if GetKeyState("Shift")
        Send, {Blind}{Shift Up}
    if GetKeyState("Ctrl")
        Send, {Blind}{Ctrl Up}
    if GetKeyState("Alt")
        Send, {Blind}{Alt Up}
    if GetKeyState("LWin")
        Send, {Blind}{LWin Up}
    if GetKeyState("RWin")
        Send, {Blind}{RWin Up}
    if GetKeyState("Space")
        Send, {Blind}{Space Up}
    if GetKeyState("Enter")
        Send, {Blind}{Enter Up}
}

; ============================================
; MOUSE BLOCKING HOTKEYS
; ============================================

#If (BlockMouseActive)
*LButton::return
*RButton::return
*MButton::return
*WheelUp::return
*WheelDown::return
#If