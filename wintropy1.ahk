;------------------------------------------------------------------------------;
; Wintropy                                                                     ;
; Save and Restore Windows Position                                            ;
;                                                                              ;
; Original code from: https://github.com/rwese/DockWin                         ;
;------------------------------------------------------------------------------;

#Requires AutoHotkey v2.0+
#SingleInstance force
#Warn

;------------------------------------------------------------------------------;
; Global Variables                                                             ;
;------------------------------------------------------------------------------;

AppName := StrSplit(A_Scriptname, ".")[1]
AppVersion := "2.4.10"
IconName := A_IsCompiled? (AppName . ".exe"): (AppName . ".ico")
IniName := AppName . ".ini"

LogName := AppName . ".log"
LogLevel := 3                                                   ; 0) OFF, 1) ERROR, 2) WARN, 3) INFO, 4) DEBUG, 5) TRACE

;------------------------------------------------------------------------------;
; Main                                                                         ;
;------------------------------------------------------------------------------;

{
  SendMode("Input")
  SetWorkingDir(A_ScriptDir)

  SetTitleMatchMode(2)
  SetTitleMatchMode("fast")
  DetectHiddenWindows(false)

  ; Parse command line
  NeedAdmin := ParseCommandLine(A_Args, &LogLevel)

  ; LogError("LogLevel = " . LogLevel . ", A_IsAdmin = " . A_IsAdmin . ", NeedAdmin = " . NeedAdmin)

  ; If we are required to run elevated and we are not already, we restart
  if (!A_IsAdmin && NeedAdmin) {
    RestartAsAdministrator(A_Args)
  }
  TrayMenu()

  return
}

;------------------------------------------------------------------------------;
; Hotstrings & Hotkeys                                                         ;
;------------------------------------------------------------------------------;

:?*:a'\::à
:?*:a'/::á
:?*:A'\::À
:?*:A'/::Á

:?*:e'\::è
:?*:e'/::é
:?*:E'\::È
:?*:E'/::É

:?*:i'\::ì
:?*:i'/::í
:?*:I'\::Ì
:?*:I'/::Í

:?*:o'\::ò
:?*:o'/::ó
:?*:O'\::Ò
:?*:O'/::Ó

:?*:u'\::ù
:?*:u'/::ú
:?*:U'\::Ù
:?*:U'/::Ú

#0:: WindowsRestore()                                           ; <Win> + 0
#+0:: WindowsCapture()                                          ; <Win> + <Shift> + 0
#F12:: ModernStandby()                                          ; <Win> + <F12>
^+4:: SendInput("€")                                            ; <Ctrl> + <Shift> + 4

#HotIf MouseIsOver("ahk_class Shell_TrayWnd")
WheelUp:: Send("{Volume_Up}")
MButton:: Send("{Volume_Mute}")
WheelDown:: Send("{Volume_Down}")

;------------------------------------------------------------------------------;
; Functions                                                                    ;
;------------------------------------------------------------------------------;

Log(level, msg)
{
  global LogName, LogLevel
  static id := ["ERROR: ", " WARN: ", " INFO: ", "DEBUG: ", "TRACE: "]

  if (level < 1 || level > 5 || level > LogLevel) {
    return
  }

  FileAppend(FormatTime("", "yyyy.MM.dd HH:mm:ss ") . Id[level] . msg . "`r`n", LogName)  
}
  
LogError(msg)
{
  Log(1, msg)
}

LogWarning(msg)
{ 
  Log(2, msg)
}

LogInfo(msg)
{ 
  Log(3, msg)
}

LogDebug(msg)
{
  Log(4, msg)
}

LogTrace(msg)
{
  Log(5, msg)
}

;-------------------------------------------------------------------------------

MouseIsOver(winTitle)
{
  MouseGetPos(,, &winId)
  return (WinExist(winTitle " ahk_id " winId))
}

;-------------------------------------------------------------------------------

TrayMenu()
;
; Setup Windows taskbar tray menu
;
{
  global AppName, AppVersion, IconName, IniName

  tipStr     := AppName . " " . AppVersion . "`nCapture and Restore Windows Position"
  restoreStr := "Restore`t(Win+0)"
  captureStr := "Capture`t(Win+Shift+0)"
  editStr    := "Edit " . IniName
  exitStr    := "Exit"

  TraySetIcon(IconName, 1)
  A_IconTip := tipStr
  tray := A_TrayMenu
  tray.Delete()

  tray.Add(restoreStr, WindowsRestore)
  tray.Add(captureStr, WindowsCapture)
  tray.Add(editStr, IniEdit)
  tray.Add()                                                    ; Separator
  tray.Add(exitStr, ApplicationExit)

  tray.Default := restoreStr
  tray.ClickCount := 1
}

;-------------------------------------------------------------------------------

ModernStandby()
;
; Put the system in Modern Standby (S0)
;
{
  Sleep(500)
  
  ; WM_SYSCOMMAND, SC_MONITORPOWER, -1: On / 1: Low-power / 2: Off
  SendMessage(0x112, 0xF170, 2, , "Program Manager")
}

;-------------------------------------------------------------------------------

ApplicationExit(*)
;
; Exit application
;
{
  ExitApp()
}

;-------------------------------------------------------------------------------

IniEdit(*)
;
; Edit ini file
;
{
  global AppName, IniName
  
  try {
    Run("notepad " . IniName)
  }
  catch as err {
    MsgBox("Exception running `"notepad " . IniName . "`"`nSpecifically: " . err.message, AppName, 16)
  }
}

;-------------------------------------------------------------------------------

WindowsRestore(*)
;
; Restore windows listed in the ini file
;
{
  global AppName, IniName
  static Params := "title x y height width show path args dir verb"

  ; Make sure we have an .ini file
  if (!FileExist(IniName)) {
    MsgBox("`"" . IniName . "`" does not exist", AppName, 16)
    return
  }

  SavedActiveWindow := WinGetTitle("A")
  
  SectionToFind := SectionHeader()
  SectionFound := false
 
  loop read, IniName {
    ; trim spaces
    iniLine := Trim(A_LoopReadLine)

    ; skip comments and empty lines
    if (iniLine = "" || SubStr(iniLine, 1, 1) = ";") {
      continue 
    }

    if (!SectionFound) {
      ; Read through file until correct section found
      if (iniLine = SectionToFind) {
        SectionFound := true
      }
      continue
    }    

    ; Exit when next section is reached
    if (SectionFound && SubStr(iniLine, 1, 5) != "title") {
      break
    }
   
    Win_title := ""
    Win_x := 0
    Win_y := 0
    Win_width := 0
    Win_height := 0
    Win_show := 0                                               ; -1) minimized, 0) normal, 1) maximized
    Win_path := ""
    Win_args := ""
    Win_dir := ""
    Win_verb := ""

    loop parse, iniLine, "CSV" {
      if (EqualPos := InStr(A_LoopField, "=")) {
        Var := Trim(SubStr(A_LoopField, 1, EqualPos - 1))
        Val := Trim(SubStr(A_LoopField, EqualPos + 1))
        if (InStr(Params, Var)) {
          ; Remove any surrounding double quotes
          if (SubStr(Val, 1, 1) = "`"") {
            Val := SubStr(Val, 2, StrLen(Val) - 2)
          }
          Win_%Var% := Val  
        }
      }
    }
    
    ; Check if program is already running, if not, start it
    if (!WinExist(Win_title) && (Win_path != "")) {
      try {
        ; Run Win_path  
        ShellRun(Win_path, Win_args, Win_dir, Win_verb)

        ; Give some time for the program to launch
        Sleep(1000)
      }
      catch as err {
        MsgBox("Exception running " . Win_path . "`nSpecifically: " . err.message, AppName, 16)
      }
    }

    if ((Win_show = 1) && WinExist(Win_title)) {  
      WinRestore()
      WinActivate()
      WinMove(Win_x, Win_y, Win_width, Win_height, "A")
      WinMove(Win_x, Win_y, Win_width, Win_height, "A")
      WinMaximize("A")
    }
    else if ((Win_show = -1) && (StrLen(Win_title) > 0) && WinExist(Win_title)) {  
      WinRestore()
      WinActivate()
      WinMove(Win_x, Win_y, Win_width, Win_height, "A")
      WinMove(Win_x, Win_y, Win_width, Win_height, "A")
      WinMinimize("A")
    }
    else if ((StrLen(Win_title) > 0) && WinExist(Win_title)) {  
      WinRestore()
      WinActivate()
      WinMove(Win_x, Win_y, Win_width, Win_height, "A")
      WinMove(Win_x, Win_y, Win_width, Win_height, "A")
    }
  }

  if (!SectionFound) {
    MsgBox("Section does not exist in " . IniName . "`n`nLooking for: " . SectionToFind, AppName, 48)
  }

  ; Restore window that was active at beginning of script
  WinActivate(SavedActiveWindow)
}

;-------------------------------------------------------------------------------

WindowsCapture(*)
;
; Create an ini file with a list of currently open windows
;
{
  global AppName, IniName
  
  msgResult := MsgBox("Save windows position (it will append to " . IniName . ")?", AppName, 36)
  if (msgResult = "NO") {
    return
  }

  SavedActiveWindow := WinGetTitle("A")

  IniFile := FileOpen(IniName, "a")
  if (!IsObject(IniFile)) {
    MsgBox("Can't open `"" . IniName . "`" for writing.", AppName, 16)
    return
  }

  IniLine := SectionHeader() . "`r`n"
  IniFile.Write(IniLine)

  ; Loop through all windows on the entire system
  ids := WinGetlist(,,"Program Manager",)
  for (this_id in ids) {
    this_class := WinGetClass("ahk_id " . this_id)
    this_title := WinGetTitle("ahk_id " . this_id)
    Win_show := WinGetminmax("ahk_class " . this_class)
    WinActivate("ahk_id " . this_id)
    WinGetPos(&Win_x, &Win_y, &Win_width, &Win_height, "A")

    if ((StrLen(this_title) > 0) && (this_title != "Start")) {
      IniLine := "title=`"" this_title . "`",x=" Win_x . ",y=" Win_y . ",width=" Win_width . ",height=" Win_height . ",show=" Win_show . ",path=`"`",args=`"`",dir=`"`",verb=`"`"`r`n"
      IniFile.Write(IniLine)
    }
  
    ; Re-minimize any windows that were minimised before we started
    if (Win_show = -1) {
      WinMinimize("A")
    }
  }

  ; Add blank line after section
  IniFile.write("`r`n")
  IniFile.Close()

  ; Restore active window
  WinActivate(SavedActiveWindow)
}

;-------------------------------------------------------------------------------

SectionHeader()
;
; Create standardized section header for later retrieval
;
{
  mc := MonitorGetCount()
  mp := MonitorGetPrimary()
  WinGetPos(&dx, &dy, &dw, &dh, "Program Manager")

  DesktopMap(&dc, &da)
 
  return ("desktop" . da . ": monitors=" . mc . ",primary=" . mp . ",position=" . dx . "," . dy . "," . dw . "," . dh) 
}

;-------------------------------------------------------------------------------

DesktopMap(&desktopCount, &desktopCurrent)
;
; This function examines the registry to build a list of the current virtual desktops and which one we're currently on.
; List of desktops appears to be in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops
; On Windows 11 the current desktop UUID appears to be in the same location
; On previous versions in HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\1\VirtualDesktops
;
{
  desktopCount := 1
  desktopCurrent := 1

  ; Get the current desktop UUID. Length should be 32 always, but there's no guarantee this couldn't change in a later Windows release so we check.
  idLength := 32
  sessionId := GetSessionId()
  if (sessionId) {
    desktopId := RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops", "CurrentVirtualDesktop")
    if (A_LastError) {
      desktopId := RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\SessionInfo\" . sessionId . "\VirtualDesktops", "CurrentVirtualDesktop")
    }
    
    if (desktopId) {
      idLength := StrLen(desktopId)
    }
  }

  ; Get a list of the UUIDs for all virtual desktops on the system
  desktopList := RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VirtualDesktops", "VirtualDesktopIDs")
  if (desktopList) {
    desktopListLength := StrLen(desktopList)
    ; Figure out how many virtual desktops there are
    desktopCount := floor(desktopListLength / idLength)
  }

  ; Parse the REG_DATA string that stores the array of UUID's for virtual desktops in the registry.
  i := 0
  while (desktopId && i < desktopCount) {
    startPos := (i * idLength) + 1
    desktopIter := SubStr(desktopList, startPos, idLength)

    ; Break out if we find a match in the list. If we didn't find anything, keep the
    ; old guess and pray we're still correct :-D.
    if (desktopIter = desktopId) {
      desktopCurrent := i + 1
      break
    }
    i++
  }
}

;-------------------------------------------------------------------------------

GetSessionId()
;
; This functions finds out ID of current session.
;
{
  sessionId := 0

  processId := DllCall("GetCurrentProcessId", "UInt")
  DllCall("ProcessIdToSessionId", "UInt", processId, "UInt*", &sessionId)

  return (sessionId)
}

;-------------------------------------------------------------------------------

ShellRun(prms*)
;
; Parameters:
;   1: application to launch
;   2: command line parameters
;   3: working directory for the new process
;   4: verb, for example pass "RunAs" to run as administrator
;   5: suggestion to the application about how to show its window - see the msdn link for possible values
;
;   Shell.ShellExecute(File [, Arguments, Directory, Operation, Show])
;   http://msdn.microsoft.com/en-us/library/windows/desktop/gg537745
;
; Usage Examples:
;   ShellRun("Taskmgr.exe")                     ; Task manager
;   ShellRun("Notepad.exe", A_ScriptFullPath)   ; Open a file with notepad
;   ShellRun("Notepad.exe",,,"RunAs")           ; Open untitled notepad as administrator
;
; Credits:
;   https://github.com/Lexikos/AutoHotkey-Release/blob/master/installer/source/Lib/ShellRun.ahk
;   https://www.autohotkey.com/boards/viewtopic.php?f=82&t=78190
;
{
  shellWindows := ComObject("Shell.Application").Windows

  ; SWC_DESKTOP, SWFO_NEEDDISPATCH
  desktop := shellWindows.FindWindowSW(0, 0, 8, 0, 1)
   
  ; Retrieve top-level browser object: SID_STopLevelBrowser, IID_IShellBrowser 
  tlb := ComObjQuery(desktop, "{4C96BE40-915C-11CF-99D3-00AA004AE837}", "{000214E2-0000-0000-C000-000000000046}")
    
  ; IShellBrowser.QueryActiveShellView -> IShellView VT_UNKNOWN
  ComCall(15, tlb, "ptr*", sv := ComValue(13, 0))
    
  ; Define IID_IDispatch.
  NumPut("int64", 0x20400, "int64", 0x46000000000000C0, IID_IDispatch := Buffer(16))
   
  ; IShellView.GetItemObject -> IDispatch (object which implements IShellFolderViewDual) VT_DISPATCH
  ComCall(15, sv, "uint", 0, "ptr", IID_IDispatch, "ptr*", sfvd := ComValue(9, 0))
   
  ; Get Shell object
  shell := sfvd.Application
   
  ; IShellDispatch2.ShellExecute
  shell.ShellExecute(prms*)
}

;-------------------------------------------------------------------------------

RestartAsAdministrator(args)
;
; We restart with Administrator priviledges
;
{
  ; Re-join arguments
  argsLine := ""
  for (param in args) {
    argsLine := argsLine . " " . param
  }

  fullCommandLine := DllCall("GetCommandLine", "str")

  if !(A_IsAdmin || RegExMatch(fullCommandLine, " /restart(?!\S)")) {
    try {
      if (A_IsCompiled) {
        Run('*RunAs "' . A_ScriptFullPath . '" /restart' . argsLine)
      }
      else {
        Run('*RunAs "' . A_AhkPath . '" /restart "' . A_ScriptFullPath . '"' . argsLine)
      }
    }
    
    ExitApp()
  }
}

;-------------------------------------------------------------------------------

ParseCommandLine(args, &logLev)
;
; Parse command line: wintropy /log=x /admin
;
;   x = 0,1,2,3,4,5
;
; Return true is we need to run as administrator
;
{
  admin := false

  for (param in args) {
    if (InStr(param, "/log=") = 1) {
      val := SubStr(param, 6, 1)
      if (val >= "0" && val <= "5") {
        logLev := (Ord(val) - Ord("0"))
      }
    }
    
    if (param = "/admin") {
      admin := true
    }
  }

  return (admin)
}

;------------------------------------------------------------------------------;
; EOF                                                                          ;
;------------------------------------------------------------------------------;