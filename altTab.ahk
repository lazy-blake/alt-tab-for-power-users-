; Smart Alt+Letter Window Switcher for AutoHotkey v2 
; Press Alt + first letter of program to jump directly to that window
; Press Ctrl+Alt+L to list all windows (for debugging)

#Requires AutoHotkey v2.0
#SingleInstance Force
SendMode "Input"
SetWorkingDir A_ScriptDir

; Function to check if a window should be considered
IsValidWindow(windowID) {
    try {
        ; Check if window exists
        if (!WinExist("ahk_id " windowID))
            return false
            
        ; Check if window has a title
        title := WinGetTitle("ahk_id " windowID)
        if (title == "")
            return false
            
        ; Check window style (exclude some system windows)
        style := WinGetStyle("ahk_id " windowID)
        if (!(style & 0x10000000))  ; WS_VISIBLE
            return false
            
        ; Get window class and exclude some system windows
        class := WinGetClass("ahk_id " windowID)
        excludedClasses := ["Shell_TrayWnd", "DV2ControlHost", "MsgrIMEWindowClass", "SysShadow", "Button"]
        for excludedClass in excludedClasses {
            if (class == excludedClass)
                return false
        }
        
        return true
    } catch {
        return false
    }
}

; Function to get clean program name from title or process name
GetProgramName(windowID, title) {
    ; First check for specific app patterns by title (for UWP/Store apps)
    titlePatterns := Map(
        "WhatsApp", "WhatsApp",
        "DaVinci Resolve", "DaVinci Resolve",
        "Microsoft Store", "Store", 
        "Calculator", "Calculator",
        "Settings", "Settings",
        "Mail", "Mail",
        "Calendar", "Calendar",
        "Photos", "Photos",
        "Movies & TV", "Movies"
    )
    
    ; Check if title matches any known patterns
    for pattern, friendlyName in titlePatterns {
        if (RegExMatch(title, "i)" . pattern)) {
            return friendlyName
        }
    }
    
    ; Try to get the process name
    try {
        processName := WinGetProcessName("ahk_id " windowID)
        ; Remove .exe extension
        processName := RegExReplace(processName, "\.exe$", "", &count)
        
        ; Map common process names to readable names
        processMap := Map(
            "chrome", "Chrome",
            "firefox", "Firefox", 
            "msedge", "Edge",
            "Discord", "Discord",
            "Spotify", "Spotify",
            "notepad", "Notepad",
            "Code", "VS Code",
            "explorer", "Explorer",
            "cmd", "Command Prompt",
            "powershell", "PowerShell",
            "WindowsTerminal", "Terminal",
            "zen", "Zen Browser",
            "Resolve", "DaVinci Resolve"
        )
        
        ; Check if we have a mapped name
        for processKey, mappedName in processMap {
            if (RegExMatch(processName, "i)" . processKey)) {
                return mappedName
            }
        }
        
        ; Skip generic container processes and use title cleaning instead
        genericProcesses := ["ApplicationFrameHost", "dwm", "winlogon", "csrss", "svchost"]
        for genericProcess in genericProcesses {
            if (RegExMatch(processName, "i)^" . genericProcess . "$")) {
                ; Fall through to title cleaning
                break
            }
        }
        
        ; If process name looks good, use it
        if (processName != "" && !RegExMatch(processName, "^(dwm|winlogon|csrss|svchost)")) {
            return processName
        }
    } catch {
        ; Fall back to title cleaning
    }
    
    ; Fall back to cleaning the title
    cleanTitle := title
    
    ; Common patterns to clean up
    patterns := [
        "^(.+) - Google Chrome$",
        "^(.+) - Chrome$", 
        "^(.+) - Mozilla Firefox$",
        "^(.+) - Firefox$",
        "^(.+) - Microsoft Edge$",
        "^(.+) - Discord$",
        "^(.+) - Spotify$",
        "^(.+) - Notepad$",
        "^(.+) - Visual Studio Code$",
        "^(.+) - Zen Browser$",
        "^\[.+\] (.+)$",  ; Remove [brackets] prefix
        "^(.+) \([0-9]+\)$",  ; Remove (1), (2) etc
        "^(.+) \(Normal\)$",  ; Remove (Normal) status
        "^(.+) \(Maximized\)$",  ; Remove (Maximized) status
        "^(.+) \(Minimized\)$"  ; Remove (Minimized) status
    ]
    
    for pattern in patterns {
        if (RegExMatch(cleanTitle, pattern, &match)) {
            cleanTitle := match[1]
            break
        }
    }
    
    return Trim(cleanTitle)
}

; Debug function to list all windows
ListAllWindows() {
    windows := WinGetList()
    output := "All Valid Windows:`n`n"
    count := 0
    
    for windowID in windows {
        if (IsValidWindow(windowID)) {
            try {
                title := WinGetTitle("ahk_id " windowID)
                class := WinGetClass("ahk_id " windowID)
                processName := WinGetProcessName("ahk_id " windowID)
                programName := GetProgramName(windowID, title)
                firstLetter := SubStr(programName, 1, 1)
                minMax := WinGetMinMax("ahk_id " windowID)
                status := minMax == -1 ? " (Minimized)" : minMax == 1 ? " (Maximized)" : " (Normal)"
                
                output .= count+1 . ". " . title . status . "`n"
                output .= "   Process: " . processName . "`n"
                output .= "   Program: " . programName . " | First Letter: " . firstLetter . "`n"
                output .= "   Class: " . class . "`n`n"
                count++
            } catch {
                continue
            }
        }
    }
    
    ; Show in a message box
    MsgBox(output, "Window List (" count " windows)")
}

; function to switch windows
SwitchToWindowByLetter(letter) {
    windows := WinGetList()
    matchingWindows := []
    
    ; Loop through all windows
    for windowID in windows {
        if (IsValidWindow(windowID)) {
            try {
                title := WinGetTitle("ahk_id " windowID)
                programName := GetProgramName(windowID, title)
                firstLetter := SubStr(programName, 1, 1)
                
                ; Check if first letter matches (case insensitive)
                if (firstLetter != "" && RegExMatch(firstLetter, "i)^" . letter . "$")) {
                    matchingWindows.Push({
                        id: windowID, 
                        title: title, 
                        programName: programName
                    })
                }
            } catch {
                continue
            }
        }
    }
    
    ; If we found matching windows
    if (matchingWindows.Length > 0) {
        ; Get currently active window
        try {
            currentWindow := WinGetID("A")
        } catch {
            currentWindow := 0
        }
        
        ; Find current window in our matches
        currentIndex := 0
        for index, window in matchingWindows {
            if (window.id == currentWindow) {
                currentIndex := index
                break
            }
        }
        
        ; Switch to next matching window (or first if current not in matches)
        nextIndex := currentIndex >= matchingWindows.Length ? 1 : currentIndex + 1
        targetWindow := matchingWindows[nextIndex]
        
        ; Activate the target window
        try {
            ; If window is minimized, restore it first
            minMax := WinGetMinMax("ahk_id " targetWindow.id)
            if (minMax == -1) {
                WinRestore("ahk_id " targetWindow.id)
            }
            
            WinActivate("ahk_id " targetWindow.id)
            
            ; Show tooltip with program name and status
            status := minMax == -1 ? " (was minimized)" : ""
            ToolTip("→ " . targetWindow.programName . status . " (" . matchingWindows.Length . " matches)")
            SetTimer(() => ToolTip(), -2000)
        } catch {
            ToolTip("Could not switch to window")
            SetTimer(() => ToolTip(), -1000)
        }
    } else {
        ; Show debug info about what windows we found
        ToolTip('No windows found starting with "' . letter . '". Press Ctrl+Alt+L to see all windows.')
        SetTimer(() => ToolTip(), -3000)
    }
}

; to list all windows
^!l::ListAllWindows()

; Hotkeys for Alt + Letter combinations
!a::SwitchToWindowByLetter("a")
!b::SwitchToWindowByLetter("b")
!c::SwitchToWindowByLetter("c")
!d::SwitchToWindowByLetter("d")
!e::SwitchToWindowByLetter("e")
!f::SwitchToWindowByLetter("f")
!g::SwitchToWindowByLetter("g")
!h::SwitchToWindowByLetter("h")
!i::SwitchToWindowByLetter("i")
!j::SwitchToWindowByLetter("j")
!k::SwitchToWindowByLetter("k")
!l::SwitchToWindowByLetter("l")
!m::SwitchToWindowByLetter("m")
!n::SwitchToWindowByLetter("n")
!o::SwitchToWindowByLetter("o")
!p::SwitchToWindowByLetter("p")
!q::SwitchToWindowByLetter("q")
!r::SwitchToWindowByLetter("r")
!s::SwitchToWindowByLetter("s")
!t::SwitchToWindowByLetter("t")
!u::SwitchToWindowByLetter("u")
!v::SwitchToWindowByLetter("v")
!w::SwitchToWindowByLetter("w")
!x::SwitchToWindowByLetter("x")
!y::SwitchToWindowByLetter("y")
!z::SwitchToWindowByLetter("z")

; Numbers
!1::SwitchToWindowByLetter("1")
!2::SwitchToWindowByLetter("2")
!3::SwitchToWindowByLetter("3")
!4::SwitchToWindowByLetter("4")
!5::SwitchToWindowByLetter("5")
!6::SwitchToWindowByLetter("6")
!7::SwitchToWindowByLetter("7")
!8::SwitchToWindowByLetter("8")
!9::SwitchToWindowByLetter("9")
!0::SwitchToWindowByLetter("0")

; Exit hotkey
^!q::ExitApp()