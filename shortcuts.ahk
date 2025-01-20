logFile := "C:\Users\100700092024\OneDrive - CGM\Documents\AutoHotkey\logfile.txt"

; Function to log messages
Log(message) {
    global logFile
    FileDelete, %logFile% ;
    FileAppend, %A_Now% - %message%`n, %logFile%
}

; Function to get the path of the active Explorer window
GetActiveExplorerPath() {
    for window in ComObjCreate("Shell.Application").Windows
    {
        Log("Inspecting window: " . window.FullName . " with HWND: " . window.hwnd)
        if (InStr(window.FullName, "explorer.exe") && window.hwnd = WinExist("A"))
        {
            Log("Match found: " . window.Document.Folder.Self.Path)
            return window.Document.Folder.Self.Path
        }
    }
    Log("No matching Explorer window found")
    return ""
}

; Open Microsoft Edge
^!b::
Log("Opening Microsoft Edge")
Run, msedge.exe
return

; Open Windows Terminal
^!t::
Log("Opening Windows Terminal")
Run, wt.exe
return

; Open GlazeWM
^!k::
Log("Opening Glaze Window Manager")
Run, "C:\Program Files\glzr.io\GlazeWM\glazewm.exe"
return

; Open current folder with Windows Terminal
^!z::
Log("Attempting to open current folder with Windows Terminal")
    ; Get the path of the active Explorer window
    WinGetClass, winClass, A
    Log("Active window class: " . winClass)
    if (winClass = "CabinetWClass") ; If the active window is an Explorer window
    {
        folderPath := GetActiveExplorerPath()
        Log("Retrieved folder path: " . folderPath)
        if (folderPath != "")
        {
            ; Check if the folder path ends with a backslash and add one if necessary
            if (SubStr(folderPath, 0) != "\") {
                folderPath .= "\"
            }
            Run, wt.exe --startingDirectory "%folderPath%
            Log("Opened folder in Windows Terminal: " . folderPath)
        }
        else
        {
            Log("Failed to retrieve folder path")
        }
    }
    else
    {
        Log("Active window is not an Explorer window")
    }
return

; Open current folder with Visual Studio Code
^!x::
Log("Attempting to open current folder with Visual Studio Code")
    ; Get the path of the active Explorer window
    WinGetClass, winClass, A
    Log("Active window class: " . winClass)
    if (winClass = "CabinetWClass") ; If the active window is an Explorer window
    {
        folderPath := GetActiveExplorerPath()
        Log("Retrieved folder path: " . folderPath)
        if (folderPath != "")
        {
            ; Check if the folder path ends with a backslash and add one if necessary
            if (SubStr(folderPath, 0) != "\") {
                folderPath .= "\"
            }
            Run, "code" "%folderPath%"
            Log("Opened folder in Visual Studio Code: " . folderPath)
        }
        else
        {
            Log("Failed to retrieve folder path")
        }
    }
    else
    {
        Log("Active window is not an Explorer window")
    }
return

; Create a new desktop and go back to the previous virtual desktop
; COMPLETELY USELESS
; I wanted to move the currently focused window on a new desktop but I did not manage to do it
^!d::
Log("Creating new virtual desktop")
    Send, ^#d  ; Create a new virtual desktop (Win+Ctrl+D)
    Sleep, 300
    Send, ^#{Left}  ; Win+Ctrl+Left Arrow to return to the previous desktop
    Log("Returned to previous virtual desktop")
