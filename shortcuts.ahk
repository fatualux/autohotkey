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

; Open current folder with Visual Studio Code
^!x::
Log("Attempting to open current folder with VS Code")
    ; Get the path of the active Explorer window
    WinGetClass, winClass, A
    Log("Active window class: " . winClass)
    if (winClass = "CabinetWClass") ; If the active window is an Explorer window
    {
        WinGet, activePath, ProcessPath, % "ahk_id" WinActive("A")
        Log("Active window process path: " . activePath)
        if (InStr(activePath, "explorer.exe"))
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
                Log("Opened folder in VS Code: " . folderPath)
            }
            else
            {
                Log("Failed to retrieve folder path")
            }
        }
        else
        {
            Log("Active window process is not explorer.exe")
        }
    }
    else
    {
        Log("Active window is not an Explorer window")
    }
return

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

; Move focused window to a new virtual desktop
^!d::
Log("Creating new virtual desktop")
    ; Create a new virtual desktop
    Send, ^#d ; Create a new virtual desktop
    Sleep, 500 ; Wait for the new virtual desktop to be created
    Log("Switched to new virtual desktop")
    Send, ^#Right ; Switch to the newly created virtual desktop
    Sleep, 500 ; Wait for the switch to complete
    Log("Moved to new virtual desktop")

    ; Move back to the original desktop
    Send, ^#Left
    Sleep, 500 ; Wait for the switch to complete
    Log("Moved back to the original desktop")

    ; Open Task View
    Send, #Tab
    Sleep, 500 ; Wait for Task View to open
    Log("Opened Task View")

    ; Move the focused window to the new virtual desktop
    MouseMove, 100, 100 ; Move the mouse to the center of the screen (adjust as needed)
    Sleep, 100
    Log("Mouse moved to center")
    Click ; Click to focus the window in Task View
    Sleep, 100
    Log("Clicked to focus window in Task View")
    Send, ^#Right ; Move the focused window to the new virtual desktop
    Sleep, 500
    Log("Moved focused window to the new virtual desktop")
    Send, #Tab ; Exit Task View
    Log("Exited Task View")
return

; Test retrieving the process path of the active window
f1::
    WinGet, activePath, ProcessPath, % "ahk_id" WinActive("A")
    msgbox % activePath
return