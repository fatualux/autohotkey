logFile := "C:\Users\100700092024\OneDrive - CGM\Documents\AutoHotkey\logfile.txt"

; Function to log messages
Log(message) {
    global logFile
    FileDelete, %logFile%
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

; Open Brave
^!u::
Log("Opening Brave")
Run, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Brave.lnk"
return

; Open Windows Terminal
^!t::
Log("Opening Windows Terminal")
Run, wt.exe
return

; Open GlazeWM
^!k::
Log("Opening Glaze Window Manager")

; Open Startup folder
Run, "C:\Users\100700092024\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\autoGlazeWM.bat"

; Open GlazeWM
Run, "C:\Program Files\glzr.io\GlazeWM\glazewm.exe"
return

; Function to check if a process exists by its PID
ProcessExist(pid) {
    Process, Exist, %pid%
    return ErrorLevel
}

; Open current folder with Visual Studio Code or last document if no window is opened
^!x::
Log("Attempting to open current folder with Visual Studio Code")
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
        Run, "code" ; Open Visual Studio Code with the last document
        Log("Opened Visual Studio Code with last document")
        }
    }
    else
    {
        Log("Active window is not an Explorer window")
    Run, "code" ; Open Visual Studio Code with the last document
    Log("Opened Visual Studio Code with last document")
    }
return

; Open Windows Explorer with Ctrl + H
^!h::
Log("Opening Windows Explorer")
Run, explorer.exe "C:\Users\100700092024\Downloads"
return

; Close focused window with Ctrl + Q
^q::
Log("Closing focused window")
WinClose, A
return

; Open Windows Terminal in current directory with Ctrl + Alt + Z
^!z::
Log("Opening Windows Terminal in current directory")
WinGetClass, winClass, A
if (winClass = "CabinetWClass") ; Check if it's an Explorer window
{
    folderPath := GetActiveExplorerPath()
    if (folderPath != "")
    {
        ; Run Windows Terminal with the current folder path
        Run, wt.exe -d "%folderPath%"
        Log("Opened Windows Terminal in: " . folderPath)
    }
    else
    {
        Log("Failed to retrieve folder path for Windows Terminal")
    }
}
else
{
    Log("Active window is not an Explorer window")
}
return

; Open PowerShell with user-selected directory using Ctrl + Alt + S
^!s::
Log("Prompting user to select a directory for PowerShell")
; Open a file dialog to select a directory
FileSelectFolder, selectedFolder, , 3, Select a folder to open in PowerShell:
if (selectedFolder) {
    Log("User selected folder: " . selectedFolder)
    ; Run PowerShell with the selected directory
    Run, powershell.exe -NoExit -Command "Set-Location -Path '%selectedFolder%'"
    Log("Opened PowerShell in: " . selectedFolder)
} else {
    Log("No folder selected")
}
return