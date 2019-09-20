tell application "WeChat" to activate
tell application "System Events"
    tell process "WeChat"
        click menu item "查找…" of menu "编辑" of menu bar item "编辑" of menu bar 1
        keystroke "Zyd18566231115"
        delay 0.5
        key code 76
        key code 48 using {command down}
        key code 48 using {command down}
        key code 9 using {command down}
        -- key code 76
    end tell
end tell