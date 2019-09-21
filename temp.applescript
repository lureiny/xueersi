tell application "WeChat" to activate
tell application "System Events"
    tell process "WeChat"
        click menu item "查找…" of menu "编辑" of menu bar item "编辑" of menu bar 1
        keystroke "Zyd18566231115"
        key code 76
        key code 48 using {command down}
        delay 0.1
        key code 48 using {command down}
        delay 0.1
        set the clipboard to "您好，李四本次小测的成绩为7 
错题解析： 
1:第一题解析
6:第六题解析
7:第七题解析
"
        key code 9 using {command down}
        key code 76
        
    end tell
end tell