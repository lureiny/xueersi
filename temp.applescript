tell application "WeChat" to activate
tell application "System Events"
    tell process "WeChat"
        set the clipboard to "Zyd18566231115"
        click menu item "查找…" of menu "编辑" of menu bar item "编辑" of menu bar 1
        key code 9 using {command down}
        delay 0.3
        key code 76
        key code 48 using {command down}
        delay 0.1
        key code 48 using {command down}
        delay 0.1
        set the clipboard to "您好，赵燕迪。测试数据：5 
解析测试： 
1:2:解析测试
3:2:解析测试
5:2:解析测试
7:2:解析测试
9:2:解析测试

打扰大家了，抱歉，一会可能还会有几次
"
        key code 9 using {command down}
        key code 76
        
    end tell
end tell