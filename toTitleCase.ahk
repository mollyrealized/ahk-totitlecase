#InstallKeybdHook
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance force
#Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^+t::  ; Ctrl+Shift+T hotkey
    oldClipboard := ClipboardAll
    Clipboard := ""
    Send, ^c
    ClipWait, 2
    if ErrorLevel
    {
        MsgBox, Failed to get text.
        return
    }
    
    converted := ToTitleCase(Clipboard)
    Clipboard := converted
    Send, ^v
    
    Clipboard := oldClipboard
return

ToTitleCase(text) {
    smallWords := "a|an|and|as|at|but|by|en|for|if|in|nor|of|on|or|per|the|to|v\.?|vs\.?|via"
    
    words := []
    word := ""
    Loop, Parse, text
    {
        if A_LoopField in %A_Space%,:–—-
        {
            if (word != "")
                words.Push(word)
            words.Push(A_LoopField)
            word := ""
        }
        else
            word .= A_LoopField
    }
    if (word != "")
        words.Push(word)

    result := ""
    for index, current in words
    {
        if (index = 1 or index = words.Length() or not RegExMatch(current, "i)^(" . smallWords . ")$")
            or (index > 1 and words[index-1] = ":")
            or (index < words.Length() and words[index+1] = ":")
            or (index < words.Length() and words[index+1] = "-" and (index = 1 or words[index-1] != "-")))
        {
            ; Capitalize
            StringUpper, current, current, T
        }
        else if (RegExMatch(SubStr(current, 2), "[A-Z]|\.\.")
            or (index < words.Length() and words[index+1] = ":" and words[index+2] != ""))
        {
            ; Keep original capitalization
        }
        else
        {
            ; Lowercase
            StringLower, current, current
        }
        result .= current
    }

    return result
}