; kanakanji.ahk - IME自動変換スクリプト（確実性重視版）
; 速度より確実性を優先し、IME状態を厳密に管理
#NoEnv
#SingleInstance Force
SetBatchLines, -1
SetWorkingDir %A_ScriptDir%

; IME.ahkをインクルード（同じディレクトリに配置必須）
#Include IME.ahk

; ============================================================
; ログ設定
; ============================================================
global LOG_FILE := "kanakanji_debug.log"

if (A_Args.Length() >= 2 && A_Args[2] != "")
{
    LOG_FILE := A_Args[2]
}

global LOG_ENABLED := true

if (LOG_ENABLED)
{
    FileDelete, %LOG_FILE%
    LogWrite("=== IME Conversion Debug Log Started (RELIABILITY MODE) ===")
    LogWrite("Timestamp: " . A_Now)
    LogWrite("Log file: " . LOG_FILE)
}

; ============================================================
; グローバル設定値（確実性重視：すべての待機時間を延長）
; ============================================================
global SLEEP_IME_ACTIVATE := 500          ; 300 → 500ms
global SLEEP_MODE_CHANGE := 500           ; 入力モード変更後の待機
global SLEEP_BASE_INPUT := 500            ; 300 → 500ms
global SLEEP_PER_CHAR := 80               ; 50 → 80ms
global SLEEP_AFTER_CONVERT := 800         ; 500 → 800ms
global SLEEP_AFTER_CONFIRM := 400         ; 200 → 400ms
global SLEEP_CLIPBOARD := 300             ; 150 → 300ms
global CLIPBOARD_TIMEOUT := 50            ; 30 → 50

; キーストローク間の遅延（確実性重視で増加）
global KEY_DELAY := 20                    ; 10 → 20ms
global KEY_PRESS_DURATION := 20           ; 10 → 20ms

global notepadHwnd := 0
global conversionCount := 0

; IME変換モード定数
global IME_CMODE_NATIVE := 1              ; ひらがな
global IME_CMODE_ALPHANUMERIC := 9        ; 全角英数
global IME_CMODE_KATAKANA := 25           ; 全角カタカナ

; コマンドライン引数から待ち時間を設定
if (A_Args.Length() >= 1 && A_Args[1] != "" && A_Args[1] != "default")
{
    customSleep := A_Args[1]
    if customSleep is integer
    {
        SLEEP_AFTER_CONVERT := customSleep
        LogWrite("Custom SLEEP_AFTER_CONVERT = " . customSleep . "ms")
    }
}

; ログに設定値を記録
LogWrite("--- Configuration (RELIABILITY MODE) ---")
LogWrite("SLEEP_IME_ACTIVATE: " . SLEEP_IME_ACTIVATE . "ms")
LogWrite("SLEEP_MODE_CHANGE: " . SLEEP_MODE_CHANGE . "ms")
LogWrite("SLEEP_BASE_INPUT: " . SLEEP_BASE_INPUT . "ms")
LogWrite("SLEEP_PER_CHAR: " . SLEEP_PER_CHAR . "ms")
LogWrite("SLEEP_AFTER_CONVERT: " . SLEEP_AFTER_CONVERT . "ms")
LogWrite("SLEEP_AFTER_CONFIRM: " . SLEEP_AFTER_CONFIRM . "ms")
LogWrite("KEY_DELAY: " . KEY_DELAY . "ms")
LogWrite("KEY_PRESS_DURATION: " . KEY_PRESS_DURATION . "ms")
LogWrite("----------------------------------------")

; Notepadを準備
PrepareNotepad()

; メイン処理ループ
Loop
{
    input := ReadStdIn()
    if (input = "")
        break
    
    conversionCount++
    LogWrite("")
    LogWrite("### Conversion #" . conversionCount . " START ###")
    LogWrite("Input line: [" . input . "]")
    LogWrite("Input length: " . StrLen(input) . " characters")
    
    result := ConvertWithIME(input)
    
    LogWrite("Output result: [" . result . "]")
    LogWrite("### Conversion #" . conversionCount . " END ###")
    
    ; 標準出力に結果を出力
    FileAppend, %result%`n, *, UTF-8
}

LogWrite("")
LogWrite("=== IME Conversion Debug Log Completed ===")
LogWrite("Total conversions: " . conversionCount)

ExitApp

; ============================================================
; ログ出力関数
; ============================================================
LogWrite(message)
{
    global LOG_FILE, LOG_ENABLED
    if (!LOG_ENABLED)
        return
    
    timestamp := A_Hour . ":" . A_Min . ":" . A_Sec . "." . A_MSec
    logLine := "[" . timestamp . "] " . message . "`n"
    FileAppend, %logLine%, %LOG_FILE%, UTF-8
}

; ============================================================
; IME入力モードを取得する関数
; ============================================================
GetIMEMode(winTitle)
{
    mode := IME_GetConvMode(winTitle)
    
    ; モードを文字列に変換（デバッグ用）
    if (mode = 0)
        return "OFF/Direct"
    else if (mode = 1)
        return "Hiragana"
    else if (mode = 9)
        return "FullAlpha"
    else if (mode = 25)
        return "FullKatakana"
    else
        return "Unknown(" . mode . ")"
}

; ============================================================
; IME入力モードを「ひらがな」に強制設定する関数
; ============================================================
ForceHiraganaMode(winTitle)
{
    global IME_CMODE_NATIVE, SLEEP_MODE_CHANGE
    
    LogWrite(">>> ForceHiraganaMode: Starting")
    
    ; 現在のモードを確認
    currentMode := IME_GetConvMode(winTitle)
    LogWrite(">>> ForceHiraganaMode: Current mode = " . GetIMEMode(winTitle) . " (" . currentMode . ")")
    
    ; ひらがなモードでない場合は設定
    if (currentMode != IME_CMODE_NATIVE)
    {
        LogWrite(">>> ForceHiraganaMode: Mode is NOT Hiragana, forcing change")
        IME_SetConvMode(IME_CMODE_NATIVE, winTitle)
        Sleep, %SLEEP_MODE_CHANGE%
        
        ; 設定後のモードを確認
        newMode := IME_GetConvMode(winTitle)
        LogWrite(">>> ForceHiraganaMode: New mode = " . GetIMEMode(winTitle) . " (" . newMode . ")")
        
        if (newMode != IME_CMODE_NATIVE)
        {
            LogWrite(">>> ForceHiraganaMode: WARNING - Failed to set Hiragana mode!")
            return false
        }
    }
    else
    {
        LogWrite(">>> ForceHiraganaMode: Already in Hiragana mode")
    }
    
    LogWrite(">>> ForceHiraganaMode: Completed successfully")
    return true
}

; ============================================================
; Notepadを準備する関数
; ============================================================
PrepareNotepad()
{
    global notepadHwnd
    
    LogWrite(">>> PrepareNotepad: Starting")
    
    ; 既存のNotepadを検索
    WinGet, existingWindows, List, ahk_exe notepad.exe
    
    if (existingWindows > 0)
    {
        notepadHwnd := existingWindows1
        LogWrite(">>> PrepareNotepad: Using existing Notepad (HWND: " . notepadHwnd . ")")
    }
    Else
    {
        LogWrite(">>> PrepareNotepad: Launching new Notepad")
        Run, notepad.exe
        WinWait, ahk_exe notepad.exe, , 5
        if ErrorLevel
        {
            LogWrite(">>> PrepareNotepad: ERROR - Failed to launch Notepad")
            FileAppend, ERROR: Failed to launch Notepad`n, *, UTF-8
            ExitApp, 1
        }
        WinGet, notepadHwnd, ID, ahk_exe notepad.exe
        LogWrite(">>> PrepareNotepad: Notepad launched (HWND: " . notepadHwnd . ")")
    }
    
    ; ウィンドウをアクティブ化
    WinActivate, ahk_id %notepadHwnd%
    WinWaitActive, ahk_id %notepadHwnd%, , 5
    if ErrorLevel
    {
        LogWrite(">>> PrepareNotepad: ERROR - Failed to activate Notepad")
        FileAppend, ERROR: Failed to activate Notepad`n, *, UTF-8
        ExitApp, 1
    }
    
    LogWrite(">>> PrepareNotepad: Notepad activated")
    
    ; 初期状態：内容をクリア
    Send, ^a
    Sleep, 100
    Send, {Delete}
    Sleep, 100
    
    LogWrite(">>> PrepareNotepad: Completed")
}

; ============================================================
; IMEで変換する関数（確実性重視版）
; ============================================================
ConvertWithIME(text)
{
    global notepadHwnd
    global SLEEP_IME_ACTIVATE, SLEEP_MODE_CHANGE, SLEEP_BASE_INPUT, SLEEP_PER_CHAR
    global SLEEP_AFTER_CONVERT, SLEEP_AFTER_CONFIRM
    global SLEEP_CLIPBOARD, CLIPBOARD_TIMEOUT
    global KEY_DELAY, KEY_PRESS_DURATION
    
    startTime := A_TickCount
    winTitle := "ahk_id " . notepadHwnd
    
    ; ウィンドウの存在確認
    if !WinExist(winTitle)
    {
        LogWrite("STEP 0: Notepad window not found, re-preparing")
        PrepareNotepad()
        winTitle := "ahk_id " . notepadHwnd
    }
    
    ; アクティブ化
    LogWrite("STEP 1: Activating Notepad window")
    WinActivate, %winTitle%
    WinWaitActive, %winTitle%, , 5
    LogWrite("STEP 1: Window activated")
    Sleep, 200  ; 追加の安定化待機
    
    ; 内容をクリア
    LogWrite("STEP 2: Clearing content")
    Send, ^a
    Sleep, 100
    Send, {Delete}
    Sleep, 100
    LogWrite("STEP 2: Content cleared")
    
    ; IMEをオンにする
    LogWrite("STEP 3: Turning IME ON")
    imeStateBefore := IME_GET(winTitle)
    LogWrite("STEP 3: IME state before: " . imeStateBefore)
    
    if (imeStateBefore = 0)
    {
        IME_SET(1, winTitle)
        Sleep, %SLEEP_IME_ACTIVATE%
    }
    
    imeStateAfter := IME_GET(winTitle)
    LogWrite("STEP 3: IME state after: " . imeStateAfter . " (waited " . SLEEP_IME_ACTIVATE . "ms)")
    
    if (imeStateAfter = 0)
    {
        LogWrite("STEP 3: ERROR - Failed to turn IME ON")
        FileAppend, ERROR: Failed to turn IME ON`n, *, UTF-8
        return text  ; エラー時は入力をそのまま返す
    }
    
    ; ★★★ 重要：入力モードを「ひらがな」に強制設定 ★★★
    LogWrite("STEP 3.5: Forcing Hiragana input mode")
    modeSetSuccess := ForceHiraganaMode(winTitle)
    if (!modeSetSuccess)
    {
        LogWrite("STEP 3.5: ERROR - Failed to set Hiragana mode")
        FileAppend, ERROR: Failed to set Hiragana mode`n, *, UTF-8
        return text
    }
    
    ; さらに安定化のための追加待機
    Sleep, 300
    LogWrite("STEP 3.5: Additional stabilization wait completed")
    
    ; ひらがなを入力
    LogWrite("STEP 4: Sending input text")
    LogWrite("STEP 4: Input text: [" . text . "]")
    LogWrite("STEP 4: Text length: " . StrLen(text) . " characters")
    LogWrite("STEP 4: KEY_DELAY=" . KEY_DELAY . "ms, KEY_PRESS_DURATION=" . KEY_PRESS_DURATION . "ms")
    
    SetKeyDelay, %KEY_DELAY%, %KEY_PRESS_DURATION%
    SendMode Event
    
    inputStartTime := A_TickCount
    Send, %text%
    inputEndTime := A_TickCount
    inputDuration := inputEndTime - inputStartTime
    
    SendMode Input
    
    LogWrite("STEP 4: Input sent (took " . inputDuration . "ms)")
    
    ; 入力完了を待つ（動的待ち時間 + 大幅な余裕）
    textLength := StrLen(text)
    dynamicSleep := SLEEP_BASE_INPUT + (textLength * SLEEP_PER_CHAR)
    
    ; 最低1200ms、最大5000msに設定（確実性重視）
    if (dynamicSleep < 1200)
        dynamicSleep := 1200
    if (dynamicSleep > 5000)
        dynamicSleep := 5000
    
    LogWrite("STEP 5: Waiting for input completion")
    LogWrite("STEP 5: Calculated wait time: " . dynamicSleep . "ms (base=" . SLEEP_BASE_INPUT . "ms + " . textLength . " chars * " . SLEEP_PER_CHAR . "ms)")
    Sleep, %dynamicSleep%
    LogWrite("STEP 5: Wait completed")
    
    ; IMEの状態を再確認（ON/OFF）
    LogWrite("STEP 6: Verifying IME ON/OFF state")
    imeState := IME_GET(winTitle)
    LogWrite("STEP 6: Current IME state: " . imeState)
    
    if (imeState = 0)
    {
        LogWrite("STEP 6: WARNING - IME is OFF, turning it back ON")
        IME_SET(1, winTitle)
        Sleep, %SLEEP_IME_ACTIVATE%
        imeState := IME_GET(winTitle)
        LogWrite("STEP 6: IME state after re-activation: " . imeState)
    }
    
    ; ★★★ 重要：入力モードを再確認して「ひらがな」に戻す ★★★
    LogWrite("STEP 6.5: Re-verifying and forcing Hiragana mode")
    currentMode := GetIMEMode(winTitle)
    LogWrite("STEP 6.5: Current input mode: " . currentMode)
    
    modeSetSuccess := ForceHiraganaMode(winTitle)
    if (!modeSetSuccess)
    {
        LogWrite("STEP 6.5: WARNING - Failed to re-set Hiragana mode")
    }
    
    ; 追加の安全待機（確実性重視）
    LogWrite("STEP 7: Additional safety wait (500ms)")
    Sleep, 500
    LogWrite("STEP 7: Safety wait completed")
    
    ; スペースキーで変換
    LogWrite("STEP 8: Sending SPACE key for conversion")
    Send, {Space}
    Sleep, %SLEEP_AFTER_CONVERT%
    LogWrite("STEP 8: SPACE sent, waited " . SLEEP_AFTER_CONVERT . "ms")
    
    ; Enterで確定
    LogWrite("STEP 9: Sending ENTER key for confirmation")
    Send, {Enter}
    Sleep, %SLEEP_AFTER_CONFIRM%
    LogWrite("STEP 9: ENTER sent, waited " . SLEEP_AFTER_CONFIRM . "ms")
    
    ; 結果をクリップボードにコピー
    LogWrite("STEP 10: Copying result to clipboard")
    Clipboard := ""
    Send, ^a
    Sleep, 200
    Send, ^c
    Sleep, %SLEEP_CLIPBOARD%
    
    ; クリップボードの内容を取得（確実性重視でタイムアウト増加）
    clipboardRetrieved := false
    Loop, %CLIPBOARD_TIMEOUT%
    {
        if (Clipboard != "")
        {
            clipboardRetrieved := true
            LogWrite("STEP 10: Clipboard retrieved after " . A_Index . " attempts")
            break
        }
        Sleep, 100  ; 50 → 100ms
    }
    
    if (!clipboardRetrieved)
    {
        LogWrite("STEP 10: WARNING - Failed to retrieve clipboard after " . CLIPBOARD_TIMEOUT . " attempts")
    }
    
    result := Clipboard
    LogWrite("STEP 10: Clipboard content: [" . result . "]")
    LogWrite("STEP 10: Result length: " . StrLen(result) . " characters")
    
    ; 結果の分析
    if (InStr(result, " "))
    {
        LogWrite("STEP 10: WARNING - Result contains SPACE character(s)")
    }
    
    if (result = text)
    {
        LogWrite("STEP 10: WARNING - Result equals input (no conversion occurred)")
    }
    
    ; IMEをオフにする
    LogWrite("STEP 11: Turning IME OFF")
    IME_SET(0, winTitle)
    Sleep, 200
    imeStateFinal := IME_GET(winTitle)
    LogWrite("STEP 11: IME state final: " . imeStateFinal)
    
    endTime := A_TickCount
    totalTime := endTime - startTime
    LogWrite("STEP 12: Conversion completed (total time: " . totalTime . "ms)")
    
    return result
}

; ============================================================
; 標準入力から1行読み取る関数
; ============================================================
ReadStdIn()
{
    static hStdIn := DllCall("GetStdHandle", "int", -10, "ptr")
    VarSetCapacity(buf, 8192)
    
    if !DllCall("ReadFile", "ptr", hStdIn, "ptr", &buf, "uint", 8192, "uint*", bytesRead, "ptr", 0)
        return ""
    
    if (bytesRead = 0)
        return ""
    
    result := StrGet(&buf, bytesRead, "UTF-8")
    result := RegExReplace(result, "\r?\n$", "")
    
    return result
}
