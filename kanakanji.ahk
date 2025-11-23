; kanakanji.ahk - IME自動変換スクリプト（デバッグ版 - ログファイル指定対応）
; 詳細なログ出力でMicrosoft IMEの動作を追跡
#NoEnv
#SingleInstance Force
SetBatchLines, -1
SetWorkingDir %A_ScriptDir%

; IME.ahkをインクルード（同じディレクトリに配置必須）
#Include IME.ahk

; ============================================================
; ログ設定
; ============================================================
; コマンドライン引数からログファイル名を取得
; A_Args[1]: sleep_convert または "default"
; A_Args[2]: log_file名
global LOG_FILE := "kanakanji_debug.log"  ; デフォルト値

if (A_Args.Length() >= 2 && A_Args[2] != "")
{
    LOG_FILE := A_Args[2]
}

global LOG_ENABLED := true

; ログファイルを初期化
if (LOG_ENABLED)
{
    FileDelete, %LOG_FILE%
    LogWrite("=== IME Conversion Debug Log Started ===")
    LogWrite("Timestamp: " . A_Now)
    LogWrite("Log file: " . LOG_FILE)
}

; ============================================================
; グローバル設定値（Microsoft IME用デフォルト値）
; ============================================================
global SLEEP_IME_ACTIVATE := 300
global SLEEP_BASE_INPUT := 300
global SLEEP_PER_CHAR := 50
global SLEEP_AFTER_CONVERT := 500
global SLEEP_AFTER_CONFIRM := 200
global SLEEP_CLIPBOARD := 150
global CLIPBOARD_TIMEOUT := 30

global KEY_DELAY := 10
global KEY_PRESS_DURATION := 10

global notepadHwnd := 0
global conversionCount := 0

; コマンドライン引数から待ち時間を設定
; A_Args[1]: sleep_convert または "default"
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
LogWrite("--- Configuration ---")
LogWrite("SLEEP_IME_ACTIVATE: " . SLEEP_IME_ACTIVATE . "ms")
LogWrite("SLEEP_BASE_INPUT: " . SLEEP_BASE_INPUT . "ms")
LogWrite("SLEEP_PER_CHAR: " . SLEEP_PER_CHAR . "ms")
LogWrite("SLEEP_AFTER_CONVERT: " . SLEEP_AFTER_CONVERT . "ms")
LogWrite("SLEEP_AFTER_CONFIRM: " . SLEEP_AFTER_CONFIRM . "ms")
LogWrite("KEY_DELAY: " . KEY_DELAY . "ms")
LogWrite("KEY_PRESS_DURATION: " . KEY_PRESS_DURATION . "ms")
LogWrite("---------------------")

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
    WinWaitActive, ahk_id %notepadHwnd%, , 3
    if ErrorLevel
    {
        LogWrite(">>> PrepareNotepad: ERROR - Failed to activate Notepad")
        FileAppend, ERROR: Failed to activate Notepad`n, *, UTF-8
        ExitApp, 1
    }
    
    LogWrite(">>> PrepareNotepad: Notepad activated")
    
    ; 初期状態：内容をクリア
    Send, ^a
    Sleep, 50
    Send, {Delete}
    Sleep, 50
    
    LogWrite(">>> PrepareNotepad: Completed")
}

; ============================================================
; IMEで変換する関数（詳細ログ付き）
; ============================================================
ConvertWithIME(text)
{
    global notepadHwnd
    global SLEEP_IME_ACTIVATE, SLEEP_BASE_INPUT, SLEEP_PER_CHAR
    global SLEEP_AFTER_CONVERT, SLEEP_AFTER_CONFIRM
    global SLEEP_CLIPBOARD, CLIPBOARD_TIMEOUT
    global KEY_DELAY, KEY_PRESS_DURATION
    
    startTime := A_TickCount
    
    ; ウィンドウの存在確認
    if !WinExist("ahk_id " . notepadHwnd)
    {
        LogWrite("STEP 0: Notepad window not found, re-preparing")
        PrepareNotepad()
    }
    
    ; アクティブ化
    LogWrite("STEP 1: Activating Notepad window")
    WinActivate, ahk_id %notepadHwnd%
    WinWaitActive, ahk_id %notepadHwnd%, , 2
    LogWrite("STEP 1: Window activated")
    
    ; 内容をクリア
    LogWrite("STEP 2: Clearing content")
    Send, ^a
    Sleep, 50
    Send, {Delete}
    Sleep, 50
    LogWrite("STEP 2: Content cleared")
    
    ; IMEをオンにする
    LogWrite("STEP 3: Turning IME ON")
    imeStateBefore := IME_GET("ahk_id " . notepadHwnd)
    LogWrite("STEP 3: IME state before: " . imeStateBefore)
    
    IME_SET(1, "ahk_id " . notepadHwnd)
    Sleep, %SLEEP_IME_ACTIVATE%
    
    imeStateAfter := IME_GET("ahk_id " . notepadHwnd)
    LogWrite("STEP 3: IME state after: " . imeStateAfter . " (waited " . SLEEP_IME_ACTIVATE . "ms)")
    
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
    
    ; 入力完了を待つ（動的待ち時間）
    textLength := StrLen(text)
    dynamicSleep := SLEEP_BASE_INPUT + (textLength * SLEEP_PER_CHAR)
    
    if (dynamicSleep < 800)
        dynamicSleep := 800
    if (dynamicSleep > 3000)
        dynamicSleep := 3000
    
    LogWrite("STEP 5: Waiting for input completion")
    LogWrite("STEP 5: Calculated wait time: " . dynamicSleep . "ms (base=" . SLEEP_BASE_INPUT . "ms + " . textLength . " chars * " . SLEEP_PER_CHAR . "ms)")
    Sleep, %dynamicSleep%
    LogWrite("STEP 5: Wait completed")
    
    ; IMEの状態を再確認
    LogWrite("STEP 6: Verifying IME state")
    imeState := IME_GET("ahk_id " . notepadHwnd)
    LogWrite("STEP 6: Current IME state: " . imeState)
    
    if (imeState = 0)
    {
        LogWrite("STEP 6: WARNING - IME is OFF, turning it back ON")
        IME_SET(1, "ahk_id " . notepadHwnd)
        Sleep, 200
        imeState := IME_GET("ahk_id " . notepadHwnd)
        LogWrite("STEP 6: IME state after re-activation: " . imeState)
    }
    
    ; 追加の安全待機
    LogWrite("STEP 7: Additional safety wait (200ms)")
    Sleep, 200
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
    Sleep, 100
    Send, ^c
    Sleep, %SLEEP_CLIPBOARD%
    
    ; クリップボードの内容を取得
    clipboardRetrieved := false
    Loop, %CLIPBOARD_TIMEOUT%
    {
        if (Clipboard != "")
        {
            clipboardRetrieved := true
            LogWrite("STEP 10: Clipboard retrieved after " . A_Index . " attempts")
            break
        }
        Sleep, 50
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
    IME_SET(0, "ahk_id " . notepadHwnd)
    Sleep, 100
    imeStateFinal := IME_GET("ahk_id " . notepadHwnd)
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
