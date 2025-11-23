; kanakanji.ahk - IME自動変換スクリプト（Microsoft IME最適化版 - 入力速度調整）
; 同一Notepadウィンドウを使い回して効率的に処理
#NoEnv
#SingleInstance Force
SetBatchLines, -1
SetWorkingDir %A_ScriptDir%

; IME.ahkをインクルード（同じディレクトリに配置必須）
#Include IME.ahk

; ============================================================
; グローバル設定値（Microsoft IME用デフォルト値）
; ============================================================
global SLEEP_IME_ACTIVATE := 300      ; IME起動待ち
global SLEEP_AFTER_INPUT := 500       ; 入力後待ち（400→500に増加）
global SLEEP_AFTER_CONVERT := 500     ; 変換後待ち（最重要）
global SLEEP_AFTER_CONFIRM := 200     ; 確定後待ち
global SLEEP_CLIPBOARD := 150         ; クリップボード待ち
global CLIPBOARD_TIMEOUT := 30        ; クリップボード取得タイムアウト（回数）

; キーストローク間の遅延設定（Microsoft IME対策）
global KEY_DELAY := 10                ; 各キー入力間の遅延（ミリ秒）
global KEY_PRESS_DURATION := 10       ; 各キーの押下時間（ミリ秒）

; Notepad関連
global notepadHwnd := 0

; コマンドライン引数から待ち時間を設定（オプション）
; 使用例: AutoHotkey.exe kanakanji.ahk 700
; これで SLEEP_AFTER_CONVERT を 700ms に上書き
if (A_Args.Length() >= 1 && A_Args[1] != "")
{
    customSleep := A_Args[1]
    if customSleep is integer
    {
        SLEEP_AFTER_CONVERT := customSleep
        FileAppend, INFO: Custom SLEEP_AFTER_CONVERT = %customSleep%ms`n, *, UTF-8
    }
}

; Notepadを準備
PrepareNotepad()

; メイン処理ループ：標準入力から1行ずつ読み取って変換
Loop
{
    input := ReadStdIn()
    if (input = "")
        break
    
    ; IMEで変換
    result := ConvertWithIME(input)
    
    ; 標準出力に結果を出力
    FileAppend, %result%`n, *, UTF-8
}

ExitApp

; ============================================================
; Notepadを準備する関数
; ============================================================
PrepareNotepad()
{
    global notepadHwnd
    
    ; 既存のNotepadを検索
    WinGet, existingWindows, List, ahk_exe notepad.exe
    
    if (existingWindows > 0)
    {
        ; 既存のNotepadを使用
        notepadHwnd := existingWindows1
    }
    Else
    {
        ; 新規にNotepadを起動
        Run, notepad.exe
        WinWait, ahk_exe notepad.exe, , 5
        if ErrorLevel
        {
            FileAppend, ERROR: Failed to launch Notepad`n, *, UTF-8
            ExitApp, 1
        }
        WinGet, notepadHwnd, ID, ahk_exe notepad.exe
    }
    
    ; ウィンドウをアクティブ化
    WinActivate, ahk_id %notepadHwnd%
    WinWaitActive, ahk_id %notepadHwnd%, , 3
    if ErrorLevel
    {
        FileAppend, ERROR: Failed to activate Notepad`n, *, UTF-8
        ExitApp, 1
    }
    
    ; 初期状態：内容をクリア
    Send, ^a
    Sleep, 50
    Send, {Delete}
    Sleep, 50
}

; ============================================================
; IMEで変換する関数（入力速度を抑制）
; ============================================================
ConvertWithIME(text)
{
    global notepadHwnd
    global SLEEP_IME_ACTIVATE, SLEEP_AFTER_INPUT, SLEEP_AFTER_CONVERT
    global SLEEP_AFTER_CONFIRM, SLEEP_CLIPBOARD, CLIPBOARD_TIMEOUT
    global KEY_DELAY, KEY_PRESS_DURATION
    
    ; ウィンドウの存在確認
    if !WinExist("ahk_id " . notepadHwnd)
    {
        ; ウィンドウが閉じられた場合は再準備
        PrepareNotepad()
    }
    
    ; アクティブ化
    WinActivate, ahk_id %notepadHwnd%
    WinWaitActive, ahk_id %notepadHwnd%, , 2
    
    ; 内容をクリア
    Send, ^a
    Sleep, 50
    Send, {Delete}
    Sleep, 50
    
    ; ① IMEをオンにする
    IME_SET(1, "ahk_id " . notepadHwnd)
    Sleep, %SLEEP_IME_ACTIVATE%
    
    ; ② ひらがなを入力（速度を抑制）
    ; SendInputではなくSendを使用し、キーストローク間に遅延を設定
    SetKeyDelay, %KEY_DELAY%, %KEY_PRESS_DURATION%
    SendMode Event  ; より確実な送信モード
    
    Send, %text%
    
    ; SendModeを元に戻す
    SendMode Input
    
    ; 入力完了を十分待つ
    Sleep, %SLEEP_AFTER_INPUT%
    
    ; ③ スペースキーで変換（第一候補）
    Send, {Space}
    Sleep, %SLEEP_AFTER_CONVERT%
    
    ; ④ Enterで確定
    Send, {Enter}
    Sleep, %SLEEP_AFTER_CONFIRM%
    
    ; ⑤ 結果をクリップボードにコピー
    Clipboard := ""
    Send, ^a
    Sleep, 100
    Send, ^c
    Sleep, %SLEEP_CLIPBOARD%
    
    ; クリップボードの内容を取得（タイムアウト対策）
    Loop, %CLIPBOARD_TIMEOUT%
    {
        if (Clipboard != "")
            break
        Sleep, 50
    }
    
    result := Clipboard
    
    ; IMEをオフにする（次の入力のため）
    IME_SET(0, "ahk_id " . notepadHwnd)
    Sleep, 100
    
    return result
}

; ============================================================
; 標準入力から1行読み取る関数
; ============================================================
ReadStdIn()
{
    static hStdIn := DllCall("GetStdHandle", "int", -10, "ptr")
    VarSetCapacity(buf, 8192)
    
    ; ファイルから読み取り
    if !DllCall("ReadFile", "ptr", hStdIn, "ptr", &buf, "uint", 8192, "uint*", bytesRead, "ptr", 0)
        return ""
    
    if (bytesRead = 0)
        return ""
    
    ; UTF-8として文字列に変換
    result := StrGet(&buf, bytesRead, "UTF-8")
    
    ; 改行コードを除去
    result := RegExReplace(result, "\r?\n$", "")
    
    return result
}
