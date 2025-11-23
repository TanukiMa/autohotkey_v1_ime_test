; ime_test_universal.ahk
; Microsoft IME、Mozc、Google日本語入力すべてで動作
#Include %A_ScriptDir%\IME.ahk
; コマンドライン引数から平仮名を取得
hiragana := %1%

if (hiragana = "") {
    FileAppend, ERROR: 引数が必要です, *
    ExitApp
}

; メモ帳を開く
Run, notepad.exe
WinWait, ahk_class Notepad, , 10
if ErrorLevel {
    FileAppend, ERROR: メモ帳起動失敗, *
    ExitApp
}

; ウィンドウをアクティブにして最大化
WinActivate, ahk_class Notepad
WinMaximize, ahk_class Notepad
Sleep, 800

; IMEをIME.ahkで制御してONにする
if (IME_GET("A") != 1) {
    IME_SET(0,"A")
    Sleep, 200
    IME_SET(1,"A")
}
Sleep, 600

; 平仮名をクリップボード経由で入力
Clipboard := %hiragana%
ClipWait, 2
if ErrorLevel {
    FileAppend, ERROR: クリップボード失敗, *
    WinClose, ahk_class Notepad
    Send, n
    ExitApp
}

Send, ^v
Sleep, 1000

; 全選択
Send, ^a
Sleep, 400

; Spaceキーで変換
Send, {Space}
Sleep, 1200

; Enterで確定
Send, {Enter}
Sleep, 600

; 結果を取得
Send, ^a
Sleep, 400
Send, ^c
Sleep, 600

; クリップボードの内容を標準出力に出力
FileAppend, %Clipboard%, *

; メモ帳を閉じる
WinClose, ahk_class Notepad
Sleep, 400
Send, n

ExitApp