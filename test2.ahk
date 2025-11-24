#NoEnv
SendMode Input
SetWorkingDir %A_ScriptDir%

; ===== 設定: 使用しているIME名を指定してください =====
IME_NAME := "Microsoft IME"  ; または "Google日本語入力" / "Mozc" / "Mozc4med"

; ===== IME制御関数 =====
IME_GET(WinTitle:="A") {
    WinGet, hWnd, ID, %WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", "Uint", hWnd, "Uint")
    Return SendMessage(0x283, 0x005, 0, "", "ahk_id " DefaultIMEWnd)
}

IME_SET(SetSts:=0, WinTitle:="A") {
    WinGet, hWnd, ID, %WinTitle%
    DefaultIMEWnd := DllCall("imm32\ImmGetDefaultIMEWnd", "Uint", hWnd, "Uint")
    SendMessage, 0x283, 0x006, SetSts, "", ahk_id %DefaultIMEWnd%
}

; ===== メイン処理 =====
^+t::  ; Ctrl+Shift+T で実行
    ; ファイルパス設定
    inputFile := A_ScriptDir . "\input.txt"
    outputFile := A_ScriptDir . "\output_" . IME_NAME . ".txt"
    
    ; 入力ファイルの存在確認
    IfNotExist, %inputFile%
    {
        MsgBox, エラー: input.txt が見つかりません。`nスクリプトと同じフォルダに input.txt を作成してください。
        Return
    }
    
    ; テスト文を読み込み
    testSentences := []
    Loop, Read, %inputFile%
    {
        line := Trim(A_LoopReadLine)
        if (line != "")
            testSentences.Push(line)
    }
    
    if (testSentences.Length() = 0) {
        MsgBox, エラー: input.txt にテスト文が含まれていません。
        Return
    }
    
    ; 出力ファイルを初期化
    FileDelete, %outputFile%
    timestamp := A_YYYY . "-" . A_MM . "-" . A_DD . " " . A_Hour . ":" . A_Min . ":" . A_Sec
    FileAppend, IME変換テスト結果`n, %outputFile%
    FileAppend, IME: %IME_NAME%`n, %outputFile%
    FileAppend, 実行日時: %timestamp%`n, %outputFile%
    FileAppend, テスト文数: %testSentences.Length()%`n`n, %outputFile%
    FileAppend, ========================================`n`n, %outputFile%
    
    MsgBox, 4, IMEテスト開始, IME: %IME_NAME%`nテスト文数: %testSentences.Length()%`n`nテストを開始しますか？
    IfMsgBox No
        Return
    
    Sleep, 1000
    
    ; メモ帳を起動
    Run, notepad.exe
    WinWait, ahk_class Notepad, , 5
    if ErrorLevel {
        MsgBox, エラー: メモ帳を起動できませんでした。
        Return
    }
    
    WinActivate, ahk_class Notepad
    Sleep, 500
    
    ; 各テスト文を実行
    For index, sentence in testSentences {
        ; IMEをONにする
        IME_SET(1, "ahk_class Notepad")
        Sleep, 300
        
        ; 平仮名を入力（ローマ字入力）
        SendInput, %sentence%
        Sleep, 400
        
        ; スペースキーで変換
        Send, {Space}
        Sleep, 400
        
        ; Enterで確定
        Send, {Enter}
        Sleep, 300
        
        ; 改行を追加
        Send, {Enter}
        Sleep, 100
    }
    
    ; 全選択してクリップボードにコピー
    Send, ^a
    Sleep, 200
    Send, ^c
    Sleep, 300
    
    ; クリップボードの内容を解析して結果ファイルに保存
    resultText := Clipboard
    resultLines := StrSplit(resultText, "`n", "`r")
    
    lineIndex := 1
    For sentenceIndex, sentence in testSentences {
        if (lineIndex <= resultLines.Length()) {
            converted := Trim(resultLines[lineIndex])
            FileAppend, [%sentenceIndex%]`n, %outputFile%
            FileAppend, 入力: %sentence%`n, %outputFile%
            FileAppend, 変換: %converted%`n`n, %outputFile%
            lineIndex += 2  ; 次の行（空行をスキップ）
        }
    }
    
    ; メモ帳を閉じる（保存しない）
    WinClose, ahk_class Notepad
    Sleep, 300
    Send, n  ; 保存しないを選択
    
    MsgBox, テスト完了！`n結果は以下に保存されました：`n%outputFile%
    
    ; 結果ファイルを開く
    Run, %outputFile%
Return

; Escキーで終了
Esc::ExitApp
