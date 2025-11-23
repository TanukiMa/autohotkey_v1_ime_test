#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kanakanji.py - IMEを使用してひらがなを漢字に変換するツール（デバッグ版）

使用方法:
    python kanakanji.py input.txt -o output.txt --debug
    python kanakanji.py input.txt -o output.txt --debug --log msime_debug.log
"""

import argparse
import subprocess
import sys
import os
import shutil
import time
from pathlib import Path
from datetime import datetime


def find_autohotkey():
    """
    AutoHotkey.exeのパスを検索する
    
    Returns:
        str: AutoHotkey.exeのフルパス、見つからない場合はNone
    """
    common_paths = [
        r"C:\Program Files\AutoHotkey\AutoHotkey.exe",
        r"C:\Program Files\AutoHotkey\v1.1.37.02\AutoHotkey.exe",
        r"C:\Program Files\AutoHotkey\v1.1.36.02\AutoHotkey.exe",
        r"C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\AutoHotkey\AutoHotkey.exe"),
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
    
    ahk_path = shutil.which("AutoHotkey.exe")
    if ahk_path:
        return ahk_path
    
    return None


def log_debug(message, debug=False):
    """デバッグメッセージを出力"""
    if debug:
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[{timestamp}] {message}", file=sys.stderr)


def convert_with_ime(input_file, output_file, ahk_script="kanakanji.ahk", 
                     sleep_convert=None, log_file=None, debug=False):
    """
    入力ファイルからひらがなを読み込み、AutoHotkeyスクリプトでIME変換
    
    Args:
        input_file (str): 入力ファイルパス（ひらがな、1行1フレーズ）
        output_file (str): 出力ファイルパス
        ahk_script (str): AutoHotkeyスクリプトのパス
        sleep_convert (int): 変換処理後の待ち時間（ミリ秒）
        log_file (str): ログファイル名
        debug (bool): デバッグモード
    """
    
    log_debug("=== Python Debug Log Started ===", debug)
    
    # デフォルトログファイル名
    if log_file is None:
        log_file = "kanakanji_debug.log"
    
    log_debug(f"Log file: {log_file}", debug)
    
    # AutoHotkeyの実行ファイルパスを検索
    ahk_exe = find_autohotkey()
    if not ahk_exe:
        print("Error: AutoHotkey.exe が見つかりません", file=sys.stderr)
        print("https://www.autohotkey.com/ からダウンロードしてインストールしてください", file=sys.stderr)
        sys.exit(1)
    
    log_debug(f"Using AutoHotkey: {ahk_exe}", debug)
    
    # AHKスクリプトの存在確認
    if not os.path.exists(ahk_script):
        print(f"Error: {ahk_script} が見つかりません", file=sys.stderr)
        sys.exit(1)
    
    log_debug(f"Using AHK script: {os.path.abspath(ahk_script)}", debug)
    
    # IME.ahkの存在確認
    ime_ahk_path = os.path.join(os.path.dirname(os.path.abspath(ahk_script)), "IME.ahk")
    if not os.path.exists(ime_ahk_path):
        print(f"Error: IME.ahk が見つかりません", file=sys.stderr)
        print(f"Expected location: {ime_ahk_path}", file=sys.stderr)
        sys.exit(1)
    
    log_debug(f"Found IME.ahk: {ime_ahk_path}", debug)
    
    # 入力ファイルを読み込み
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            input_lines = [line.strip() for line in f if line.strip()]
    except FileNotFoundError:
        print(f"Error: 入力ファイル '{input_file}' が見つかりません", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading input file: {e}", file=sys.stderr)
        sys.exit(1)
    
    if not input_lines:
        print("Warning: 入力ファイルが空です", file=sys.stderr)
        sys.exit(1)
    
    log_debug(f"Loaded {len(input_lines)} lines from input file", debug)
    
    print(f"Processing {len(input_lines)} lines...")
    if sleep_convert:
        print(f"Custom convert sleep time: {sleep_convert}ms")
        log_debug(f"Custom convert sleep: {sleep_convert}ms", debug)
    else:
        print(f"Using default settings (Microsoft IME optimized)")
    
    if debug:
        print(f"Debug mode enabled - detailed logs will be written to {log_file}")
    
    print("(処理中はNotepadウィンドウが前面に表示されます)")
    print()
    
    results = []
    
    try:
        # AutoHotkeyプロセスを起動
        # 引数: [ahk_exe, script, sleep_convert, log_file]
        cmd = [ahk_exe, ahk_script]
        
        # 第1引数: sleep_convert（指定がない場合は "default"）
        if sleep_convert:
            cmd.append(str(sleep_convert))
        else:
            cmd.append("default")
        
        # 第2引数: log_file
        cmd.append(log_file)
        
        log_debug(f"Starting AHK process: {' '.join(cmd)}", debug)
        
        process = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8'
        )
        
        log_debug("AHK process started", debug)
        
        # 各行を処理
        for i, line in enumerate(input_lines, 1):
            log_debug(f"--- Processing line {i}/{len(input_lines)} ---", debug)
            log_debug(f"Input: [{line}] (length: {len(line)})", debug)
            
            # 進捗表示
            print(f"[{i:3d}/{len(input_lines)}] Converting: {line[:40]:<40}", end='\r')
            
            # AHKスクリプトに送信
            send_time = time.time()
            process.stdin.write(line + '\n')
            process.stdin.flush()
            log_debug(f"Sent to AHK at {send_time:.3f}", debug)
            
            # 結果を受信
            receive_start = time.time()
            result = process.stdout.readline().strip()
            receive_time = time.time()
            
            # INFO行をスキップ
            while result.startswith("INFO:"):
                log_debug(f"AHK Info: {result}", debug)
                result = process.stdout.readline().strip()
            
            log_debug(f"Received from AHK at {receive_time:.3f} (took {(receive_time - send_time):.3f}s)", debug)
            log_debug(f"Output: [{result}] (length: {len(result)})", debug)
            
            # 結果の検証
            if ' ' in result and result != line:
                log_debug(f"WARNING: Output contains unexpected spaces", debug)
            
            if result == line:
                log_debug(f"WARNING: Output equals input (no conversion occurred)", debug)
            
            results.append(result)
            
            if debug:
                print(f"\n  Input:  {line}")
                print(f"  Output: {result}")
        
        # プロセスを終了
        log_debug("Closing AHK process", debug)
        process.stdin.close()
        process.wait(timeout=10)
        log_debug("AHK process closed", debug)
        
        print("\n")
        print("✓ Conversion completed.")
        
    except subprocess.TimeoutExpired:
        process.kill()
        print("\nError: AutoHotkey process timeout", file=sys.stderr)
        log_debug("ERROR: Process timeout", debug)
        sys.exit(1)
    except KeyboardInterrupt:
        process.kill()
        print("\n\nInterrupted by user", file=sys.stderr)
        log_debug("ERROR: Interrupted by user", debug)
        sys.exit(1)
    except Exception as e:
        print(f"\nError during conversion: {e}", file=sys.stderr)
        log_debug(f"ERROR: {e}", debug)
        if process:
            process.kill()
        sys.exit(1)
    
    # 結果を出力ファイルに保存
    try:
        log_debug(f"Writing results to {output_file}", debug)
        with open(output_file, 'w', encoding='utf-8') as f:
            for result in results:
                f.write(result + '\n')
        print(f"✓ Results saved to: {output_file}")
        log_debug(f"Results written successfully", debug)
    except Exception as e:
        print(f"Error writing output file: {e}", file=sys.stderr)
        log_debug(f"ERROR writing output: {e}", debug)
        sys.exit(1)
    
    if debug:
        print(f"\n✓ Debug log saved to: {log_file}")
        log_debug("=== Python Debug Log Completed ===", debug)


def main():
    parser = argparse.ArgumentParser(
        description='IMEを使用してひらがなを漢字に変換するツール（デバッグ版）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  # 通常実行
  python kanakanji.py input.txt -o output.txt
  
  # デバッグモード（デフォルトログファイル: kanakanji_debug.log）
  python kanakanji.py input.txt -o output.txt --debug
  
  # デバッグモード + カスタムログファイル
  python kanakanji.py input.txt -o output.txt --debug --log msime_debug.log
  
  # 異なるIMEでテスト（ログを分ける）
  python kanakanji.py input.txt -o output_msime.txt --debug --log msime.log
  python kanakanji.py input.txt -o output_google.txt --debug --log google.log --sleep-convert 300
  python kanakanji.py input.txt -o output_mozc.txt --debug --log mozc.log --sleep-convert 400
  
デバッグモード:
  --debug を指定すると、以下の詳細情報がログに記録されます：
  - Python側: 標準エラー出力に詳細ログ
  - AHK側: 指定したログファイルに詳細ログ
    - 各処理ステップのタイムスタンプ
    - IME状態の変化
    - キー入力のタイミング
    - クリップボードの内容
    - Sleep時間の詳細
        """
    )
    
    parser.add_argument(
        'input_file',
        help='入力ファイル（ひらがな、1行1フレーズ）'
    )
    parser.add_argument(
        '-o', '--output',
        required=True,
        help='出力ファイル（変換結果）'
    )
    parser.add_argument(
        '--ahk-script',
        default='kanakanji.ahk',
        help='AutoHotkeyスクリプトのパス（デフォルト: kanakanji.ahk）'
    )
    parser.add_argument(
        '--sleep-convert',
        type=int,
        metavar='MILLISECONDS',
        help='変換処理後の待ち時間（ミリ秒）'
    )
    parser.add_argument(
        '--log',
        metavar='LOGFILE',
        help='ログファイル名（デフォルト: kanakanji_debug.log）'
    )
    parser.add_argument(
        '--debug',
        action='store_true',
        help='デバッグモード（詳細ログを出力）'
    )
    
    args = parser.parse_args()
    
    convert_with_ime(
        args.input_file,
        args.output,
        args.ahk_script,
        args.sleep_convert,
        args.log,
        args.debug
    )


if __name__ == '__main__':
    main()
