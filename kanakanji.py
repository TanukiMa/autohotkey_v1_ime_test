#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kanakanji.py - IMEを使用してひらがなを漢字に変換するツール（完全修正版）
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
    """AutoHotkey.exeのパスを検索"""
    common_paths = [
        r"C:\Program Files\AutoHotkey\AutoHotkey.exe",
        r"C:\Program Files\AutoHotkey\v1.1.37.02\AutoHotkey.exe",
        r"C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe",
        os.path.expanduser(r"~\AppData\Local\Programs\AutoHotkey\AutoHotkey.exe"),
    ]
    
    for path in common_paths:
        if os.path.exists(path):
            return path
    
    return shutil.which("AutoHotkey.exe")


def log_debug(message, debug=False):
    """デバッグメッセージを出力"""
    if debug:
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[{timestamp}] {message}", file=sys.stderr)
        sys.stderr.flush()


def load_input_lines(input_file, max_length=100, debug=False):
    """
    入力ファイルを読み込み、1行ずつ処理
    
    Args:
        input_file: 入力ファイルパス
        max_length: 1行の最大文字数
        debug: デバッグモード
    
    Returns:
        有効な行のリスト
    """
    input_lines = []
    
    try:
        with open(input_file, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                # 改行を除去（\r\nと\nの両方に対応）
                line = line.rstrip('\r\n')
                
                # 空行とコメント行をスキップ
                if not line or line.startswith('#'):
                    log_debug(f"Line {line_num}: Skipped (empty or comment)", debug)
                    continue
                
                # 文字数チェック
                if len(line) > max_length:
                    log_debug(f"Line {line_num}: WARNING - Too long ({len(line)} chars), truncating to {max_length}", debug)
                    line = line[:max_length]
                
                # 制御文字を除去
                line = ''.join(char for char in line if ord(char) >= 32 or char in '\t\n\r')
                
                if line:  # 空でない場合のみ追加
                    input_lines.append(line)
                    log_debug(f"Line {line_num}: Loaded [{line}] ({len(line)} chars)", debug)
                else:
                    log_debug(f"Line {line_num}: Skipped (became empty after cleaning)", debug)
    
    except FileNotFoundError:
        print(f"Error: 入力ファイル '{input_file}' が見つかりません", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error reading input file: {e}", file=sys.stderr)
        sys.exit(1)
    
    return input_lines


def convert_with_ime(input_file, output_file, ahk_script="kanakanji.ahk", 
                     sleep_convert=None, log_file=None, max_length=100, ime_mode="hiragana", debug=False):
    """IME変換処理"""
    
    log_debug("=== Python Debug Log Started ===", debug)
    
    if log_file is None:
        log_file = "kanakanji_debug.log"
    
    valid_ime_modes = {"hiragana", "fullalpha", "katakana", "direct"}
    if ime_mode not in valid_ime_modes:
        print(f"Error: Unsupported IME mode '{ime_mode}'. Choose from {', '.join(sorted(valid_ime_modes))}.", file=sys.stderr)
        sys.exit(1)
    
    log_debug(f"Target IME mode: {ime_mode}", debug)
    
    # AutoHotkeyの実行ファイルパスを検索
    ahk_exe = find_autohotkey()
    if not ahk_exe:
        print("Error: AutoHotkey.exe が見つかりません", file=sys.stderr)
        sys.exit(1)
    
    log_debug(f"Using AutoHotkey: {ahk_exe}", debug)
    
    # AHKスクリプトの存在確認
    if not os.path.exists(ahk_script):
        print(f"Error: {ahk_script} が見つかりません", file=sys.stderr)
        sys.exit(1)
    
    # IME.ahkの存在確認
    ime_ahk_path = os.path.join(os.path.dirname(os.path.abspath(ahk_script)), "IME.ahk")
    if not os.path.exists(ime_ahk_path):
        print(f"Error: IME.ahk が見つかりません", file=sys.stderr)
        sys.exit(1)
    
    # 入力ファイルを読み込み（改善版）
    input_lines = load_input_lines(input_file, max_length, debug)
    
    if not input_lines:
        print("Warning: 入力ファイルに有効な行がありません", file=sys.stderr)
        sys.exit(1)
    
    log_debug(f"Loaded {len(input_lines)} valid lines from input file", debug)
    
    print(f"Processing {len(input_lines)} lines (max {max_length} chars per line)...")
    if sleep_convert:
        print(f"Custom convert sleep time: {sleep_convert}ms")
    print(f"Target IME mode: {ime_mode}")
    
    if debug:
        print(f"Debug mode enabled - logs will be written to {log_file}")
    
    print("\n" + "="*60)
    print("重要: Notepadが起動したら、以下の設定を行ってください：")
    print("  1. IMEをONにする")
    print("  2. 入力モードを「ひらがな」に設定")
    print("  3. 10秒後に自動的に処理が開始されます")
    print("="*60 + "\n")
    
    results = []
    
    try:
        # AutoHotkeyプロセスを起動
        cmd = [ahk_exe, ahk_script]
        
        if sleep_convert:
            cmd.append(str(sleep_convert))
        else:
            cmd.append("default")
        
        cmd.append(log_file)
        cmd.append(ime_mode)
        
        log_debug(f"Starting AHK process: {' '.join(cmd)}", debug)
        
        process = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding='utf-8',
            bufsize=1,
            universal_newlines=True
        )
        
        log_debug("AHK process started", debug)
        
        # 各行を処理
        for i, line in enumerate(input_lines, 1):
            log_debug(f"--- Processing line {i}/{len(input_lines)} ---", debug)
            log_debug(f"Input: [{line}] (length: {len(line)})", debug)
            
            # 進捗表示
            print(f"[{i:3d}/{len(input_lines)}] Converting: {line[:50]:<50}", end='', flush=True)
            
            # AHKスクリプトに送信
            send_time = time.time()
            try:
                # 重要：改行コードは\nのみを送信
                process.stdin.write(line + '\n')
                process.stdin.flush()
                log_debug(f"Sent to AHK: [{line}]", debug)
            except BrokenPipeError:
                print("\nError: AHK process terminated unexpectedly", file=sys.stderr)
                log_debug("ERROR: Broken pipe", debug)
                break
            
            # 結果を受信
            receive_start = time.time()
            try:
                result = process.stdout.readline()
                if not result:
                    print("\nError: No output from AHK", file=sys.stderr)
                    log_debug("ERROR: EOF from AHK stdout", debug)
                    break
                
                # 改行を除去
                result = result.rstrip('\r\n')
                receive_time = time.time()
                
                # INFO行をスキップ
                while result.startswith("INFO:"):
                    log_debug(f"AHK Info: {result}", debug)
                    result = process.stdout.readline().rstrip('\r\n')
                
            except Exception as e:
                print(f"\nError reading from AHK: {e}", file=sys.stderr)
                log_debug(f"ERROR: {e}", debug)
                break
            
            log_debug(f"Received: [{result}] (took {(receive_time - send_time):.3f}s)", debug)
            
            # 結果の検証
            warnings = []
            if ' ' in result and result != line:
                warnings.append("spaces")
            if result == line:
                warnings.append("no conversion")
            if not result:
                warnings.append("empty")
            
            if warnings:
                print(f" [WARNING: {', '.join(warnings)}]", end='')
            
            print()  # 改行
            
            results.append(result if result else line)  # 空の場合は元の入力を保存
            
            if debug and result:
                print(f"  Input:  {line}")
                print(f"  Output: {result}")
        
        # プロセスを終了
        log_debug("Closing AHK process", debug)
        process.stdin.close()
        process.wait(timeout=10)
        log_debug("AHK process closed", debug)
        
        print("\n✓ Conversion completed.")
        
    except subprocess.TimeoutExpired:
        process.kill()
        print("\nError: Process timeout", file=sys.stderr)
        sys.exit(1)
    except KeyboardInterrupt:
        process.kill()
        print("\n\nInterrupted by user", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        if process:
            process.kill()
        sys.exit(1)
    
    # 結果を出力
    try:
        with open(output_file, 'w', encoding='utf-8', newline='\n') as f:
            for result in results:
                f.write(result + '\n')
        print(f"✓ Results saved to: {output_file}")
    except Exception as e:
        print(f"Error writing output: {e}", file=sys.stderr)
        sys.exit(1)
    
    if debug:
        print(f"\n✓ Debug log: {log_file}")


def main():
    parser = argparse.ArgumentParser(
        description='IMEを使用してひらがなを漢字に変換（完全修正版）',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python kanakanji.py input.txt -o output.txt --debug
  
入力ファイル形式:
  - 1行1フレーズ（ひらがなのみ）
  - 空行・コメント行（#）は無視
  - 1行の最大文字数: 100文字（デフォルト）
        """
    )
    
    parser.add_argument('input_file', help='入力ファイル')
    parser.add_argument('-o', '--output', required=True, help='出力ファイル')
    parser.add_argument('--ahk-script', default='kanakanji.ahk', help='AHKスクリプト')
    parser.add_argument('--sleep-convert', type=int, help='変換待ち時間（ms）')
    parser.add_argument('--log', help='ログファイル名')
    parser.add_argument('--max-length', type=int, default=100, help='1行の最大文字数（デフォルト: 100）')
    parser.add_argument('--ime-mode', choices=['hiragana', 'fullalpha', 'katakana', 'direct'], default='hiragana',
                        help='強制するIME入力モード（デフォルト: hiragana）')
    parser.add_argument('--debug', action='store_true', help='デバッグモード')
    
    args = parser.parse_args()
    
    convert_with_ime(
        args.input_file,
        args.output,
        args.ahk_script,
        args.sleep_convert,
        args.log,
        args.max_length,
        args.ime_mode,
        args.debug
    )


if __name__ == '__main__':
    main()
