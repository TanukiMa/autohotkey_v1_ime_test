# ime_comparison_test.py
import subprocess
import csv
import time
import os
import sys
from datetime import datetime

def find_autohotkey():
    """AutoHotkeyの実行ファイルを探す"""
    possible_paths = [
        r"C:\Program Files\AutoHotkey\AutoHotkey.exe",
        r"C:\Program Files\AutoHotkey\AutoHotkeyU64.exe",
        r"C:\Program Files\AutoHotkey\AutoHotkeyU32.exe",
        r"C:\Program Files (x86)\AutoHotkey\AutoHotkey.exe",
    ]
    
    for path in possible_paths:
        if os.path.exists(path):
            return path
    
    print("エラー: AutoHotkeyが見つかりません")
    print("https://www.autohotkey.com/ からダウンロードしてください")
    return None

def test_ime_conversion(ahk_path, hiragana, max_retries=2):
    """IME変換テストを実行"""
    for attempt in range(max_retries):
        try:
            process = subprocess.Popen(
                [ahk_path, 'ime_test_universal.ahk', hiragana],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            
            stdout, stderr = process.communicate(timeout=40)
            result = stdout.decode('utf-8', errors='ignore').strip()
            
            # エラーチェック
            if result.startswith('ERROR:'):
                print(f"    ⚠ {result}")
                if attempt < max_retries - 1:
                    print(f"    リトライ中... ({attempt + 1}/{max_retries})")
                    time.sleep(2)
                    continue
                return None
            
            return result
            
        except subprocess.TimeoutExpired:
            print(f"    ⚠ タイムアウト")
            process.kill()
            if attempt < max_retries - 1:
                print(f"    リトライ中... ({attempt + 1}/{max_retries})")
                time.sleep(2)
                continue
            return None
        except Exception as e:
            print(f"    ⚠ エラー: {e}")
            return None
    
    return None

def main():
    # AutoHotkeyのパス確認
    ahk_path = find_autohotkey()
    if not ahk_path:
        sys.exit(1)

    # スクリプトファイル確認
    if not os.path.exists('ime_test_universal.ahk'):
        print("エラー: ime_test_universal.ahk が見つかりません")
        sys.exit(1)

    # 引数チェック（1行1文ファイル）
    if len(sys.argv) < 2:
        print(f"使用方法: python {os.path.basename(sys.argv[0])} <1行1文の入力ファイル> [-o 出力ファイル]")
        sys.exit(1)
    input_file = sys.argv[1]
    if not os.path.exists(input_file):
        print(f"エラー: 入力ファイルがありません: {input_file}")
        sys.exit(1)

    # 出力ファイルオプション解析 (-o / --output)
    output_file = None
    i = 2
    while i < len(sys.argv):
        arg = sys.argv[i]
        if arg in ('-o', '--output'):
            if i + 1 < len(sys.argv):
                output_file = sys.argv[i + 1]
                i += 2
                continue
            else:
                print("エラー: -o/--output の後にファイル名を指定してください")
                sys.exit(1)
        i += 1

    print("=" * 70)
    print("IME変換テスト (ファイル入力モード)")
    print("=" * 70)
    print(f"AutoHotkey: {ahk_path}")
    print(f"入力ファイル: {input_file}")
    if output_file:
        print(f"出力ファイル: {output_file}")
    print()

    # 1行ずつ処理
    results = []
    with open(input_file, 'r', encoding='utf-8') as f:
        for idx, line in enumerate(f, 1):
            hiragana = line.strip()
            if not hiragana:
                continue
            print(f"[{idx}] 入力: {hiragana}", end=' ', flush=True)
            result = test_ime_conversion(ahk_path, hiragana)
            if result:
                converted = result != hiragana
                status = "✓" if converted else "×"
                print(f"→ {result} {status}")
                results.append({'input': hiragana, 'output': result})
            else:
                print("→ (取得失敗) ×")
                results.append({'input': hiragana, 'output': ''})
            time.sleep(0.5)

    if output_file:
        try:
            with open(output_file, 'w', encoding='utf-8', newline='') as out:
                writer = csv.writer(out)
                writer.writerow(['input', 'output'])
                for r in results:
                    writer.writerow([r['input'], r['output']])
            print(f"\n結果を書き出しました: {output_file}")
        except Exception as e:
            print(f"\nエラー: 出力ファイルへの書き込みに失敗しました: {e}")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n中断されました")
        sys.exit(0)
