"""
Kindleアプリのページを全てスクリーンショットしてPDFに変換するスクリプト
"""

import time
import sys
import re
from PIL import Image
import pyautogui
import pygetwindow as gw
from pathlib import Path
import numpy as np
from datetime import datetime
import hashlib

# Windows API
try:
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("警告: pywin32がインストールされていません。")


class KindleCapture:
    def __init__(self):
        self.window = None
        self.images = []
        self.image_hashes = []
        self.output_dir = Path("kindle_output")
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")

        # 設定
        self.page_turn_wait = 0.8  # ページ送り後の待機（短縮）
        self.initial_wait = 2.0   # 開始時の待機
        self.page_turn_key = 'left'  # ページ送りキー（日本語:left, 英語:right）

    def setup_output_dir(self, book_name):
        """出力フォルダを本の名前で作成"""
        self.output_dir = Path("kindle_output") / book_name
        self.output_dir.mkdir(parents=True, exist_ok=True)
        (self.output_dir / "screenshots").mkdir(exist_ok=True)

    def find_kindle_window(self):
        print("\nKindleウィンドウを検索中...")
        for window in gw.getAllWindows():
            try:
                if window.title and "kindle" in window.title.lower():
                    exclude = ["cursor", "chrome", "edge", "firefox", "code", "vscode"]
                    if not any(x in window.title.lower() for x in exclude):
                        if window.width > 0 and window.height > 0:
                            print(f"  発見: {window.title}")
                            self.window = window
                            return window
            except:
                continue
        print("[エラー] Kindleウィンドウが見つかりません")
        return None

    def get_book_name(self):
        """ウィンドウタイトルから書籍名を取得"""
        if not self.window:
            return "kindle_book"
        title = self.window.title
        print(f"  ウィンドウタイトル: {title}")  # デバッグ用

        # "Kindle" や "Amazon" を除去
        name = title
        name = re.sub(r'\s*[-–]\s*Kindle.*$', '', name, flags=re.IGNORECASE)
        name = re.sub(r'\s*[-–]\s*Amazon.*$', '', name, flags=re.IGNORECASE)
        name = re.sub(r'^Kindle\s*[-–]\s*', '', name, flags=re.IGNORECASE)
        name = re.sub(r'^Amazon\s*[-–]\s*', '', name, flags=re.IGNORECASE)
        name = re.sub(r'Kindle for PC', '', name, flags=re.IGNORECASE)
        name = re.sub(r'Kindle', '', name, flags=re.IGNORECASE)

        # ファイル名に使えない文字を除去
        name = re.sub(r'[\\/:*?"<>|]', '', name)
        name = name.strip(' -–')

        return name if name else "kindle_book"

    def get_content_region(self):
        left = max(0, self.window.left)
        top = max(0, self.window.top)
        width = self.window.width + min(0, self.window.left)
        height = self.window.height + min(0, self.window.top)
        top_margin = max(60, min(120, int(height * 0.08)))
        bottom_margin = max(20, min(50, int(height * 0.03)))
        return (left, top + top_margin, width, height - top_margin - bottom_margin)

    def activate_window(self):
        if not self.window:
            return
        if WIN32_AVAILABLE:
            try:
                hwnd = win32gui.FindWindow(None, self.window.title)
                if hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    win32gui.SetFocus(hwnd)
            except:
                pass
        try:
            self.window.activate()
        except:
            pass
        time.sleep(0.3)

    def take_screenshot(self):
        left, top, width, height = self.get_content_region()
        return pyautogui.screenshot(region=(left, top, width, height))

    def compute_hash(self, image):
        small = image.resize((50, 50)).convert('L')
        return hashlib.md5(np.array(small).tobytes()).hexdigest()

    def turn_page(self):
        pyautogui.press(self.page_turn_key)
        time.sleep(self.page_turn_wait)
        return True

    def capture_pages(self, page_count):
        book_name = self.get_book_name()
        self.setup_output_dir(book_name)
        print(f"  保存先: {self.output_dir}")
        self.images = []
        self.image_hashes = []
        same_count = 0

        print(f"\n{'='*50}")
        print(f"最大 {page_count} ページをキャプチャします")
        print(f"{'='*50}")
        print("3秒後に開始... Kindleを前面にしてください！")

        for i in range(3, 0, -1):
            print(f"  {i}...", end="\r")
            time.sleep(1)
        print("  開始！")

        self.activate_window()
        time.sleep(self.initial_wait)

        for page_num in range(1, page_count + 1):
            screenshot = self.take_screenshot()
            current_hash = self.compute_hash(screenshot)

            # 重複チェック（直前と同じなら）
            if self.image_hashes and current_hash == self.image_hashes[-1]:
                same_count += 1
                print(f"  [ページ {page_num}] 同じページ検出 ({same_count}回目)")
                if same_count >= 3:
                    print("\n最後のページに到達しました")
                    break
            else:
                same_count = 0
                self.images.append(screenshot)
                self.image_hashes.append(current_hash)

                # 保存
                path = self.output_dir / "screenshots" / f"page_{len(self.images):04d}.png"
                screenshot.save(path)
                print(f"  [ページ {len(self.images)}/{page_count}] OK")

            if page_num < page_count:
                self.turn_page()

        return self.images

    def save_pdf(self):
        if not self.images:
            print("画像がありません")
            return None

        book_name = self.get_book_name()
        pdf_path = self.output_dir / f"{book_name}.pdf"
        print(f"\nPDF作成中: {pdf_path}")

        rgb_images = [img.convert('RGB') for img in self.images]
        rgb_images[0].save(
            pdf_path,
            "PDF",
            resolution=150.0,
            save_all=True,
            append_images=rgb_images[1:] if len(rgb_images) > 1 else []
        )

        print(f"[OK] PDF作成完了！")
        return pdf_path


def main():
    print("=" * 50)
    print("Kindle to PDF Converter")
    print("=" * 50)

    capture = KindleCapture()

    if not capture.find_kindle_window():
        print("\nKindleアプリを起動して本を開いてから再実行してください")
        return

    print(f"\n[OK] {capture.window.title}")

    # 本の種類を選択
    print("\n本の種類を選んでください:")
    print("  1: 日本語/縦書き（←キーで進む）")
    print("  2: 英語/横書き（→キーで進む）")
    choice = input("番号を入力 [1]: ").strip()
    if choice == "2":
        capture.page_turn_key = 'right'
        print("→ 英語/横書きモード")
    else:
        capture.page_turn_key = 'left'
        print("→ 日本語/縦書きモード")

    # ページ数（自動で最後まで）
    page_count = 9999
    print("\n最後のページまで自動でキャプチャします")

    # キャプチャ実行
    images = capture.capture_pages(page_count)

    if images:
        pdf_path = capture.save_pdf()
        print(f"\n{'='*50}")
        print(f"完了！ {len(images)}ページ")
        print(f"PDF: {pdf_path}")
        print(f"スクショ: {capture.output_dir / 'screenshots'}")
        print(f"{'='*50}")


if __name__ == "__main__":
    pyautogui.FAILSAFE = False
    main()
