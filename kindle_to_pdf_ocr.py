"""
Kindleアプリのページを全てスクリーンショットしてPDFに変換し、OCRで文字起こしするスクリプト
改善版v2: 最後のページ自動検出、ページ飛ばし防止、OCR縦書き対応
"""

import time
import os
import sys
import json
import threading
import re
from PIL import Image, ImageEnhance, ImageFilter
import pyautogui
import pygetwindow as gw
from pathlib import Path
import numpy as np
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, asdict, field
from typing import Optional, List, Tuple
import hashlib

# Windows APIを使用して確実にウィンドウをアクティブ化
try:
    import win32gui
    import win32con
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print("警告: pywin32がインストールされていません。")

# OCRライブラリのインポート
try:
    import easyocr
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False

try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except ImportError:
    TESSERACT_AVAILABLE = False


@dataclass
class Config:
    """設定クラス"""
    # ページ送り関連
    page_turn_wait: float = 1.0  # ページ送り後の待機時間（長めに設定）
    activation_wait: float = 0.5  # ウィンドウアクティブ化後の待機時間
    initial_wait: float = 2.0  # 最初のページ読み込み待機時間
    page_turn_retry: int = 5  # ページ送りのリトライ回数

    # 最後のページ検出
    auto_stop_on_last_page: bool = True  # 最後のページで自動停止
    same_page_threshold: int = 3  # 同じページが続いたら停止する回数

    # OCR関連
    ocr_engine: str = "easyocr"  # "easyocr" or "tesseract"
    ocr_languages: List[str] = None  # OCR言語 ["ja", "en"]
    ocr_parallel_workers: int = 2  # 並列OCR処理のワーカー数（少なめに）
    vertical_text: bool = True  # 縦書きテキスト対応

    # 画像処理関連
    image_preprocessing: bool = True  # OCR前処理を行うか
    contrast_factor: float = 1.3  # コントラスト調整係数
    sharpness_factor: float = 1.2  # シャープネス調整係数

    # 出力関連
    output_dir: str = "kindle_output"  # 出力ディレクトリ
    save_intermediate: bool = True  # 中間ファイルを保存するか
    pdf_resolution: float = 150.0  # PDF解像度

    # 重複検出（厳密に）
    similarity_threshold: float = 0.90  # 画像類似度の閾値（高めに設定）
    use_hash_comparison: bool = True  # ハッシュ比較も使用

    def __post_init__(self):
        if self.ocr_languages is None:
            self.ocr_languages = ["ja", "en"]

    def save(self, path: str):
        """設定をJSONファイルに保存"""
        data = asdict(self)
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    @classmethod
    def load(cls, path: str) -> 'Config':
        """JSONファイルから設定を読み込み"""
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return cls(**data)


class KindleCapture:
    """Kindleスクリーンショット取得クラス"""

    def __init__(self, config: Config):
        self.config = config
        self.window = None
        self.images: List[Image.Image] = []
        self.image_hashes: List[str] = []  # 画像ハッシュを保存
        self.output_dir = Path(config.output_dir)
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.consecutive_same_pages = 0  # 連続して同じページだった回数

    def setup_output_dir(self):
        """出力ディレクトリを作成"""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        if self.config.save_intermediate:
            (self.output_dir / "screenshots").mkdir(exist_ok=True)

    def find_kindle_window(self) -> Optional[object]:
        """Kindleアプリのウィンドウを検索"""
        try:
            all_windows = gw.getAllWindows()
            kindle_windows = []

            print("\nKindleウィンドウを検索中...")

            for window in all_windows:
                try:
                    if window.title:
                        title_lower = window.title.lower()

                        if "kindle" in title_lower:
                            exclude_keywords = [
                                "cursor", "brave", "chrome", "edge", "firefox",
                                "safari", "python", "cmd", "powershell",
                                "visual studio", "code", "vscode"
                            ]
                            should_exclude = any(kw in title_lower for kw in exclude_keywords)

                            if not should_exclude and window.width > 0 and window.height > 0:
                                kindle_windows.append(window)
                                print(f"  発見: {window.title}")
                except Exception:
                    continue

            if not kindle_windows:
                print("\n[エラー] Kindleウィンドウが見つかりません。")
                return None

            # 優先順位で選択
            for window in kindle_windows:
                if "kindle for pc" in window.title.lower():
                    self.window = window
                    return window

            for window in kindle_windows:
                if len(window.title) > 10:
                    self.window = window
                    return window

            self.window = kindle_windows[0]
            return kindle_windows[0]

        except Exception as e:
            print(f"[エラー] ウィンドウ検索エラー: {e}")
            return None

    def get_content_region(self) -> Tuple[int, int, int, int]:
        """コンテンツ領域を取得"""
        if not self.window:
            raise ValueError("ウィンドウが設定されていません")

        left = max(0, self.window.left)
        top = max(0, self.window.top)
        width = self.window.width + min(0, self.window.left)
        height = self.window.height + min(0, self.window.top)

        top_margin = max(60, min(120, int(height * 0.08)))
        bottom_margin = max(20, min(50, int(height * 0.03)))

        return (left, top + top_margin, width, height - top_margin - bottom_margin)

    def activate_window(self):
        """ウィンドウをアクティブ化"""
        if not self.window:
            return

        if WIN32_AVAILABLE:
            try:
                hwnd = win32gui.FindWindow(None, self.window.title)
                if hwnd:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    time.sleep(0.1)
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.1)
                    win32gui.SetFocus(hwnd)
            except Exception:
                pass

        try:
            self.window.activate()
        except Exception:
            pass

        time.sleep(self.config.activation_wait)

        # コンテンツ領域の中央をクリック
        left, top, width, height = self.get_content_region()
        center_x = left + width // 2
        center_y = top + height // 2
        pyautogui.click(center_x, center_y)
        time.sleep(0.2)

    def take_screenshot(self) -> Image.Image:
        """スクリーンショットを取得"""
        left, top, width, height = self.get_content_region()
        return pyautogui.screenshot(region=(left, top, width, height))

    def compute_image_hash(self, image: Image.Image) -> str:
        """画像のハッシュ値を計算"""
        # 小さくリサイズしてハッシュ計算（高速化）
        small = image.resize((50, 50)).convert('L')
        return hashlib.md5(np.array(small).tobytes()).hexdigest()

    def images_are_similar(self, img1: Image.Image, img2: Image.Image) -> bool:
        """2つの画像が類似しているか判定（より厳密に）"""
        try:
            # ハッシュ比較（高速）
            if self.config.use_hash_comparison:
                hash1 = self.compute_image_hash(img1)
                hash2 = self.compute_image_hash(img2)
                if hash1 == hash2:
                    return True

            # ピクセル比較（詳細）
            img1_resized = img1.resize((150, 150))
            img2_resized = img2.resize((150, 150))

            arr1 = np.array(img1_resized).astype(float)
            arr2 = np.array(img2_resized).astype(float)

            diff = np.abs(arr1 - arr2)
            similarity = 1.0 - (np.mean(diff) / 255.0)

            return similarity >= self.config.similarity_threshold
        except Exception:
            return False

    def is_duplicate_page(self, image: Image.Image) -> bool:
        """既にキャプチャ済みのページかチェック"""
        if not self.images:
            return False

        current_hash = self.compute_image_hash(image)

        # 直近5ページと比較
        for prev_hash in self.image_hashes[-5:]:
            if current_hash == prev_hash:
                return True

        # 最後のページと詳細比較
        if self.images_are_similar(self.images[-1], image):
            return True

        return False

    def turn_page(self) -> Tuple[bool, Image.Image]:
        """
        ページを送る
        Returns: (ページが変わったか, 新しいページのスクリーンショット)
        """
        before_screenshot = self.take_screenshot()
        before_hash = self.compute_image_hash(before_screenshot)

        for attempt in range(self.config.page_turn_retry):
            # ウィンドウを確実にアクティブ化
            self.activate_window()
            time.sleep(0.2)

            # 左矢印キーでページ送り
            pyautogui.press('left')

            # ページ読み込み待機（徐々に長くする）
            wait_time = self.config.page_turn_wait + (attempt * 0.3)
            time.sleep(wait_time)

            # 新しいスクリーンショットを取得
            after_screenshot = self.take_screenshot()
            after_hash = self.compute_image_hash(after_screenshot)

            # ハッシュが異なる = ページが変わった
            if before_hash != after_hash:
                # さらにピクセル比較で確認
                if not self.images_are_similar(before_screenshot, after_screenshot):
                    return (True, after_screenshot)

            print(f"    リトライ {attempt + 1}/{self.config.page_turn_retry}...", end="\r")

        # 最大リトライ後も変わらない = 最後のページの可能性
        return (False, self.take_screenshot())

    def capture_pages(self, page_count: int) -> List[Image.Image]:
        """指定されたページ数をキャプチャ"""
        self.setup_output_dir()
        self.images = []
        self.image_hashes = []
        self.consecutive_same_pages = 0

        print(f"\n{'='*60}")
        print(f"最大 {page_count} ページをキャプチャします...")
        print(f"（最後のページを検出したら自動停止）")
        print(f"{'='*60}")
        print("中断する場合は Ctrl+C を押してください")

        # カウントダウン
        for i in range(3, 0, -1):
            print(f"  {i}秒後に開始...", end="\r")
            time.sleep(1)
        print("  開始！" + " " * 20)

        # ウィンドウをアクティブ化
        self.activate_window()
        time.sleep(self.config.initial_wait)

        try:
            page_num = 1
            while page_num <= page_count:
                self.activate_window()
                time.sleep(0.3)

                # スクリーンショット取得
                screenshot = self.take_screenshot()

                # 重複チェック
                if self.is_duplicate_page(screenshot):
                    self.consecutive_same_pages += 1
                    print(f"  [ページ {page_num}] 重複検出 ({self.consecutive_same_pages}回目)")

                    if (self.config.auto_stop_on_last_page and
                        self.consecutive_same_pages >= self.config.same_page_threshold):
                        print(f"\n[INFO] 最後のページに到達したため自動停止します。")
                        break
                else:
                    self.consecutive_same_pages = 0

                    # 新しいページを保存
                    self.images.append(screenshot)
                    self.image_hashes.append(self.compute_image_hash(screenshot))

                    # 中間保存
                    if self.config.save_intermediate:
                        screenshot_path = self.output_dir / "screenshots" / f"page_{len(self.images):04d}.png"
                        screenshot.save(screenshot_path)

                    print(f"  [ページ {len(self.images)}/{page_count}] キャプチャ完了")

                # 最後のページでなければページ送り
                if page_num < page_count:
                    page_changed, _ = self.turn_page()

                    if not page_changed:
                        self.consecutive_same_pages += 1
                        if (self.config.auto_stop_on_last_page and
                            self.consecutive_same_pages >= self.config.same_page_threshold):
                            print(f"\n[INFO] ページが変わらないため自動停止します。")
                            break

                page_num += 1

        except KeyboardInterrupt:
            print(f"\n中断されました。{len(self.images)}ページまでキャプチャ。")

        return self.images


class ImagePreprocessor:
    """OCR用画像前処理クラス"""

    def __init__(self, config: Config):
        self.config = config

    def preprocess_for_kindle(self, image: Image.Image) -> Image.Image:
        """Kindle特化の前処理"""
        if image.mode != 'RGB':
            image = image.convert('RGB')

        # コントラスト調整
        enhancer = ImageEnhance.Contrast(image)
        enhanced = enhancer.enhance(self.config.contrast_factor)

        # シャープネス調整
        enhancer = ImageEnhance.Sharpness(enhanced)
        enhanced = enhancer.enhance(self.config.sharpness_factor)

        return enhanced


class OCRProcessor:
    """OCR処理クラス（縦書き対応・文字連結改善）"""

    def __init__(self, config: Config):
        self.config = config
        self.preprocessor = ImagePreprocessor(config)
        self.reader = None
        self._init_lock = threading.Lock()

    def _init_easyocr(self):
        """EasyOCRリーダーを初期化"""
        if self.reader is None:
            with self._init_lock:
                if self.reader is None:
                    print("  EasyOCRを初期化中（初回は時間がかかります）...")
                    self.reader = easyocr.Reader(
                        self.config.ocr_languages,
                        gpu=False,
                        verbose=False
                    )
                    print("  [OK] EasyOCR初期化完了")

    def clean_ocr_text(self, text_blocks: List[Tuple]) -> str:
        """
        OCR結果を読みやすい形式に整形
        text_blocks: EasyOCRの結果 [(bbox, text, confidence), ...]
        """
        if not text_blocks:
            return ""

        # 各テキストブロックの位置情報を取得
        blocks_with_pos = []
        for block in text_blocks:
            bbox = block[0]
            text = block[1]
            confidence = block[2] if len(block) > 2 else 1.0

            # バウンディングボックスの中心座標を計算
            x_center = sum(p[0] for p in bbox) / 4
            y_center = sum(p[1] for p in bbox) / 4

            # バウンディングボックスの幅と高さ
            x_coords = [p[0] for p in bbox]
            y_coords = [p[1] for p in bbox]
            width = max(x_coords) - min(x_coords)
            height = max(y_coords) - min(y_coords)

            blocks_with_pos.append({
                'text': text,
                'x': x_center,
                'y': y_center,
                'width': width,
                'height': height,
                'confidence': confidence,
                'top': min(y_coords),
                'left': min(x_coords)
            })

        # 縦書き判定（縦長の文字が多い場合）
        is_vertical = False
        if self.config.vertical_text:
            vertical_count = sum(1 for b in blocks_with_pos if b['height'] > b['width'] * 1.5)
            is_vertical = vertical_count > len(blocks_with_pos) * 0.3

        if is_vertical:
            # 縦書き: 右から左、上から下にソート
            blocks_with_pos.sort(key=lambda b: (-b['left'], b['top']))
        else:
            # 横書き: 上から下、左から右にソート
            blocks_with_pos.sort(key=lambda b: (b['top'], b['left']))

        # テキストを連結
        lines = []
        current_line = []
        prev_block = None
        line_threshold = 30  # 同じ行とみなす閾値（ピクセル）

        for block in blocks_with_pos:
            text = block['text'].strip()
            if not text:
                continue

            if prev_block is None:
                current_line.append(text)
            else:
                # 同じ行かどうか判定
                if is_vertical:
                    same_line = abs(block['left'] - prev_block['left']) < line_threshold
                else:
                    same_line = abs(block['top'] - prev_block['top']) < line_threshold

                if same_line:
                    current_line.append(text)
                else:
                    # 新しい行
                    lines.append(''.join(current_line))
                    current_line = [text]

            prev_block = block

        # 最後の行を追加
        if current_line:
            lines.append(''.join(current_line))

        # 行を整形して返す
        return '\n'.join(lines)

    def ocr_single_image(self, image: Image.Image, page_num: int) -> Tuple[int, str]:
        """1枚の画像をOCR処理"""
        try:
            # 前処理
            processed = self.preprocessor.preprocess_for_kindle(image)
            img_array = np.array(processed)

            if self.config.ocr_engine == "easyocr" and EASYOCR_AVAILABLE:
                self._init_easyocr()

                # detail=1 で詳細な結果を取得
                results = self.reader.readtext(img_array, detail=1)

                # 整形して連結
                text = self.clean_ocr_text(results)

            elif TESSERACT_AVAILABLE:
                # Tesseractの場合は縦書き対応オプションを使用
                if self.config.vertical_text:
                    custom_config = r'--psm 5 -l jpn+eng'  # 縦書きモード
                else:
                    custom_config = r'--psm 6 -l jpn+eng'  # 横書きモード
                text = pytesseract.image_to_string(processed, config=custom_config)
            else:
                text = "[OCRエンジンが利用できません]"

            return (page_num, text)
        except Exception as e:
            return (page_num, f"[OCRエラー: {e}]")

    def format_output_text(self, results: dict) -> str:
        """出力テキストを見やすく整形"""
        output_lines = []

        for page_num in sorted(results.keys()):
            text = results[page_num]

            # 空白行を削除して整形
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            formatted_text = '\n'.join(lines)

            output_lines.append(f"\n{'='*50}")
            output_lines.append(f"ページ {page_num}")
            output_lines.append(f"{'='*50}\n")
            output_lines.append(formatted_text)

        return '\n'.join(output_lines)

    def process_images(self, images: List[Image.Image], output_path: str) -> bool:
        """画像リストをOCR処理"""
        print(f"\n{'='*60}")
        print("OCR処理を開始します...")
        print(f"{'='*60}")

        if not EASYOCR_AVAILABLE and not TESSERACT_AVAILABLE:
            print("[エラー] OCRライブラリがインストールされていません。")
            print("  pip install easyocr")
            return False

        results = {}

        # EasyOCRは並列処理に向かないので、逐次処理を推奨
        # （モデルのメモリ使用量が大きい）
        if self.config.ocr_engine == "easyocr" and EASYOCR_AVAILABLE:
            self._init_easyocr()

        for i, img in enumerate(images, 1):
            print(f"  [OCR {i}/{len(images)}] 処理中...", end="")
            page_num, text = self.ocr_single_image(img, i)
            results[page_num] = text
            # 抽出文字数を表示
            char_count = len(text.replace('\n', '').replace(' ', ''))
            print(f" 完了 ({char_count}文字)")

        # 見やすい形式で保存
        formatted_text = self.format_output_text(results)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(formatted_text)

        total_chars = sum(len(t.replace('\n', '').replace(' ', '')) for t in results.values())
        print(f"\n[OK] OCR完了: {output_path}")
        print(f"  合計文字数: {total_chars}文字")

        return True


class PDFConverter:
    """PDF変換クラス"""

    def __init__(self, config: Config):
        self.config = config

    def convert(self, images: List[Image.Image], output_path: str):
        """画像リストをPDFに変換"""
        if not images:
            raise ValueError("画像がありません")

        print(f"\nPDFを作成中: {output_path}")

        rgb_images = [img.convert('RGB') for img in images]

        rgb_images[0].save(
            output_path,
            "PDF",
            resolution=self.config.pdf_resolution,
            save_all=True,
            append_images=rgb_images[1:] if len(rgb_images) > 1 else []
        )

        print(f"[OK] PDF作成完了: {output_path}")


def get_page_count() -> int:
    """ページ数を入力"""
    while True:
        try:
            count = int(input("最大ページ数を入力してください（最後のページで自動停止します）: "))
            if count > 0:
                return count
            print("1以上の数値を入力してください。")
        except ValueError:
            print("有効な数値を入力してください。")


def main():
    """メイン処理"""
    print("=" * 60)
    print("Kindle to PDF & OCR Converter (改善版 v2)")
    print("=" * 60)
    print("\n機能:")
    print("  - 最後のページを自動検出して停止")
    print("  - ページ飛ばし防止（厳密な重複チェック）")
    print("  - 縦書きテキスト対応OCR")
    print("  - 見やすい形式でテキスト出力")

    # 設定を読み込みまたは新規作成
    config_path = "kindle_config.json"
    if Path(config_path).exists():
        try:
            config = Config.load(config_path)
            print(f"\n設定を読み込みました: {config_path}")
        except Exception:
            config = Config()
            print(f"\n設定ファイルが読み込めないためデフォルト設定を使用")
    else:
        config = Config()
        config.save(config_path)
        print(f"\nデフォルト設定を保存しました: {config_path}")

    # コマンドライン引数からページ数を取得
    page_count = None
    if len(sys.argv) > 1:
        try:
            page_count = int(sys.argv[1])
            print(f"コマンドライン引数からページ数: {page_count}")
        except ValueError:
            print(f"警告: 無効なページ数 '{sys.argv[1]}'")

    # Kindleウィンドウを検索
    capture = KindleCapture(config)
    kindle_window = capture.find_kindle_window()

    if not kindle_window:
        print("\n[ヒント] Kindleアプリを起動し、本を開いてから再実行してください。")
        return

    print(f"\n[OK] Kindleウィンドウ: {kindle_window.title}")
    print(f"  位置: ({kindle_window.left}, {kindle_window.top})")
    print(f"  サイズ: {kindle_window.width} x {kindle_window.height}")

    # 確認
    if page_count is None:
        print("\nEnterで続行、Ctrl+Cで中断")
        try:
            input()
        except KeyboardInterrupt:
            print("\n中断されました。")
            return
        page_count = get_page_count()

    try:
        # キャプチャ
        images = capture.capture_pages(page_count)

        if not images:
            print("キャプチャされた画像がありません。")
            return

        print(f"\n合計 {len(images)} ページをキャプチャしました。")

        # PDF変換
        output_dir = Path(config.output_dir)
        pdf_path = output_dir / f"kindle_{capture.session_id}.pdf"

        converter = PDFConverter(config)
        converter.convert(images, str(pdf_path))

        # OCR処理
        text_path = output_dir / f"kindle_{capture.session_id}.txt"
        ocr_processor = OCRProcessor(config)
        ocr_success = ocr_processor.process_images(images, str(text_path))

        # 結果表示
        print(f"\n{'='*60}")
        print("[OK] すべての処理が完了しました！")
        print(f"{'='*60}")
        print(f"  PDF: {pdf_path}")
        if ocr_success:
            print(f"  テキスト: {text_path}")
        if config.save_intermediate:
            print(f"  スクリーンショット: {output_dir / 'screenshots'}")
        print(f"\n出力フォルダ: {output_dir.absolute()}")

    except Exception as e:
        print(f"\nエラー: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    pyautogui.FAILSAFE = False
    main()
