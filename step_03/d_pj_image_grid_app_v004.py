import io
import json
import logging
import os
import sys
import tempfile
from dataclasses import asdict, dataclass, field
from functools import lru_cache
from typing import Any, Dict, List, Optional, Tuple, Union

from PIL import Image, UnidentifiedImageError
from PyQt6.QtCore import QRectF, QSize, Qt, QThread, pyqtSignal
from PyQt6.QtGui import (QColor, QDragEnterEvent, QDropEvent, QImage, QPainter,
                         QPen, QPixmap)
from PyQt6.QtWidgets import (QApplication, QCheckBox, QColorDialog, QComboBox,
                             QDialog, QDialogButtonBox, QDoubleSpinBox,
                             QFileDialog, QFrame, QGridLayout, QGroupBox,
                             QHBoxLayout, QLabel, QLineEdit, QListWidget,
                             QListWidgetItem, QMainWindow, QMenu, QMenuBar,
                             QMessageBox, QProgressDialog, QPushButton,
                             QScrollArea, QSpinBox, QSplitter, QVBoxLayout,
                             QWidget)
from reportlab.lib.pagesizes import A3, A4
from reportlab.pdfgen import canvas

# ロギングの設定
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# 定数定義
MM_TO_PT: float = 2.83465  # ミリメートルからポイントへの変換係数
DEFAULT_ROW_HEIGHT_MM: float = 100.0
DEFAULT_COL_WIDTH_MM: float = 100.0
DEFAULT_GRID_WIDTH: int = 1
DEFAULT_PREVIEW_HEIGHT: int = 600
THUMBNAIL_SIZE: Tuple[int, int] = (200, 200)  # プレビュー用のサムネイルサイズ
SETTINGS_FILE: str = "grid_settings.json"  # 設定ファイルのパス

# サポートする画像形式
SUPPORTED_IMAGE_FORMATS = {
    "Images": "*.png *.jpg *.jpeg",
    "PSD Files": "*.psd",
    "PDF Files": "*.pdf",
    "All Supported Files": "*.png *.jpg *.jpeg *.psd *.pdf"
}

try:
    import psd_tools
    PSD_SUPPORT = True
    logger.info("PSD形式のサポートが有効です")
except ImportError:
    PSD_SUPPORT = False
    logger.warning("psd-toolsがインストールされていません。PSD形式はサポートされません。")

try:
    import fitz  # PyMuPDF
    PDF_SUPPORT = True
    logger.info("PDF形式のサポートが有効です")
except ImportError:
    PDF_SUPPORT = False
    logger.warning("PyMuPDFがインストールされていません。PDF形式はサポートされません。")

@dataclass
class GridSettings:
    # ... (GridSettings クラス - 変更なし)
    """グリッド設定を保持するデータクラス"""
    row_height_mm: float = DEFAULT_ROW_HEIGHT_MM
    col_width_mm: float = DEFAULT_COL_WIDTH_MM
    grid_line_visible: bool = True
    grid_color: QColor = field(default_factory=lambda: QColor(0, 0, 0))
    grid_width: int = DEFAULT_GRID_WIDTH
    page_size: Tuple[float, float] = A4

    def to_dict(self) -> Dict[str, Any]:
        """設定を辞書形式に変換"""
        settings_dict = asdict(self)
        # QColorをRGB値のタプルに変換
        settings_dict['grid_color'] = self.grid_color.getRgb()
        # ページサイズを文字列に変換
        settings_dict['page_size'] = 'A4' if self.page_size == A4 else 'A3'
        return settings_dict

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'GridSettings':
        """辞書から設定を復元"""
        # QColorを復元
        if 'grid_color' in data:
            color_data = data['grid_color']
            data['grid_color'] = QColor(*color_data)

        # ページサイズを復元
        if 'page_size' in data:
            data['page_size'] = A4 if data['page_size'] == 'A4' else A3

        return cls(**data)

    def save_to_file(self, file_path: str = SETTINGS_FILE) -> None:
        """設定をファイルに保存"""
        try:
            # 既存の設定ファイルをバックアップ
            if os.path.exists(file_path):
                backup_path = f"{file_path}.backup"
                import shutil
                shutil.copy2(file_path, backup_path)

            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, indent=4)
            logger.info(f"設定を保存しました: {file_path}")
        except Exception as e:
            logger.error(f"設定の保存中にエラーが発生しました: {e}")
            # バックアップから復元
            backup_path = f"{file_path}.backup"
            if os.path.exists(backup_path):
                try:
                    import shutil
                    shutil.copy2(backup_path, file_path)
                    logger.info("設定ファイルをバックアップから復元しました")
                except Exception as restore_error:
                    logger.error(f"設定ファイルの復元中にエラーが発生しました: {restore_error}")
            raise

    @classmethod
    def load_from_file(cls, file_path: str = SETTINGS_FILE) -> 'GridSettings':
        """設定をファイルから読み込み"""
        try:
            if not os.path.exists(file_path):
                logger.info(f"設定ファイルが存在しません: {file_path}")
                return cls()

            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return cls.from_dict(data)
        except json.JSONDecodeError as e:
            logger.error(f"設定ファイルの形式が不正です: {e}")
            # 破損した設定ファイルをバックアップ
            backup_path = f"{file_path}.backup"
            try:
                if os.path.exists(file_path):
                    import shutil
                    shutil.copy2(file_path, backup_path)
                    os.remove(file_path)
            except Exception as backup_error:
                logger.error(f"設定ファイルのバックアップ中にエラーが発生しました: {backup_error}")
            return cls()
        except Exception as e:
            logger.error(f"設定の読み込み中にエラーが発生しました: {e}")
            return cls()


class ImageProcessor:
    # ... (ImageProcessor クラス - _load_psd() メソッドを修正)
    """画像処理クラス"""

    def __init__(self):
        self.cmyk_profile_path = None
        self.color_conversion_intent = 'perceptual'  # perceptual, relative, saturation, absolute
        logger.info("ImageProcessor initialized")

    def set_cmyk_profile(self, profile_path: str) -> None:
        """CMYKプロファイルを設定"""
        if os.path.exists(profile_path):
            self.cmyk_profile_path = profile_path
            logger.info(f"CMYKプロファイルを設定: {profile_path}")
        else:
            logger.error(f"CMYKプロファイルが見つかりません: {profile_path}")
            raise ValueError(f"ICCプロファイルが見つかりません: {profile_path}")

    def set_color_conversion_intent(self, intent: str) -> None:
        """色変換方法を設定"""
        valid_intents = ['perceptual', 'relative', 'saturation', 'absolute']
        if intent not in valid_intents:
            logger.error(f"無効な色変換方法です: {intent}")
            raise ValueError(f"無効な色変換方法です: {intent}")
        self.color_conversion_intent = intent
        logger.info(f"色変換方法を設定: {intent}")

    @staticmethod
    def convert_to_cmyk(image: Image.Image, profile_path: str = None,
                       intent: str = 'perceptual') -> Image.Image:
        """画像をCMYK形式に変換（ICCプロファイル対応）"""
        try:
            from PIL import ImageCms

            if profile_path and os.path.exists(profile_path):
                # ICCプロファイルを使用した変換
                srgb_profile = ImageCms.createProfile("sRGB")
                cmyk_profile = ImageCms.getOpenProfile(profile_path)

                if image.mode == 'RGBA':
                    # アルファチャンネルを白背景で合成
                    background = Image.new('RGB', image.size, (255, 255, 255))
                    background.paste(image, mask=image.split()[3])
                    image = background

                if image.mode != 'RGB':
                    image = image.convert('RGB')

                # 色変換方法を設定
                if intent == 'perceptual':
                    intent = ImageCms.INTENT_PERCEPTUAL
                elif intent == 'relative':
                    intent = ImageCms.INTENT_RELATIVE_COLORIMETRIC
                elif intent == 'saturation':
                    intent = ImageCms.INTENT_SATURATION
                else:  # absolute
                    intent = ImageCms.INTENT_ABSOLUTE_COLORIMETRIC

                # 正確なプロファイル変換を適用
                return ImageCms.profileToProfile(
                    image, srgb_profile, cmyk_profile,
                    outputMode='CMYK',
                    intent=intent
                )
            else:
                # 従来の変換方法にフォールバック
                if image.mode == 'CMYK':
                    return image
                elif image.mode == 'RGBA':
                    background = Image.new('RGB', image.size, (255, 255, 255))
                    background.paste(image, mask=image.split()[3])
                    image = background
                return image.convert('CMYK')

        except ImportError:
            # ImageCmsが利用できない場合は従来の変換方法を使用
            if image.mode == 'CMYK':
                return image
            elif image.mode == 'RGBA':
                background = Image.new('RGB', image.size, (255, 255, 255))
                background.paste(image, mask=image.split()[3])
                image = background
            return image.convert('CMYK')

    @staticmethod
    def load_image(file_path: str) -> Optional[Image.Image]:
        """画像ファイルを読み込む"""
        try:
            logger.info(f"画像の読み込みを開始: {file_path}")

            # ファイル拡張子を取得
            ext = os.path.splitext(file_path)[1].lower()

            if ext == '.psd' and PSD_SUPPORT:
                logger.info("PSDファイルを読み込み")
                return ImageProcessor._load_psd(file_path)
            elif ext == '.pdf' and PDF_SUPPORT:
                logger.info("PDFファイルを読み込み")
                return ImageProcessor._load_pdf(file_path)
            else:
                # 通常の画像ファイル
                logger.info("通常の画像ファイルを読み込み")
                return Image.open(file_path)
        except Exception as e:
            logger.error(f"画像の読み込みに失敗: {file_path}, エラー: {e}", exc_info=True)
            return None

    @staticmethod
    def _load_psd(file_path: str) -> Optional[Image.Image]:
        """PSDファイルを読み込む"""
        psd_image = None  # psd_image を初期化
        try:
            logger.info(f"PSDファイルを読み込み開始: {file_path}")
            with psd_tools.PSDImage.open(file_path) as psd: # with ステートメントを使用
                psd_image = psd # psd_image に代入 (finally ブロックで使用するため)

                # レイヤー情報をログ出力
                logger.info(f"レイヤー数: {len(psd)}")
                for i, layer in enumerate(psd):
                    logger.info(f"レイヤー {i}: 名前={layer.name}, 可視={layer.visible}, サイズ={layer.size}")

                # レイヤー選択ダイアログを表示
                dialog = PSDLayerDialog(psd)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    selected_index = dialog.get_selected_layer_index()
                    if selected_index is not None:
                        logger.info(f"選択されたレイヤー: {selected_index}")
                        try:
                            # 選択されたレイヤーを取得
                            if len(psd) > selected_index:
                                selected_layer = psd[selected_index]
                                logger.info(f"選択されたレイヤーの情報: 名前={selected_layer.name}, 可視={selected_layer.visible}, サイズ={selected_layer.size}")

                                # レイヤーを合成
                                logger.info("レイヤーの合成を開始")
                                composite = selected_layer.composite()
                                logger.info(f"合成完了: サイズ={composite.size}, モード={composite.mode}")
                                return composite
                            else:
                                raise ValueError("選択されたレイヤーが存在しません")
                        except Exception as e:
                            logger.error(f"レイヤーの合成中にエラーが発生: {e}", exc_info=True)
                            return None
                else:
                    logger.info("レイヤー選択がキャンセルされました")
                    return None

        except Exception as e:
            logger.error(f"PSDファイルの読み込みに失敗: {file_path}, エラー: {e}", exc_info=True)
            return None


    @staticmethod
    def _load_pdf(file_path: str) -> Optional[Image.Image]:
        # ... (_load_pdf() メソッド - 変更なし)
        """PDFファイルから画像を抽出"""
        try:
            doc = fitz.open(file_path)
            # 最初のページから画像を抽出
            page = doc[0]
            image_list = page.get_images()

            if not image_list:
                logger.warning(f"PDFファイルに画像が見つかりません: {file_path}")
                return None

            # 最初の画像を取得
            xref = image_list[0][0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]

            # バイトデータからPIL Imageを作成
            return Image.open(io.BytesIO(image_bytes))
        except Exception as e:
            logger.error(f"PDFファイルからの画像抽出に失敗: {file_path}, エラー: {e}", exc_info=True)
            return None

    def process_image(self, img_path: str, target_size: tuple) -> Optional[Image.Image]:
        # ... (process_image() メソッド - 変更なし)
        """画像を処理"""
        try:
            logger.info(f"画像の処理を開始: {img_path}")

            # 画像を読み込む
            img = self.load_image(img_path)
            if img is None:
                logger.error(f"画像の読み込みに失敗: {img_path}")
                return None

            # 画像の情報をログ出力
            logger.info(f"画像情報 - サイズ: {img.size}, モード: {img.mode}")

            # 画像をリサイズ
            img.thumbnail(target_size, Image.Resampling.LANCZOS)
            logger.info(f"画像をリサイズ: {target_size}")

            # CMYKプロファイルが設定されている場合は変換
            if self.cmyk_profile_path:
                logger.info("CMYK変換を実行")
                img = self.convert_to_cmyk(
                    img,
                    self.cmyk_profile_path,
                    self.color_conversion_intent
                )
                logger.info(f"CMYK変換完了 - モード: {img.mode}")

            return img

        except Exception as e:
            logger.error(f"画像処理エラー: {str(e)}", exc_info=True)
            return None


class PDFGenerationThread(QThread):
    # ... (PDFGenerationThread クラス - 変更なし)
    """PDF生成をバックグラウンドで実行するスレッド"""
    finished = pyqtSignal(str, str)  # 成功時に一時ファイルパスとディレクトリを送信
    error = pyqtSignal(str)     # エラー時にメッセージを送信
    progress = pyqtSignal(int)  # 進捗状況を送信

    def __init__(self, image_paths: List[str], settings: GridSettings):
        super().__init__()
        self.image_paths = image_paths
        self.settings = settings
        self.temp_dir = None
        self.image_processor = ImageProcessor()
        # CMYK設定を初期化
        self.cmyk_profile_path = None
        self.color_conversion_intent = 'perceptual'

    def set_cmyk_profile(self, profile_path: str) -> None:
        """CMYKプロファイルを設定"""
        self.cmyk_profile_path = profile_path
        self.image_processor.set_cmyk_profile(profile_path)

    def set_color_conversion_intent(self, intent: str) -> None:
        """色変換方法を設定"""
        self.color_conversion_intent = intent
        self.image_processor.set_color_conversion_intent(intent)

    def run(self) -> None:
        try:
            self.temp_dir = tempfile.mkdtemp()  # 一時ディレクトリを作成
            file_path = os.path.join(self.temp_dir, "output.pdf")
            pdf = canvas.Canvas(file_path, pagesize=self.settings.page_size)

            page_width, page_height = self.settings.page_size

            # 行と列の数を計算
            col_width_pt = self.settings.col_width_mm * MM_TO_PT
            row_height_pt = self.settings.row_height_mm * MM_TO_PT
            cols = max(1, int(page_width / col_width_pt))
            rows = max(1, int(page_height / row_height_pt))

            total_cells = rows * cols
            processed_cells = 0

            # 画像の配置
            for row in range(rows):
                for col in range(cols):
                    cell_index = row * cols + col
                    img_index = cell_index % len(self.image_paths) if self.image_paths else 0

                    if self.image_paths:
                        try:
                            self._process_image(pdf, self.image_paths[img_index],
                                              row, col, col_width_pt, row_height_pt,
                                              page_height, self.temp_dir)
                        except Exception as e:
                            logger.error(f"画像の処理中にエラーが発生しました: {self.image_paths[img_index]}, エラー: {e}")
                            continue

                    processed_cells += 1
                    progress = int((processed_cells / total_cells) * 100)
                    self.progress.emit(progress)

            # グリッド線の描画
            if self.settings.grid_line_visible:
                self._draw_grid_lines(pdf, cols, rows, col_width_pt, row_height_pt,
                                    page_width, page_height)

            pdf.save()
            self.finished.emit(file_path, self.temp_dir)

        except Exception as e:
            logger.error(f"PDF生成中にエラーが発生しました: {e}")
            self.error.emit(str(e))

    def __del__(self):
        """デストラクタ：一時ディレクトリの削除"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                import shutil
                shutil.rmtree(self.temp_dir)
            except Exception as e:
                logger.error(f"一時ディレクトリの削除中にエラーが発生しました: {e}")

    def _process_image(self, pdf: canvas.Canvas, img_path: str, row: int, col: int,
                     col_width_pt: float, row_height_pt: float, page_height: float,
                     temp_dir: str) -> None:
        # ... (_process_image() メソッド - 変更なし)
        """個々の画像を処理してPDFに配置する"""
        try:
            logger.info(f"画像の処理を開始: {img_path}")

            # 画像を読み込む
            img = self.image_processor.process_image(img_path, (int(col_width_pt), int(row_height_pt)))
            if img is None:
                logger.error(f"画像の処理に失敗: {img_path}")
                return
            # CMYK形式に変換 (PDFGenerationThread では CMYK 変換しない)
            # CMYK 変換は ImageProcessor.process_image() で行う
            img_cmyk = img


            # TIFFとして保存（CMYK対応）
            temp_img_path = os.path.join(temp_dir, f"temp_{row}_{col}.tif")
            img_cmyk.save(temp_img_path, format='TIFF', compression='lzw')
            logger.info(f"一時ファイルを保存: {temp_img_path}")

            # 画像を配置
            x = col * col_width_pt
            y = page_height - (row + 1) * row_height_pt

            # 画像のアスペクト比を保持
            img_width, img_height = img_cmyk.size
            aspect_ratio = img_width / img_height

            if aspect_ratio > 1:
                # 横長の画像
                new_width = col_width_pt
                new_height = new_width / aspect_ratio
                y += (row_height_pt - new_height) / 2
            else:
                # 縦長の画像
                new_height = row_height_pt
                new_width = new_height * aspect_ratio
                x += (col_width_pt - new_width) / 2

            # 画像を配置
            pdf.drawImage(temp_img_path, x, y, width=new_width, height=new_height)
            logger.info(f"画像を配置: ({x}, {y}), サイズ: {new_width}x{new_height}")

        except Exception as e:
            logger.error(f"画像の処理中にエラーが発生: {img_path}, エラー: {e}", exc_info=True)
            raise

    def _draw_grid_lines(self, pdf: canvas.Canvas, cols: int, rows: int,
                        col_width_pt: float, row_height_pt: float,
                        page_width: float, page_height: float) -> None:
        # ... (_draw_grid_lines() メソッド - 変更なし)
        """グリッド線を描画する"""
        r, g, b = (self.settings.grid_color.red() / 255.0,
                  self.settings.grid_color.green() / 255.0,
                  self.settings.grid_color.blue() / 255.0)
        pdf.setStrokeColorRGB(r, g, b)
        pdf.setLineWidth(self.settings.grid_width)

        # 垂直線
        for col in range(cols + 1):
            x = col * col_width_pt
            pdf.line(x, 0, x, page_height)

        # 水平線
        for row in range(rows + 1):
            y = page_height - row * row_height_pt
            pdf.line(0, y, page_width, y)


class ImageGridApp(QMainWindow):
    # ... (ImageGridApp クラス - _create_thumbnail() メソッドを修正、processed_images_cache を追加)
    def __init__(self):
        super().__init__()
        self.image_paths: List[str] = []
        self.image_processor = ImageProcessor()  # ImageProcessorのインスタンスを作成
        self.processed_images_cache: Dict[str, Image.Image] = {} # 画像キャッシュを追加 # 追記
        try:
            self.settings = GridSettings.load_from_file()  # 設定を読み込み
        except Exception as e:
            logger.error(f"設定の読み込みに失敗しました: {e}")
            self.settings = GridSettings()  # デフォルト設定を使用
            QMessageBox.warning(
                self,
                "設定読み込みエラー",
                "設定ファイルの読み込みに失敗しました。\nデフォルト設定で起動します。"
            )
        self.preview_labels: List[QLabel] = []
        self.pdf_thread: Optional[PDFGenerationThread] = None
        self.progress_dialog: Optional[QProgressDialog] = None

        # メインウィジェットの作成
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # UIの初期化
        self.initUI()

    def initUI(self) -> None:
        # ... (initUI() メソッド - 変更なし)
        """UIコンポーネントの初期化"""
        # メニューバーの初期化
        self._init_menubar()

        # メインレイアウトを QVBoxLayout に変更（QSplitter を配置するため）
        main_layout = QVBoxLayout(self.central_widget)

        # スプリッターの作成
        main_splitter = QSplitter(Qt.Orientation.Horizontal)

        # 左パネル (設定項目)
        controls_panel = QWidget()
        controls_layout = QVBoxLayout(controls_panel)

        # 画像操作グループ
        image_group = QGroupBox("画像操作")
        image_layout = QVBoxLayout()
        self._init_image_controls(image_layout)
        image_group.setLayout(image_layout)
        controls_layout.addWidget(image_group)

        # グリッド設定グループ
        grid_group = QGroupBox("グリッド設定")
        grid_layout = QVBoxLayout()
        self._init_grid_controls(grid_layout)
        grid_group.setLayout(grid_layout)
        controls_layout.addWidget(grid_group)

        # ページ設定グループ
        page_group = QGroupBox("ページ設定")
        page_layout = QVBoxLayout()
        self._init_page_controls(page_layout)
        page_group.setLayout(page_layout)
        controls_layout.addWidget(page_group)

        # PDF生成グループ
        pdf_group = QGroupBox("PDF出力")
        pdf_layout = QVBoxLayout()
        self._init_pdf_controls(pdf_layout)
        pdf_group.setLayout(pdf_layout)
        controls_layout.addWidget(pdf_group)

        controls_panel.setLayout(controls_layout)

        # 右パネル (プレビュー)
        preview_panel = QWidget()
        preview_layout = QVBoxLayout(preview_panel)
        self._init_preview_area(preview_layout)
        preview_panel.setLayout(preview_layout)

        # スプリッターに左右パネルを追加
        main_splitter.addWidget(controls_panel)
        main_splitter.addWidget(preview_panel)

        # スプリッターの初期サイズを設定（左:右 = 1:2）
        main_splitter.setStretchFactor(0, 1)  # 左パネル
        main_splitter.setStretchFactor(1, 2)  # 右パネル

        # メインレイアウトにスプリッターを追加
        main_layout.addWidget(main_splitter)

        self.setAcceptDrops(True)
        self.setWindowTitle("画像グリッド作成ツール")
        self.resize(1000, 700)  # ウィンドウサイズをさらに大きく

        self.update_preview()

    def _init_menubar(self) -> None:
        # ... (_init_menubar() メソッド - 変更なし)
        """メニューバーの初期化"""
        menubar = self.menuBar()
        settings_menu = menubar.addMenu("設定")

        # 設定リセットメニュー項目
        reset_settings_action = settings_menu.addAction("設定をリセット")
        reset_settings_action.triggered.connect(self.reset_settings)

    def reset_settings(self) -> None:
        # ... (reset_settings() メソッド - 変更なし)
        """設定をリセットする"""
        reply = QMessageBox.question(
            self,
            "設定のリセット",
            "設定を初期状態にリセットします。\nよろしいですか？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            try:
                # 設定ファイルのバックアップを作成
                if os.path.exists(SETTINGS_FILE):
                    backup_path = f"{SETTINGS_FILE}.backup"
                    import shutil
                    shutil.copy2(SETTINGS_FILE, backup_path)

                # 設定ファイルを削除
                if os.path.exists(SETTINGS_FILE):
                    os.remove(SETTINGS_FILE)

                # デフォルト設定を再読み込み
                self.settings = GridSettings()

                # UIを更新
                self.row_height_spinbox.setValue(self.settings.row_height_mm)
                self.col_width_spinbox.setValue(self.settings.col_width_mm)
                self.grid_line_checkbox.setChecked(self.settings.grid_line_visible)
                self.grid_width_spinbox.setValue(self.settings.grid_width)
                self.page_size_combo.setCurrentText("A4" if self.settings.page_size == A4 else "A3")

                # プレビューを更新
                self.update_preview()

                QMessageBox.information(
                    self,
                    "完了",
                    "設定をリセットしました。\nアプリケーションを再起動すると、設定が反映されます。"
                )
            except Exception as e:
                logger.error(f"設定のリセットに失敗しました: {e}")
                QMessageBox.critical(
                    self,
                    "エラー",
                    f"設定のリセットに失敗しました: {e}"
                )

    def _init_image_controls(self, layout: QVBoxLayout) -> None:
        # ... (_init_image_controls() メソッド - 変更なし)
        """画像関連のコントロールを初期化"""
        self.btn_add_images = QPushButton('画像を追加')
        self.btn_add_images.clicked.connect(self.load_images)
        layout.addWidget(self.btn_add_images)

    def _init_grid_controls(self, layout: QVBoxLayout) -> None:
        # ... (_init_grid_controls() メソッド - 変更なし)
        """グリッド設定関連のコントロールを初期化"""
        # 行の高さ設定
        self.row_height_spinbox = QDoubleSpinBox()
        self.row_height_spinbox.setRange(10.0, 297.0)
        self.row_height_spinbox.setValue(self.settings.row_height_mm)
        self.row_height_spinbox.setSuffix(" mm")
        self.row_height_spinbox.valueChanged.connect(self.update_grid)
        layout.addWidget(QLabel("行の高さ:"))
        layout.addWidget(self.row_height_spinbox)

        # 列の幅設定
        self.col_width_spinbox = QDoubleSpinBox()
        self.col_width_spinbox.setRange(10.0, 210.0)
        self.col_width_spinbox.setValue(self.settings.col_width_mm)
        self.col_width_spinbox.setSuffix(" mm")
        self.col_width_spinbox.valueChanged.connect(self.update_grid)
        layout.addWidget(QLabel("列の幅:"))
        layout.addWidget(self.col_width_spinbox)

        # グリッド線の設定
        self._init_grid_line_controls(layout)

    def _init_page_controls(self, layout: QVBoxLayout) -> None:
        # ... (_init_page_controls() メソッド - 変更なし)
        """ページ設定関連のコントロールを初期化"""
        # ページサイズ選択
        self.page_size_combo = QComboBox()
        self.page_size_combo.addItems(["A4", "A3"])
        self.page_size_combo.setCurrentText("A4" if self.settings.page_size == A4 else "A3")
        self.page_size_combo.currentTextChanged.connect(self.update_page_size)
        layout.addWidget(QLabel("用紙サイズ:"))
        layout.addWidget(self.page_size_combo)

    def _init_pdf_controls(self, layout: QVBoxLayout) -> None:
        # ... (_init_pdf_controls() メソッド - 変更なし)
        """PDF出力関連のコントロールを初期化"""
        # PDF生成ボタン
        self.btn_generate_pdf = QPushButton('PDFを作成')
        self.btn_generate_pdf.clicked.connect(self.generate_pdf)
        layout.addWidget(self.btn_generate_pdf)

    def _init_grid_line_controls(self, layout: QVBoxLayout) -> None:
        # ... (_init_grid_line_controls() メソッド - 変更なし)
        """グリッド線関連のコントロールを初期化"""
        self.grid_line_checkbox = QCheckBox("グリッド線を表示")
        self.grid_line_checkbox.setChecked(self.settings.grid_line_visible)
        self.grid_line_checkbox.stateChanged.connect(self.update_grid)
        layout.addWidget(self.grid_line_checkbox)

        self.grid_color_btn = QPushButton("グリッド線の色")
        self.grid_color_btn.clicked.connect(self.select_grid_color)
        layout.addWidget(self.grid_color_btn)

        self.grid_width_spinbox = QSpinBox()
        self.grid_width_spinbox.setRange(1, 5)
        self.grid_width_spinbox.setValue(self.settings.grid_width)
        self.grid_width_spinbox.valueChanged.connect(self.update_grid)
        layout.addWidget(QLabel("線の太さ:"))
        layout.addWidget(self.grid_width_spinbox)

    def _init_preview_area(self, layout: QVBoxLayout) -> None:
        # ... (_init_preview_area() メソッド - 変更なし)
        """プレビューエリアを初期化"""
        # プレビューリフレッシュボタン
        self.refresh_preview_button = QPushButton("プレビューを更新")
        self.refresh_preview_button.clicked.connect(self.update_preview)
        self.refresh_preview_button.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #cccccc;
                border-radius: 3px;
                padding: 5px 10px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        layout.addWidget(self.refresh_preview_button)

        # プレビューエリア
        self.preview_area_scroll = QScrollArea()
        self.preview_area_grid = QGridLayout()
        self.preview_area_grid.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_area_widget = QWidget()
        self.preview_area_widget.setLayout(self.preview_area_grid)
        self.preview_area_scroll.setWidget(self.preview_area_widget)
        self.preview_area_scroll.setWidgetResizable(True)
        layout.addWidget(self.preview_area_scroll)

    def load_images(self):
        # ... (load_images() メソッド - processed_images_cache にキャッシュ)
        """画像ファイルを選択して読み込む"""
        # サポートされているファイル形式のフィルターを作成
        file_filter = ";;".join([f"{name} ({pattern})" for name, pattern in SUPPORTED_IMAGE_FORMATS.items()])

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "画像を選択",
            "",
            file_filter
        )

        if files:
            # 各ファイルを処理
            for file_path in files:
                try:
                    # 画像を読み込んで検証
                    img = self.image_processor.process_image(
                        file_path,
                        (int(self.settings.col_width_mm), int(self.settings.row_height_mm))
                    )
                    if img is not None:
                        self.image_paths.append(file_path)
                        self.processed_images_cache[file_path] = img # キャッシュ # 追記
                    else:
                        QMessageBox.warning(
                            self,
                            "警告",
                            f"画像の読み込みに失敗しました: {file_path}"
                        )
                except Exception as e:
                    logger.error(f"画像の読み込み中にエラーが発生しました: {file_path}, エラー: {e}")
                    QMessageBox.warning(
                        self,
                        "エラー",
                        f"画像の読み込み中にエラーが発生しました: {file_path}\n{str(e)}"
                    )

            # プレビューを一度だけ更新
            self.update_preview()

    def update_grid(self):
        # ... (update_grid() メソッド - 変更なし)
        """グリッド設定を更新"""
        self.settings.row_height_mm = self.row_height_spinbox.value()
        self.settings.col_width_mm = self.col_width_spinbox.value()
        self.settings.grid_line_visible = self.grid_line_checkbox.isChecked()
        self.settings.grid_width = self.grid_width_spinbox.value()
        self.update_preview()

    def update_page_size(self, size_text):
        # ... (update_page_size() メソッド - 変更なし)
        """ページサイズを更新"""
        if size_text == "A4":
            self.settings.page_size = A4
            self.row_height_spinbox.setRange(10.0, 297.0)  # A4の高さ制限
            self.col_width_spinbox.setRange(10.0, 210.0)   # A4の幅制限
        elif size_text == "A3":
            self.settings.page_size = A3
            self.row_height_spinbox.setRange(10.0, 420.0)  # A3の高さ制限
            self.col_width_spinbox.setRange(10.0, 297.0)   # A3の幅制限
        self.update_preview()

    @lru_cache(maxsize=100)
    def _create_thumbnail(self, img_path: str) -> QPixmap: # 引数を img_path から img に変更 # 変更
        """画像のサムネイルを生成（キャッシュ付き）"""
        try:
            # キャッシュから PIL Image を取得 # 変更
            img = self.processed_images_cache.get(img_path)
            if img is None: # 念のため None チェック
                logger.error(f"キャッシュミス: {img_path}")
                return QPixmap()

            # PIL ImageをQPixmapに変換 (RGB変換は ImageProcessor.process_image() で行うように変更)
            data = img.convert('RGB').tobytes("raw", "RGB") # RGB 変換を明示的に
            qim = QImage(data, img.size[0], img.size[1], QImage.Format.Format_RGB888)
            pixmap = QPixmap.fromImage(qim)

            # サムネイルサイズにリサイズ
            thumbnail = pixmap.scaled(THUMBNAIL_SIZE[0], THUMBNAIL_SIZE[1],
                                    Qt.AspectRatioMode.KeepAspectRatio,
                                    Qt.TransformationMode.SmoothTransformation)
            return thumbnail
        except Exception as e:
            logger.error(f"サムネイルの生成に失敗しました: {img_path}, エラー: {e}")
            return QPixmap()


    def update_preview(self):
        # ... (update_preview() メソッド - _create_thumbnail() 呼び出しを修正)
        """プレビューを更新"""
        try:
            # 既存のプレビューをクリア
            for label in self.preview_labels:
                label.clear()
                label.setParent(None)
            self.preview_labels = []

            # グリッドコンテナのクリア
            while self.preview_area_grid.count():
                item = self.preview_area_grid.takeAt(0)
                if item.widget():
                    item.widget().deleteLater()

            if not self.image_paths:
                # 画像がない場合は初期メッセージを表示
                message_label = QLabel("画像をドラッグ＆ドロップするか、\n「画像を追加」ボタンで画像を選択してください。")
                message_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                message_label.setStyleSheet("""
                    QLabel {
                        color: #666666;
                        font-size: 14px;
                        padding: 20px;
                        background-color: #f5f5f5;
                        border: 2px dashed #cccccc;
                        border-radius: 5px;
                    }
                """)
                self.preview_area_grid.addWidget(message_label)
                return

            # プレビューのサイズを計算（A4/A3の比率を保持）
            preview_height = DEFAULT_PREVIEW_HEIGHT
            preview_width = int(preview_height * (self.settings.page_size[0] / self.settings.page_size[1]))

            # プレビュー用のフレームを作成（用紙を模したフレーム）
            self.preview_frame = QFrame()
            self.preview_frame.setFixedSize(preview_width, preview_height)
            self.preview_frame.setFrameShape(QFrame.Shape.Box)
            self.preview_frame.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border: 1px solid #cccccc;
                    border-radius: 2px;
                }
            """)

            # 行と列の数を計算
            col_width_pt = self.settings.col_width_mm * MM_TO_PT
            row_height_pt = self.settings.row_height_mm * MM_TO_PT
            cols = max(1, int(self.settings.page_size[0] / col_width_pt))
            rows = max(1, int(self.settings.page_size[1] / row_height_pt))

            # プレビューでのセルサイズを計算
            cell_width = preview_width / (self.settings.page_size[0] / col_width_pt)
            cell_height = preview_height / (self.settings.page_size[1] / row_height_pt)

            # 画像を描画するためのpaintEventを設定
            def paint_preview(event):
                painter = QPainter(self.preview_frame)
                painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)

                try:
                    # 画像の描画
                    for row in range(rows):
                        for col in range(cols):
                            cell_index = row * cols + col
                            img_index = cell_index % len(self.image_paths) if self.image_paths else 0

                            if self.image_paths:
                                img_path = self.image_paths[img_index]
                                # キャッシュから Image.Image オブジェクトを取得 # 変更
                                img = self.processed_images_cache.get(img_path)
                                if img is None: # 念のため None チェック
                                    logger.error(f"キャッシュミス (paintEvent): {img_path}")
                                    continue # キャッシュミス時はスキップ

                                thumbnail = self._create_thumbnail(img_path) # _create_thumbnail に img ではなく img_path を渡す (キャッシュキーとして使用)

                                # セルのサイズとアスペクト比を計算
                                cell_rect_width = cell_width
                                cell_rect_height = cell_height
                                cell_aspect = cell_rect_width / cell_rect_height

                                # 画像のアスペクト比を計算
                                img_aspect = thumbnail.width() / thumbnail.height()

                                # アスペクト比に基づいてサイズを調整
                                if img_aspect > cell_aspect:
                                    new_width = cell_rect_width
                                    new_height = cell_rect_width / img_aspect
                                else:
                                    new_height = cell_rect_height
                                    new_width = cell_rect_height * img_aspect

                                # セル内での位置を計算（センタリング）
                                x = col * cell_width + (cell_width - new_width) / 2
                                y = row * cell_height + (cell_height - new_height) / 2

                                # 画像を描画
                                target_rect = QRectF(x, y, new_width, new_height)
                                source_rect = QRectF(thumbnail.rect())
                                painter.drawPixmap(target_rect, thumbnail, source_rect)

                    # グリッド線の描画
                    if self.settings.grid_line_visible:
                        pen = QPen(self.settings.grid_color)
                        pen.setWidth(self.settings.grid_width)
                        painter.setPen(pen)

                        # 垂直線
                        for col in range(cols + 1):
                            x = col * cell_width
                            painter.drawLine(int(x), 0, int(x), preview_height)

                        # 水平線
                        for row in range(rows + 1):
                            y = row * cell_height
                            painter.drawLine(0, int(y), preview_width, int(y))
                finally:
                    painter.end()

            # paintEventを設定
            self.preview_frame.paintEvent = paint_preview

            # プレビューフレームをレイアウトに追加
            self.preview_area_grid.addWidget(self.preview_frame)
            self.preview_frame.update()

        except Exception as e:
            logger.error(f"プレビューの更新中にエラーが発生しました: {e}", exc_info=True)
            QMessageBox.warning(
                self,
                "エラー",
                f"プレビューの更新中にエラーが発生しました: {str(e)}"
            )

    def select_grid_color(self):
        # ... (select_grid_color() メソッド - 変更なし)
        """グリッド線の色を選択するダイアログを表示"""
        color = QColorDialog.getColor(self.settings.grid_color, self, "グリッド線の色を選択")
        if color.isValid():
            self.settings.grid_color = color
            self.update_preview()

    def generate_pdf(self):
        # ... (generate_pdf() メソッド - 変更なし)
        """PDFを生成"""
        if not self.image_paths:
            QMessageBox.warning(self, "警告", "画像が追加されていません。")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "PDFを保存", "", "PDF Files (*.pdf)"
        )
        if not file_path:
            return  # ユーザーがキャンセルした場合
        if not file_path.lower().endswith('.pdf'):
            file_path += '.pdf'

        try:
            logger.info(f"PDF生成を開始: {file_path}")

            # 進捗ダイアログの作成
            self.progress_dialog = QProgressDialog("PDFを生成中...", "キャンセル", 0, 100, self)
            self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
            self.progress_dialog.setAutoClose(True)
            self.progress_dialog.setAutoReset(True)

            # PDF生成スレッドの作成と開始
            self.pdf_thread = PDFGenerationThread(
                self.image_paths,
                self.settings
            )

            # CMYK設定を渡す
            if self.image_processor.cmyk_profile_path:
                logger.info(f"CMYKプロファイルを設定: {self.image_processor.cmyk_profile_path}")
                self.pdf_thread.set_cmyk_profile(self.image_processor.cmyk_profile_path)

            logger.info(f"色変換方法を設定: {self.image_processor.color_conversion_intent}")
            self.pdf_thread.set_color_conversion_intent(self.image_processor.color_conversion_intent)

            # シグナルの接続
            self.pdf_thread.finished.connect(
                lambda temp_path, temp_dir: self.on_pdf_generation_finished(temp_path, temp_dir, file_path)
            )
            self.pdf_thread.error.connect(self.on_pdf_generation_error)
            self.pdf_thread.progress.connect(self.progress_dialog.setValue)
            self.progress_dialog.canceled.connect(self.pdf_thread.terminate)

            # スレッドの開始
            self.pdf_thread.start()
            self.progress_dialog.show()

            logger.info("PDF生成スレッドを開始")

        except Exception as e:
            logger.error(f"PDF生成の初期化中にエラーが発生: {e}", exc_info=True)
            QMessageBox.critical(
                self,
                "エラー",
                f"PDF生成の初期化中にエラーが発生しました:\n{str(e)}"
            )

    def on_pdf_generation_finished(self, temp_path: str, temp_dir: str, final_path: str):
        # ... (on_pdf_generation_finished() メソッド - 変更なし)
        """PDF生成完了時の処理"""
        try:
            # 一時ファイルを最終保存先にコピー
            import shutil
            shutil.copy2(temp_path, final_path)
            QMessageBox.information(self, "完了", f"PDFを作成しました: {final_path}")
        except Exception as e:
            logger.error(f"PDFの保存中にエラーが発生しました: {e}")
            QMessageBox.critical(self, "エラー", f"PDFの保存中にエラーが発生しました: {e}")
        finally:
            # 一時ディレクトリの削除
            try:
                shutil.rmtree(temp_dir)
            except Exception as e:
                logger.error(f"一時ディレクトリの削除中にエラーが発生しました: {e}")

    def on_pdf_generation_error(self, error_message: str):
        # ... (on_pdf_generation_error() メソッド - 変更なし)
        """PDF生成エラー時の処理"""
        QMessageBox.critical(self, "エラー", f"PDFの生成中にエラーが発生しました: {error_message}")

    def dragEnterEvent(self, event: QDragEnterEvent):
        # ... (dragEnterEvent() メソッド - 変更なし)
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        # ... (dropEvent() メソッド - processed_images_cache にキャッシュ)
        """ドラッグ＆ドロップ時のイベント処理"""
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith((".png", ".jpg", ".jpeg", ".psd", ".pdf")):
                try:
                    # 画像を読み込んで検証
                    img = self.image_processor.process_image(
                        file_path,
                        (int(self.settings.col_width_mm), int(self.settings.row_height_mm))
                    )
                    if img is not None:
                        self.image_paths.append(file_path)
                        self.processed_images_cache[file_path] = img # キャッシュ # 追記
                    else:
                        QMessageBox.warning(
                            self,
                            "警告",
                            f"画像の読み込みに失敗しました: {file_path}"
                        )
                except Exception as e:
                    logger.error(f"画像の読み込み中にエラーが発生しました: {file_path}, エラー: {e}")
                    QMessageBox.warning(
                        self,
                        "エラー",
                        f"画像の読み込みに失敗しました: {file_path}\n{str(e)}"
                    )
        self.update_preview()

    def closeEvent(self, event: Any) -> None:
        # ... (closeEvent() メソッド - 変更なし)
        """アプリケーション終了時の処理"""
        try:
            self.settings.save_to_file()  # 設定を保存
            self._create_thumbnail.cache_clear()  # サムネイルキャッシュをクリア
        except Exception as e:
            logger.error(f"設定の保存中にエラーが発生しました: {e}")
            QMessageBox.warning(self, "警告", "設定の保存に失敗しました。")
        super().closeEvent(event)

    def _create_settings_group(self):
        # ... (_create_settings_group() メソッド - 変更なし)
        """設定グループの作成"""
        settings_group = QGroupBox("設定")
        settings_layout = QVBoxLayout()

        # グリッド設定
        grid_layout = QHBoxLayout()
        grid_layout.addWidget(QLabel("行の高さ(mm):"))
        self.row_height_spinbox = QDoubleSpinBox()
        self.row_height_spinbox.setRange(10, 1000)
        self.row_height_spinbox.setValue(150)
        self.row_height_spinbox.setSingleStep(1)
        grid_layout.addWidget(self.row_height_spinbox)

        grid_layout.addWidget(QLabel("列の幅(mm):"))
        self.col_width_spinbox = QDoubleSpinBox()
        self.col_width_spinbox.setRange(10, 1000)
        self.col_width_spinbox.setValue(150)
        self.col_width_spinbox.setSingleStep(1)
        grid_layout.addWidget(self.col_width_spinbox)

        settings_layout.addLayout(grid_layout)

        # ページ設定
        page_layout = QHBoxLayout()
        page_layout.addWidget(QLabel("ページサイズ:"))
        self.page_size_combo = QComboBox()
        self.page_size_combo.addItems(["A4", "A3"])
        self.page_size_combo.currentTextChanged.connect(self.update_page_size)
        page_layout.addWidget(self.page_size_combo)

        settings_layout.addLayout(page_layout)

        # CMYK設定
        cmyk_layout = QHBoxLayout()
        cmyk_layout.addWidget(QLabel("CMYKプロファイル:"))
        self.cmyk_profile_path = QLineEdit()
        self.cmyk_profile_path.setReadOnly(True)
        cmyk_layout.addWidget(self.cmyk_profile_path)

        browse_btn = QPushButton("参照")
        browse_btn.clicked.connect(self._browse_cmyk_profile)
        cmyk_layout.addWidget(browse_btn)

        settings_layout.addLayout(cmyk_layout)

        # 色変換方法
        color_intent_layout = QHBoxLayout()
        color_intent_layout.addWidget(QLabel("色変換方法:"))
        self.color_intent_combo = QComboBox()
        self.color_intent_combo.addItems(["知覚的", "相対的", "彩度優先", "絶対的"])
        self.color_intent_combo.currentIndexChanged.connect(self._update_color_intent)
        color_intent_layout.addWidget(self.color_intent_combo)

        settings_layout.addLayout(color_intent_layout)

        settings_group.setLayout(settings_layout)
        return settings_group

    def _browse_cmyk_profile(self):
        # ... (_browse_cmyk_profile() メソッド - 変更なし)
        """CMYKプロファイルを選択"""
        file_path, _ = QFileDialog.getOpenFileNames(
            self,
            "CMYKプロファイルを選択",
            "",
            "ICCプロファイル (*.icc *.icm)"
        )
        if file_path:
            self.cmyk_profile_path.setText(file_path)
            self.image_processor.set_cmyk_profile(file_path)

    def _update_color_intent(self, index: int):
        # ... (_update_color_intent() メソッド - 変更なし)
        """色変換方法を更新"""
        intents = {
            0: 'perceptual',    # 知覚的
            1: 'relative',      # 相対的
            2: 'saturation',    # 彩度優先
            3: 'absolute'       # 絶対的
        }
        self.image_processor.set_color_conversion_intent(intents[index])


class PSDLayerDialog(QDialog):
    # ... (PSDLayerDialog クラス - 変更なし)
    """PSDレイヤー選択ダイアログ"""
    def __init__(self, psd: psd_tools.PSDImage, parent=None):
        super().__init__(parent)
        self.psd = psd
        self.selected_layer_index = None
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.initUI()

    def initUI(self):
        self.setWindowTitle("PSDレイヤー選択")
        self.setModal(True)
        layout = QVBoxLayout(self)

        # レイヤーリスト
        self.layer_list = QListWidget()
        self.layer_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)

        # レイヤー情報を追加
        for i, layer in enumerate(self.psd):
            layer_name = layer.name if layer.name else f"レイヤー {i+1}"
            visible_icon = "👁" if layer.visible else "👁‍🗨"
            size_info = f" ({layer.size[0]}x{layer.size[1]})"
            item = QListWidgetItem(f"{visible_icon} {layer_name}{size_info}")
            item.setData(Qt.ItemDataRole.UserRole, i)  # レイヤーインデックスを保存
            self.layer_list.addItem(item)
            logger.info(f"レイヤーリストに追加: {layer_name}{size_info}")

        layout.addWidget(QLabel("レイヤーを選択してください:"))
        layout.addWidget(self.layer_list)

        # ボタン
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        # デフォルトで最初のレイヤーを選択
        if self.layer_list.count() > 0:
            self.layer_list.setCurrentRow(0)
            logger.info("最初のレイヤーを選択")

        # ダイアログのサイズを設定
        self.resize(400, 500)
        logger.info("PSDレイヤー選択ダイアログを初期化完了")

    def get_selected_layer_index(self) -> Optional[int]:
        """選択されたレイヤーのインデックスを返す"""
        current_item = self.layer_list.currentItem()
        if current_item:
            index = current_item.data(Qt.ItemDataRole.UserRole)
            logger.info(f"選択されたレイヤーインデックス: {index}")
            return index
        logger.warning("レイヤーが選択されていません")
        return None

    def closeEvent(self, event):
        """ダイアログが閉じられる時の処理"""
        logger.info("レイヤー選択ダイアログを閉じます")
        self.reject()  # キャンセルとして扱う
        super().closeEvent(event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ImageGridApp()
    ex.show()
    sys.exit(app.exec())