import sys
import os
import re
import base64
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QLineEdit, 
                             QFileDialog, QProgressBar, QTextEdit, QGroupBox,
                             QMessageBox, QCheckBox, QListWidget, QListWidgetItem,
                             QSplitter, QFrame, QScrollArea, QSpinBox, QComboBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl, QTimer, QMarginsF
from PyQt5.QtWebEngineWidgets import QWebEngineView, QWebEnginePage, QWebEngineSettings
from PyQt5.QtGui import QPageLayout, QPageSize, QDragEnterEvent, QDropEvent, QFont, QIcon, QPalette, QColor
import markdown
from markdown.extensions import fenced_code, tables, toc, nl2br, sane_lists
import tempfile
import json


class FileItem(QFrame):
    """ファイルリストの個別アイテム"""
    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.selected = False
        self.parent_converter = parent
        self.setup_ui()
        
    def setup_ui(self):
        self.setFixedHeight(30)
        self.setFrameStyle(QFrame.StyledPanel)
        self.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                margin: 1px;
            }
            QFrame:hover {
                background-color: #f0f8ff;
                border-color: #3498db;
            }
        """)
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(8)
        
        # ファイル名ラベル
        self.file_label = QLabel(Path(self.file_path).name)
        self.file_label.setToolTip(self.file_path)
        self.file_label.setStyleSheet("QLabel { color: #2c3e50; font-size: 12px; border: none; }")
        layout.addWidget(self.file_label, 1)
        
        # 削除ボタン
        self.delete_btn = QPushButton("×")
        self.delete_btn.setFixedSize(20, 20)
        self.delete_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)
        self.delete_btn.clicked.connect(self.delete_file)
        layout.addWidget(self.delete_btn, 0)
        
    def delete_file(self):
        if self.parent_converter:
            self.parent_converter.remove_single_file(self.file_path)
            
    def mousePressEvent(self, event):
        if self.parent_converter:
            self.parent_converter.select_file_item(self)
        super().mousePressEvent(event)
        
    def set_selected(self, selected):
        self.selected = selected
        if selected:
            self.setStyleSheet("""
                QFrame {
                    background-color: #3498db;
                    border: 1px solid #2980b9;
                    border-radius: 4px;
                    margin: 1px;
                }
            """)
            self.file_label.setStyleSheet("QLabel { color: white; font-size: 12px; border: none; }")
        else:
            self.setStyleSheet("""
                QFrame {
                    background-color: white;
                    border: 1px solid #e0e0e0;
                    border-radius: 4px;
                    margin: 1px;
                }
                QFrame:hover {
                    background-color: #f0f8ff;
                    border-color: #3498db;
                }
            """)
            self.file_label.setStyleSheet("QLabel { color: #2c3e50; font-size: 12px; border: none; }")


class MarkdownToPdfConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_files = []  # 複数ファイル対応
        self.current_file_index = 0
        self.page_break_marker = "<!-- pagebreak -->"
        self.init_ui()
        self.setup_webengine()
        self.apply_professional_styling()
        
    def setup_webengine(self):
        """WebEngineの設定を最適化"""
        settings = QWebEngineSettings.defaultSettings()
        settings.setAttribute(QWebEngineSettings.JavascriptEnabled, True)
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls, True)
        settings.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(QWebEngineSettings.PluginsEnabled, True)
        # エラーログを抑制
        settings.setAttribute(QWebEngineSettings.ErrorPageEnabled, False)
        settings.setAttribute(QWebEngineSettings.JavascriptCanOpenWindows, False)
        settings.setAttribute(QWebEngineSettings.JavascriptCanAccessClipboard, False)
        # カラー印刷のための設定
        settings.setAttribute(QWebEngineSettings.PrintElementBackgrounds, True)
        # 追加のカラー印刷設定
        settings.setAttribute(QWebEngineSettings.AllowRunningInsecureContent, True)
        
    def init_ui(self):
        self.setWindowTitle("Markdown to PDF Converter")
        self.setGeometry(100, 100, 1200, 800)
        
        # アイコンを設定
        icon_path = Path(__file__).parent / "img" / "icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        # メインウィジェット
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # メインレイアウト（左右分割）
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)
        
        # スプリッター
        splitter = QSplitter(Qt.Horizontal)
        main_layout.addWidget(splitter)
        
        # 左側パネル（設定エリア）
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)
        
        # 右側パネル（プレビューエリア）
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)
        
        # スプリッターの初期サイズ設定（左:右 = 2:3）
        splitter.setSizes([400, 600])
        
        # ドラッグアンドドロップを有効化
        self.setAcceptDrops(True)
        
    def create_left_panel(self):
        """左側の設定パネルを作成"""
        left_widget = QWidget()
        left_widget.setMaximumWidth(450)
        left_widget.setMinimumWidth(350)
        layout = QVBoxLayout(left_widget)
        layout.setSpacing(15)
        
        # タイトル
        title_label = QLabel("Markdown to PDF Converter")
        title_label.setObjectName("titleLabel")
        layout.addWidget(title_label)
        
        # 入力ファイル選択部
        input_group = self.create_input_group()
        layout.addWidget(input_group)
        
        # 出力設定部
        output_group = self.create_output_group()
        layout.addWidget(output_group)
        
        # 変換オプション
        options_group = self.create_options_group()
        layout.addWidget(options_group)
        
        # 変換ボタン
        self.convert_btn = QPushButton("PDFに変換")
        self.convert_btn.setObjectName("convertButton")
        self.convert_btn.clicked.connect(self.convert_to_pdf)
        self.convert_btn.setEnabled(False)
        self.convert_btn.setMinimumHeight(45)
        layout.addWidget(self.convert_btn)
        
        # プログレスバー
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setMinimumHeight(25)
        layout.addWidget(self.progress_bar)
        
        # ステータスラベル
        self.status_label = QLabel("ファイルを選択してください")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        
        # スペーサー
        layout.addStretch()
        
        return left_widget
        
    def create_input_group(self):
        """入力ファイル選択グループを作成"""
        input_group = QGroupBox("入力ファイル")
        input_group.setObjectName("inputGroup")
        layout = QVBoxLayout()
        
        # ファイル選択ボタン
        file_select_layout = QHBoxLayout()
        self.browse_input_btn = QPushButton("📁 ファイルを選択")
        self.browse_input_btn.setObjectName("browseButton")
        self.browse_input_btn.clicked.connect(self.browse_input_files)
        self.browse_input_btn.setMinimumHeight(35)
        
        self.clear_files_btn = QPushButton("🗑 クリア")
        self.clear_files_btn.setObjectName("clearButton")
        self.clear_files_btn.clicked.connect(self.clear_files)
        self.clear_files_btn.setMinimumHeight(35)
        self.clear_files_btn.setEnabled(False)
        
        file_select_layout.addWidget(self.browse_input_btn)
        file_select_layout.addWidget(self.clear_files_btn)
        layout.addLayout(file_select_layout)
        
        # ドラッグ&ドロップエリア
        self.drop_area = QLabel("📄 ここにMarkdownファイルを\nドラッグ&ドロップ")
        self.drop_area.setObjectName("dropArea")
        self.drop_area.setAlignment(Qt.AlignCenter)
        self.drop_area.setMinimumHeight(80)
        self.drop_area.setStyleSheet("""
            QLabel#dropArea {
                border: 2px dashed #cccccc;
                border-radius: 8px;
                background-color: #f8f9fa;
                color: #6c757d;
                font-size: 14px;
            }
            QLabel#dropArea:hover {
                border-color: #007bff;
                background-color: #e3f2fd;
            }
        """)
        layout.addWidget(self.drop_area)
        
        # ファイルリスト（スクロールエリア）
        self.file_scroll = QScrollArea()
        self.file_scroll.setMaximumHeight(150)
        self.file_scroll.setMinimumHeight(100)
        self.file_scroll.setWidgetResizable(True)
        self.file_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.file_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # ファイルリストのコンテナ
        self.file_container = QWidget()
        self.file_layout = QVBoxLayout(self.file_container)
        self.file_layout.setContentsMargins(0, 0, 0, 0)
        self.file_layout.setSpacing(2)
        self.file_layout.addStretch()  # 下部にスペーサーを追加
        
        self.file_scroll.setWidget(self.file_container)
        self.file_scroll.setObjectName("fileScrollArea")
        layout.addWidget(self.file_scroll)
        
        # ファイルアイテムのリスト
        self.file_items = []
        
        input_group.setLayout(layout)
        return input_group
        
    def create_output_group(self):
        """出力設定グループを作成"""
        output_group = QGroupBox("出力設定")
        output_group.setObjectName("outputGroup")
        layout = QVBoxLayout()
        
        # 出力先フォルダ
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("出力フォルダ:"))
        self.output_folder = QLineEdit()
        self.output_folder.setPlaceholderText("出力先フォルダを選択...")
        self.browse_folder_btn = QPushButton("📁")
        self.browse_folder_btn.setMaximumWidth(40)
        self.browse_folder_btn.clicked.connect(self.browse_output_folder)
        folder_layout.addWidget(self.output_folder)
        folder_layout.addWidget(self.browse_folder_btn)
        layout.addLayout(folder_layout)
        
        # 改ページマーカー設定
        pagebreak_layout = QHBoxLayout()
        pagebreak_layout.addWidget(QLabel("改ページマーカー:"))
        self.pagebreak_input = QLineEdit(self.page_break_marker)
        self.pagebreak_input.setMaximumWidth(200)
        pagebreak_layout.addWidget(self.pagebreak_input)
        pagebreak_layout.addStretch()
        layout.addLayout(pagebreak_layout)
        
        output_group.setLayout(layout)
        return output_group
        
    def create_options_group(self):
        """変換オプショングループを作成"""
        options_group = QGroupBox("変換オプション")
        options_group.setObjectName("optionsGroup")
        layout = QVBoxLayout()
        
        # 目次オプション
        self.include_toc = QCheckBox("目次を含める")
        self.include_toc.setChecked(True)
        layout.addWidget(self.include_toc)
        
        # 上書き確認オプション
        self.confirm_overwrite = QCheckBox("上書き時に確認する")
        self.confirm_overwrite.setChecked(True)
        layout.addWidget(self.confirm_overwrite)
        
        options_group.setLayout(layout)
        return options_group
        
    def create_right_panel(self):
        """右側のプレビューパネルを作成"""
        right_widget = QWidget()
        layout = QVBoxLayout(right_widget)
        layout.setSpacing(5)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # プレビューヘッダー（高さを制限）
        header_widget = QWidget()
        header_widget.setMaximumHeight(40)
        header_widget.setMinimumHeight(40)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)
        
        preview_label = QLabel("プレビュー")
        preview_label.setObjectName("previewLabel")
        header_layout.addWidget(preview_label)
        
        # ページ制御
        self.page_controls = QWidget()
        self.page_controls.setMaximumHeight(35)
        page_layout = QHBoxLayout(self.page_controls)
        page_layout.setContentsMargins(0, 0, 0, 0)
        
        self.prev_page_btn = QPushButton("◀")
        self.prev_page_btn.setMaximumWidth(30)
        self.prev_page_btn.setMaximumHeight(30)
        self.prev_page_btn.setEnabled(False)
        self.prev_page_btn.clicked.connect(self.prev_page)
        
        self.page_info = QLabel("- / -")
        self.page_info.setAlignment(Qt.AlignCenter)
        self.page_info.setMinimumWidth(60)
        self.page_info.setMaximumHeight(30)
        
        self.next_page_btn = QPushButton("▶")
        self.next_page_btn.setMaximumWidth(30)
        self.next_page_btn.setMaximumHeight(30)
        self.next_page_btn.setEnabled(False)
        self.next_page_btn.clicked.connect(self.next_page)
        
        # ズームコントロール
        self.zoom_out_btn = QPushButton("🔍-")
        self.zoom_out_btn.setMaximumWidth(30)
        self.zoom_out_btn.setMaximumHeight(30)
        self.zoom_out_btn.clicked.connect(self.zoom_out)
        
        self.zoom_info = QLabel("100%")
        self.zoom_info.setAlignment(Qt.AlignCenter)
        self.zoom_info.setMinimumWidth(50)
        self.zoom_info.setMaximumHeight(30)
        
        self.zoom_in_btn = QPushButton("🔍+")
        self.zoom_in_btn.setMaximumWidth(30)
        self.zoom_in_btn.setMaximumHeight(30)
        self.zoom_in_btn.clicked.connect(self.zoom_in)
        
        self.current_zoom = 1.0
        
        page_layout.addWidget(self.prev_page_btn)
        page_layout.addWidget(self.page_info)
        page_layout.addWidget(self.next_page_btn)
        page_layout.addWidget(QLabel("|"))
        page_layout.addWidget(self.zoom_out_btn)
        page_layout.addWidget(self.zoom_info)
        page_layout.addWidget(self.zoom_in_btn)
        
        header_layout.addStretch()
        header_layout.addWidget(self.page_controls)
        
        layout.addWidget(header_widget)
        
        # プレビューエリア（残りの全スペースを使用）
        self.web_view = QWebEngineView()
        self.web_view.setObjectName("webView")
        self.web_view.setMinimumHeight(300)
        # 横長表示を防ぐために最大幅を設定
        self.web_view.setSizePolicy(self.web_view.sizePolicy().horizontalPolicy(), self.web_view.sizePolicy().verticalPolicy())
        layout.addWidget(self.web_view, 1)  # ストレッチファクター1を設定
        
        return right_widget
        
    def apply_professional_styling(self):
        """プロフェッショナルなスタイリングを適用"""
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            
            QLabel#titleLabel {
                font-size: 18px;
                font-weight: bold;
                color: #2c3e50;
                padding: 10px 0;
                border-bottom: 2px solid #3498db;
                margin-bottom: 10px;
            }
            
            QGroupBox {
                font-weight: bold;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: white;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                color: #2c3e50;
                background-color: white;
            }
            
            QPushButton#convertButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: bold;
                padding: 12px;
            }
            
            QPushButton#convertButton:hover {
                background-color: #2980b9;
            }
            
            QPushButton#convertButton:pressed {
                background-color: #1c5985;
            }
            
            QPushButton#convertButton:disabled {
                background-color: #bdc3c7;
                color: #7f8c8d;
            }
            
            QPushButton#browseButton {
                background-color: #27ae60;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                padding: 8px 16px;
            }
            
            QPushButton#browseButton:hover {
                background-color: #229954;
            }
            
            QPushButton#clearButton {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 4px;
                font-weight: bold;
                padding: 8px 16px;
            }
            
            QPushButton#clearButton:hover {
                background-color: #c0392b;
            }
            
            QLineEdit {
                border: 2px solid #bdc3c7;
                border-radius: 4px;
                padding: 8px;
                font-size: 13px;
                background-color: white;
            }
            
            QLineEdit:focus {
                border-color: #3498db;
            }
            
            QScrollArea#fileScrollArea {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: white;
            }
            
            QScrollArea#fileScrollArea QScrollBar:vertical {
                border: none;
                background: #f0f0f0;
                width: 12px;
                border-radius: 6px;
            }
            
            QScrollArea#fileScrollArea QScrollBar::handle:vertical {
                background: #bdc3c7;
                border-radius: 6px;
                min-height: 20px;
            }
            
            QScrollArea#fileScrollArea QScrollBar::handle:vertical:hover {
                background: #95a5a6;
            }
            
            QLabel#statusLabel {
                color: #34495e;
                font-size: 12px;
                padding: 8px;
                background-color: #ecf0f1;
                border-radius: 4px;
            }
            
            QLabel#previewLabel {
                font-size: 16px;
                font-weight: bold;
                color: #2c3e50;
            }
            
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 4px;
                text-align: center;
                font-weight: bold;
            }
            
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 2px;
            }
            
            QCheckBox {
                spacing: 8px;
                font-size: 13px;
            }
            
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            
            QCheckBox::indicator:unchecked {
                border: 2px solid #bdc3c7;
                border-radius: 3px;
                background-color: white;
            }
            
            QCheckBox::indicator:checked {
                border: 2px solid #3498db;
                border-radius: 3px;
                background-color: #3498db;
            }
            
            QWebEngineView#webView {
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                background-color: #f0f0f0;
            }
            

        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_area.setStyleSheet("""
                QLabel#dropArea {
                    border: 2px dashed #007bff;
                    border-radius: 8px;
                    background-color: #e3f2fd;
                    color: #0056b3;
                    font-size: 14px;
                }
            """)
            
    def dragLeaveEvent(self, event):
        self.drop_area.setStyleSheet("""
            QLabel#dropArea {
                border: 2px dashed #cccccc;
                border-radius: 8px;
                background-color: #f8f9fa;
                color: #6c757d;
                font-size: 14px;
            }
            QLabel#dropArea:hover {
                border-color: #007bff;
                background-color: #e3f2fd;
            }
        """)
            
    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        markdown_files = [f for f in files if f.lower().endswith(('.md', '.markdown'))]
        
        if markdown_files:
            self.add_files(markdown_files)
            
        self.drop_area.setStyleSheet("""
            QLabel#dropArea {
                border: 2px dashed #cccccc;
                border-radius: 8px;
                background-color: #f8f9fa;
                color: #6c757d;
                font-size: 14px;
            }
            QLabel#dropArea:hover {
                border-color: #007bff;
                background-color: #e3f2fd;
            }
        """)
            
    def browse_input_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Markdownファイルを選択", "", 
            "Markdown Files (*.md *.markdown);;All Files (*.*)"
        )
        if file_paths:
            self.add_files(file_paths)
            
    def add_files(self, file_paths):
        """ファイルをリストに追加"""
        for file_path in file_paths:
            if file_path not in self.current_files:
                self.current_files.append(file_path)
                
                # ファイルアイテムを作成
                file_item = FileItem(file_path, self)
                self.file_items.append(file_item)
                
                # スペーサーの前に挿入
                self.file_layout.insertWidget(self.file_layout.count() - 1, file_item)
        
        self.update_ui_state()
        
        # 最初のファイルを選択してプレビュー
        if self.file_items:
            self.select_file_item(self.file_items[0])
            
    def clear_files(self):
        """ファイルリストをクリア"""
        self.current_files.clear()
        
        # ファイルアイテムを削除
        for item in self.file_items:
            self.file_layout.removeWidget(item)
            item.deleteLater()
        self.file_items.clear()
        
        self.current_file_index = 0
        self.web_view.setHtml("")
        self.update_ui_state()
        
    def remove_single_file(self, file_path):
        """単一ファイルを削除"""
        if file_path in self.current_files:
            # 削除するファイルのインデックスを取得
            remove_index = self.current_files.index(file_path)
            
            # リストから削除
            self.current_files.remove(file_path)
            
            # ファイルアイテムを削除
            for i, item in enumerate(self.file_items):
                if item.file_path == file_path:
                    self.file_layout.removeWidget(item)
                    item.deleteLater()
                    self.file_items.pop(i)
                    break
            
            # 現在のファイルインデックスを調整
            if remove_index <= self.current_file_index:
                self.current_file_index = max(0, self.current_file_index - 1)
            
            # プレビューを更新
            if self.file_items:
                # 有効なインデックスに調整
                if self.current_file_index >= len(self.file_items):
                    self.current_file_index = len(self.file_items) - 1
                
                self.select_file_item(self.file_items[self.current_file_index])
            else:
                self.web_view.setHtml("")
                self.current_file_index = 0
            
            self.update_ui_state()
         
    def select_file_item(self, selected_item):
        """ファイルアイテムを選択"""
        # 全てのアイテムの選択を解除
        for item in self.file_items:
            item.set_selected(False)
        
        # 選択されたアイテムを選択状態にする
        selected_item.set_selected(True)
        
        # インデックスを更新
        self.current_file_index = self.file_items.index(selected_item)
        
        # プレビューを更新
        self.load_markdown_file(selected_item.file_path)
        
    def update_ui_state(self):
        """UIの状態を更新"""
        has_files = len(self.current_files) > 0
        self.convert_btn.setEnabled(has_files)
        self.clear_files_btn.setEnabled(has_files)
        
        if has_files:
            self.status_label.setText(f"{len(self.current_files)}個のファイルが選択されています")
            self.drop_area.setText("📄 追加のファイルを\nドラッグ&ドロップ")
        else:
            self.status_label.setText("ファイルを選択してください")
            self.drop_area.setText("📄 ここにMarkdownファイルを\nドラッグ&ドロップ")
            
    def browse_output_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "出力フォルダを選択")
        if folder_path:
            self.output_folder.setText(folder_path)
            
    def load_markdown_file(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # デフォルトの出力フォルダを設定
            if not self.output_folder.text():
                self.output_folder.setText(str(Path(file_path).parent))
            
            # プレビューを更新
            html_content = self.markdown_to_html(content)
            
            # ベースURLを設定（ローカル画像の読み込み用）
            base_url = QUrl.fromLocalFile(str(Path(file_path).parent) + '/')
            self.web_view.setHtml(html_content, base_url)
            
            # ファイル名を更新
            file_name = Path(file_path).stem
            self.status_label.setText(f"プレビュー中: {file_name}.md")
            
            # ページ情報を更新（簡略化）
            self.update_page_info()
            
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"ファイルの読み込みに失敗しました:\n{str(e)}")
            
    def update_page_info(self):
        """ページ情報を更新"""
        if self.current_files:
            current_index = self.current_file_index + 1
            total_files = len(self.current_files)
            self.page_info.setText(f"{current_index} / {total_files}")
            
            # ページ制御ボタンの状態を更新
            self.prev_page_btn.setEnabled(current_index > 1)
            self.next_page_btn.setEnabled(current_index < total_files)
        else:
            self.page_info.setText("- / -")
            self.prev_page_btn.setEnabled(False)
            self.next_page_btn.setEnabled(False)
            
    def prev_page(self):
        """前のファイルに移動"""
        if self.current_file_index > 0:
            self.select_file_item(self.file_items[self.current_file_index - 1])
            
    def next_page(self):
        """次のファイルに移動"""
        if self.current_file_index < len(self.file_items) - 1:
            self.select_file_item(self.file_items[self.current_file_index + 1])
            
    def zoom_in(self):
        """ズームイン"""
        if self.current_zoom < 3.0:
            self.current_zoom += 0.1
            self.web_view.setZoomFactor(self.current_zoom)
            self.zoom_info.setText(f"{int(self.current_zoom * 100)}%")
            
    def zoom_out(self):
        """ズームアウト"""
        if self.current_zoom > 0.3:
            self.current_zoom -= 0.1
            self.web_view.setZoomFactor(self.current_zoom)
            self.zoom_info.setText(f"{int(self.current_zoom * 100)}%")
            
    def markdown_to_html(self, markdown_content):
        """MarkdownをHTMLに変換"""
        # 改ページマーカーを処理
        page_break_marker = self.pagebreak_input.text() or self.page_break_marker
        markdown_content = markdown_content.replace(
            page_break_marker, 
            '<div class="page-break" style="page-break-after: always;"></div>'
        )
        
        # Markdown拡張機能の設定
        extensions = [
            'fenced_code',
            'tables',
            'toc' if self.include_toc.isChecked() else None,
            'nl2br',
            'sane_lists',
            'codehilite',
            'attr_list',
            'def_list',
            'footnotes',
            'md_in_html',
            'meta',
        ]
        extensions = [ext for ext in extensions if ext]
        
        # MarkdownをHTMLに変換
        md = markdown.Markdown(extensions=extensions)
        body_html = md.convert(markdown_content)
        
        # HTML生成（Mermaidを簡素化）
        html_template = f'''
<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="color-scheme" content="light">
    <meta name="print-color-adjust" content="exact">
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Helvetica Neue', 'Hiragino Sans', Arial, 'Meiryo', sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 900px;
            margin: 20px auto;
            padding: 20px;
            background: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 4px;
        }}
        h1, h2, h3, h4, h5, h6 {{
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }}
        h1 {{ font-size: 2em; border-bottom: 1px solid #eaecef; padding-bottom: .3em; }}
        h2 {{ font-size: 1.5em; border-bottom: 1px solid #eaecef; padding-bottom: .3em; }}
        h3 {{ font-size: 1.25em; }}
        code {{
            background-color: rgba(27,31,35,.05);
            border-radius: 3px;
            font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace;
            font-size: 85%;
            margin: 0;
            padding: .2em .4em;
        }}
        pre {{
            background-color: #f6f8fa;
            border-radius: 3px;
            font-size: 85%;
            line-height: 1.45;
            overflow: auto;
            padding: 16px;
        }}
        pre code {{
            background-color: transparent;
            border: 0;
            display: inline;
            line-height: inherit;
            margin: 0;
            overflow: visible;
            padding: 0;
            word-wrap: normal;
        }}
        table {{
            border-collapse: collapse;
            margin: 16px 0;
            width: 100%;
            display: table;
        }}
        table th, table td {{
            border: 1px solid #dfe2e5;
            padding: 6px 13px;
        }}
        table th {{
            background-color: #f6f8fa;
            font-weight: 600;
        }}
        blockquote {{
            border-left: 4px solid #dfe2e5;
            color: #6a737d;
            padding-left: 16px;
            margin: 16px 0;
        }}
        img {{
            max-width: 100%;
            height: auto;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }}
        .mermaid {{
            text-align: center;
            margin: 16px 0;
            padding: 16px;
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 4px;
        }}
        /* 印刷用スタイル */
        @media print {{
            body {{
                margin: 0;
                padding: 10mm;
                font-size: 10pt;
                color: #333 !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            .page-break {{
                page-break-after: always;
                break-after: page;
            }}
            h1, h2, h3 {{
                page-break-after: avoid;
                color: #333 !important;
            }}
            pre, table, blockquote {{
                page-break-inside: avoid;
            }}
            code {{
                background-color: rgba(27,31,35,.05) !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            pre {{
                background-color: #f6f8fa !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            table th {{
                background-color: #f6f8fa !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            blockquote {{
                border-left: 4px solid #dfe2e5 !important;
                color: #6a737d !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            .mermaid {{
                background-color: #f8f9fa !important;
                border: 1px solid #e9ecef !important;
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
            img {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
                color-adjust: exact !important;
            }}
            * {{
                -webkit-print-color-adjust: exact !important;
                print-color-adjust: exact !important;
            }}
        }}
    </style>
</head>
<body>
    {body_html}
    <script>
        // Mermaidコードブロックを処理（エラーハンドリング付き）
        try {{
            if (typeof document !== 'undefined' && document.querySelectorAll) {{
                document.querySelectorAll('pre code.language-mermaid').forEach(function(element) {{
                    try {{
                        const mermaidCode = element.textContent;
                        const mermaidDiv = document.createElement('div');
                        mermaidDiv.className = 'mermaid';
                        mermaidDiv.innerHTML = '<pre style="text-align: left; background: #f8f9fa; padding: 10px; border-radius: 4px;"><code>' + 
                                              mermaidCode.replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</code></pre>';
                        if (element.parentNode && element.parentNode.parentNode) {{
                            element.parentNode.parentNode.replaceChild(mermaidDiv, element.parentNode);
                        }}
                    }} catch (e) {{
                        console.log('Mermaid processing error:', e);
                    }}
                }});
            }}
        }} catch (e) {{
            console.log('Document processing error:', e);
        }}
    </script>
</body>
</html>
'''
        return html_template
        
    def convert_to_pdf(self):
        if not self.current_files or not self.output_folder.text():
            QMessageBox.warning(self, "警告", "入力ファイルと出力先を指定してください")
            return
            
        # 複数ファイルの連続変換を開始
        self.current_conversion_index = 0
        self.total_files = len(self.current_files)
        self.conversion_errors = []
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.convert_btn.setEnabled(False)
        
        self.convert_next_file()
        
    def convert_next_file(self):
        """次のファイルを変換"""
        if self.current_conversion_index >= self.total_files:
            # すべてのファイルの変換が完了
            self.on_all_conversions_finished()
            return
            
        current_file = self.current_files[self.current_conversion_index]
        file_name = Path(current_file).stem
        
        self.status_label.setText(f"変換中 ({self.current_conversion_index + 1}/{self.total_files}): {file_name}.md")
        
        # 全体進捗を更新（ファイル開始時点で基本進捗を設定）
        overall_progress = int((self.current_conversion_index * 100) / self.total_files)
        self.progress_bar.setValue(overall_progress)
        
        # 出力パスを生成
        output_path = Path(self.output_folder.text()) / f"{file_name}.pdf"
        
        # 上書き確認
        if output_path.exists() and self.confirm_overwrite.isChecked():
            reply = QMessageBox.question(
                self, "上書き確認", 
                f"ファイル '{output_path.name}' は既に存在します。\n上書きしますか？",
                QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
            )
            
            if reply == QMessageBox.Cancel:
                self.on_all_conversions_finished()
                return
            elif reply == QMessageBox.No:
                # このファイルをスキップして次へ
                self.current_conversion_index += 1
                self.convert_next_file()
                return
        
        try:
            # MarkdownをHTMLに変換
            with open(current_file, 'r', encoding='utf-8') as f:
                content = f.read()
            html_content = self.markdown_to_html(content)
            
            # 全体進捗を更新（HTML変換完了）
            file_progress = int((self.current_conversion_index * 100) / self.total_files)
            step_progress = int(20 / self.total_files)  # 各ファイルの20%進捗
            self.progress_bar.setValue(file_progress + step_progress)
            
            # ベースURLを設定
            base_url = QUrl.fromLocalFile(str(Path(current_file).parent) + '/')
            
            # WebEngineでHTMLをロード
            self.current_output_path = str(output_path)
            self.web_view.loadFinished.connect(self.on_html_loaded)
            self.web_view.setHtml(html_content, base_url)
            
        except Exception as e:
            self.conversion_errors.append(f"{file_name}.md: {str(e)}")
            self.current_conversion_index += 1
            self.convert_next_file()
            
    def on_html_loaded(self, success):
        """HTMLのロードが完了したらPDFを生成"""
        if not success:
            file_name = Path(self.current_files[self.current_conversion_index]).stem
            self.conversion_errors.append(f"{file_name}.md: HTMLのロードに失敗しました")
            self.current_conversion_index += 1
            self.convert_next_file()
            return
            
        try:
            # シグナルを切断
            self.web_view.loadFinished.disconnect()
            
            # 全体進捗を更新（HTMLロード完了）
            file_progress = int((self.current_conversion_index * 100) / self.total_files)
            step_progress = int(40 / self.total_files)  # 各ファイルの40%進捗
            self.progress_bar.setValue(file_progress + step_progress)
            
            # Mermaidのレンダリング完了を待つ（exe化で時間がかかる場合があるため延長）
            QTimer.singleShot(3000, self.generate_pdf)
            
        except Exception as e:
            file_name = Path(self.current_files[self.current_conversion_index]).stem
            self.conversion_errors.append(f"{file_name}.md: {str(e)}")
            self.current_conversion_index += 1
            self.convert_next_file()
            
    def generate_pdf(self):
        """PDFを生成（プリンターを使わない方法）"""
        try:
            # 全体進捗を更新（PDF生成開始）
            file_progress = int((self.current_conversion_index * 100) / self.total_files)
            step_progress = int(60 / self.total_files)  # 各ファイルの60%進捗
            self.progress_bar.setValue(file_progress + step_progress)
            
            output_path = self.current_output_path
            
            # WebEnginePageのprintToPdfを使用（カラー印刷設定付き）
            try:
                # ページレイアウトの設定（カラー印刷対応）
                page_layout = QPageLayout()
                page_layout.setPageSize(QPageSize(QPageSize.A4))
                page_layout.setOrientation(QPageLayout.Portrait)
                
                # printToPdfの正しい使用方法（レイアウト指定でカラー印刷）
                self.web_view.page().printToPdf(output_path, page_layout)
                # PDF生成完了を待つ（10秒後にチェック - exe化や複数ファイル処理で時間がかかる）
                QTimer.singleShot(10000, lambda: self.check_pdf_generation(output_path))
            except Exception as pdf_error:
                # プリンターエラーが発生した場合
                file_name = Path(self.current_files[self.current_conversion_index]).stem
                self.conversion_errors.append(f"{file_name}.md: PDF生成エラー - {str(pdf_error)}")
                self.current_conversion_index += 1
                self.convert_next_file()
            
        except Exception as e:
            file_name = Path(self.current_files[self.current_conversion_index]).stem
            self.conversion_errors.append(f"{file_name}.md: {str(e)}")
            self.current_conversion_index += 1
            self.convert_next_file()
            
    def check_pdf_generation(self, output_path):
        """PDF生成の成功をチェック"""
        try:
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                if file_size > 1024:  # 1KB以上あれば成功とみなす
                    self.on_single_conversion_finished(True, output_path)
                else:
                    # ファイルサイズが小さい場合、さらに5秒待つ
                    QTimer.singleShot(5000, lambda: self.check_pdf_generation_final(output_path))
            else:
                # ファイルが存在しない場合、さらに5秒待つ
                QTimer.singleShot(5000, lambda: self.check_pdf_generation_final(output_path))
        except Exception as e:
            file_name = Path(self.current_files[self.current_conversion_index]).stem
            self.conversion_errors.append(f"{file_name}.md: PDF生成チェックエラー - {str(e)}")
            self.current_conversion_index += 1
            self.convert_next_file()
            
    def check_pdf_generation_final(self, output_path):
        """PDF生成の最終チェック（より寛容な判定）"""
        try:
            if os.path.exists(output_path):
                file_size = os.path.getsize(output_path)
                if file_size > 0:  # ファイルが存在して0バイトでなければ成功とみなす
                    # ファイルサイズを確認してより詳細なログを出力
                    file_name = Path(self.current_files[self.current_conversion_index]).stem
                    print(f"PDF生成成功: {file_name}.pdf (サイズ: {file_size} bytes)")
                    self.on_single_conversion_finished(True, output_path)
                else:
                    # 0バイトファイルの場合は失敗
                    file_name = Path(self.current_files[self.current_conversion_index]).stem
                    self.conversion_errors.append(f"{file_name}.md: PDFファイルが空です（0バイト）")
                    self.current_conversion_index += 1
                    self.convert_next_file()
            else:
                # ファイルが存在しない場合は失敗
                file_name = Path(self.current_files[self.current_conversion_index]).stem
                self.conversion_errors.append(f"{file_name}.md: PDFファイルが生成されませんでした")
                self.current_conversion_index += 1
                self.convert_next_file()
        except Exception as e:
            file_name = Path(self.current_files[self.current_conversion_index]).stem
            self.conversion_errors.append(f"{file_name}.md: PDF生成最終チェックエラー - {str(e)}")
            self.current_conversion_index += 1
            self.convert_next_file()
            
    def on_single_conversion_finished(self, success, output_path):
        """単一ファイルの変換完了時の処理"""
        if not success:
            file_name = Path(self.current_files[self.current_conversion_index]).stem
            self.conversion_errors.append(f"{file_name}.md: 変換に失敗しました")
        
        # 全体進捗を更新（ファイル完了）
        file_progress = int(((self.current_conversion_index + 1) * 100) / self.total_files)
        self.progress_bar.setValue(file_progress)
        
        # 次のファイルに進む
        self.current_conversion_index += 1
        self.convert_next_file()
        
    def on_all_conversions_finished(self):
        """すべてのファイルの変換完了時の処理"""
        self.progress_bar.setValue(100)
        self.convert_btn.setEnabled(True)
        
        success_count = self.total_files - len(self.conversion_errors)
        
        if len(self.conversion_errors) == 0:
            # すべて成功
            self.status_label.setText(f"すべてのファイルの変換が完了しました ({success_count}個)")
            QMessageBox.information(
                self, "変換完了", 
                f"すべてのファイルの変換が完了しました。\n\n"
                f"変換されたファイル数: {success_count}\n"
                f"出力先: {self.output_folder.text()}"
            )
        else:
            # 一部または全部失敗
            self.status_label.setText(f"変換完了: 成功 {success_count}個, 失敗 {len(self.conversion_errors)}個")
            error_message = f"変換結果:\n成功: {success_count}個\n失敗: {len(self.conversion_errors)}個\n\n"
            error_message += "エラー詳細:\n" + "\n".join(self.conversion_errors[:5])
            if len(self.conversion_errors) > 5:
                error_message += f"\n... 他 {len(self.conversion_errors) - 5}個のエラー"
                
            QMessageBox.warning(self, "変換完了（エラーあり）", error_message)
            
        # プログレスバーを非表示にする
        QTimer.singleShot(2000, lambda: self.progress_bar.setVisible(False))


def main():
    # 非推奨警告を抑制
    import warnings
    warnings.filterwarnings("ignore", category=DeprecationWarning)
    
    # OpenGLエラー対策: ソフトウェアレンダリングを強制
    os.environ['QT_OPENGL'] = 'software'
    os.environ['QT_ANGLE_PLATFORM'] = 'software'
    os.environ['QTWEBENGINE_CHROMIUM_FLAGS'] = '--disable-gpu --disable-software-rasterizer --disable-gpu-compositing --disable-gpu-rasterization'
    
    # アプリケーションの設定
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    # OpenGL問題を回避するための追加設定
    QApplication.setAttribute(Qt.AA_UseSoftwareOpenGL, True)
    
    app = QApplication(sys.argv)
    app.setApplicationName("Markdown to PDF Converter")
    
    converter = MarkdownToPdfConverter()
    converter.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()