#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Echos — Private LLM client with encryption, RAG, and chat management
# Copyright (c) 2026 Vecsai
# Distributed under the MIT License. See LICENSE file for details.
# Project: https://github.com/Vecsai/Echos

import sys
import json
import requests
import time
import hashlib
import base64
import io
import zipfile
import subprocess
import re
import tempfile
import importlib.util
import os
import threading
import ast
import logging
logger = logging.getLogger(__name__)
import shutil
print(shutil.which('dot'))

from contextlib import contextmanager
from docx.shared import Inches, RGBColor


try:
    import graphviz
    GRAPHVIZ_AVAILABLE = True
except ImportError:
    graphviz = None
    GRAPHVIZ_AVAILABLE = False


try:
    from pptx import Presentation
    PPTX_AVAILABLE = True
except ImportError:
    Presentation = None
    PPTX_AVAILABLE = False


try:
    from sentence_transformers import SentenceTransformer
    import faiss
    import numpy as np
    RAG_AVAILABLE = True
except ImportError:
    SentenceTransformer = None
    faiss = None
    np = None
    RAG_AVAILABLE = False


try:
    import matplotlib
    matplotlib.use('QtAgg')
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    matplotlib = None
    FigureCanvas = None
    Figure = None
    MATPLOTLIB_AVAILABLE = False


try:
    from pygments import highlight
    from pygments.lexers import get_lexer_by_name, TextLexer
    from pygments.formatters import HtmlFormatter
    PYGMENTS_AVAILABLE = True
except ImportError:
    highlight = None
    get_lexer_by_name = None
    TextLexer = None
    HtmlFormatter = None
    PYGMENTS_AVAILABLE = False


try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    Fernet = None
    CRYPTO_AVAILABLE = False


try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    Image = None
    PIL_AVAILABLE = False


from typing import List, Dict, Any, Optional

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTextBrowser, QTextEdit, QPushButton, QLabel, QLineEdit, QSpinBox,
    QDoubleSpinBox, QCheckBox, QFileDialog, QMessageBox, QTreeWidget,
    QTreeWidgetItem, QFormLayout, QGroupBox, QDialog, QListWidget,
    QListWidgetItem, QInputDialog, QMenu, QSplitter, QScrollArea, QButtonGroup, QFontComboBox, QFrame, QSizePolicy,
    QKeySequenceEdit, QHeaderView, QDateEdit, QGraphicsOpacityEffect, QProgressBar
)
from PySide6.QtCore import Qt, QThread, Signal, QEvent, QSize, QTimer, QUrl, QPoint, QRegularExpression, \
    QRect, QDate, QPropertyAnimation, QEasingCurve, QFileSystemWatcher

from PySide6.QtGui import QTextCursor, QIcon, QColor, QBrush, QPixmap, QPalette, QTextDocument, QDesktopServices, QFont, \
    QSyntaxHighlighter, QTextCharFormat, QKeySequence, QPainter, QPen, QAction

from PySide6.QtPrintSupport import QPrinter
from datetime import datetime

# === Патч для .exe: автоматический поиск ресурсов в sys._MEIPASS ===
if hasattr(sys, '_MEIPASS'):
    import os as _os
    _MEIPASS = sys._MEIPASS

    # Патч QIcon
    _orig_qicon_init = QIcon.__init__
    def _patched_qicon_init(self, *args, **kwargs):
        new_args = []
        for a in args:
            if isinstance(a, str) and not _os.path.isabs(a):
                # Сначала пробуем путь из _MEIPASS
                meipass_path = _os.path.join(_MEIPASS, a)
                if _os.path.exists(meipass_path):
                    a = meipass_path
            new_args.append(a)
        _orig_qicon_init(self, *new_args, **kwargs)
    QIcon.__init__ = _patched_qicon_init

    # Патч QPixmap
    _orig_qpixmap_init = QPixmap.__init__
    def _patched_qpixmap_init(self, *args, **kwargs):
        new_args = []
        for a in args:
            if isinstance(a, str) and not _os.path.isabs(a):
                meipass_path = _os.path.join(_MEIPASS, a)
                if _os.path.exists(meipass_path):
                    a = meipass_path
            new_args.append(a)
        _orig_qpixmap_init(self, *new_args, **kwargs)
    QPixmap.__init__ = _patched_qpixmap_init

    print(f"[PATCH] Ресурсы будут искаться в: {_MEIPASS}")

try:
    import markdown
    MARKDOWN_AVAILABLE = True
except ImportError:
    markdown = None
    MARKDOWN_AVAILABLE = False


try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError:
    Document = None
    DOCX_AVAILABLE = False


try:
    from PyPDF2 import PdfReader
    PDF_AVAILABLE = True
except ImportError:
    PdfReader = None
    PDF_AVAILABLE = False


try:
    import pypdfium2 as pdfium
    PDF_IMAGES_AVAILABLE = True
except ImportError:
    pdfium = None
    PDF_IMAGES_AVAILABLE = False


# === Информация о программе ===
APP_VERSION = "5.0.0"
GITHUB_REPO = "Vecsai/Echos"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
print(GITHUB_API_URL)


class ChatWorker(QThread):
    """Диалог настройки внешнего вида чата (цвет фона или обои).
    Позволяет выбрать цвет, загрузить изображение и обрезать его для использования в качестве фона.
    Основной класс"""

    chunk_received = Signal(str)
    finished_ok = Signal(str, str)  # (answer, reasoning)
    error_occurred = Signal(str)
    stopped = Signal()
    retry_attempt = Signal(int, int, str)  # (попытка, всего, ошибка)

    def __init__(self, api_base: str, model: str, messages: List[Dict[str, str]],
                 params: Dict[str, Any], stream: bool = True):
        super().__init__()
        self.api_base = api_base.rstrip('/')
        self.model = model
        self.messages = messages
        self.params = params
        self.stream = stream
        self._stop_requested = False
        self._response = None

        # Настройки буферизации
        self.buffer = ""
        self.buffer_limit = 150

    def stop(self):
        self._stop_requested = True
        if self._response:
            try:
                self._response.close()
            except Exception:
                pass

    def _flush_buffer(self):
        if self.buffer:
            self.chunk_received.emit(self.buffer)
            self.buffer = ""

    def run(self):
        url = f"{self.api_base}/chat/completions"
        headers = {"Content-Type": "application/json"}
        payload = {
            "model": self.model,
            "messages": self.messages,
            **self.params,
            "stream": self.stream
        }

        max_retries = 3
        retry_delay = 2
        attempt = 0

        while attempt < max_retries:
            try:
                if self.stream:
                    with requests.post(
                            url, headers=headers, json=payload, stream=True,
                            timeout=(10, 120)  # (connect timeout, read timeout)
                    ) as resp:
                        self._response = resp
                        if resp.status_code != 200:
                            self.error_occurred.emit(f"HTTP {resp.status_code}: {resp.text}")
                            return

                        full_text = ""
                        full_reasoning = ""

                        try:
                            for line in resp.iter_lines():
                                if self._stop_requested:
                                    self.stopped.emit()
                                    return
                                if not line:
                                    continue
                                line = line.decode('utf-8')
                                if line.startswith("data: "):
                                    data_str = line[6:]
                                    if data_str.strip() == "[DONE]":
                                        break
                                    try:
                                        data = json.loads(data_str)
                                        delta = data.get("choices", [{}])[0].get("delta", {})
                                        content = delta.get("content", "")
                                        reasoning = delta.get("reasoning_content") or delta.get("reasoning") or ""

                                        if reasoning:
                                            full_reasoning += reasoning
                                        if content:
                                            full_text += content
                                            self.buffer += content
                                            if len(self.buffer) >= self.buffer_limit:
                                                self._flush_buffer()

                                    except json.JSONDecodeError:
                                        continue
                        except (requests.exceptions.ChunkedEncodingError,
                                requests.exceptions.ConnectionError,
                                requests.exceptions.ReadTimeout,
                                ConnectionResetError,
                                BrokenPipeError,
                                OSError,
                                AttributeError) as e:
                            # Если пользователь нажал «Стоп» — это не ошибка, а нормальная остановка
                            if self._stop_requested:
                                self.stopped.emit()
                                return

                            print(f"[ChatWorker] Обрыв соединения: {type(e).__name__}: {e}")
                            if full_text.strip():
                                self._flush_buffer()
                                self.finished_ok.emit(full_text, full_reasoning)
                                return
                            else:
                                self.error_occurred.emit(
                                    f"Соединение прервано на середине ответа.\n"
                                    f"Причина: {type(e).__name__}\n\n"
                                    "Проверь сеть и попробуй снова."
                                )
                                return

                        self._flush_buffer()
                        self._response = None

                        if not full_text.strip() and full_reasoning.strip():
                            self.finished_ok.emit("", full_reasoning)
                        elif full_text.strip():
                            self.finished_ok.emit(full_text, full_reasoning)
                        else:
                            self.error_occurred.emit("Модель вернула пустой ответ.")
                            return
                    return  # успех

                else:
                    resp = requests.post(url, headers=headers, json=payload, timeout=300)
                    self._response = resp
                    if resp.status_code != 200:
                        self.error_occurred.emit(f"HTTP {resp.status_code}: {resp.text}")
                        return
                    if self._stop_requested:
                        self.stopped.emit()
                        return

                    data = resp.json()
                    message = data.get("choices", [{}])[0].get("message", {})
                    content = message.get("content", "")
                    reasoning = message.get("reasoning_content") or message.get("reasoning") or ""
                    self._response = None

                    if not content.strip() and reasoning.strip():
                        self.finished_ok.emit("", reasoning)
                    elif content.strip():
                        self.finished_ok.emit(content, reasoning)
                    else:
                        self.error_occurred.emit("Модель вернула пустой ответ.")
                        return
                    return  # успех

            except requests.exceptions.RequestException as e:
                attempt += 1
                if self._stop_requested:
                    self.stopped.emit()
                    return
                if attempt < max_retries:
                    self.retry_attempt.emit(attempt, max_retries, str(e))
                    time.sleep(retry_delay)
                    retry_delay *= 2
                    continue
                else:
                    self.error_occurred.emit(
                        f"Не удалось получить ответ после {max_retries} попыток.\n\n"
                        f"Последняя ошибка: {str(e)}"
                    )
                    return


class RAGManager:
    """Менеджер базы знаний (RAG).
    Хранит текстовые фрагменты документов, строит эмбеддинги с помощью sentence-transformers и FAISS, выполняет поиск релевантных фрагментов по запросу.
    Основной класс"""

    def __init__(self):
        self.documents = []  # список текстовых фрагментов
        self.embeddings = None  # numpy матрица эмбеддингов
        self.index = None  # faiss индекс
        self.model = None
        if RAG_AVAILABLE:
            try:
                # Используем лёгкую модель эмбеддингов
                self.model = SentenceTransformer('all-MiniLM-L6-v2')
            except Exception as e:
                print(f"Не удалось загрузить модель эмбеддингов: {e}")
                self.model = None

    def add_document(self, text: str, chunk_size: int = 1000):
        """Разбивает текст на чанки и сохраняет их."""

        chunks = [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]
        self.documents.extend(chunks)
        if self.model and RAG_AVAILABLE:
            # Пересчитываем эмбеддинги
            self._rebuild_index()

    def _rebuild_index(self):
        if not self.documents or not self.model:
            return
        embeddings = self.model.encode(self.documents, convert_to_numpy=True)
        dimension = embeddings.shape[1]
        self.index = faiss.IndexFlatL2(dimension)
        self.index.add(embeddings)
        self.embeddings = embeddings

    def search(self, query: str, top_k: int = 3):
        """Ищет наиболее релевантные чанки. Возвращает список текстов."""

        if not self.documents:
            return []
        if self.model and self.index:
            query_vec = self.model.encode([query], convert_to_numpy=True)
            distances, indices = self.index.search(query_vec, top_k)
            # Порог: чем меньше расстояние, тем ближе фрагмент. Отсекаем далёкие
            threshold = 1.0
            result = []
            for dist, i in zip(distances[0], indices[0]):
                if i < len(self.documents) and dist < threshold:
                    result.append(self.documents[i])
            return result
        else:

            query_words = set(query.lower().split())
            scored = []
            for i, doc in enumerate(self.documents):
                doc_words = set(doc.lower().split())
                score = len(query_words & doc_words)
                scored.append((score, i))
            scored.sort(reverse=True)
            return [self.documents[i] for _, i in scored[:top_k]]

    def add_document_smart(self, text: str, max_chunk_size: int = 1000):
        """
        Разбивает текст на смысловые блоки (по заголовкам, абзацам, спискам)
        и добавляет их в документы. Если блок слишком большой, дополнительно
        разрезается по предложениям.
        """

        if not text or not text.strip():
            return

        # Предварительная обработка: разбиваем на строки
        lines = text.splitlines()
        blocks = []
        current_block = []
        current_len = 0

        def flush_block():
            nonlocal current_block, current_len
            if current_block:
                block_text = "\n".join(current_block).strip()
                if block_text:
                    blocks.append(block_text)
                current_block = []
                current_len = 0

        for line in lines:
            stripped = line.strip()
            if stripped.startswith("#") or stripped.startswith("---") or stripped.startswith("***"):
                flush_block()
                current_block.append(line)
                current_len += len(line)
                continue

            if not stripped:
                flush_block()
                continue

            current_block.append(line)
            current_len += len(line)

            if current_len >= max_chunk_size:
                flush_block()
        flush_block()

        final_chunks = []
        for block in blocks:
            if len(block) <= max_chunk_size:
                final_chunks.append(block)
            else:
                sub_chunks = self._split_by_sentences(block, max_chunk_size)
                final_chunks.extend(sub_chunks)

        self.documents.extend(final_chunks)
        if self.model and RAG_AVAILABLE:
            self._rebuild_index()

    def _split_by_sentences(self, text: str, max_size: int) -> List[str]:
        """Разбивает длинный текст на части по предложениям, не превышая max_size."""

        import re
        sentences = re.split(r'(?<=[.!?])\s+', text)
        chunks = []
        current = ""
        for sent in sentences:
            if len(current) + len(sent) + 1 <= max_size:
                current = (current + " " + sent).strip()
            else:
                if current:
                    chunks.append(current)
                current = sent
        if current:
            chunks.append(current)
        return chunks


class ChatAppearanceDialog(QDialog):
    """Диалог настройки внешнего вида чата (цвет фона или обои).
    Позволяет выбрать цвет, загрузить изображение и обрезать его для использования в качестве фона."""

    def __init__(self, parent=None, current_bg_color=None, current_bg_image=None,
                 current_bg_image_original=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Wallpaper_2.png"))
        self.setWindowTitle("Настройка обоев чата")
        self.setModal(True)
        self.selected_color = current_bg_color
        self.selected_image_path = current_bg_image_original or current_bg_image  # для отображения используем исходное
        self.result_image_path = None  # сюда сохраним итоговый путь (обрезанный или исходный)
        self.result_image_original = None  # сюда сохраним исходный путь
        self.init_ui()

        if current_bg_color:
            for btn, color in zip(self.color_buttons, self.pastel_colors):
                if color == current_bg_color:
                    btn.setChecked(True)
                    break

        if self.selected_image_path and os.path.exists(self.selected_image_path):
            self.show_image_preview(self.selected_image_path)
        else:
            self.crop_label.setText("Нет изображения")
            self.btn_remove_image.setVisible(False)
            self.update_remove_button_position()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Выбор цвета
        layout.addWidget(QLabel("Выберите цвет:"))
        color_layout = QHBoxLayout()
        self.pastel_colors = ["#FFFFFF", "#FFF9C4", "#F8BBD0", "#E1BEE7", "#BBDEFB", "#C8E6C9", "#E0E0E0"]
        self.color_group = QButtonGroup(self)
        self.color_group.setExclusive(True)
        self.color_buttons = []
        for color in self.pastel_colors:
            btn = QPushButton()
            btn.setCheckable(True)
            btn.setFixedSize(30, 30)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {color};
                    border-radius: 15px;
                    border: 2px solid transparent;
                }}
                QPushButton:checked {{
                    border: 2px solid #333;
                }}
            """)
            self.color_group.addButton(btn)
            btn.clicked.connect(lambda checked, c=color: self.select_color(c))
            self.color_buttons.append(btn)
            color_layout.addWidget(btn)
        color_layout.addStretch()
        layout.addLayout(color_layout)

        # Выбор изображения
        btn_choose_image = QPushButton("Выбрать изображение")
        btn_choose_image.clicked.connect(self.choose_image)
        layout.addWidget(btn_choose_image)

        # Превью изображения с выделением области
        self.preview_container = QWidget(self)
        self.preview_container.setMinimumSize(300, 200)
        layout.addWidget(self.preview_container)

        self.crop_label = CropLabel(self.preview_container)
        self.crop_label.setStyleSheet("border: 1px solid #ccc; background: #f0f0f0;")
        self.crop_label.setMinimumSize(300, 200)
        self.crop_label.setAlignment(Qt.AlignCenter)
        self.crop_label.setText("Перетащите изображение\nили нажмите «Выбрать»")

        self.btn_remove_image = QPushButton(self.preview_container)
        self.btn_remove_image.setIcon(QIcon("Image/Close.png"))
        self.btn_remove_image.setIconSize(QSize(12, 12))
        self.btn_remove_image.setFixedSize(18, 18)
        self.btn_remove_image.setStyleSheet("""
            QPushButton {
                border-radius: 9px;
                background-color: rgba(255, 0, 0, 0.7);
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(200, 0, 0, 0.9);
            }
        """)
        self.btn_remove_image.setVisible(False)
        self.btn_remove_image.clicked.connect(self.remove_image)

        # Кнопка очистки выделения
        btn_clear_sel = QPushButton("Очистить выделение")
        btn_clear_sel.clicked.connect(self.crop_label.clear_selection)
        layout.addWidget(btn_clear_sel)

        # Кнопки
        btn_box = QHBoxLayout()
        btn_default = QPushButton("По умолчанию")
        btn_default.clicked.connect(self.reset_to_default)
        btn_save = QPushButton("Сохранить")
        btn_save.clicked.connect(self.accept)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(btn_default)
        btn_box.addStretch()
        btn_box.addWidget(btn_save)
        btn_box.addWidget(btn_cancel)
        layout.addLayout(btn_box)

        self.update_remove_button_position()

    def choose_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите изображение", "",
                                                   "Изображения (*.png *.jpg *.jpeg *.bmp *.gif *.webp)")
        if file_path:
            self.selected_image_path = os.path.abspath(file_path)
            self.result_image_original = self.selected_image_path
            self.show_image_preview(self.selected_image_path)
            self.selected_color = None
            self.color_group.setExclusive(False)
            for btn in self.color_buttons:
                btn.setChecked(False)
            self.color_group.setExclusive(True)
            self.update_remove_button_position()

    def accept(self):
        if self.selected_image_path:
            cropped = self.crop_label.get_cropped_pixmap()
            if cropped and not cropped.isNull() and self.crop_label._selection_rect:
                # Если было выделение, сохраняем обрезанное во временный файл
                import tempfile
                temp_path = os.path.join(tempfile.gettempdir(), f"wallpaper_{int(time.time())}.png")
                cropped.save(temp_path, "PNG")
                self.result_image_path = temp_path
                self.result_image_original = self.selected_image_path
            else:
                # Если выделения нет, используем исходное изображение как есть
                self.result_image_path = self.selected_image_path
                self.result_image_original = self.selected_image_path
        else:
            self.result_image_path = None
            self.result_image_original = None
        super().accept()

    def update_remove_button_position(self):
        if hasattr(self, 'crop_label') and hasattr(self, 'btn_remove_image'):
            x = self.crop_label.width() - self.btn_remove_image.width() - 2
            y = 2
            self.btn_remove_image.move(x, y)

    def select_color(self, color):
        self.selected_color = color
        self.selected_image_path = None
        self.crop_label.setPixmap(QPixmap())
        self.crop_label.setText("Нет изображения")
        self.btn_remove_image.setVisible(False)
        self.update_remove_button_position()

    def set_image_from_path(self, file_path):
        self.selected_image_path = os.path.abspath(file_path)
        self.show_image_preview(self.selected_image_path)
        self.selected_color = None
        self.color_group.setExclusive(False)
        for btn in self.color_buttons:
            btn.setChecked(False)
        self.color_group.setExclusive(True)
        self.update_remove_button_position()

    def show_image_preview(self, path):
        pixmap = QPixmap(path)
        if not pixmap.isNull():
            self.crop_label.set_pixmap(pixmap)
            self.crop_label.setText("")
            self.btn_remove_image.setVisible(True)
        else:
            self.crop_label.setPixmap(QPixmap())
            self.crop_label.setText("Ошибка загрузки")
            self.btn_remove_image.setVisible(False)
        self.update_remove_button_position()

    def remove_image(self):
        self.selected_image_path = None
        self.crop_label.setPixmap(QPixmap())
        self.crop_label.setText("Нет изображения")
        self.btn_remove_image.setVisible(False)
        self.update_remove_button_position()

    def reset_to_default(self):
        self.selected_color = None
        self.selected_image_path = None
        self.crop_label.setPixmap(QPixmap())
        self.crop_label.setText("Нет изображения")
        self.btn_remove_image.setVisible(False)
        self.color_group.setExclusive(False)
        for btn in self.color_buttons:
            btn.setChecked(False)
        self.color_group.setExclusive(True)


class CropLabel(QLabel):
    """Метка для выбора области изображения (обрезка обоев).
    Поддерживает выделение прямоугольной области мышью."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setMouseTracking(True)
        self._original_pixmap = None
        self._scaled_pixmap = None
        self._selection_rect = None
        self._origin = None
        self._drawing = False
        self.setMinimumSize(300, 200)

    def set_pixmap(self, pixmap: QPixmap):
        self._original_pixmap = pixmap
        self._selection_rect = None
        self._update_scaled_pixmap()

    def _update_scaled_pixmap(self):
        if self._original_pixmap:
            self._scaled_pixmap = self._original_pixmap.scaled(
                self.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.setPixmap(self._scaled_pixmap)
        else:
            self._scaled_pixmap = None
            self.setPixmap(QPixmap())

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._update_scaled_pixmap()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._scaled_pixmap:
            self._origin = event.position().toPoint()
            self._selection_rect = QRect(self._origin, QPoint())
            self._drawing = True
            self.update()

    def mouseMoveEvent(self, event):
        if self._drawing:
            self._selection_rect = QRect(self._origin, event.position().toPoint()).normalized()
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self._drawing:
            self._drawing = False
            if self._selection_rect and (self._selection_rect.width() < 5 or self._selection_rect.height() < 5):
                self._selection_rect = None
            self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        if self._selection_rect and self._scaled_pixmap:
            painter = QPainter(self)
            pen = QPen(QColor(0, 120, 215), 2)
            painter.setPen(pen)
            painter.drawRect(self._selection_rect)

    def get_cropped_pixmap(self) -> QPixmap:
        """Возвращает вырезанный фрагмент исходного изображения согласно выделению.
           Если выделения нет, возвращает исходное изображение целиком."""

        if not self._selection_rect or not self._original_pixmap or not self._scaled_pixmap:
            return self._original_pixmap

        # Получаем фактическую область содержимого QLabel (без рамки)
        content_rect = self.contentsRect()

        # Вычисляем смещение масштабированного изображения внутри содержимого
        # (из-за центрирования по умолчанию)
        offset_x = (content_rect.width() - self._scaled_pixmap.width()) // 2
        offset_y = (content_rect.height() - self._scaled_pixmap.height()) // 2

        # Преобразуем координаты выделения (в системе виджета) в координаты изображения
        # Вычитаем смещение, чтобы получить координаты относительно самого изображения
        x_in_pixmap = self._selection_rect.x() - offset_x
        y_in_pixmap = self._selection_rect.y() - offset_y

        # Масштабируем в координаты исходного изображения
        scale_x = self._original_pixmap.width() / self._scaled_pixmap.width()
        scale_y = self._original_pixmap.height() / self._scaled_pixmap.height()
        x = int(x_in_pixmap * scale_x)
        y = int(y_in_pixmap * scale_y)
        w = int(self._selection_rect.width() * scale_x)
        h = int(self._selection_rect.height() * scale_y)

        # Ограничиваем, чтобы не выйти за границы исходного изображения
        x = max(0, min(x, self._original_pixmap.width() - 1))
        y = max(0, min(y, self._original_pixmap.height() - 1))
        w = min(w, self._original_pixmap.width() - x)
        h = min(h, self._original_pixmap.height() - y)

        return self._original_pixmap.copy(x, y, w, h)

    def clear_selection(self):
        self._selection_rect = None
        self.update()


class ChatTreeWidget(QTreeWidget):
    """Дерево чатов и папок.
    Поддерживает drag&drop, подсветку папок при перетаскивании, управление закреплением."""

    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self._highlighted_item = None
        self._highlighted_original_brush = None
        self._dragged_item = None
        self.setMouseTracking(True)

        # Включаем drag & drop
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)

        self.setStyleSheet("""
            QTreeView::drop-indicator {
                background: #1a73e8;
                border: 1px solid #0d47a1;
            }
        """)

    def _clear_highlight(self):
        if self._highlighted_item:
            if self._highlighted_original_brush is not None:
                self._highlighted_item.setBackground(0, self._highlighted_original_brush)
            else:
                self._highlighted_item.setBackground(0, QBrush())
            self._highlighted_item = None
            self._highlighted_original_brush = None

    def _highlight_item(self, item):
        if item is None or item == self._highlighted_item:
            return
        self._clear_highlight()
        data = item.data(0, Qt.UserRole)
        if data and data.get('type') == 'folder':
            current_brush = item.background(0)
            self._highlighted_original_brush = current_brush
            self._highlighted_item = item
            item.setBackground(0, QColor("#d0e8ff"))

    def _get_separator_item(self):
        """Возвращает элемент-разделитель, если он есть."""

        for i in range(self.topLevelItemCount()):
            it = self.topLevelItem(i)
            if it.data(0, Qt.UserRole).get('type') == 'separator':
                return it
        return None

    def startDrag(self, supportedActions):
        self._dragged_item = self.currentItem()
        super().startDrag(supportedActions)

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-qabstractitemmodeldatalist"):
            event.acceptProposedAction()
        else:
            event.ignore()
        self._clear_highlight()

    def dragMoveEvent(self, event):
        pos = event.position().toPoint()
        item = self.itemAt(pos)
        dragged = self.currentItem() or self._dragged_item

        # Проверяем допустимость сброса
        if not self._can_drop_on(dragged, item):
            self._clear_highlight()
            event.ignore()
            return

        # Подсвечиваем папку, если курсор над ней
        if item is not None:
            data = item.data(0, Qt.UserRole)
            if data and data.get('type') == 'folder':
                self._highlight_item(item)
            else:
                self._clear_highlight()
        else:
            self._clear_highlight()

        event.setDropAction(Qt.MoveAction)
        event.accept()
        event.acceptProposedAction()

    def _can_drop_on(self, dragged, target):
        if dragged is None or target is None:
            return True  # корень разрешён
        dragged_data = dragged.data(0, Qt.UserRole)
        target_data = target.data(0, Qt.UserRole)
        if dragged_data and target_data:
            if dragged_data.get('type') == 'folder' and target_data.get('type') == 'folder':
                return False  # папку в папку нельзя
        return True

    def dragLeaveEvent(self, event):
        self._clear_highlight()
        super().dragLeaveEvent(event)
        self._dragged_item = None

    def dropEvent(self, event):
        self._clear_highlight()
        dragged_item = self._dragged_item if self._dragged_item else self.currentItem()
        drop_pos = event.position().toPoint()
        target_item = self.itemAt(drop_pos)

        # Запрещаем сброс папки на папку
        if not self._can_drop_on(dragged_item, target_item):
            event.ignore()
            self._dragged_item = None
            return

        # Стандартное перемещение
        event.setDropAction(Qt.MoveAction)
        super().dropEvent(event)

        if dragged_item:
            # Определяем новый статус закрепления
            new_pinned = False
            if target_item is not None:
                target_data = target_item.data(0, Qt.UserRole)
                if target_data and target_data.get('type') == 'folder':
                    new_pinned = target_data.get('pinned', False)
                else:
                    # Если цель не папка, оставляем текущий статус
                    new_pinned = self.main_window.is_item_pinned(dragged_item)
            else:
                # Сброс в корень
                separator = self._get_separator_item()
                if separator is None:
                    new_pinned = False
                else:
                    sep_rect = self.visualItemRect(separator)
                    if drop_pos.y() < sep_rect.top():
                        new_pinned = True  # выше разделителя — закрепляем
                    else:
                        new_pinned = False  # ниже — открепляем

            # Обновляем флаг напрямую, без дополнительных перемещений
            data = dragged_item.data(0, Qt.UserRole)
            data['pinned'] = new_pinned
            dragged_item.setData(0, Qt.UserRole, data)

            # Обновляем данные чата или папки в списках
            if data['type'] == 'chat':
                filepath = data['filepath']
                for chat in self.main_window.chats:
                    if chat['filepath'] == filepath:
                        chat['pinned'] = new_pinned
                        self.main_window.save_chat_by_filepath(filepath)
                        break
            elif data['type'] == 'folder':
                folder_name = data['name']
                for folder in self.main_window.folders:
                    if isinstance(folder, dict) and folder.get('name') == folder_name:
                        folder['pinned'] = new_pinned
                        self.main_window.save_folders()
                        break

            # Перестраиваем дерево отложенно, чтобы не тормозить интерфейс
            QTimer.singleShot(0, self.main_window.on_tree_structure_changed)

        self._dragged_item = None
        event.acceptProposedAction()


class ChatTextBrowser(QTextBrowser):
    """Область отображения сообщений чата (наследник QTextBrowser).
    Включает обработку drag&drop для прикрепления файлов."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.scroll_down_button = None
        self.stop_button = None
        self.btn_chat_settings = None
        self.btn_toggle_reasoning = None
        self._bg_color = None
        self._bg_image_path = None
        self._bg_pixmap = None
        self.search_panel = None
        self.main_window = None

        # === ВАЖНО: включаем приём drops на самом виджете и его viewport ===
        self.setAcceptDrops(True)
        self.viewport().setAcceptDrops(True)

    # === Переопределяем dragEnterEvent для viewport ===
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    # === Переопределяем dragMoveEvent для viewport (чтобы разрешить drop) ===
    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    # === Переопределяем dropEvent для viewport ===
    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and self.main_window:
            for url in urls:
                if url.isLocalFile():
                    path = url.toLocalFile()
                    if os.path.exists(path):
                        self.main_window.attachments.append(path)
            self.main_window.update_attachments_ui()
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.scroll_down_button:
            self.scroll_down_button.update_position()
        if self.stop_button:
            self.stop_button.update_position()
        if self.btn_chat_settings:
            self.btn_chat_settings.update_position()
        if self.btn_toggle_reasoning:
            self.btn_toggle_reasoning.update_position()
        if self.search_panel and self.main_window:
            self.main_window.update_search_panel_position()


class CodeHighlighter(QSyntaxHighlighter):
    """Синтаксическая подсветка кода в поле ввода (для языков программирования)."""

    def __init__(self, document):
        super().__init__(document)
        self.highlighting_rules = []

        # Форматы для подсветки
        keyword_format = QTextCharFormat()
        keyword_format.setForeground(QColor("#569CD6"))
        keyword_format.setFontWeight(QFont.Bold)

        string_format = QTextCharFormat()
        string_format.setForeground(QColor("#CE9178"))

        comment_format = QTextCharFormat()
        comment_format.setForeground(QColor("#6A9955"))
        comment_format.setFontItalic(True)

        # Простейшие правила: ключевые слова, строки, комментарии
        keywords = ["def", "class", "if", "else", "elif", "for", "while", "return",
                    "import", "from", "try", "except", "finally", "with", "as",
                    "True", "False", "None", "and", "or", "not", "in", "is"]
        for kw in keywords:
            pattern = QRegularExpression(f"\\b{kw}\\b")
            self.highlighting_rules.append((pattern, keyword_format))

        # Строки в двойных или одинарных кавычках
        self.highlighting_rules.append(
            (QRegularExpression(r'"[^"\\]*(\\.[^"\\]*)*"'), string_format)
        )
        self.highlighting_rules.append(
            (QRegularExpression(r"'[^'\\]*(\\.[^'\\]*)*'"), string_format)
        )

        # Комментарии (до конца строки)
        self.highlighting_rules.append(
            (QRegularExpression(r"#[^\n]*"), comment_format)
        )

    def highlightBlock(self, text):
        for pattern, fmt in self.highlighting_rules:
            match_iterator = pattern.globalMatch(text)
            while match_iterator.hasNext():
                match = match_iterator.next()
                self.setFormat(match.capturedStart(), match.capturedLength(), fmt)


class PlainPasteTextEdit(QTextEdit):
    """Текстовый редактор для ввода сообщений.
    Вставляет только чистый текст, обрабатывает вставку изображений и файлов, имеет собственное контекстное меню."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = None

    def insertFromMimeData(self, source):
        added_attachments = False

        # Обработка изображения
        if source.hasImage():
            if self.main_window:
                image = source.imageData()
                if image:
                    import tempfile
                    temp_dir = tempfile.gettempdir()
                    file_path = os.path.join(temp_dir, f"pasted_image_{int(time.time())}.png")
                    image.save(file_path, "PNG")
                    self.main_window.attachments.append(file_path)
                    self.main_window.update_attachments_ui()
                    self.main_window.statusBar().showMessage("Изображение добавлено как вложение")
                    added_attachments = True

        # Обработка файлов/URL
        if source.hasUrls():
            if self.main_window:
                for url in source.urls():
                    if url.isLocalFile():
                        path = url.toLocalFile()
                        if os.path.exists(path):
                            self.main_window.attachments.append(path)
                            added_attachments = True
                if added_attachments:
                    self.main_window.update_attachments_ui()
                    self.main_window.statusBar().showMessage("Файлы добавлены как вложения")

        # Если были добавлены вложения, не вставляем текст
        if added_attachments:
            return

        # Обычный текст – вставляем как чистый текст
        if source.hasText():
            self.insertPlainText(source.text())
        else:
            super().insertFromMimeData(source)

    def _get_context_icon(self, name):
        """Возвращает иконку для контекстного меню с учётом темы."""

        if self.main_window and hasattr(self.main_window, 'settings') and self.main_window.settings.get("theme", "light") == "dark":
            folder = "Image_Context_menu_white"
            base, ext = os.path.splitext(name)
            filename = f"{base}_White{ext}"
        else:
            folder = "Image_Context_menu"
            filename = name
        path = os.path.join(folder, filename)
        if self.main_window and hasattr(self.main_window, 'resource_path'):
            full_path = self.main_window.resource_path(path)
        else:
            full_path = path
        return QIcon(full_path)

    def contextMenuEvent(self, event):
        """Заменяем стандартное контекстное меню на своё."""

        menu = QMenu(self)

        has_selection = self.textCursor().hasSelection()
        clipboard = QApplication.clipboard()
        can_paste = clipboard.mimeData().hasText() or clipboard.mimeData().hasUrls()

        # Копировать
        copy_action = menu.addAction(self._get_context_icon("Copy.png"), "Копировать")
        copy_action.setEnabled(has_selection)
        copy_action.triggered.connect(self.copy)

        # Вырезать
        cut_action = menu.addAction(self._get_context_icon("Cut.png"), "Вырезать")
        cut_action.setEnabled(has_selection)
        cut_action.triggered.connect(self.cut)

        # Вставить
        paste_action = menu.addAction(self._get_context_icon("Paste.png"), "Вставить")
        paste_action.setEnabled(can_paste)
        paste_action.triggered.connect(self.paste)

        # Удалить
        delete_action = menu.addAction(self._get_context_icon("Delete.png"), "Удалить")
        delete_action.setEnabled(has_selection)
        delete_action.triggered.connect(self.delete_selected)

        menu.addSeparator()

        # Выбрать все
        select_all_action = menu.addAction(self._get_context_icon("Select_all.png"), "Выбрать все")
        select_all_action.triggered.connect(self.selectAll)

        menu.addSeparator()

        # Вернуть
        undo_action = menu.addAction(self._get_context_icon("Undo.png"), "Вернуть")
        undo_action.setEnabled(self.document().isUndoAvailable())
        undo_action.triggered.connect(self.undo)

        # Повторить
        redo_action = menu.addAction(self._get_context_icon("Redo.png"), "Повторить")
        redo_action.setEnabled(self.document().isRedoAvailable())
        redo_action.triggered.connect(self.redo)

        menu.addSeparator()

        # Сохранить промпт (если есть текст)
        text = self.toPlainText().strip()
        if text and self.main_window:
            save_prompt_action = menu.addAction(self._get_context_icon("Save_promt.png"), "Сохранить промпт")
            save_prompt_action.triggered.connect(lambda: self.main_window.save_prompt(text))

        menu.exec(event.globalPos())

    def delete_selected(self):
        cursor = self.textCursor()
        if cursor.hasSelection():
            cursor.removeSelectedText()
            self.setTextCursor(cursor)


class BusyOverlay(QWidget):
    """Полупрозрачный оверлей с надписью 'Загрузка…'."""

    def __init__(self, parent):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents, False)
        self.setAttribute(Qt.WA_NoSystemBackground)
        self.setStyleSheet("background-color: rgba(0, 0, 0, 120);")
        self.hide()

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)

        self.label = QLabel("Загрузка…")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("""
            QLabel {
                color: white;
                background-color: rgba(0, 0, 0, 180);
                padding: 14px 28px;
                border-radius: 8px;
                font-size: 16px;
            }
        """)
        layout.addWidget(self.label, alignment=Qt.AlignCenter)

    def show_overlay(self, parent_widget, text="Загрузка…"):
        self.label.setText(text)
        self.setParent(parent_widget)
        self.setGeometry(parent_widget.rect())
        self.raise_()
        self.show()
        QApplication.processEvents()

    def hide_overlay(self):
        self.hide()
        QApplication.processEvents()


class LoadingSplash(QWidget):
    """Сплэш-экран с прогрессом загрузки."""

    def __init__(self, dark=False):
        super().__init__()
        self.setWindowFlags(Qt.SplashScreen | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setFixedSize(420, 260)

        # === Цвета под тему ===
        if dark:
            bg = "#2b2b2b"
            title_color = "#ffffff"
            status_color = "#cccccc"
            bar_border = "#555"
            bar_bg = "#1e1e1e"
            bar_chunk = "#D8A7B1"
            bar_text = "#ffffff"
        else:
            bg = "#f5f5f5"
            title_color = "#000000"
            status_color = "#666666"
            bar_border = "#cccccc"
            bar_bg = "#ffffff"
            bar_chunk = "#D8A7B1"
            bar_text = "#000000"

        self.setStyleSheet(f"""
            QWidget {{
                background-color: {bg};
                border-radius: 10px;
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)

        # Логотип
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setStyleSheet("background: transparent;")
        logo_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)),
            "Icon/Icon.png"
        )
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path)
            if not pixmap.isNull():
                pixmap = pixmap.scaled(96, 96, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                logo_label.setPixmap(pixmap)
        layout.addWidget(logo_label, alignment=Qt.AlignCenter)

        # Название
        title = QLabel("Echos")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(
            f"color: {title_color}; font-size: 26px; font-weight: bold; background: transparent;"
        )
        layout.addWidget(title)

        # Статус
        self.status = QLabel("Загрузка…")
        self.status.setAlignment(Qt.AlignCenter)
        self.status.setStyleSheet(
            f"color: {status_color}; font-size: 12px; background: transparent;"
        )
        layout.addWidget(self.status)

        # Прогресс-бар
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFixedWidth(340)
        self.progress.setFixedHeight(18)
        self.progress.setStyleSheet(f"""
            QProgressBar {{
                border: 1px solid {bar_border};
                border-radius: 6px;
                background: {bar_bg};
                text-align: center;
                color: {bar_text};
                font-size: 11px;
            }}
            QProgressBar::chunk {{
                background: {bar_chunk};
                border-radius: 5px;
            }}
        """)
        layout.addWidget(self.progress, alignment=Qt.AlignCenter)

        # Центрируем
        screen = QApplication.primaryScreen().availableGeometry()
        self.move(
            (screen.width() - self.width()) // 2,
            (screen.height() - self.height()) // 2
        )

    def set_progress(self, current, total, text=None):
        if total > 0:
            percent = int(current * 100 / total)
            self.progress.setValue(percent)
        if text is not None:
            self.status.setText(text)
        QApplication.processEvents()


class MessageWidget(QFrame):
    """Виджет одного сообщения в чате (рамка с текстом, кнопками действий, вложениями).
    Может отображать роль (пользователь/ассистент), время, reasoning, закрепление и звезду."""

    edit_clicked = Signal(int)
    copy_clicked = Signal(int)
    regenerate_clicked = Signal(int)
    unpin_clicked = Signal(int)
    context_menu_requested = Signal(int, QPoint, object)
    attachment_clicked = Signal(dict)

    def __init__(self, role: str, content: str = "", time_str: str = "", reasoning: str = "",
                 pinned: bool = False, idx: int = None, starred: bool = False, parent=None,
                 image_cache: dict = None, show_actions: bool = True,  comment: str = "", reaction: str = None, quote: str = None, raw_content: str = None):
        super().__init__(parent)
        self.image_cache = image_cache if image_cache is not None else {}

        self.role = role
        self.idx = idx
        self.pinned = pinned
        self.text_color = "#000000"
        self.content_html = content
        self.content_text = content
        self.raw_content = raw_content if raw_content is not None else content
        self.decrypt_func = None
        self.starred = starred
        self.comment_text = comment
        self.reaction = reaction
        self.quote = quote

        if role == "user":
            self.bg_color = "#e3f2fd"
            self.label_color = "#1a73e8"
            self.label_text = "👤 Пользователь"
            self.alignment = Qt.AlignRight
        else:
            self.bg_color = "#f1f1f1"
            self.label_color = "#188038"
            self.label_text = "🤖 Ассистент"
            self.alignment = Qt.AlignLeft

        self.setStyleSheet(f"""
            QFrame {{
                background-color: {self.bg_color};
                border-radius: 18px;
                border: none;
            }}
        """)

        self.setMaximumWidth(600)
        self.setMinimumWidth(100)
        self.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Fixed)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(10, 6, 10, 6)
        main_layout.setSpacing(4)

        header_layout = QHBoxLayout()
        header_layout.setSpacing(4)

        self.role_label = QLabel(self.label_text)
        self.role_label.setStyleSheet(f"color: {self.label_color}; font-weight: 600; background: transparent;")
        header_layout.addWidget(self.role_label)

        self.star_label = QLabel("⭐")
        self.star_label.setStyleSheet("background: transparent; color: #f5b942; font-size: 14px;")
        if not self.starred:
            self.star_label.hide()
        header_layout.addWidget(self.star_label)

        if time_str:
            self.time_label = QLabel(time_str)
            self.time_label.setStyleSheet("color: #999; font-size: 0.8em; background: transparent;")
            header_layout.addWidget(self.time_label)

        header_layout.addStretch()

        self.reaction_label = QLabel(self.reaction if self.reaction else "")
        self.reaction_label.setStyleSheet("background: transparent; font-size: 14px;")
        if not self.reaction:
            self.reaction_label.hide()
        header_layout.addWidget(self.reaction_label)

        # Кнопки действий
        if show_actions:
            button_style = """
                QPushButton { background: transparent; border: none; color: #1a73e8; font-size: 14px; }
                QPushButton:hover { background-color: rgba(0,0,0,0.05); border-radius: 12px; }
            """
            self.btn_edit = QPushButton()
            self.btn_edit.setIcon(QIcon("Image/Edit.png"))
            self.btn_edit.setIconSize(QSize(16, 16))
            self.btn_edit.setFixedSize(24, 24)
            self.btn_edit.setCursor(Qt.PointingHandCursor)
            self.btn_edit.setStyleSheet(button_style)
            self.btn_edit.clicked.connect(lambda: self.edit_clicked.emit(self.idx))
            header_layout.addWidget(self.btn_edit)

            self.btn_copy = QPushButton()
            self.btn_copy.setIcon(QIcon("Image/Copy.png"))
            self.btn_copy.setIconSize(QSize(16, 16))
            self.btn_copy.setFixedSize(24, 24)
            self.btn_copy.setCursor(Qt.PointingHandCursor)
            self.btn_copy.setStyleSheet(button_style)
            self.btn_copy.clicked.connect(lambda: self.copy_clicked.emit(self.idx))
            header_layout.addWidget(self.btn_copy)

            if role == "assistant":
                self.btn_regenerate = QPushButton()
                self.btn_regenerate.setIcon(QIcon("Image/Reboot.png"))
                self.btn_regenerate.setIconSize(QSize(16, 16))
                self.btn_regenerate.setFixedSize(24, 24)
                self.btn_regenerate.setCursor(Qt.PointingHandCursor)
                self.btn_regenerate.setStyleSheet(button_style)
                self.btn_regenerate.clicked.connect(lambda: self.regenerate_clicked.emit(self.idx))
                header_layout.addWidget(self.btn_regenerate)

            if pinned:
                self.btn_unpin = QPushButton()
                self.btn_unpin.setIcon(QIcon("Image/Close_pin.png"))
                self.btn_unpin.setIconSize(QSize(16, 16))
                self.btn_unpin.setFixedSize(24, 24)
                self.btn_unpin.setCursor(Qt.PointingHandCursor)
                self.btn_unpin.setStyleSheet(button_style)
                self.btn_unpin.clicked.connect(lambda: self.unpin_clicked.emit(self.idx))
                header_layout.addWidget(self.btn_unpin)

        main_layout.addLayout(header_layout)

        if reasoning and role == "assistant":
            self.reasoning_label = QLabel(reasoning)
            self.reasoning_label.setWordWrap(True)
            self.reasoning_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self.reasoning_label.setStyleSheet("""
                background-color: #fff9c4;
                border-left: 4px solid #fbc02d;
                padding: 4px 6px;
                border-radius: 4px;
                font-style: italic;
                color: #555;
            """)
            self.reasoning_label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
            main_layout.addWidget(self.reasoning_label)

        self.content_label = QLabel(content)
        self.content_label.setWordWrap(True)
        self.content_label.setTextFormat(Qt.RichText)
        self.content_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.content_label.setStyleSheet(f"color: {self.text_color}; background: transparent;")
        self.content_label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        main_layout.addWidget(self.content_label)
        if self.comment_text:
            self.comment_label = QLabel("📝 " + self.comment_text)
            self.comment_label.setWordWrap(True)
            self.comment_label.setStyleSheet(
                "color: #795548; background: #f9f2e7; border-left: 3px solid #a1887f; "
                "padding: 2px 4px; margin-top: 4px;"
            )
            self.comment_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            main_layout.addWidget(self.comment_label)
        else:
            self.comment_label = None

        if self.quote:
            self.quote_label = QLabel(self.quote)
            self.quote_label.setWordWrap(True)
            self.quote_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
            self.quote_label.setStyleSheet("""
                    background-color: #f0f0f0;
                    border-left: 4px solid #1a73e8;
                    padding: 6px 8px;
                    margin-top: 4px;
                    color: #333;
                    font-style: italic;
                """)
            main_layout.addWidget(self.quote_label)
        else:
            self.quote_label = None

        self.attachments_layout = QVBoxLayout()
        self.attachments_layout.setContentsMargins(0, 4, 0, 0)
        self.attachments_layout.setSpacing(4)
        main_layout.addLayout(self.attachments_layout)

        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.emit_context_menu)
        for child in self.findChildren(QLabel):
            child.setContextMenuPolicy(Qt.NoContextMenu)

    def apply_search_highlight(self, query: str, is_assistant: bool = False,
                               markdown_func=None, escape_func=None):
        """Подсвечивает ТОЛЬКО совпадающие фрагменты в тексте сообщения."""

        # Если запрос пустой — восстанавливаем исходный вид
        if not query or not query.strip():
            if hasattr(self, 'content_html'):
                self.content_label.setText(self.content_html)
            return

        query = query.strip()
        import re

        START = "\uE000"
        END = "\uE001"

        # Ищем все совпадения в ЧИСТОМ тексте (raw_content)
        pattern = re.compile(re.escape(query), re.IGNORECASE)
        marked = pattern.sub(lambda m: START + m.group(0) + END, self.raw_content)

        # Преобразуем в HTML, сохраняя маркеры
        if is_assistant and markdown_func:
            html = markdown_func(marked)
        elif escape_func:
            html = escape_func(marked).replace('\n', '<br>')
        else:
            html = marked

        # На случай если markdown превратил маркеры в HTML-энтити
        html = html.replace("&#xE000;", START).replace("&#xE001;", END)
        html = html.replace("&#57344;", START).replace("&#57345;", END)

        # Оборачиваем только совпадения в подсветку
        html = html.replace(START, '<span style="background-color: #FB8C00; color: #000000;">')
        html = html.replace(END, '</span>')

        self.content_label.setText(html)

    def emit_context_menu(self, pos):
        self.context_menu_requested.emit(self.idx, self.mapToGlobal(pos), self)

    def set_content_text(self, text: str):
        self.content_text = text
        self.content_label.setText(text)

    def set_attachments(self, attachments: list):
        # Очищаем старые виджеты
        while self.attachments_layout.count():
            item = self.attachments_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for att in attachments:
            if att.get('type') == 'image':
                path = att.get('path')
                if not path or not os.path.exists(path):
                    continue

                # Кнопка с миниатюрой изображения
                btn = QPushButton()
                btn.setFixedSize(180, 180)
                btn.setIconSize(QSize(160, 160))
                btn.setStyleSheet("""
                    QPushButton {
                        border: 1px solid #ccc;
                        border-radius: 8px;
                        background: #f5f5f5;
                        color: #000;
                    }
                    QPushButton:hover {
                        background: #e0e0e0;
                    }
                """)
                btn.setCursor(Qt.PointingHandCursor)
                btn.clicked.connect(lambda checked, a=att: self.attachment_clicked.emit(a))
                self.attachments_layout.addWidget(btn)

                # Асинхронная загрузка миниатюры
                QTimer.singleShot(0, lambda b=btn, p=path: self._load_image_into_button(b, p))

            elif att.get('type') == 'file_text':
                # Кнопка с именем файла и иконкой документа
                btn = QPushButton(f"📄 {att.get('name', 'Файл')}")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setStyleSheet("""
                    QPushButton {
                        text-align: left;
                        padding: 8px;
                        background: #f0f0f0;
                        border: 1px solid #ccc;
                        border-radius: 6px;
                        color: #000;
                    }
                    QPushButton:hover {
                        background: #e0e0e0;
                    }
                """)
                btn.clicked.connect(lambda checked, a=att: self.attachment_clicked.emit(a))
                self.attachments_layout.addWidget(btn)

            elif att.get('type') == 'file':
                # Кнопка с именем файла и иконкой файла
                btn = QPushButton(f"📁 {att.get('name', 'Файл')}")
                btn.setCursor(Qt.PointingHandCursor)
                btn.setStyleSheet("""
                    QPushButton {
                        text-align: left;
                        padding: 8px;
                        background: #f9f9f9;
                        border: 1px solid #ccc;
                        border-radius: 6px;
                    }
                    QPushButton:hover {
                        background: #e0e0e0;
                    }
                """)
                btn.clicked.connect(lambda checked, a=att: self.attachment_clicked.emit(a))
                self.attachments_layout.addWidget(btn)

        self.attachments_layout.invalidate()
        self.adjustSize()

    def _load_image_into_button(self, button: QPushButton, image_path: str):
        """Загружает миниатюру изображения в кнопку."""

        preview = self._load_image_preview(image_path, max_size=160)
        if preview.isNull():
            button.setText("Ошибка загрузки")
            return
        button.setIcon(QIcon(preview))
        button.setText("")

    def _load_image_async(self, label: QLabel, image_path: str):
        preview = self._load_image_preview(image_path, max_size=400)
        if preview.isNull():
            label.setText("Ошибка загрузки")
            return

        if label.width() > 0 and preview.width() > label.width():
            preview = preview.scaledToWidth(label.width(), Qt.SmoothTransformation)

        label.setPixmap(preview)
        label.setStyleSheet("")
        label.setText("")
        label.adjustSize()
        label.updateGeometry()
        self.attachments_layout.invalidate()
        self.adjustSize()

    def _load_image_preview(self, image_path: str, max_size: int = 400) -> QPixmap:
        if image_path in self.image_cache:
            return self.image_cache[image_path]

        # Проверяем, зашифрован ли файл — если да, кэш миниатюр не используем
        skip_cache = False
        try:
            with open(image_path, 'rb') as f:
                head = f.read(4)
            if head.startswith(b"ENC:"):
                skip_cache = True
        except OSError:
            return QPixmap()

        # Кэш миниатюр только для незашифрованных
        thumb_path = None
        if not skip_cache and hasattr(self, 'thumbnails_dir') and self.thumbnails_dir:
            try:
                mtime = os.path.getmtime(image_path)
            except OSError:
                mtime = 0
            hash_name = hashlib.md5(f"{image_path}_{mtime}".encode('utf-8')).hexdigest() + ".jpg"
            thumb_path = os.path.join(self.thumbnails_dir, hash_name)
            if os.path.exists(thumb_path):
                cached = QPixmap(thumb_path)
                if not cached.isNull():
                    self.image_cache[image_path] = cached
                    return cached

        # Читаем файл
        try:
            with open(image_path, 'rb') as f:
                raw = f.read()
        except OSError:
            return QPixmap()

        # Расшифровываем, если нужно
        if raw.startswith(b"ENC:") and self.decrypt_func:
            try:
                raw = self.decrypt_func(raw)
            except Exception as e:
                print(f"Ошибка расшифровки {image_path}: {e}")
                return QPixmap()

        # Загружаем QPixmap
        thumb = QPixmap()
        if PIL_AVAILABLE:
            from PIL import Image
            try:
                img = Image.open(io.BytesIO(raw))
                img.thumbnail((max_size, max_size))
                buffer = io.BytesIO()
                img.convert("RGB").save(buffer, format="JPEG", quality=80)
                buffer.seek(0)
                thumb.loadFromData(buffer.read(), "JPEG")
                if not skip_cache and thumb_path and not thumb.isNull():
                    thumb.save(thumb_path, "JPEG", 80)
            except Exception:
                thumb.loadFromData(raw)
        else:
            thumb.loadFromData(raw)

        if not thumb.isNull():
            self.image_cache[image_path] = thumb
        return thumb

    def resizeEvent(self, event):
        super().resizeEvent(event)
        available_width = self.width() - 20
        if hasattr(self, 'reasoning_label'):
            self.reasoning_label.setMaximumWidth(available_width)
        self.content_label.setMaximumWidth(available_width)
        self.role_label.setMaximumWidth(available_width)
        if hasattr(self, 'time_label'):
            self.time_label.setMaximumWidth(available_width)


class TagsDialog(QDialog):
    """Диалог управления тегами чата.
    Отображает список глобальных тегов, позволяет добавлять, удалять и выбирать теги для текущего чата."""

    def __init__(self, all_tags, current_tags, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Теги")
        self.setModal(True)
        self.resize(400, 350)
        self.all_tags = list(all_tags)
        self.current_tags = list(current_tags)

        layout = QVBoxLayout(self)

        # Текущий тег
        self.current_tag_label = QLabel(
            "Текущий тег: " + (", ".join(self.current_tags) if self.current_tags else "нет"))
        layout.addWidget(self.current_tag_label)

        # Список всех тегов
        self.tags_list = QListWidget()
        self.refresh_list()
        self.tags_list.itemDoubleClicked.connect(self.on_item_double_clicked)

        # Если тёмная тема, делаем текст элементов коричневым
        if parent and hasattr(parent, 'settings') and parent.settings.get("theme", "light") == "dark":
            self.tags_list.setStyleSheet("QListWidget::item { color: #964B00; }")

        layout.addWidget(self.tags_list)

        # Кнопки
        btn_box = QHBoxLayout()
        btn_apply = QPushButton("Выбрать")
        btn_apply.clicked.connect(self.apply_selected_tag)
        btn_remove = QPushButton("Удалить тег")
        btn_remove.clicked.connect(self.remove_selected_tag)
        btn_reset = QPushButton("Сбросить")
        btn_reset.clicked.connect(self.reset_current_tag)
        btn_box.addWidget(btn_apply)
        btn_box.addWidget(btn_remove)
        btn_box.addWidget(btn_reset)
        layout.addLayout(btn_box)

        # Добавление нового тега
        add_layout = QHBoxLayout()
        self.edit_new_tag = QLineEdit()
        self.edit_new_tag.setPlaceholderText("Новый тег...")
        btn_add = QPushButton("Добавить")
        btn_add.clicked.connect(self.add_new_tag)
        add_layout.addWidget(self.edit_new_tag)
        add_layout.addWidget(btn_add)
        layout.addLayout(add_layout)

        # Отмена
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        layout.addWidget(btn_cancel)

    def refresh_list(self):
        """Обновляет список тегов, не изменяя all_tags."""
        self.tags_list.clear()
        for tag in self.all_tags:
            item = QListWidgetItem(tag)
            if tag in self.current_tags:
                item.setBackground(QColor("#d3f0d3"))
            self.tags_list.addItem(item)

    def on_item_double_clicked(self, item):
        self.apply_tag(item)
        self.accept()

    def apply_tag(self, item):
        tag = item.text()
        self.current_tags = [tag]
        self.current_tag_label.setText(f"Текущий тег: {tag}")
        self.refresh_list()

    def apply_selected_tag(self):
        item = self.tags_list.currentItem()
        if item:
            self.apply_tag(item)
            self.accept()

    def remove_selected_tag(self):
        item = self.tags_list.currentItem()
        if not item:
            return
        tag = item.text()
        reply = QMessageBox.question(self, "Удалить тег", f"Удалить тег '{tag}' из глобального списка?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return
        if tag in self.all_tags:
            self.all_tags.remove(tag)
        if tag in self.current_tags:
            self.current_tags.remove(tag)
        self.current_tag_label.setText("Текущий тег: " + (", ".join(self.current_tags) if self.current_tags else "нет"))
        self.refresh_list()

    def reset_current_tag(self):
        self.current_tags = []
        self.current_tag_label.setText("Текущий тег: нет")
        self.refresh_list()
        self.accept()

    def add_new_tag(self):
        new_tag = self.edit_new_tag.text().strip().lstrip('#')
        if new_tag and new_tag not in self.all_tags:
            self.all_tags.append(new_tag)
            self.edit_new_tag.clear()
            self.refresh_list()

    def get_current_tags(self):
        return self.current_tags

    def get_all_tags(self):
        return self.all_tags


class PluginAPI:
    """API для плагинов, предоставляющий доступ к некоторым функциям главного окна (логирование, получение текущего чата)."""

    def __init__(self, main_window):
        self.main_window = main_window

    def log(self, message):
        """Плагин может писать в лог программы."""

        print(f"[Плагин] {message}")

    def get_current_chat_title(self):
        """Получить название текущего чата."""

        if 0 <= self.main_window.current_chat_index < len(self.main_window.chats):
            return self.main_window.chats[self.main_window.current_chat_index].get('title', '')
        return None


class PluginsManagerDialog(QDialog):
    """Диалог управления плагинами (просмотр, удаление, открытие папки)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Plagins.png"))
        self.setWindowTitle("Управление плагинами")
        self.setModal(True)
        self.resize(500, 350)
        self.plugins_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plugins")
        self.init_ui()
        self.refresh_list()

    def init_ui(self):
        layout = QVBoxLayout(self)

        label = QLabel("Список плагинов (.py файлы в папке plugins):")
        layout.addWidget(label)

        self.list_widget = QListWidget()
        layout.addWidget(self.list_widget)

        btn_box = QHBoxLayout()
        btn_refresh = QPushButton("Обновить")
        btn_refresh.clicked.connect(self.refresh_list)
        btn_open_dir = QPushButton("Открыть папку")
        btn_open_dir.clicked.connect(self.open_dir)
        btn_delete = QPushButton("Удалить выбранный")
        btn_delete.clicked.connect(self.delete_selected)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)

        btn_reload = QPushButton("Перезагрузить")
        btn_reload.clicked.connect(self.on_reload_plugins)
        btn_box.addWidget(btn_reload)

        btn_box.addWidget(btn_refresh)
        btn_box.addWidget(btn_open_dir)
        btn_box.addWidget(btn_delete)
        btn_box.addStretch()
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def on_reload_plugins(self):
        parent = self.parent()
        if parent and hasattr(parent, 'reload_plugins'):
            parent.reload_plugins()
            self.refresh_list()  # обновляем список отображаемых плагинов

    def refresh_list(self):
        self.list_widget.clear()
        if os.path.isdir(self.plugins_dir):
            for f in os.listdir(self.plugins_dir):
                if f.endswith('.py'):
                    self.list_widget.addItem(f)

    def open_dir(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.plugins_dir))

    def delete_selected(self):
        item = self.list_widget.currentItem()
        if not item:
            return
        filepath = os.path.join(self.plugins_dir, item.text())
        reply = QMessageBox.question(self, "Удалить плагин", f"Удалить файл {item.text()}?",
                                     QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                os.remove(filepath)
                self.refresh_list()
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось удалить: {e}")


class AboutDialog(QDialog):
    """Окно «О программе» с информацией о версии, описанием и ссылками."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Programs.png"))
        self.setWindowTitle("О программе")
        self.setModal(True)
        self.resize(480, 420)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)
        layout.setSpacing(8)

        # === Логотип с фолбэками ===
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)

        base_dir = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            "Icon/Icon.png",
            "Image/Icon.png",
            "Icon/Icon.ico",
            "Image/Chat_2.png",
        ]
        pixmap = QPixmap()
        for path in candidates:
            full_path = os.path.join(base_dir, path)
            if os.path.exists(full_path):
                pixmap = QPixmap(full_path)
                if not pixmap.isNull():
                    break
        if not pixmap.isNull():
            logo_label.setPixmap(pixmap.scaled(96, 96, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        layout.addWidget(logo_label)

        # Название
        title = QLabel("Echos")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 22px; font-weight: bold;")
        layout.addWidget(title)

        # Версия
        version = QLabel(f"Версия: {APP_VERSION}")
        version.setAlignment(Qt.AlignCenter)
        version.setStyleSheet("color: #666;")
        layout.addWidget(version)

        layout.addSpacing(10)

        # Описание
        desc = QLabel(
            "Мощный клиент для локальных моделей LM Studio.\n"
            "Поддержка чатов, папок, тегов, закладок, шифрования, RAG и плагинов."
        )
        desc.setWordWrap(True)
        desc.setAlignment(Qt.AlignCenter)
        layout.addWidget(desc)

        layout.addSpacing(10)

        # Ссылки
        link_site = QLabel('<a href="https://github.com/Vecsai/Echos">GitHub проекта</a>')
        link_site.setOpenExternalLinks(True)
        link_site.setAlignment(Qt.AlignCenter)
        layout.addWidget(link_site)

        link_support = QLabel('<a href="https://github.com/Vecsai/Echos/issues">Сообщить о проблеме</a>')
        link_support.setOpenExternalLinks(True)
        link_support.setAlignment(Qt.AlignCenter)
        layout.addWidget(link_support)


class MainWindow(QMainWindow):
    """Главное окно приложения.
    Координирует все элементы интерфейса, управляет чатами, папками, настройками, горячими клавишами, отправкой сообщений, уведомлениями и другими функциями.
    Головной класс"""

    # ======================================================================
    #  ИНИЦИАЛИЗАЦИЯ И НАСТРОЙКА
    # ======================================================================

    def __init__(self, splash=None):
        super().__init__()
        self._splash = splash
        self.setWindowTitle("Echos")
        self.resize(1200, 900)
        self.temp_file_text = ""

        # Инициализация иконок (должна быть до использования)
        self.icon_folder_closed = QIcon("Image/Close_folder.png")
        self.icon_folder_open = QIcon("Image/Open_folder.png")
        self.icon_chat = QIcon("Image/Chat.png")

        # Загрузка иконок (обязательно до init_ui)
        self.icon_attach = QIcon(self.resource_path("Image/Attach.png"))
        self.icon_send = QIcon(self.resource_path("Image/Send.png"))
        self.icon_pin = QIcon(self.resource_path("Image/Pin.png"))
        self.icon_close = QIcon(self.resource_path("Image/Close.png"))
        self.icon_down = QIcon(self.resource_path("Image/Down.png"))
        self.icon_stop = QIcon(self.resource_path("Image/Stop_generation.png"))
        self.icon_setting_chat = QIcon(self.resource_path("Image/Setting_chat.png"))
        self.icon_wallpaper = QIcon(self.resource_path("Image/Wallpaper.png"))
        self.icon_search = QIcon(self.resource_path("Image/Search.png"))
        self.icon_trash = QIcon(self.resource_path("Image/Trash.png"))
        self.icon_statistics = QIcon(self.resource_path("Image/Statistics.png"))
        self.icon_reflection = QIcon(self.resource_path("Image/Reflection.png"))
        self.icon_promt = QIcon(self.resource_path("Image/Save_promt.png"))
        self.icon_json = QIcon(self.resource_path("Image/Json.png"))
        self.icon_pdf = QIcon(self.resource_path("Image/Pdf.png"))
        self.icon_zip = QIcon(self.resource_path("Image/Zip.png"))
        self.icon_pin_prompt = QIcon(self.resource_path("Image/Pin_promt.png"))
        self.icon_search_global = QIcon(self.resource_path("Image/Search.png"))
        self.icon_add_chat = QIcon(self.resource_path("Image/Add_chat.png"))
        self.icon_add_folder = QIcon(self.resource_path("Image/Add_folder.png"))
        self.icon_sort = QIcon(self.resource_path("Image/Sorting.png"))
        self.icon_markdown = QIcon(self.resource_path("Image/Markdown.png"))
        self.icon_archive = QIcon(self.resource_path("Image/Archive.png"))
        self.icon_archive_2 = QIcon(self.resource_path("Image/Archive_2.png"))
        self.icon_settings = QIcon(self.resource_path("Image/Settings_3.png"))
        self.icon_import_1 = QIcon(self.resource_path("Icon/Icon.png"))
        self.icon_import = QIcon(self.resource_path("Image/Import.png"))
        self.icon_export = QIcon(self.resource_path("Image/Export.png"))
        self.icon_collapse_expand = QIcon(self.resource_path("Image/Collapse_expand.png"))
        self.icon_bookmarks = QIcon(self.resource_path("Image/Bookmarks.png"))
        self.icon_clear_history = QIcon(self.resource_path("Image/Clear_history.png"))
        self.icon_repeat_request = QIcon(self.resource_path("Image/Repeat_equest.png"))
        self.icon_summary_dialogue = QIcon(self.resource_path("Image/Summary_dialogue.png"))
        self.icon_chat_parameters = QIcon(self.resource_path("Image/Chat_parameters.png"))
        self.icon_commands = QIcon(self.resource_path("Image/Commands.png"))

        # Белые иконки (из папки Image_White)
        self.icon_bookmarks_white = QIcon(self.resource_path("Image_White/Bookmarks_white.png"))
        self.icon_archive_white = QIcon(self.resource_path("Image_White/Archive_White.png"))
        self.icon_attach_white = QIcon(self.resource_path("Image_White/Attach_White.png"))
        self.icon_chat_white = QIcon(self.resource_path("Image_White/Chat_White.png"))
        self.icon_close_folder_white = QIcon(self.resource_path("Image_White/Close_folder_White.png"))
        self.icon_collapse_expand_white = QIcon(self.resource_path("Image_White/Collapse_expand_White.png"))
        self.icon_import_white = QIcon(self.resource_path("Image_White/Import_White.png"))
        self.icon_down_white = QIcon(self.resource_path("Image_White/Down_White.png"))
        self.icon_export_white = QIcon(self.resource_path("Image_White/Export_White.png"))
        self.icon_html_white = QIcon(self.resource_path("Image_White/Html_White.png"))
        self.icon_json_white = QIcon(self.resource_path("Image_White/Json_White.png"))
        self.icon_markdown_white = QIcon(self.resource_path("Image_White/Markdown_White.png"))
        self.icon_open_folder_white = QIcon(self.resource_path("Image_White/Open_folder_White.png"))
        self.icon_pdf_white = QIcon(self.resource_path("Image_White/Pdf_White.png"))
        self.icon_reflection_white = QIcon(self.resource_path("Image_White/Reflection_White.png"))
        self.icon_save_promt_white = QIcon(self.resource_path("Image_White/Save_promt_White.png"))
        self.icon_search_white = QIcon(self.resource_path("Image_White/Search_White.png"))
        self.icon_send_white = QIcon(self.resource_path("Image_White/Send_White.png"))
        self.icon_settings_3_white = QIcon(self.resource_path("Image_White/Settings_3_White.png"))
        self.icon_statistics_white = QIcon(self.resource_path("Image_White/Statistics_White.png"))
        self.icon_theme_white = QIcon(self.resource_path("Image_White/Theme_White.png"))
        self.icon_trash_white = QIcon(self.resource_path("Image_White/Trash_White.png"))
        self.icon_upward_d_white = QIcon(self.resource_path("Image_White/Upward_d_White.png"))
        self.icon_wallpaper_white = QIcon(self.resource_path("Image_White/Wallpaper_White.png"))
        self.icon_zip_white = QIcon(self.resource_path("Image_White/Zip_White.png"))
        self.icon_setting_chat_white = QIcon(self.resource_path("Image_White/Setting_chat_White.png"))
        self.icon_add_chat_white = QIcon(self.resource_path("Image_White/Add_chat_White.png"))
        self.icon_add_folder_white = QIcon(self.resource_path("Image_White/Add_folder_White.png"))
        self.icon_sort_white = QIcon(self.resource_path("Image_White/Sorting_White.png"))
        self.icon_clear_history_white = QIcon(self.resource_path("Image_White/Clear_history_White.png"))
        self.icon_repeat_request_white = QIcon(self.resource_path("Image_White/Repeat_equest_White.png"))
        self.icon_summary_dialogue_white = QIcon(self.resource_path("Image_White/Summary_dialogue_White.png"))
        self.icon_chat_parameters_white = QIcon(self.resource_path("Image_White/Chat_parameters_White.png"))
        self.icon_commands_white = QIcon(self.resource_path("Image_White/Commands_White.png"))

        self.icon_folder_closed = QIcon("Image/Close_folder.png")
        self.icon_folder_open = QIcon("Image/Open_folder.png")
        self.icon_chat = QIcon("Image/Chat.png")
        self.icon_folder_closed_white = QIcon("Image_White/Close_folder_White.png")
        self.icon_folder_open_white = QIcon("Image_White/Open_folder_White.png")
        self.icon_chat_white = QIcon("Image_White/Chat_White.png")

        # Иконка для кнопки переключения темы (обычная)
        self.icon_theme = QIcon(self.resource_path("Image/Theme.png"))

        self.icon_html = QIcon(self.resource_path("Image/Html.png"))
        self.show_reasoning = True
        self.trash_items = self.load_trash()
        self.prompts = self.load_prompts()

        self.lazy_load_limit = 30  # сколько сообщений подгружаем за раз
        self.loaded_count = {}  # filepath -> уже загружено
        self.load_more_widget = None  # ссылка на кнопку "Показать предыдущие"
        self.image_cache = {}
        self.setup_logging()
        self.rag_manager = RAGManager()
        self.rag_folder = "RAG_folder"  # имя папки
        self.rag_watcher = None

        self.worker: Optional[ChatWorker] = None
        self.thinking_cursor = None
        self.attachments = []  # временный список путей прикреплённых файлов
        self.message_positions = []  # список (start, end, msg_index) для текущего чата
        self._auto_scrolling = False
        self._skip_scroll_restore = False
        self.preview_texts = {}  # словарь для хранения текстов предпросмотра
        self.preview_counter = 0  # счётчик уникальных ID предпросмотра
        self.preview_images = {}  # словарь для хранения изображений предпросмотра
        self.preview_image_counter = 0  # счётчик уникальных ID изображений
        self.archived_chats = []  # список чатов в архиве (словари как в self.chats)
        self._pending_tail = None
        self._is_regenerating = False
        self._skip_scroll_once = False
        self.thinking_widget = None
        self.current_response_widget = None
        self.rag_manager = RAGManager()
        self.message_widgets = []  # список кортежей (filepath, msg_idx, widget)
        self.pending_scroll_target = None  # (filepath, msg_idx) или None
        self._restore_scroll_value = None
        self.deduplicate_trash()


        self.last_backup_hash = None

        self.unread_chats = set()  # для уведомлений
        self.original_chats_order = None  # для сброса сортировки
        self.is_generating = False
        self.generation_chat_index = -1
        self.generation_chat_filepath = None
        self.settings = {
            "api_base": "http://localhost:1234/v1",
            "model": "local-model",
            "max_tokens": 4096,
            "temperature": 0.7,
            "top_p": 0.9,
            "top_k": 40,
            "repeat_penalty": 1.1,
            "stream": True,
            "max_context_tokens": 4096,
            "system_prompt": "",
            "multimodal": False,
            "open_last_chat_on_startup": False,
            "auto_backup": True,
            "auto_backup_interval_min": 5,
            "backup_on_start": True,
            "max_backups": 5,
            "lm_studio_path": "",
            "lm_studio_auto_start": False,
            "encrypt_chats": False,
            "max_images_per_message": 4,
            "max_prompt_chars": 8000,  # Максимальная длина текста промпта в символах (включая текстовые вложения)
            "chat_font_family": "Arial",
            "chat_font_size": 10,
            "input_font_family": "Arial",
            "input_font_size": 10,
            "sound_notifications": True,
            "toast_notifications": True,
            "encryption_salt": "",
            "encryption_kdf_iterations": 200000,
            "max_attachment_size_mb": 20,
            "rag_threshold": 1.0,
            "sound_file": "Sound/Sound_1.MP3",
            "all_tags": [],
            "hotkeys": {
                "new_chat": "Ctrl+N",
                "new_folder": "Ctrl+Shift+N",
                "close_chat": "Ctrl+W",
                "search": "Ctrl+F",
                "import_deepseek": "Ctrl+I",
                "export": "Ctrl+E",
                "settings": "Ctrl+P",
                "send_message": "Ctrl+Return",
                "toggle_theme": "Ctrl+T",
                "toggle_sidebar": "Ctrl+B",
                "clear_chat": "Ctrl+Shift+L",
                "starred_messages": "Ctrl+Shift+B",
                "statistics": "Ctrl+Shift+S",
                "prompts": "Ctrl+Alt+P",
                "focus_input": "Ctrl+L",
                "next_chat": "Ctrl+Tab",
                "prev_chat": "Ctrl+Shift+Tab",
                "chat_settings": "Ctrl+Shift+P",
            },
            "theme": "light"
        }
        self.plugins = []
        self.load_plugins()
        self.config_file = "chat_config.json"
        self.load_settings()
        self._fernet = None
        self._session_unlocked = set()
        self.rag_enabled = self.settings.get("rag_enabled", False)
        if self.rag_enabled:
            self.ensure_rag_folder()
            self.load_rag_documents_from_folder()
            self.start_rag_watcher()

        self.chats_dir = "chats"
        os.makedirs(self.chats_dir, exist_ok=True)
        self.docs_dir = "Documentation"
        os.makedirs(self.docs_dir, exist_ok=True)
        self.thumbnails_dir = "attachments/thumbnails"
        os.makedirs(self.thumbnails_dir, exist_ok=True)
        self.folders_file = "folders.json"
        self.scroll_positions_file = "scroll_positions.json"
        self.folders = self.load_folders()  # список имён папок
        self.folders = [f if isinstance(f, dict) else {"name": str(f), "color": "#FFFFFF", "pinned": False} for f in
                        self.folders]

        self.chats = []  # список словарей: filepath, title, messages, folder
        self._restoring_scroll = False
        self.current_chat_index = -1  # индекс в self.chats текущего чата
        self.chat_scroll_positions = {}  # filepath -> значение verticalScrollBar()
        self._selecting_chat = False  # флаг для предотвращения повторного выбора чата
        self._restoring_cursor = False
        self.chat_scroll_positions = {}

        self.plugins = []
        self.load_plugins()

        self._backup_in_progress = False
        # Сопоставления для быстрого доступа к элементам дерева
        self.chat_items = {}  # filepath -> QTreeWidgetItem чата
        self.folder_items = {}  # имя папки -> QTreeWidgetItem папки

        self.init_ui()
        self.busy_overlay = BusyOverlay(self)
        self.load_chats_from_disk()
        self.apply_theme()
        self.update_system_theme()

        # Загружаем последний активный чат, если это разрешено
        last_chat_filepath = self.load_last_chat_filepath()
        if self.settings.get("open_last_chat_on_startup", False) and last_chat_filepath:
            if not self.select_chat_by_filepath(last_chat_filepath):
                self.current_chat_index = -1
                self.show_empty_chat_placeholder()
        else:
            self.current_chat_index = -1
            self.show_empty_chat_placeholder()

        # === ВОССТАНОВЛЕНИЕ ЧЕРНОВИКА ===
        self.restore_draft()  # добавили вызов восстановления черновика
        # Если шифрование включено или есть соль, запросим мастер-пароль
        if self.settings.get("encrypt_chats", False) and self.settings.get("encryption_salt"):
            try:
                self._get_fernet()  # запросит пароль и установит self._fernet
            except RuntimeError as e:
                QMessageBox.warning(self, "Ошибка", str(e))

        # Инициализация таймера автосохранения
        self.backup_timer = QTimer(self)
        interval_ms = self.settings.get("auto_backup_interval_min", 5) * 60 * 1000
        self.backup_timer.timeout.connect(self.create_backup)
        self.backup_timer.start(interval_ms)

        # Бэкап при старте, если включено
        if self.settings.get("backup_on_start", True):
            QTimer.singleShot(1000, self.create_backup)

        # Автозапуск LM Studio, если включено
        if self.settings.get("lm_studio_auto_start", False):
            QTimer.singleShot(1000, self.launch_lm_studio)

        self.refresh_unread_ui()
        # Показываем ошибки загрузки после старта (когда сплэш уже закрыт)
        if getattr(self, '_pending_errors', None):
            QTimer.singleShot(500, self._show_pending_load_errors)

        # Проверка обновлений через 5 секунд после старта (молча)
        QTimer.singleShot(5000, lambda: self.check_for_updates(silent=True))

        # Проверка LM Studio через 7 секунд
        QTimer.singleShot(7000, self._check_lm_studio_on_startup)

    def _show_pending_load_errors(self):
        errors = getattr(self, '_pending_errors', None)
        if not errors:
            return
        self._pending_errors = None
        msg = "\n".join(errors[:10])
        if len(errors) > 10:
            msg += f"\n…и ещё {len(errors) - 10}"
        QMessageBox.warning(
            self,
            "Часть чатов не загружена",
            f"Повреждённые файлы перемещены в папку 'corrupted':\n\n{msg}\n\n"
            "Ты можешь попробовать восстановить их из папки 'backups' вручную."
        )

    def refresh_unread_ui(self):
        """Обновляет заголовок окна и иконку в панели задач на основе количества непрочитанных чатов."""
        count = len(self.unread_chats)

        # === Заголовок окна ===
        base_title = "Echos"
        if count > 0:
            self.setWindowTitle(f"{base_title} ({count})")
        else:
            self.setWindowTitle(base_title)

        # === Иконка в панели задач (Windows) ===
        self.update_taskbar_icon()


    @contextmanager
    def busy(self, text="Загрузка…"):
        """Показывает оверлей на время выполнения блока."""
        self.busy_overlay.show_overlay(self, text)
        QApplication.processEvents()
        try:
            yield
        finally:
            self.busy_overlay.hide_overlay()


    def update_taskbar_icon(self):
        """Рисует иконку окна с числом непрочитанных (красный кружок в углу)."""
        from PySide6.QtGui import QPainter, QColor, QFont

        # Путь к базовой иконке
        icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Icon/Icon.ico")
        base_icon = QIcon(icon_path)
        if base_icon.isNull():
            # fallback — используем иконку окна как есть
            return

        size = 64
        pixmap = base_icon.pixmap(size, size)
        if pixmap.isNull():
            return

        count = len(self.unread_chats)
        if count == 0:
            # Без цифры
            self.setWindowIcon(base_icon)
            return

        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)

        # Красный кружок в правом верхнем углу
        badge_size = 26
        x = size - badge_size
        y = 0
        painter.setBrush(QColor("#d32f2f"))
        painter.setPen(Qt.NoPen)
        painter.drawEllipse(x, y, badge_size, badge_size)

        # Цифра внутри кружка
        painter.setPen(QColor("white"))
        font = QFont()
        font.setBold(True)
        font.setPixelSize(18)
        painter.setFont(font)
        text = str(count) if count < 100 else "99+"
        painter.drawText(x, y, badge_size, badge_size, Qt.AlignCenter, text)
        painter.end()

        self.setWindowIcon(QIcon(pixmap))

    def render_dot_to_pixmap(self, dot_code: str) -> QPixmap:
        """Рендерит DOT-код в изображение с помощью Graphviz."""
        if not GRAPHVIZ_AVAILABLE:
            return QPixmap()
        try:
            src = graphviz.Source(dot_code)
            # Рендерим во временный PNG-файл
            output_path = src.render(filename='temp_scheme', format='png', cleanup=True)
            pixmap = QPixmap(output_path)
            if not pixmap.isNull():
                return pixmap
            return QPixmap()
        except Exception as e:
            self.logger.warning(f"Ошибка рендеринга DOT: {e}")
            return QPixmap()

    def show_diagram_from_text(self, text: str, python_code: str = ""):
        """Строит схему: сначала пытается использовать Python-код, затем DOT, затем простые стрелки."""

        if not GRAPHVIZ_AVAILABLE:
            QMessageBox.warning(
                self, "Graphviz не установлен",
                "Для отображения схем установите Graphviz:\nhttps://graphviz.org/download/\n"
                "И добавьте его в системный PATH."
            )
            return

        # 1. Если передан Python-код — строим схему из него
        if python_code:
            dot_code = self.python_code_to_dot(python_code)
            if dot_code:
                pixmap = self.render_dot_to_pixmap(dot_code)
                if not pixmap.isNull():
                    dialog = SchemeViewerDialog(pixmap, self)
                    dialog.exec()
                    return

        # 2. Пытаемся извлечь DOT-блок из текста
        dot_blocks = re.findall(r'```dot\s+(.*?)```', text, re.DOTALL)
        if dot_blocks:
            dot_code = dot_blocks[0].strip()
            pixmap = self.render_dot_to_pixmap(dot_code)
            if not pixmap.isNull():
                dialog = SchemeViewerDialog(pixmap, self)
                dialog.exec()
                return

        # 3. Если DOT нет — пытаемся извлечь рёбра из простого текста и сформировать DOT
        edges = self._extract_edges_from_text(text)
        if edges:
            nodes = set()
            for src, dst in edges:
                nodes.add(src)
                nodes.add(dst)
            dot_lines = ['digraph G {']
            for node in nodes:
                shape = self._classify_shape(node)
                dot_lines.append(f'  "{node}" [shape={shape}]')
            for src, dst in edges:
                dot_lines.append(f'  "{src}" -> "{dst}"')
            dot_lines.append('}')
            dot_code = '\n'.join(dot_lines)
            pixmap = self.render_dot_to_pixmap(dot_code)
            if not pixmap.isNull():
                dialog = SchemeViewerDialog(pixmap, self)
                dialog.exec()
                return

        QMessageBox.information(self, "Схема", "Не удалось построить схему из данного сообщения.")

    def _extract_edges_from_text(self, text):
        """Извлекает рёбра из произвольного текста (поддерживает разные стрелки)."""

        edges = []
        # Варианты стрелок: ->, →, –>, —>, -->
        pattern = re.compile(r'(?P<src>\w+|\"[^\"]+\")\s*(?:->|→|–>|—>|-->)\s*(?P<dst>\w+|\"[^\"]+\")')
        for m in pattern.finditer(text):
            src = m.group('src').strip('"')
            dst = m.group('dst').strip('"')
            edges.append((src, dst))
        return edges

    def python_code_to_dot(self, code_text: str) -> str:
        """Преобразует Python-код в DOT-граф блок-схемы."""

        import ast
        import re

        def make_id():
            make_id.counter += 1
            return f"n{make_id.counter}"

        make_id.counter = 0

        nodes = []  # список кортежей (id, label, shape)
        edges = []  # список кортежей (src_id, dst_id)

        def add_node(label, shape):
            nid = make_id()
            nodes.append((nid, label, shape))
            return nid

        def add_edge(src, dst):
            if src is not None and dst is not None:
                edges.append((src, dst))

        def process_body(body, prev_id):
            """Обрабатывает список операторов. Возвращает последний активный узел."""

            current = prev_id
            for stmt in body:
                if isinstance(stmt, ast.Assign):
                    # Проверяем, есть ли справа вызов input()
                    if isinstance(stmt.value, ast.Call) and isinstance(stmt.value.func,
                                                                       ast.Name) and stmt.value.func.id == 'input':
                        prompt = ast.unparse(stmt.value.args[0]) if stmt.value.args else ""
                        target = ast.unparse(stmt.targets[0])
                        nid = add_node(f"Ввод {target}: {prompt}", "parallelogram")
                    else:
                        targets = ", ".join([ast.unparse(t) for t in stmt.targets])
                        value = ast.unparse(stmt.value)
                        nid = add_node(f"{targets} = {value}", "box")
                    add_edge(current, nid)
                    current = nid
                elif isinstance(stmt, ast.Expr) and isinstance(stmt.value, ast.Call):
                    func = stmt.value.func
                    if isinstance(func, ast.Name):
                        if func.id == 'input':
                            prompt = ast.unparse(stmt.value.args[0]) if stmt.value.args else ""
                            nid = add_node(f"Ввод {prompt}", "parallelogram")
                            add_edge(current, nid)
                            current = nid
                        elif func.id == 'print':
                            args = ", ".join([ast.unparse(a) for a in stmt.value.args])
                            nid = add_node(f"Вывод {args}", "parallelogram")
                            add_edge(current, nid)
                            current = nid
                        # Другие вызовы игнорируем
                elif isinstance(stmt, ast.If):
                    current = self._process_if(stmt, current, add_node, add_edge, process_body)
                elif isinstance(stmt, ast.While):
                    cond_id = add_node(f"Цикл: {ast.unparse(stmt.test)}", "diamond")
                    add_edge(current, cond_id)
                    body_end = process_body(stmt.body, cond_id)
                    add_edge(body_end, cond_id)  # обратная связь
                    # Точка выхода из цикла
                    exit_id = add_node("", "point")
                    add_edge(cond_id, exit_id)
                    current = exit_id
                elif isinstance(stmt, ast.For):
                    target = ast.unparse(stmt.target)
                    iter_ = ast.unparse(stmt.iter)
                    cond_id = add_node(f"Для {target} в {iter_}", "diamond")
                    add_edge(current, cond_id)
                    body_end = process_body(stmt.body, cond_id)
                    add_edge(body_end, cond_id)
                    exit_id = add_node("", "point")
                    add_edge(cond_id, exit_id)
                    current = exit_id
                # Остальные конструкции (return, break и т.д.) можно добавить
                elif isinstance(stmt, ast.Return):
                    nid = add_node(f"Возврат {ast.unparse(stmt.value) if stmt.value else ''}", "box")
                    add_edge(current, nid)
                    current = nid
                elif isinstance(stmt, ast.Break):
                    nid = add_node("Прервать", "box")
                    add_edge(current, nid)
                    current = nid
                elif isinstance(stmt, ast.Continue):
                    nid = add_node("Продолжить", "box")
                    add_edge(current, nid)
                    current = nid
                else:
                    # Необработанные конструкции пропускаем
                    pass
            return current

        try:
            tree = ast.parse(code_text)
        except SyntaxError:
            return ""

        start_id = add_node("Начало", "ellipse")
        end_id = add_node("Конец", "ellipse")
        last_id = process_body(tree.body, start_id)
        add_edge(last_id, end_id)

        dot_lines = ["digraph G {"]
        for nid, label, shape in nodes:
            label_escaped = label.replace('"', '\\"')
            if shape == "point":
                dot_lines.append(f'  "{nid}" [shape=point, width=0.1, height=0.1, label=""];')
            else:
                dot_lines.append(f'  "{nid}" [shape={shape}, label="{label_escaped}"];')
        for src, dst in edges:
            dot_lines.append(f'  "{src}" -> "{dst}";')
        dot_lines.append("}")
        return "\n".join(dot_lines)

    def _process_if(self, if_stmt, prev_id, add_node, add_edge, process_body):
        """Обрабатывает if/elif/else, возвращает точку слияния."""

        # Создаём узел условия
        cond_id = add_node(f"Если {ast.unparse(if_stmt.test)}", "diamond")
        add_edge(prev_id, cond_id)

        # Обрабатываем основную ветку
        true_end = process_body(if_stmt.body, cond_id)

        # Если есть else/elif
        if if_stmt.orelse:
            # Если это elif
            if len(if_stmt.orelse) == 1 and isinstance(if_stmt.orelse[0], ast.If):
                # Рекурсивно обрабатываем elif, начиная с того же условия (cond_id)
                # но cond_id уже использован, нужно создать новую точку слияния позже
                elif_end = self._process_if(if_stmt.orelse[0], cond_id, add_node, add_edge, process_body)
                # Создаём точку слияния
                merge_id = add_node("", "point")
                add_edge(true_end, merge_id)
                add_edge(elif_end, merge_id)
                return merge_id
            else:
                # Обычный else
                false_end = process_body(if_stmt.orelse, cond_id)
                merge_id = add_node("", "point")
                add_edge(true_end, merge_id)
                add_edge(false_end, merge_id)
                return merge_id
        else:
            # Нет else – точка слияния всё равно нужна для выхода
            merge_id = add_node("", "point")
            add_edge(true_end, merge_id)
            return merge_id


    def extract_python_code_from_message(self, content: str) -> str:
        """Извлекает Python-код из блока ```python ... ``` или возвращает весь текст, если это код."""

        # Ищем блок python
        blocks = re.findall(r'```python\s+(.*?)```', content, re.DOTALL)
        if blocks:
            return blocks[0]
        # Если блока нет, но текст выглядит как код (содержит def, import, print и т.п.)
        # Простая эвристика: много строк с отступами и ключевыми словами
        lines = content.strip().splitlines()
        if len(lines) > 1 and any(
                line.strip().startswith(('import ', 'from ', 'def ', 'if ', 'for ', 'while ', 'print(', 'input('))
                for line in lines
        ):
            return content
        return ""

    def _classify_shape(self, node_text):
        """Определяет форму узла по его содержимому."""
        t = node_text.lower()
        if any(word in t for word in ['начало', 'конец', 'start', 'end', 'начать', 'завершить']):
            return 'ellipse'
        if any(word in t for word in ['ввод', 'вывод', 'input', 'output', 'вход', 'выход']):
            return 'parallelogram'
        if any(word in t for word in ['если', 'if', 'условие', '?', 'condition', 'проверка']):
            return 'diamond'
        # По умолчанию – прямоугольник
        return 'box'

    def setup_logging(self):
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
        log_file = os.path.join(log_dir, "app.log")

        import logging
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )
        self.logger = logging.getLogger("ChatApp")

        # Логирование необработанных исключений
        def handle_exception(exc_type, exc_value, exc_traceback):
            self.logger.error("Необработанное исключение", exc_info=(exc_type, exc_value, exc_traceback))
            try:
                QMessageBox.critical(
                    None,
                    "Непредвиденная ошибка",
                    f"Произошла ошибка:\n\n{exc_type.__name__}: {exc_value}\n\n"
                    "Подробности в logs/app.log"
                )
            except Exception:
                pass


        sys.excepthook = handle_exception

    def load_settings(self):
        if os.path.exists(self.config_file):
            data = self._safe_json_load(self.config_file, None)
            if isinstance(data, dict):
                self.settings.update(data)
            else:
                self.logger.warning("chat_config.json имеет неверный формат, использую настройки по умолчанию.")
                print("⚠️ chat_config.json имеет неверный формат, использую настройки по умолчанию.")
        else:
            self.logger.warning("chat_config.json имеет неверный формат, использую настройки по умолчанию.")

        if "encryption_key" in self.settings:
            del self.settings["encryption_key"]
            self.save_settings()

        # Гарантируем наличие всех ключей горячих клавиш (даже если файл старый)
        default_hotkeys = {
            "new_chat": "Ctrl+N",
            "new_folder": "Ctrl+Shift+N",
            "close_chat": "Ctrl+W",
            "search": "Ctrl+F",
            "import_deepseek": "Ctrl+I",
            "export": "Ctrl+E",
            "settings": "Ctrl+P",
            "send_message": "Ctrl+Return",
            "toggle_theme": "Ctrl+T",
            "toggle_sidebar": "Ctrl+B",
            "clear_chat": "Ctrl+Shift+L",
            "starred_messages": "Ctrl+Shift+B",
            "statistics": "Ctrl+Shift+S",
            "prompts": "Ctrl+Alt+P",
            "focus_input": "Ctrl+L",
            "next_chat": "Ctrl+Tab",
            "prev_chat": "Ctrl+Shift+Tab",
            "chat_settings": "Ctrl+Shift+P",
        }
        if "hotkeys" not in self.settings:
            self.settings["hotkeys"] = {}
        for key, value in default_hotkeys.items():
            if key not in self.settings["hotkeys"]:
                self.settings["hotkeys"][key] = value

    def save_settings(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
            # Уведомляем плагины об изменении настроек
            api = PluginAPI(self)
            for plugin in self.plugins:
                try:
                    if hasattr(plugin, 'on_settings_changed'):
                        plugin.on_settings_changed(self.settings, api)
                except Exception as e:
                    self.logger.warning(f"Ошибка в on_settings_changed: {e}")
        except Exception as e:
            self.logger.error(f"Ошибка сохранения промптов: {e}")
            self.show_error("Ошибка сохранения", f"Не удалось сохранить промпты.\nПричина: {e}")

    def load_folders(self) -> List[Dict[str, Any]]:
        if os.path.exists(self.folders_file):
            try:
                with open(self.folders_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    return [{"name": str(name), "color": "#FFFFFF", "pinned": False} for name in data]
                elif isinstance(data, dict) and 'folders' in data:
                    result = []
                    for folder in data['folders']:
                        if isinstance(folder, dict):
                            result.append(folder)
                        else:
                            result.append({"name": str(folder), "color": "#FFFFFF", "pinned": False})
                    return result
            except (json.JSONDecodeError, UnicodeDecodeError) as e:
                self.logger.warning(f"Ошибка загрузки папок: {e}. Создаю пустой список.")
                print(f"⚠️ Ошибка загрузки папок: {e}. Создаю пустой список.")
            except Exception as e:
                self.logger.warning(f"Неизвестная ошибка при чтении папок: {e}. Создаю пустой список.")
                print(f"⚠️ Неизвестная ошибка при чтении папок: {e}. Создаю пустой список.")
        return []

    def save_folders(self):
        try:
            with open(self.folders_file, 'w', encoding='utf-8') as f:
                json.dump({"folders": self.folders}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.error(f"Ошибка сохранения папок: {e}")
            self.show_error("Ошибка сохранения", f"Не удалось сохранить папки.\nПричина: {e}")

    def load_trash(self):
        if os.path.exists('trash.json'):
            data = self._safe_json_load('trash.json', [])
            if isinstance(data, list):
                self.trash_items = data
                self.deduplicate_trash()
                return self.trash_items
            else:
                print("⚠️ trash.json повреждён, корзина пуста.")
                return []
        return []

    def save_trash(self):
        self.deduplicate_trash()
        try:
            with open('trash.json', 'w', encoding='utf-8') as f:
                json.dump(self.trash_items, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.error(f"Ошибка сохранения корзины: {e}")
            self.show_error("Ошибка сохранения", f"Не удалось сохранить корзину.\nПричина: {e}")

    def load_prompts(self):
        if os.path.exists('prompts.json'):
            data = self._safe_json_load('prompts.json', [])
            if isinstance(data, list):
                return data
            else:
                self.logger.warning("prompts.json повреждён, список промптов пуст.")
                print("⚠️ prompts.json повреждён, список промптов пуст.")
                return []
        return []

    def save_prompts(self):
        """Сохраняет промпты в файл prompts.json."""

        try:
            with open('prompts.json', 'w', encoding='utf-8') as f:
                json.dump(self.prompts, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.error(f"Ошибка сохранения промптов: {e}")
            self.show_error("Ошибка сохранения", f"Не удалось сохранить промпты.\nПричина: {e}")

    def load_last_chat_filepath(self):
        if os.path.exists("last_session.json"):
            data = self._safe_json_load("last_session.json", None)
            if isinstance(data, dict):
                return data.get("last_chat")
        return None

    def save_last_chat_filepath(self):
        """Сохраняет путь к последнему активному чату."""

        if self.current_chat_index >= 0 and self.current_chat_index < len(self.chats):
            filepath = self.chats[self.current_chat_index]['filepath']
            try:
                with open("last_session.json", 'w', encoding='utf-8') as f:
                    json.dump({"last_chat": filepath}, f)
            except Exception as e:
                self.logger.warning(f"Не удалось сохранить last_session: {e}")

    def load_scroll_positions(self) -> dict:
        if os.path.exists(self.scroll_positions_file):
            try:
                with open(self.scroll_positions_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception as e:
                self.logger.warning(f"Не удалось загрузить позиции скролла: {e}")
        return {}

    def save_scroll_positions(self):
        try:
            with open(self.scroll_positions_file, 'w', encoding='utf-8') as f:
                json.dump(self.chat_scroll_positions, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.warning(f"Не удалось сохранить позиции скролла: {e}")

    def load_cursor_positions(self):
        if os.path.exists("cursor_positions.json"):
            try:
                with open("cursor_positions.json", 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
            except Exception as e:
                self.logger.warning(f"Не удалось загрузить позиции курсора: {e}")
        return {}

    def save_cursor_positions(self):
        try:
            with open("cursor_positions.json", 'w', encoding='utf-8') as f:
                json.dump(self.chat_cursor_positions, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.warning(f"Не удалось сохранить позиции курсора: {e}")

    def resource_path(self, relative_path):
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

    def _safe_json_load(self, filepath, default):
        """Безопасно загружает JSON. Если файл повреждён, возвращает default."""

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError, FileNotFoundError) as e:
            print(f"⚠️ Ошибка загрузки {filepath}: {e}. Использую значение по умолчанию.")
            return default
        except Exception as e:
            print(f"⚠️ Неизвестная ошибка при чтении {filepath}: {e}. Использую значение по умолчанию.")
            return default

    def deduplicate_trash(self):
        """Удаляет дубликаты записей в корзине по ключу (filepath для чатов, name для папок)."""

        seen_chat = set()
        seen_folder = set()
        unique = []
        for entry in self.trash_items:
            if entry.get('type') == 'chat':
                key = entry.get('filepath')
                if key and key not in seen_chat:
                    seen_chat.add(key)
                    unique.append(entry)
            elif entry.get('type') == 'folder':
                key = entry.get('name')
                if key and key not in seen_folder:
                    seen_folder.add(key)
                    unique.append(entry)
            else:
                # Неизвестный тип – оставляем как есть
                unique.append(entry)
        self.trash_items = unique

    def _hash_password(self, password: str) -> str:
        return hashlib.sha256(password.encode('utf-8')).hexdigest()

    def format_chat_title(self, chat):
        return chat.get('title', '')

    def update_chat_item_text(self, chat):
        filepath = chat['filepath']
        item = self.chat_items.get(filepath)
        if not item:
            return

        # Колонка 0: название с префиксами
        title = chat.get('title', '')
        prefix = ""
        if chat.get('incognito', False):
            prefix += "🕶 "
        if filepath in self.unread_chats:
            prefix += "🔴 "
        item.setText(0, prefix + title)

        # Колонка 2: теги (синим цветом)
        tags = chat.get('tags', [])
        if tags:
            tags_str = ' '.join(['#' + tag for tag in tags])
            item.setText(2, tags_str)
            item.setForeground(2, QBrush(QColor("#1a73e8")))
        else:
            item.setText(2, "")

    def edit_chat_tags(self, filepath):
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        chat = self.chats[idx]
        dialog = TagsDialog(self.settings.get("all_tags", []), chat.get('tags', []), self)
        result = dialog.exec()
        # Всегда сохраняем глобальный список тегов (в нём могут быть теги, не привязанные к чатам)
        self.settings["all_tags"] = dialog.get_all_tags()
        self.save_settings()
        # Применяем изменения к чату только если окно закрыто через "Выбрать" или "Сбросить"
        if result == QDialog.Accepted:
            chat['tags'] = list(dialog.get_current_tags())
            self.save_chat_by_filepath(filepath)
            self.update_chat_item_text(chat)
            self.statusBar().showMessage("Теги обновлены")

    def remove_chat_tag(self, filepath, tag):
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        chat = self.chats[idx]
        tags = chat.get('tags', [])
        if tag in tags:
            tags.remove(tag)
            chat['tags'] = tags
            self.save_chat_by_filepath(filepath)
            self.update_chat_item_text(chat)
            self.statusBar().showMessage(f"Тег '#{tag}' удалён")

    def save_draft(self, text: str):
        """Сохраняет черновик в файл draft.txt."""

        try:
            with open("draft.txt", "w", encoding="utf-8") as f:
                f.write(text)
        except Exception as e:
            print(f"Не удалось сохранить черновик: {e}")

    def load_draft(self) -> str:
        """Загружает черновик из файла draft.txt. Если файла нет, возвращает пустую строку."""

        if os.path.exists("draft.txt"):
            try:
                with open("draft.txt", "r", encoding="utf-8") as f:
                    return f.read()
            except Exception as e:
                print(f"Не удалось прочитать черновик: {e}")
        return ""

    def restore_draft(self):
        """Восстанавливает черновик в поле ввода."""

        draft = self.load_draft()
        if draft.strip():
            self.input_edit.setPlainText(draft)
            self.statusBar().showMessage("Восстановлен неотправленный черновик")

    # ======================================================================
    #  ИНТЕРФЕЙС И ИНИЦИАЛИЗАЦИЯ UI
    # ======================================================================

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Используем QSplitter для возможности изменения ширины панелей
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)

        # Левая панель (боковая)
        left_panel = QWidget()
        left_panel.setMinimumWidth(150)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(5, 5, 5, 5)
        left_layout.setSpacing(5)

        # Кнопки сверху левой панели (круглые, с иконками)
        btn_panel = QHBoxLayout()

        # Кнопка "Создать чат"
        self.btn_new_chat = QPushButton()
        self.btn_new_chat.setIcon(self.icon_add_chat)
        self.btn_new_chat.setIconSize(QSize(24, 24))
        self.btn_new_chat.setFixedSize(40, 40)
        self.btn_new_chat.setToolTip("Создать чат")
        self.btn_new_chat.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_new_chat.clicked.connect(self.create_new_chat)

        # Кнопка "Создать папку"
        self.btn_new_folder = QPushButton()
        self.btn_new_folder.setIcon(self.icon_add_folder)
        self.btn_new_folder.setIconSize(QSize(24, 24))
        self.btn_new_folder.setFixedSize(40, 40)
        self.btn_new_folder.setToolTip("Создать папку")
        self.btn_new_folder.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_new_folder.clicked.connect(self.create_new_folder)

        # Кнопка "Сортировка"
        self.btn_sort = QPushButton()
        self.btn_sort.setIcon(self.icon_sort)
        self.btn_sort.setIconSize(QSize(24, 24))
        self.btn_sort.setFixedSize(40, 40)
        self.btn_sort.setToolTip("Сортировка")
        self.btn_sort.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_sort.clicked.connect(self.show_sort_menu)

        btn_panel.addWidget(self.btn_new_chat)
        btn_panel.addWidget(self.btn_new_folder)
        btn_panel.addWidget(self.btn_sort)
        btn_panel.addStretch()
        left_layout.addLayout(btn_panel)

        # Дерево чатов
        self.chat_tree = ChatTreeWidget(self)
        self.chat_tree.setHeaderHidden(True)
        self.chat_tree.setDragDropMode(QTreeWidget.InternalMove)
        self.chat_tree.setDefaultDropAction(Qt.MoveAction)
        self.chat_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.chat_tree.customContextMenuRequested.connect(self.show_tree_context_menu)
        self.chat_tree.viewport().setContextMenuPolicy(Qt.CustomContextMenu)
        self.chat_tree.viewport().customContextMenuRequested.connect(self.show_tree_context_menu)
        self.chat_tree.itemSelectionChanged.connect(self.on_chat_selected)
        left_layout.addWidget(self.chat_tree)

        self.chat_tree.itemExpanded.connect(self.on_item_expanded)
        self.chat_tree.itemCollapsed.connect(self.on_item_collapsed)

        # Устанавливаем две колонки для иконок булавки
        self.chat_tree.setColumnCount(3)
        self.chat_tree.header().setStretchLastSection(False)
        self.chat_tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self.chat_tree.header().setSectionResizeMode(1, QHeaderView.Fixed)
        self.chat_tree.header().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.chat_tree.setColumnWidth(1, 30)

        # Правая часть: верхняя панель (импорт/экспорт/настройки), чат, ввод
        right_panel = QWidget()
        right_panel.setMinimumWidth(400)
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        # Верхняя панель правой части
        top_panel = QHBoxLayout()
        top_panel.setContentsMargins(5, 5, 5, 5)
        top_panel.setSpacing(5)

        # Кнопка скрытия/показа боковой панели (круглая с иконкой)
        self.btn_toggle_sidebar = QPushButton()
        self.btn_toggle_sidebar.setIcon(self.icon_collapse_expand)
        self.btn_toggle_sidebar.setIconSize(QSize(24, 24))
        self.btn_toggle_sidebar.setFixedSize(40, 40)
        self.btn_toggle_sidebar.setToolTip("Скрыть/показать боковую панель")
        self.btn_toggle_sidebar.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_toggle_sidebar.clicked.connect(self.toggle_sidebar)
        top_panel.addWidget(self.btn_toggle_sidebar)

        # Кнопка глобального поиска
        self.btn_global_search = QPushButton()
        self.btn_global_search.setIcon(self.icon_search)
        self.btn_global_search.setIconSize(QSize(20, 20))
        self.btn_global_search.setFixedSize(40, 40)
        self.btn_global_search.setToolTip("Поиск по всем чатам")
        self.btn_global_search.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_global_search.clicked.connect(self.open_global_search)

        top_panel.addWidget(self.btn_global_search)

        top_panel.addStretch()

        # Кнопка закладок
        self.btn_starred = QPushButton()
        self.btn_starred.setIcon(self.icon_bookmarks)
        self.btn_starred.setIconSize(QSize(20, 20))
        self.btn_starred.setFixedSize(40, 40)
        self.btn_starred.setToolTip("Закладки")
        self.btn_starred.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_starred.clicked.connect(self.show_starred_messages)
        top_panel.addWidget(self.btn_starred)

        # Кнопка импорта DeepSeek JSON (круглая)
        self.btn_import_ds = QPushButton()
        self.btn_import_ds.setIcon(self.icon_import)
        self.btn_import_ds.setIconSize(QSize(24, 24))
        self.btn_import_ds.setFixedSize(40, 40)
        self.btn_import_ds.setToolTip("Импорт DeepSeek JSON")
        self.btn_import_ds.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_import_ds.clicked.connect(self.import_deepseek_history)
        top_panel.addWidget(self.btn_import_ds)

        # Кнопка экспорта чата (круглая)
        self.btn_export = QPushButton()
        self.btn_export.setIcon(self.icon_export)
        self.btn_export.setIconSize(QSize(24, 24))
        self.btn_export.setFixedSize(40, 40)
        self.btn_export.setToolTip("Экспорт чата")
        self.btn_export.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_export.clicked.connect(self.show_export_menu)
        top_panel.addWidget(self.btn_export)

        # Кнопка настроек (круглая)
        self.btn_settings = QPushButton()
        self.btn_settings.setIcon(self.icon_settings)
        self.btn_settings.setIconSize(QSize(24, 24))
        self.btn_settings.setFixedSize(40, 40)
        self.btn_settings.setToolTip("Настройки")
        self.btn_settings.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_settings.clicked.connect(self.open_settings_dialog)
        top_panel.addWidget(self.btn_settings)

        # Кнопка архива (круглая)
        self.btn_archive = QPushButton()
        self.btn_archive.setIcon(self.icon_archive)
        self.btn_archive.setIconSize(QSize(24, 24))
        self.btn_archive.setFixedSize(40, 40)
        self.btn_archive.setToolTip("Архив чатов")
        self.btn_archive.setStyleSheet("""
                    QPushButton {
                        border-radius: 20px;
                        background-color: #f0f0f0;
                        border: 1px solid #ccc;
                    }
                    QPushButton:hover { background-color: #e0e0e0; }
                    QPushButton:pressed { background-color: #d0d0d0; }
                """)
        self.btn_archive.clicked.connect(self.show_archive_dialog)
        top_panel.addWidget(self.btn_archive)

        # Кнопка корзины
        self.btn_trash = QPushButton()
        self.btn_trash.setIcon(self.icon_trash)
        self.btn_trash.setFixedSize(40, 40)
        self.btn_trash.setToolTip("Последнее удаление")
        self.btn_trash.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_trash.clicked.connect(self.show_trash_dialog)
        top_panel.addWidget(self.btn_trash)

        right_layout.addLayout(top_panel)

        # Область чата: QScrollArea с layout для сообщений
        self.chat_scroll_area = QScrollArea()
        self._stick_to_bottom = True
        self.chat_scroll_area.verticalScrollBar().rangeChanged.connect(self._on_scroll_range_changed)
        self.chat_scroll_area.setWidgetResizable(True)
        self.chat_scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.chat_scroll_area.setStyleSheet("QScrollArea { background: transparent; }")
        self.chat_scroll_area.viewport().setStyleSheet("background: transparent;")

        self.chat_messages_widget = QWidget()
        self.chat_messages_widget.installEventFilter(self)
        self.chat_messages_widget.setStyleSheet("background: transparent;")
        self.chat_messages_layout = QVBoxLayout(self.chat_messages_widget)
        self.chat_messages_layout.setContentsMargins(8, 8, 8, 8)
        self.chat_messages_layout.setSpacing(6)
        self.chat_messages_layout.setAlignment(Qt.AlignTop)
        self.chat_scroll_area.setWidget(self.chat_messages_widget)
        self.chat_scroll_area.viewport().installEventFilter(self)

        # Включаем приём Drag&Drop для области чата
        self.chat_scroll_area.setAcceptDrops(True)
        self.chat_messages_widget.setAcceptDrops(True)
        self.chat_scroll_area.viewport().setAcceptDrops(True)

        # Устанавливаем фильтр событий на viewport и на сам chat_messages_widget
        self.chat_scroll_area.viewport().installEventFilter(self)
        self.chat_messages_widget.installEventFilter(self)
        self.installEventFilter(self)
        # Алиас для совместимости
        self.chat_browser = self.chat_scroll_area

        # ---- Панель поиска ----
        self.search_panel = QWidget(self.chat_browser)
        self.search_panel.setFixedHeight(40)
        search_layout = QHBoxLayout(self.search_panel)
        search_layout.setContentsMargins(5, 5, 5, 5)
        search_layout.setSpacing(5)

        self.search_counter_label = QLabel("")
        self.search_counter_label.setFixedHeight(24)
        self.search_counter_label.setMinimumWidth(70)
        self.search_counter_label.setAlignment(Qt.AlignCenter)
        self.search_counter_label.setStyleSheet("""
            QLabel {
                background-color: rgba(255, 255, 255, 0.7);
                border: 1px solid #ccc;
                border-radius: 8px;
                padding: 2px 6px;
                color: #333;
                font-weight: bold;
            }
        """)
        search_layout.addWidget(self.search_counter_label)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Поиск в чате...")
        self.search_edit.returnPressed.connect(self.search_next)
        self.search_edit.textChanged.connect(self._on_search_text_changed)
        search_layout.addWidget(self.search_edit)
        # Чекбокс регулярных выражений
        self.check_regex = QCheckBox("Regex")
        self.check_regex.setToolTip("Использовать регулярные выражения")
        self.check_regex.toggled.connect(self.update_search_counter)
        search_layout.addWidget(self.check_regex)

        self.btn_search_prev = QPushButton()
        self.btn_search_prev.setIcon(QIcon(self.resource_path("Image/Upward_d.png")))
        self.btn_search_prev.setFixedSize(24, 24)
        self.btn_search_prev.clicked.connect(self.search_prev)
        search_layout.addWidget(self.btn_search_prev)

        self.btn_search_next = QPushButton()
        self.btn_search_next.setIcon(QIcon(self.resource_path("Image/Downward_d.png")))
        self.btn_search_next.setFixedSize(24, 24)
        self.btn_search_next.clicked.connect(self.search_next)
        search_layout.addWidget(self.btn_search_next)

        self.btn_search_close = QPushButton()
        self.btn_search_close.setIcon(QIcon(self.resource_path("Image/Close.png")))
        self.btn_search_close.setFixedSize(24, 24)
        self.btn_search_close.clicked.connect(self.close_search_panel)
        search_layout.addWidget(self.btn_search_close)

        search_btn_style = """
            QPushButton {
                background-color: rgba(255, 255, 255, 0.7);
                border: 1px solid #ccc;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: rgba(255, 255, 255, 0.9); }
            QPushButton:pressed { background-color: rgba(255, 255, 255, 0.5); }
        """
        self.btn_search_prev.setStyleSheet(search_btn_style)
        self.btn_search_next.setStyleSheet(search_btn_style)
        self.btn_search_close.setStyleSheet(search_btn_style)

        self.btn_search_close.setStyleSheet("""
            QPushButton {
                background-color: rgba(220, 80, 80, 0.7);
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: rgba(220, 80, 80, 0.9); }
            QPushButton:pressed { background-color: rgba(220, 80, 80, 0.5); }
        """)

        self.search_panel.setVisible(False)
        self.update_search_panel_position()

        # ---- Кнопка "Вниз" ----
        self.btn_scroll_down = QPushButton(self.chat_browser)
        self.btn_scroll_down.setIcon(self.icon_down)
        self.btn_scroll_down.setIconSize(QSize(24, 24))
        self.btn_scroll_down.setFixedSize(40, 40)
        self.btn_scroll_down.setCursor(Qt.PointingHandCursor)
        self.btn_scroll_down.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: rgba(0, 0, 0, 0.6);
                border: none;
            }
            QPushButton:hover { background-color: rgba(0, 0, 0, 0.8); }
        """)
        self.btn_scroll_down.setVisible(False)
        self.btn_scroll_down.clicked.connect(self.scroll_chat_to_bottom)
        self.chat_browser.scroll_down_button = self.btn_scroll_down
        self.btn_scroll_down.update_position = lambda: self.position_scroll_down_button()
        self.chat_browser.verticalScrollBar().valueChanged.connect(self.on_chat_scrolled)
        self.position_scroll_down_button()
        self.btn_scroll_down.raise_()

        # Панель закреплённых сообщений
        self.pinned_panel = QWidget()
        self.pinned_panel.setVisible(False)
        self.pinned_panel.setStyleSheet("""
            QWidget {
                background-color: #fdf6e3;
                border: 1px solid #e0cfa0;
                border-radius: 4px;
            }
        """)
        self.pinned_layout = QHBoxLayout(self.pinned_panel)
        self.pinned_layout.setContentsMargins(8, 4, 8, 4)
        self.pinned_layout.setSpacing(6)
        right_layout.addWidget(self.pinned_panel)

        # Область чата
        right_layout.addWidget(self.chat_browser, stretch=1)

        # ---- Область прикреплённых файлов ----
        self.attachments_layout = QHBoxLayout()
        self.attachments_layout.setSpacing(5)
        self.attachments_layout.setContentsMargins(0, 0, 0, 0)
        self.attachments_widget = QWidget()
        self.attachments_widget.setLayout(self.attachments_layout)

        self.attachments_scroll = QScrollArea()
        self.attachments_scroll.setWidgetResizable(True)
        self.attachments_scroll.setFixedHeight(36)
        self.attachments_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.attachments_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.attachments_scroll.setWidget(self.attachments_widget)
        self.attachments_scroll.setVisible(False)
        right_layout.addWidget(self.attachments_scroll)

        # ---- Панель ввода ----
        input_panel = QHBoxLayout()
        input_panel.setContentsMargins(5, 5, 5, 5)
        input_panel.setSpacing(5)

        self.input_edit = PlainPasteTextEdit()
        self.input_edit.main_window = self
        self.highlighter = CodeHighlighter(self.input_edit.document())

        # Метка количества токенов
        # Метка количества токенов (в статус-баре)
        self.token_count_label = QLabel("0 токенов")
        self.token_count_label.setStyleSheet("color: #888;")
        self.statusBar().addPermanentWidget(self.token_count_label)

        self.input_edit.textChanged.connect(self.update_token_count)

        # Кнопка прикрепления
        self.btn_attach = QPushButton()
        self.btn_attach.setIcon(self.icon_attach)
        self.btn_attach.setIconSize(QSize(20, 20))
        self.btn_attach.setFixedSize(40, 40)
        self.btn_attach.setToolTip("Прикрепить файлы")
        self.btn_attach.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_attach.clicked.connect(self.attach_files)

        # Кнопка отправки
        self.btn_send = QPushButton()
        self.btn_send.setIcon(self.icon_send)
        self.btn_send.setIconSize(QSize(20, 20))
        self.btn_send.setFixedSize(40, 40)
        self.btn_send.setToolTip("Отправить")
        self.btn_send.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #D8A7B1;
                border: none;
            }
            QPushButton:hover { background-color: #C08A95; }
            QPushButton:pressed { background-color: #A87480; }
        """)
        self.btn_send.clicked.connect(self.send_message)

        # Кнопка переключения размышлений
        self.btn_toggle_reasoning = QPushButton()
        self.btn_toggle_reasoning.setIcon(self.icon_reflection)
        self.btn_toggle_reasoning.setIconSize(QSize(20, 20))
        self.btn_toggle_reasoning.setFixedSize(40, 40)
        self.btn_toggle_reasoning.setCursor(Qt.PointingHandCursor)
        self.btn_toggle_reasoning.setCheckable(True)
        self.btn_toggle_reasoning.setChecked(True)
        self.btn_toggle_reasoning.setToolTip("Показывать размышления модели")
        self.btn_toggle_reasoning.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #d0d0d0;
                border: none;
            }
            QPushButton:hover { background-color: #c0c0c0; }
            QPushButton:pressed { background-color: #b0b0b0; }
            QPushButton:checked {
                background-color: #FFD54F;
            }
            QPushButton:checked:hover { background-color: #FFCA28; }
        """)
        self.btn_toggle_reasoning.clicked.connect(self.toggle_reasoning)

        button_column = QVBoxLayout()
        button_column.setSpacing(4)
        button_column.addWidget(self.btn_attach)
        button_column.addWidget(self.btn_send)
        button_column.addWidget(self.btn_toggle_reasoning)

        # Добавляем поле ввода, метку токенов и вертикальную колонку в горизонтальный layout
        input_panel.addWidget(self.input_edit)
        input_panel.addLayout(button_column)

        right_layout.addLayout(input_panel)

        # Сборка основного сплиттера
        self.main_splitter.addWidget(left_panel)
        self.main_splitter.addWidget(right_panel)
        self.main_splitter.setStretchFactor(0, 0)
        self.main_splitter.setStretchFactor(1, 1)
        self.main_splitter.setSizes([250, 950])

        # Подключаем сигнал перемещения разделителя для сохранения размеров
        self.main_splitter.splitterMoved.connect(self.on_splitter_moved)

        # Восстанавливаем сохранённые размеры, если они есть
        if "splitter_sizes" in self.settings:
            saved = self.settings["splitter_sizes"]
            if isinstance(saved, list) and len(saved) == 2:
                self.main_splitter.setSizes(saved)

        central_layout = QVBoxLayout(central_widget)
        central_layout.setContentsMargins(0, 0, 0, 0)
        central_layout.addWidget(self.main_splitter)
        self.apply_font_settings()
        self.statusBar().showMessage("Готов")

        # Метка модели в статус-баре
        self.model_label = QLabel("")
        self.statusBar().addPermanentWidget(self.model_label)

        self.speed_label = QLabel("")
        self.statusBar().insertPermanentWidget(0, self.speed_label)

        self.speed_timer = QTimer(self)
        self.speed_timer.setInterval(500)
        self.speed_timer.timeout.connect(self.update_generation_speed)

        # Кнопка остановки генерации
        self.btn_stop_generation = QPushButton(self.chat_browser)
        self.btn_stop_generation.setIcon(self.icon_stop)
        self.btn_stop_generation.setIconSize(QSize(24, 24))
        self.btn_stop_generation.setFixedSize(40, 40)
        self.btn_stop_generation.setCursor(Qt.PointingHandCursor)
        self.btn_stop_generation.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: rgba(180, 0, 0, 0.7);
                border: none;
            }
            QPushButton:hover { background-color: rgba(220, 0, 0, 0.85); }
        """)
        self.btn_stop_generation.setVisible(False)
        self.btn_stop_generation.clicked.connect(self.stop_generation)
        self.chat_browser.stop_button = self.btn_stop_generation
        self.btn_stop_generation.update_position = lambda: self.position_stop_button()
        self.position_stop_button()
        self.btn_stop_generation.raise_()

        # Кнопка настройки чата
        self.btn_chat_settings = QPushButton(self.chat_browser)
        self.btn_chat_settings.setIcon(self.icon_setting_chat)
        self.btn_chat_settings.setIconSize(QSize(20, 20))
        self.btn_chat_settings.setFixedSize(36, 36)
        self.btn_chat_settings.setCursor(Qt.PointingHandCursor)
        self.btn_chat_settings.setStyleSheet("""
            QPushButton {
                border-radius: 18px;
                background-color: rgba(0, 0, 0, 0.4);
                border: none;
            }
            QPushButton:hover { background-color: rgba(0, 0, 0, 0.6); }
        """)
        self.btn_chat_settings.setVisible(True)
        self.btn_chat_settings.clicked.connect(self.show_chat_settings_menu)
        self.chat_browser.btn_chat_settings = self.btn_chat_settings
        self.btn_chat_settings.update_position = lambda: self.position_chat_settings_button()
        self.position_chat_settings_button()
        self.btn_chat_settings.raise_()

        # Глобальный перехват событий
        self.installEventFilter(self)

        # Кнопка переключения темы
        self.btn_theme = QPushButton()
        self.btn_theme.setIcon(self.icon_theme)
        self.btn_theme.setIconSize(QSize(24, 24))
        self.btn_theme.setFixedSize(40, 40)
        self.btn_theme.setToolTip("Переключить тему")
        self.btn_theme.setStyleSheet("""
            QPushButton {
                border-radius: 20px;
                background-color: #f0f0f0;
                border: 1px solid #ccc;
            }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """)
        self.btn_theme.clicked.connect(self.toggle_theme)
        top_panel.addWidget(self.btn_theme)

    def apply_theme(self):
        theme = self.settings.get("theme", "light")
        dark = (theme == "dark")

        if dark:
            qss = """
            QMainWindow, QDialog, QGroupBox, QTreeWidget, QListWidget, QMenu, QStatusBar,
            QLabel, QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
                background-color: #2b2b2b;
                color: #dddddd;
            }
            QTextBrowser { color: #000000; }
            QLineEdit, QSpinBox, QDoubleSpinBox, QComboBox {
                background-color: #3c3c3c;
                border: 1px solid #555;
                padding: 2px;
                border-radius: 3px;
            }
            QTreeWidget {
                background-color: #2b2b2b;
                border: 1px solid #444;
                border-radius: 3px;
            }
            QListWidget {
                background-color: #2b2b2b;
                border: 1px solid #444;
            }
            QLabel { background: transparent; }
            QGroupBox {
                border: 1px solid #555;
                border-radius: 5px;
                margin-top: 10px;
            }
            QGroupBox::title {
                color: #dddddd;
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px;
            }
            QMenu::item:selected { background-color: #4a4a4a; }
            QStatusBar { background-color: #2b2b2b; }
            QSplitter::handle { background-color: #2b2b2b; }

            QScrollBar:vertical {
                background: #2b2b2b;
                width: 10px;
                margin: 0;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: #5a5a5a;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: #2b2b2b;
            }
            QScrollBar:horizontal {
                background: #2b2b2b;
                height: 10px;
                margin: 0;
                border-radius: 5px;
            }
            QScrollBar::handle:horizontal {
                background: #5a5a5a;
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: #2b2b2b;
            }

            QPushButton {
                background-color: #3c3c3c;
                color: #dddddd;
                border: 1px solid #555;
                padding: 4px 8px;
                border-radius: 5px;
            }
            QPushButton:hover { background-color: #4a4a4a; }
            QPushButton:pressed { background-color: #555; }
            QCheckBox {
                color: #dddddd;
            }
            QRadioButton {
                color: #dddddd;
            }
            QMenu {
                background-color: #2b2b2b;
                border: 1px solid #444;
                color: #dddddd;
                padding: 4px;
            }
            QMenu::item {
                color: #dddddd;
                background-color: transparent;
                padding: 6px 20px 6px 10px;
                min-width: 160px;
            }
            QMenu::item:selected { background-color: #4a4a4a; color: #ffffff; }
            QMenu::item:disabled { color: #666666; }
            QMenu::separator {
                height: 1px;
                background: #555;
                margin: 4px 8px;
            }
            """
        else:
            qss = """
            QScrollBar:vertical {
                background: transparent;
                width: 10px;
                margin: 0;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: transparent;
            }
            QScrollBar:horizontal {
                background: transparent;
                height: 10px;
                margin: 0;
            }
            QScrollBar::handle:horizontal {
                background: #c0c0c0;
                min-width: 20px;
                border-radius: 5px;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: transparent;
            }
            QScrollBar:vertical {
                background: #f0f0f0;
                width: 10px;
                margin: 0;
                border-radius: 5px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: #f0f0f0;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar:horizontal {
                background: #f0f0f0;
                height: 10px;
                margin: 0;
                border-radius: 5px;
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: #f0f0f0;
            }
            QScrollBar::handle:horizontal {
                background: #c0c0c0;
                min-width: 20px;
                border-radius: 5px;
            }
            QMenu {
                background-color: #ffffff;
                border: 1px solid #ccc;
                color: #000000;
                padding: 4px;
            }
            QMenu::item {
                color: #000000;
                background-color: transparent;
                padding: 6px 20px 6px 10px;
                min-width: 160px;
            }
            QMenu::item:selected { background-color: #e0e0e0; color: #000000; }
            QMenu::item:disabled { color: #a0a0a0; }
            QMenu::separator {
                height: 1px;
                background: #ddd;
                margin: 4px 8px;
            }
            """
        QApplication.instance().setStyleSheet(qss)

        if dark:
            self.main_splitter.setStyleSheet("QSplitter::handle { width: 0px; background: transparent; border: none; }")
        else:
            self.main_splitter.setStyleSheet("")

        if dark:
            self.attachments_scroll.setStyleSheet("QScrollArea { background-color: #2b2b2b; }")
            self.attachments_widget.setStyleSheet("background-color: #2b2b2b;")
        else:
            self.attachments_scroll.setStyleSheet("")
            self.attachments_widget.setStyleSheet("")

        # Стиль для поля ввода (отдельно, чтобы не затрагивать QTextBrowser)
        if dark:
            self.input_edit.setStyleSheet("""
                QTextEdit {
            background-color: #1e1e1e;
            color: #dddddd;
            border: 1px solid #444;
            border-radius: 3px;
            padding: 4px;
                }
            """)
        else:
            self.input_edit.setStyleSheet("")

        round_dark = """
            QPushButton { border-radius: 20px; background-color: #3c3c3c; border: 1px solid #555; }
            QPushButton:hover { background-color: #4a4a4a; }
            QPushButton:pressed { background-color: #555; }
        """
        round_light = """
            QPushButton { border-radius: 20px; background-color: #f0f0f0; border: 1px solid #ccc; }
            QPushButton:hover { background-color: #e0e0e0; }
            QPushButton:pressed { background-color: #d0d0d0; }
        """

        standard_buttons = [
            (self.btn_new_chat, self.icon_add_chat,
             QIcon(self.resource_path("Image_White/Add_chat_White.png")) if os.path.exists(
                 self.resource_path("Image_White/Add_chat_White.png")) else self.icon_add_chat),
            (self.btn_new_folder, self.icon_add_folder,
             QIcon(self.resource_path("Image_White/Add_folder_White.png")) if os.path.exists(
                 self.resource_path("Image_White/Add_folder_White.png")) else self.icon_add_folder),
            (self.btn_sort, self.icon_sort,
             QIcon(self.resource_path("Image_White/Sorting_White.png")) if os.path.exists(
                 self.resource_path("Image_White/Sorting_White.png")) else self.icon_sort),
            (self.btn_toggle_sidebar, self.icon_collapse_expand, self.icon_collapse_expand_white),
            (self.btn_starred, self.icon_bookmarks, self.icon_bookmarks_white),
            (self.btn_global_search, self.icon_search, self.icon_search_white),
            (self.btn_import_ds, self.icon_import, self.icon_import_white),
            (self.btn_export, self.icon_export, self.icon_export_white),
            (self.btn_settings, self.icon_settings, self.icon_settings_3_white),
            (self.btn_archive, self.icon_archive, self.icon_archive_white),
            (self.btn_trash, self.icon_trash, self.icon_trash_white),
            (self.btn_theme, self.icon_theme, self.icon_theme_white),
            (self.btn_attach, self.icon_attach, self.icon_attach_white),
        ]

        special_buttons = [
            (self.btn_send, self.icon_send, self.icon_send_white),
            (self.btn_toggle_reasoning, self.icon_reflection, self.icon_reflection_white),
            (self.btn_stop_generation, self.icon_stop,
             QIcon(self.resource_path("Image_White/Stop_generation_White.png")) if os.path.exists(
                 self.resource_path("Image_White/Stop_generation_White.png")) else self.icon_stop),
            (self.btn_chat_settings, self.icon_setting_chat, self.icon_setting_chat_white),
            (self.btn_scroll_down, self.icon_down, self.icon_down_white),
        ]

        for btn, light_icon, dark_icon in standard_buttons:
            btn.setStyleSheet(round_dark if dark else round_light)
            btn.setIcon(dark_icon if dark else light_icon)

        for btn, light_icon, dark_icon in special_buttons:
            btn.setIcon(dark_icon if dark else light_icon)

        self.set_tree_items_dark(dark)

        if self.current_chat_index >= 0:
            self.apply_chat_background(self.chats[self.current_chat_index])
        else:
            self.apply_chat_background(None)
        self.chat_browser.viewport().update()
        self.chat_browser.repaint()

    def set_tree_items_dark(self, dark):
        self.chat_tree.setUpdatesEnabled(False)
        try:
            def update_item(item):
                data = item.data(0, Qt.UserRole)
                if data and data.get('type') == 'folder':
                    color = data.get('color', '#FFFFFF')
                    if dark:
                        if color == '#FFFFFF':
                            item.setBackground(0, QColor("#2b2b2b"))
                            item.setForeground(0, QColor("#c47b0e"))
                        else:
                            item.setBackground(0, QColor(color))
                            item.setForeground(0, QColor("#c47b0e"))
                    else:
                        if color:
                            item.setBackground(0, QColor(color))
                        else:
                            item.setBackground(0, QBrush())
                        item.setForeground(0, QBrush())

                    if dark:
                        if item.isExpanded():
                            item.setIcon(0, self.icon_folder_open_white)
                        else:
                            item.setIcon(0, self.icon_folder_closed_white)
                    else:
                        if item.isExpanded():
                            item.setIcon(0, self.icon_folder_open)
                        else:
                            item.setIcon(0, self.icon_folder_closed)

                elif data and data.get('type') == 'chat':
                    color = data.get('color', '#FFFFFF')
                    if dark:
                        if color == '#FFFFFF':
                            item.setBackground(0, QColor("#2b2b2b"))
                            item.setForeground(0, QColor("#c47b0e"))
                        else:
                            item.setBackground(0, QColor(color))
                            item.setForeground(0, QColor("#c47b0e"))
                    else:
                        if color:
                            item.setBackground(0, QColor(color))
                        else:
                            item.setBackground(0, QBrush())
                        item.setForeground(0, QBrush())

                    item.setIcon(0, self.icon_chat_white if dark else self.icon_chat)

                for i in range(item.childCount()):
                    update_item(item.child(i))

            for i in range(self.chat_tree.topLevelItemCount()):
                update_item(self.chat_tree.topLevelItem(i))
        finally:
            self.chat_tree.setUpdatesEnabled(True)

    def update_single_item_theme(self, item, dark):
        data = item.data(0, Qt.UserRole)
        if not data:
            return
        color = data.get('color', '#FFFFFF')

        if dark:
            if color == '#FFFFFF':
                item.setBackground(0, QColor("#2b2b2b"))
                item.setForeground(0, QColor("#c47b0e"))
            else:
                item.setBackground(0, QColor(color))
                item.setForeground(0, QColor("#c47b0e"))  # оранжевый текст
        else:
            if color:
                item.setBackground(0, QColor(color))
            else:
                item.setBackground(0, QBrush())
            item.setForeground(0, QBrush())

        if data.get('type') == 'folder':
            if dark:
                if item.isExpanded():
                    item.setIcon(0, self.icon_folder_open_white)
                else:
                    item.setIcon(0, self.icon_folder_closed_white)
            else:
                if item.isExpanded():
                    item.setIcon(0, self.icon_folder_open)
                else:
                    item.setIcon(0, self.icon_folder_closed)
        elif data.get('type') == 'chat':
            item.setIcon(0, self.icon_chat_white if dark else self.icon_chat)

    def toggle_theme(self):
        def toggle_theme(self):
            with self.busy("Применение темы…"):
                if self.settings.get("theme", "light") == "light":
                    self.settings["theme"] = "dark"
                else:
                    self.settings["theme"] = "light"
                self.save_settings()
                self.apply_theme()
                self.update_system_theme()

        if self.settings.get("theme", "light") == "light":
            self.settings["theme"] = "dark"
            if hasattr(QApplication.instance().styleHints(), 'setColorScheme'):
                QApplication.instance().styleHints().setColorScheme(Qt.ColorScheme.Dark)
        else:
            self.settings["theme"] = "light"
            if hasattr(QApplication.instance().styleHints(), 'setColorScheme'):
                QApplication.instance().styleHints().setColorScheme(Qt.ColorScheme.Light)
        self.save_settings()
        self.apply_theme()

    def update_system_theme(self):
        app = QApplication.instance()
        if hasattr(app.styleHints(), 'setColorScheme'):
            if self.settings.get("theme", "light") == "dark":
                app.styleHints().setColorScheme(Qt.ColorScheme.Dark)
            else:
                app.styleHints().setColorScheme(Qt.ColorScheme.Light)

    def apply_font_settings(self):
        """Применяет настройки шрифтов к области чата и редактору ввода."""

        chat_family = self.settings.get("chat_font_family", "Arial")
        chat_size = int(self.settings.get("chat_font_size", 10))
        input_family = self.settings.get("input_font_family", "Arial")
        input_size = int(self.settings.get("input_font_size", 10))

        chat_font = QFont(chat_family, chat_size)
        input_font = QFont(input_family, input_size)

        self.chat_browser.setFont(chat_font)
        self.input_edit.setFont(input_font)

        if self.current_chat_index >= 0:
            self._redraw_chat_with_history()

    # ======================================================================
    #  РАБОТА С ДЕРЕВОМ ЧАТОВ И ПАПОК
    # ======================================================================

    def load_chats_from_disk(self):
        # Полная очистка текущих данных перед загрузкой
        self.chats = []
        self.archived_chats = []
        self.chat_tree.clear()
        self.chat_items.clear()
        self.folder_items.clear()

        # Считаем все файлы, чтобы показывать прогресс
        if os.path.isdir(self.chats_dir):
            all_json_files = [f for f in os.listdir(self.chats_dir) if f.endswith('.json')]
        else:
            all_json_files = []
        total_files = len(all_json_files)

        if self._splash:
            self._splash.set_progress(0, total_files, "Загрузка чатов…")

        # Исключаем файлы, которые находятся в корзине
        errors = []
        trash_filepaths = set()
        if hasattr(self, 'trash_items'):
            for entry in self.trash_items:
                if entry.get('type') == 'chat':
                    trash_filepaths.add(entry.get('filepath'))

        corrupted_dir = "corrupted"
        os.makedirs(corrupted_dir, exist_ok=True)

        for idx, filename in enumerate(all_json_files):
            filepath = os.path.join(self.chats_dir, filename)
            if self._splash:
                self._splash.set_progress(
                    idx + 1, total_files,
                    f"Загрузка чатов… {idx + 1}/{total_files}"
                )
            if filepath in trash_filepaths:
                continue

            try:
                with open(filepath, 'rb') as f:
                    content = f.read()

                is_encrypted = content.startswith(b"ENC:")
                if is_encrypted:
                    # Показываем заглушку, реальная расшифровка — при клике
                    title = os.path.splitext(filename)[0]
                    chat_data = {
                        'filepath': filepath,
                        'title': title,
                        'messages': [],
                        'folder': '',
                        'color': '#FFFFFF',
                        'pinned': False,
                        'background_color': None,
                        'background_image': None,
                        'model': self.settings.get('model', 'local-model'),
                        'created_time': time.strftime('%d.%m.%Y %H:%M:%S'),
                        'last_activity_time': time.strftime('%d.%m.%Y %H:%M:%S'),
                        'total_tokens_sent': 0,
                        'total_generation_time': 0.0,
                        'background_image_original': None,
                        'encrypted': True,
                        'locked': True,
                        'password_hash': None
                    }
                    self.chats.append(chat_data)
                    self.add_chat_to_tree(chat_data)
                    continue
                else:
                    data = json.loads(content.decode('utf-8'))

                if isinstance(data, dict) and 'messages' in data:
                    title = data.get('title', os.path.splitext(filename)[0])
                    folder = data.get('folder', '')
                    messages = data['messages']
                    background_color = data.get('background_color')
                    background_image = data.get('background_image')
                    model = data.get('model', self.settings.get('model', 'local-model'))
                    created_time = data.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S'))
                    last_activity_time = data.get('last_activity_time', created_time)
                    total_tokens_sent = data.get('total_tokens_sent', 0)
                    total_generation_time = data.get('total_generation_time', 0.0)
                    system_prompt = data.get('system_prompt', '')
                    params = data.get('params', {})
                    color = data.get('color', '#FFFFFF')
                    pinned = data.get('pinned', False)
                    archived = data.get('archived', False)
                    encrypted = data.get('encrypted', is_encrypted)
                    password_hash = data.get('password_hash')
                elif isinstance(data, list):
                    messages = data
                    title = os.path.splitext(filename)[0]
                    folder = ''
                    background_color = None
                    background_image = None
                    model = self.settings.get('model', 'local-model')
                    created_time = time.strftime('%d.%m.%Y %H:%M:%S')
                    last_activity_time = created_time
                    total_tokens_sent = 0
                    total_generation_time = 0.0
                    color = '#FFFFFF'
                    pinned = False
                    archived = False
                    encrypted = is_encrypted
                    password_hash = None
                else:
                    raise ValueError("Неподдерживаемый формат файла")

                normalized_messages = []
                for msg in messages:
                    if not isinstance(msg, dict):
                        continue
                    content = msg.get('content', '')
                    if not isinstance(content, str):
                        if isinstance(content, list):
                            text_parts = []
                            for item in content:
                                if isinstance(item, dict):
                                    if 'text' in item:
                                        text_parts.append(str(item['text']))
                                    elif 'image_url' in item:
                                        text_parts.append('[Изображение]')
                                    else:
                                        text_parts.append(str(item))
                                else:
                                    text_parts.append(str(item))
                            content = ' '.join(text_parts)
                        else:
                            content = str(content)

                    attachments = msg.get('attachments', [])
                    if not attachments:
                        if 'image' in msg and msg['image']:
                            attachments.append(
                                {"type": "image", "data": msg['image'], "name": msg.get('name', 'image')})
                        if 'file_path' in msg:
                            attachments.append(
                                {"type": "file", "path": msg['file_path'], "name": msg.get('name', 'file')})

                    new_msg = msg.copy()
                    new_msg['content'] = content
                    new_msg['attachments'] = attachments
                    new_msg.pop('image', None)
                    new_msg.pop('file_path', None)
                    normalized_messages.append(new_msg)

                chat_data = {
                    'filepath': filepath,
                    'title': title,
                    'messages': normalized_messages,
                    'folder': folder,
                    'background_color': background_color,
                    'background_image': background_image,
                    'model': model,
                    'created_time': created_time,
                    'last_activity_time': last_activity_time,
                    'total_tokens_sent': total_tokens_sent,
                    'total_generation_time': total_generation_time,
                    'color': color,
                    'params': params,
                    'system_prompt': system_prompt,
                    'pinned': pinned,
                    'archived': archived,
                    'encrypted': encrypted,
                    'background_image_original': data.get('background_image_original'),
                    'tags': data.get('tags', []),
                    'password_hash': password_hash
                }

                if archived:
                    self.archived_chats.append(chat_data)
                else:
                    self.chats.append(chat_data)
                    self.add_chat_to_tree(chat_data)

            except json.JSONDecodeError as e:
                self.logger.warning(f"Файл {filename} повреждён: {e}")
                errors.append(f"{filename}: JSON повреждён")
                # Складываем в corrupted, добавляя таймстамп, чтобы не перезаписать другой битый файл
                try:
                    safe_name = f"{os.path.splitext(filename)[0]}_{int(time.time())}.json"
                    shutil.move(filepath, os.path.join(corrupted_dir, safe_name))
                except Exception as move_err:
                    self.logger.warning(f"Не удалось переместить {filename}: {move_err}")

                backup_path = self._try_restore_chat_from_backup(filename)
                if backup_path:
                    self.logger.info(f"Найден бэкап для {filename}: {backup_path}")

            except Exception as e:
                self.logger.warning(f"Ошибка загрузки {filename}: {e}")
                print(f"⚠️ Ошибка загрузки {filename}: {e}")

        for folder_data in self.folders:
            folder_name = folder_data['name']
            if folder_name not in self.folder_items:
                folder_item = QTreeWidgetItem([folder_name])
                folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
                folder_item.setData(0, Qt.UserRole, {'type': 'folder', 'name': folder_name})
                color = folder_data.get('color', '#FFFFFF')
                if color:
                    folder_item.setBackground(0, QColor(color))
                self.set_folder_icon(folder_item, expanded=False)
                self.folder_items[folder_name] = folder_item
                self.chat_tree.addTopLevelItem(folder_item)

        self.save_folders()
        tree_order = self.load_tree_order()
        if tree_order is not None:
            self.apply_tree_order(tree_order)

        # Если дерево пустое или количество элементов не совпадает с количеством чатов,
        # принудительно перестраиваем из текущих данных
        if self.chat_tree.topLevelItemCount() == 0 or len(self.chat_items) != len(self.chats):
            self.refresh_tree_from_data()

        self.reorder_tree_by_pinned()

        # === ЛОГИРОВАНИЕ ===
        if self._splash:
            self._splash.set_progress(total_files, total_files, "Применение порядка…")
        self.logger.info(f"Загружено чатов: {len(self.chats)}")
        if errors:
            self._pending_errors = errors

        self.logger.debug(f"Загружено чатов: {len(self.chats)}")
        for chat in self.chats:
            self.logger.debug(f"Чат {chat['title']} -> {chat['filepath']}")

    def _try_restore_chat_from_backup(self, filename):
        """Пытается найти файл в свежих бэкапах."""

        backup_root = "backups"
        if not os.path.isdir(backup_root):
            return None
        backups = sorted(
            [os.path.join(backup_root, d) for d in os.listdir(backup_root)
             if d.startswith("backup_") and os.path.isdir(os.path.join(backup_root, d))],
            reverse=True
        )
        for bdir in backups:
            candidate = os.path.join(bdir, "chats", filename)
            if os.path.exists(candidate):
                return candidate
        return None

    def refresh_tree_from_data(self):
        self.chat_tree.clear()
        self.folder_items.clear()
        self.chat_items.clear()
        for chat in self.chats:
            self.add_chat_to_tree(chat)
        for folder_data in self.folders:
            folder_name = folder_data['name']
            if folder_name not in self.folder_items:
                folder_item = QTreeWidgetItem([folder_name])
                folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
                folder_item.setData(0, Qt.UserRole, {'type': 'folder', 'name': folder_name})
                color = folder_data.get('color', '#FFFFFF')
                if color:
                    folder_item.setBackground(0, QColor(color))
                self.set_folder_icon(folder_item, expanded=False)
                self.folder_items[folder_name] = folder_item
                self.chat_tree.addTopLevelItem(folder_item)
        self.save_tree_order()
        self.reorder_tree_by_pinned()

    def add_chat_to_tree(self, chat_data, parent_item=None):
        filepath = chat_data['filepath']
        title = chat_data['title']
        folder = chat_data['folder']

        # Если родитель не передан и чат должен находиться в папке
        if parent_item is None and folder:
            if folder not in self.folder_items:
                folder_params = next((f for f in self.folders if isinstance(f, dict) and f.get('name') == folder), None)
                pinned = folder_params.get('pinned', False) if folder_params else False
                color = folder_params.get('color', '#FFFFFF') if folder_params else '#FFFFFF'

                folder_item = QTreeWidgetItem([folder])
                folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled | Qt.ItemIsDragEnabled)
                folder_data = {'type': 'folder', 'name': folder, 'pinned': pinned, 'color': color}
                folder_item.setData(0, Qt.UserRole, folder_data)
                if color:
                    folder_item.setBackground(0, QColor(color))
                self.set_folder_icon(folder_item, expanded=False)
                if pinned:
                    folder_item.setIcon(1, self.icon_pin)

                self.folder_items[folder] = folder_item
                self.chat_tree.addTopLevelItem(folder_item)

                if not any(isinstance(f, dict) and f.get('name') == folder for f in self.folders):
                    self.folders.append({"name": folder, "color": color, "pinned": pinned})
            parent_item = self.folder_items[folder]

        # Определяем отображаемое название с учётом тегов
        title_for_tree = chat_data.get('title', '')
        prefix = ""
        if chat_data.get('incognito', False):
            prefix += "🕶 "
        # Непрочитанность добавится позже через update_chat_item_text
        title_with_prefix = prefix + title_for_tree

        chat_item = QTreeWidgetItem([title_with_prefix])

        tags = chat_data.get('tags', [])
        if tags:
            tags_str = ' '.join(['#' + tag for tag in tags])
            chat_item.setText(2, tags_str)
            chat_item.setForeground(2, QBrush(QColor("#1a73e8")))

        chat_item = QTreeWidgetItem([title_for_tree])
        chat_item.setFlags((chat_item.flags() | Qt.ItemIsDragEnabled) & ~Qt.ItemIsDropEnabled)

        pinned = chat_data.get('pinned', False)
        color = chat_data.get('color', '#FFFFFF')
        chat_item.setData(0, Qt.UserRole, {'type': 'chat', 'filepath': filepath, 'pinned': pinned, 'color': color})
        if color:
            chat_item.setBackground(0, QColor(color))
        self.set_chat_icon(chat_item)
        if pinned:
            chat_item.setIcon(1, self.icon_pin)
        self.chat_items[filepath] = chat_item

        if parent_item:
            parent_item.addChild(chat_item)
            parent_item.setExpanded(True)
            self.set_folder_icon(parent_item, expanded=True)
        else:
            self.chat_tree.addTopLevelItem(chat_item)

        QTimer.singleShot(0, lambda: self.set_tree_items_dark(self.settings.get("theme", "light") == "dark"))

    def rebuild_chat_tree(self, sort_criterion=None):
        """Перестраивает дерево чатов с сортировкой папок и чатов внутри."""

        # Запоминаем выбранный чат, чтобы восстановить выбор после перестройки
        selected_filepath = None
        if 0 <= self.current_chat_index < len(self.chats):
            selected_filepath = self.chats[self.current_chat_index]['filepath']

        self.chat_tree.clear()
        self.chat_items.clear()
        self.folder_items.clear()

        # Функция вычисления ключа для сортировки папки
        def folder_sort_key(folder_name):
            chats_in_folder = [c for c in self.chats if c.get('folder') == folder_name]
            if sort_criterion == 'messages':
                return -sum(len(c.get('messages', [])) for c in chats_in_folder)
            elif sort_criterion == 'title':
                return folder_name.lower()
            elif sort_criterion == 'date':
                valid_dates = []
                for c in chats_in_folder:
                    date_str = c.get('created_time', '')
                    try:
                        valid_dates.append(time.mktime(time.strptime(date_str, '%d.%m.%Y %H:%M:%S')))
                    except:
                        pass
                return -max(valid_dates) if valid_dates else 0
            else:
                return 0

        # Сортировка папок верхнего уровня
        folder_items = []
        for f in self.folders:
            name = f['name']
            folder_items.append((name, f.get('color', '#FFFFFF'), f.get('pinned', False)))

        if sort_criterion is not None:
            folder_items.sort(key=lambda x: folder_sort_key(x[0]))

        # Создаём папки в отсортированном порядке
        for name, color, pinned in folder_items:
            if name not in self.folder_items:
                folder_item = QTreeWidgetItem([name])
                folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
                folder_item.setData(0, Qt.UserRole, {'type': 'folder', 'name': name, 'pinned': pinned, 'color': color})
                if color:
                    folder_item.setBackground(0, QColor(color))
                self.set_folder_icon(folder_item, expanded=False)
                if pinned:
                    folder_item.setIcon(1, self.icon_pin)
                self.folder_items[name] = folder_item
                self.chat_tree.addTopLevelItem(folder_item)

        # Функция вычисления ключа для чата
        def chat_sort_key(chat):
            if sort_criterion == 'date':
                date_str = chat.get('created_time', '')
                try:
                    return -time.mktime(time.strptime(date_str, '%d.%m.%Y %H:%M:%S'))
                except:
                    return 0
            elif sort_criterion == 'title':
                return chat.get('title', '').lower()
            elif sort_criterion == 'messages':
                return -len(chat.get('messages', []))
            else:
                return 0

        # Сортируем чаты
        chats_sorted = sorted(self.chats, key=chat_sort_key)

        # Добавляем чаты (родитель определяется по папке)
        for chat in chats_sorted:
            parent_item = self.folder_items.get(chat.get('folder', ''), None)
            self.add_chat_to_tree(chat, parent_item)

        # Применяем закрепления (это не изменит сортировку, но поднимет закреплённые вверх)
        self.reorder_tree_by_pinned()

        # Восстанавливаем выбранный чат
        if selected_filepath:
            self.select_chat_by_filepath(selected_filepath)

    def sort_chats_by(self, key):
        """Сортирует чаты по указанному ключу и перестраивает дерево."""

        def sort_key(chat):
            if key == 'date':
                date_str = chat.get('created_time', '')
                try:
                    return -time.mktime(time.strptime(date_str, '%d.%m.%Y %H:%M:%S'))
                except:
                    return 0
            elif key == 'title':
                return chat.get('title', '').lower()
            elif key == 'messages':
                return -len(chat.get('messages', []))
            else:
                return 0

        self.chats.sort(key=sort_key)
        self.rebuild_chat_tree()
        self.set_tree_items_dark(self.settings.get("theme", "light") == "dark")
        status_map = {
            'date': "по дате (сначала новые)",
            'title': "по названию (А-Я)",
            'messages': "по количеству сообщений (больше сверху)"
        }
        self.statusBar().showMessage(f"Чаты отсортированы: {status_map.get(key, key)}")

    def reorder_tree_by_pinned(self):
        """Перемещает закреплённые элементы вверх, извлекает закреплённые чаты из незакреплённых папок, добавляет разделитель."""

        self.chat_tree.setUpdatesEnabled(False)
        self.chat_tree.blockSignals(True)
        try:
            # 1. Собираем все верхнеуровневые элементы, игнорируя разделители
            top_items = []
            for i in range(self.chat_tree.topLevelItemCount()):
                item = self.chat_tree.topLevelItem(i)
                data = item.data(0, Qt.UserRole)
                if data and data.get('type') == 'separator':
                    continue
                top_items.append(item)

            # 2. Разделяем на закреплённые и обычные
            pinned_top = []
            unpinned_top = []
            for item in top_items:
                if self.is_item_pinned(item):
                    pinned_top.append(item)
                else:
                    unpinned_top.append(item)

            # 3. Извлекаем закреплённые чаты из незакреплённых папок
            for item in top_items:
                data = item.data(0, Qt.UserRole)
                if data and data.get('type') == 'folder':
                    if self.is_item_pinned(item):
                        continue
                    children = [item.child(j) for j in range(item.childCount())]
                    for child in children:
                        if self.is_item_pinned(child):
                            pinned_top.append(child)
                            item.removeChild(child)

            # 4. Очищаем верхний уровень
            while self.chat_tree.topLevelItemCount() > 0:
                self.chat_tree.takeTopLevelItem(0)

            # 5. Добавляем закреплённые
            for item in pinned_top:
                self.chat_tree.addTopLevelItem(item)

            # 6. Разделитель только если есть и закреплённые, и обычные
            if pinned_top and unpinned_top:
                separator = QTreeWidgetItem(["─────────"])
                separator.setFlags(Qt.NoItemFlags)
                separator.setForeground(0, QColor("#AAAAAA"))
                separator.setData(0, Qt.UserRole, {'type': 'separator'})
                self.chat_tree.addTopLevelItem(separator)

            # 7. Добавляем обычные
            for item in unpinned_top:
                self.chat_tree.addTopLevelItem(item)

            # 8. Иконки булавок
            for i in range(self.chat_tree.topLevelItemCount()):
                item = self.chat_tree.topLevelItem(i)
                data = item.data(0, Qt.UserRole)
                if not data or data.get('type') == 'separator':
                    continue
                if data.get('pinned', False):
                    item.setIcon(1, self.icon_pin)
                else:
                    item.setIcon(1, QIcon())

            for i in range(self.chat_tree.topLevelItemCount()):
                item = self.chat_tree.topLevelItem(i)
                data = item.data(0, Qt.UserRole)
                if data and data.get('type') == 'folder' and data.get('pinned', False):
                    self.update_children_pin_icons(item, True)

            # 9. Сохраняем порядок
            self.update_all_folder_icons()
            self.save_tree_order()
        finally:
            self.chat_tree.blockSignals(False)
            self.chat_tree.setUpdatesEnabled(True)
            # Один раз перекрашиваем — без QTimer
            self.set_tree_items_dark(self.settings.get("theme", "light") == "dark")

    def set_item_pinned(self, item, pinned):
        data = item.data(0, Qt.UserRole)
        if not data:
            return

        if data['type'] == 'chat' and not pinned:
            parent = item.parent()
            if parent:
                parent_data = parent.data(0, Qt.UserRole)
                if parent_data and parent_data.get('type') == 'folder' and parent_data.get('pinned', False):
                    # Перемещаем чат в корень
                    parent.removeChild(item)
                    self.chat_tree.addTopLevelItem(item)

                    # Обновляем folder у чата
                    filepath = data['filepath']
                    for chat in self.chats:
                        if chat['filepath'] == filepath:
                            chat['folder'] = ''
                            self.save_chat_by_filepath(filepath)
                            break

                    # Если папка опустела, сворачиваем её и меняем иконку
                    if parent.childCount() == 0:
                        parent.setExpanded(False)
                        self.set_folder_icon(parent, expanded=False)

                    # Сохраняем актуальный порядок дерева
                    self.save_tree_order()

        data['pinned'] = pinned
        item.setData(0, Qt.UserRole, data)

        if data['type'] == 'chat':
            filepath = data['filepath']
            for chat in self.chats:
                if chat['filepath'] == filepath:
                    chat['pinned'] = pinned
                    self.save_chat_by_filepath(filepath)
                    break
        elif data['type'] == 'folder':
            folder_name = data.get('name')
            for folder in self.folders:
                if isinstance(folder, dict) and folder.get('name') == folder_name:
                    folder['pinned'] = pinned
                    self.save_folders()
                    break
            # Если папка открепляется, сбрасываем закрепление у всех вложенных элементов
            if not pinned:
                self.reset_children_pinned(item)

        # Обновляем иконку булавки
        if pinned:
            item.setIcon(1, self.icon_pin)
        else:
            item.setIcon(1, QIcon())

    def toggle_pin(self, item):
        self.chat_tree.setUpdatesEnabled(False)
        try:
            new_pinned = not self.is_item_pinned(item)
            self.set_item_pinned(item, new_pinned)
            self.reorder_tree_by_pinned()
        finally:
            self.chat_tree.setUpdatesEnabled(True)
            self.set_tree_items_dark(self.settings.get("theme", "light") == "dark")

    def is_item_pinned(self, item):
        data = item.data(0, Qt.UserRole)
        if not data:
            return False
        if data.get('pinned', False):
            return True
        if data.get('type') == 'chat':
            parent = item.parent()
            if parent:
                parent_data = parent.data(0, Qt.UserRole)
                if parent_data and parent_data.get('type') == 'folder':
                    return parent_data.get('pinned', False)
        return False

    def set_item_color(self, item, color_hex):
        """Устанавливает цвет фона элемента и сохраняет в данных."""

        data = item.data(0, Qt.UserRole)
        if not data:
            return
        if color_hex is None:
            # Сброс цвета
            if self.settings.get("theme", "light") == "dark":
                item.setBackground(0, QColor("#3c3c3c"))
                item.setForeground(0, QColor("#c47b0e"))
            else:
                item.setBackground(0, QBrush())
                item.setForeground(0, QBrush())
            color_hex = "#FFFFFF"
        else:
            item.setBackground(0, QColor(color_hex))

        # Обновляем данные в самом элементе
        data['color'] = color_hex
        item.setData(0, Qt.UserRole, data)

        # Сохраняем в структуру данных
        if data['type'] == 'chat':
            filepath = data['filepath']
            for chat in self.chats:
                if chat['filepath'] == filepath:
                    chat['color'] = color_hex
                    self.save_chat_by_filepath(filepath)
                    break
        elif data['type'] == 'folder':
            folder_name = data['name']
            for folder in self.folders:
                if folder['name'] == folder_name:
                    folder['color'] = color_hex
                    self.save_folders()
                    break

    def update_all_folder_icons(self):
        """Обновляет иконки всех папок в дереве на основе их состояния раскрытия."""

        def update_recursive(item):
            data = item.data(0, Qt.UserRole)
            if data and data.get('type') == 'folder':
                if item.isExpanded():
                    item.setIcon(0, self.icon_folder_open)
                else:
                    item.setIcon(0, self.icon_folder_closed)
            for i in range(item.childCount()):
                update_recursive(item.child(i))

        for i in range(self.chat_tree.topLevelItemCount()):
            update_recursive(self.chat_tree.topLevelItem(i))

    def set_folder_icon(self, item, expanded=False):
        """Устанавливает иконку папки (закрытая/открытая)."""

        icon = self.icon_folder_open if expanded else self.icon_folder_closed
        item.setIcon(0, icon)

    def set_chat_icon(self, item):
        """Устанавливает иконку чата."""

        item.setIcon(0, self.icon_chat)

    def on_item_collapsed(self, item):
        data = item.data(0, Qt.UserRole)
        if data and data.get('type') == 'folder':
            self.set_folder_icon(item, expanded=False)
        dark = self.settings.get("theme", "light") == "dark"
        self.update_single_item_theme(item, dark)

    def on_item_expanded(self, item):
        data = item.data(0, Qt.UserRole)
        if data and data.get('type') == 'folder':
            self.set_folder_icon(item, expanded=True)
        # Обновляем тему элемента (фон, текст, иконку)
        dark = self.settings.get("theme", "light") == "dark"
        self.update_single_item_theme(item, dark)

    def rename_chat(self, item):
        filepath = item.data(0, Qt.UserRole)['filepath']
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        new_title, ok = QInputDialog.getText(self, "Переименовать чат", "Новое название:",
                                             text=self.chats[idx]['title'])
        if ok and new_title.strip():
            new_title = new_title.strip()
            self.chats[idx]['title'] = new_title
            item.setText(0, new_title)
            self.save_current_chat()

    def copy_full_dialog(self, chat):
        """Копирует весь диалог в буфер обмена в виде текста."""

        lines = [f"# {chat.get('title', 'Чат')}\n"]

        for msg in chat.get('messages', []):
            role = msg.get('role', '')
            content = msg.get('content', '')
            time_str = msg.get('time', '')

            # Если контент — список (мультимодальное сообщение)
            if isinstance(content, list):
                text_parts = []
                for item in content:
                    if isinstance(item, dict) and 'text' in item:
                        text_parts.append(item['text'])
                content = '\n'.join(text_parts)

            if role == 'user':
                header = "👤 Пользователь"
            elif role == 'assistant':
                header = "🤖 Ассистент"
            else:
                continue

            if time_str:
                header += f" ({time_str})"

            lines.append(header)
            lines.append(content)
            lines.append("")
            lines.append("---")
            lines.append("")

        text = "\n".join(lines)
        QApplication.clipboard().setText(text)
        self.statusBar().showMessage(
            f"Диалог «{chat.get('title', '')}» скопирован в буфер обмена"
        )

    def rename_folder(self, item):
        old_name = item.data(0, Qt.UserRole)['name']
        new_name, ok = QInputDialog.getText(self, "Переименовать папку", "Новое название:", text=old_name)
        if ok and new_name.strip() and new_name.strip() != old_name:
            new_name = new_name.strip()
            if new_name in self.folder_items:
                QMessageBox.warning(self, "Ошибка", "Папка с таким именем уже существует.")
                return
            for folder in self.folders:
                if isinstance(folder, dict) and folder.get('name') == old_name:
                    folder['name'] = new_name
                    break
            self.save_folders()
            self.folder_items[new_name] = self.folder_items.pop(old_name)
            self.folder_items[new_name].setData(0, Qt.UserRole, {'type': 'folder', 'name': new_name})
            self.folder_items[new_name].setText(0, new_name)
            for chat in self.chats:
                if chat['folder'] == old_name:
                    chat['folder'] = new_name
                    self.save_chat_by_filepath(chat['filepath'])
            self.save_tree_order()

    def create_new_chat(self, in_root=False):
        """Создаёт новый чат, запрашивая название у пользователя."""

        dlg = QInputDialog(self)
        dlg.setWindowTitle("Новый чат")
        dlg.setLabelText("Введите название чата:")
        dlg.setTextValue("")
        dlg.setWindowIcon(QIcon("Image/Chat_2.png"))

        if dlg.exec() != QDialog.Accepted:
            return

        title = dlg.textValue().strip()
        if not title:
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            title = f"Чат {timestamp}"

        # Определяем папку для нового чата
        folder = ''
        if not in_root:
            item = self.chat_tree.currentItem()
            if item:
                data = item.data(0, Qt.UserRole)
                if data and data.get('type') == 'chat':
                    parent = item.parent()
                    if parent:
                        folder = parent.data(0, Qt.UserRole).get('name', '')
                elif data and data.get('type') == 'folder':
                    folder = data.get('name', '')

        # Формируем имя файла на основе названия
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        if not safe_title:
            safe_title = f"Чат {time.strftime('%Y%m%d_%H%M%S')}"
        filepath = os.path.join(self.chats_dir, f"{safe_title}.json")

        if os.path.exists(filepath):
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            filepath = os.path.join(self.chats_dir, f"{safe_title}_{timestamp}.json")

        new_chat = {
            'filepath': filepath,
            'title': title,
            'messages': [],
            'folder': folder,
            'color': '#FFFFFF',
            'pinned': False,
            'background_color': None,
            'background_image': None,
            'model': self.settings.get('model', 'local-model'),
            'params': {},
            'system_prompt': '',
            'created_time': time.strftime('%d.%m.%Y %H:%M:%S'),
            'last_activity_time': time.strftime('%d.%m.%Y %H:%M:%S'),
            'total_tokens_sent': 0,
            'total_generation_time': 0.0,
            'password_hash': None,
            'background_image_original': None,
            'incognito': False,
            'encrypted': False
        }

        self.chats.append(new_chat)
        self.add_chat_to_tree(new_chat)
        dark = self.settings.get("theme", "light") == "dark"
        item = self.chat_items.get(filepath)
        if item:
            self.update_single_item_theme(item, dark)
        self.select_chat_by_filepath(filepath)
        try:
            self.save_current_chat()
            self.save_tree_order()
        except Exception as e:
            self.logger.exception("Ошибка при создании чата")
            self.show_error("Ошибка", f"Чат создан, но не удалось сохранить его на диск.\nПричина: {e}")
        self.clear_chat_display()

    def create_new_folder(self):
        """Создаёт новую папку, запрашивая имя у пользователя."""

        dlg = QInputDialog(self)
        dlg.setWindowTitle("Новая папка")
        dlg.setLabelText("Введите название папки:")
        dlg.setTextValue("")
        dlg.setWindowIcon(QIcon("Image/Folder_2.png"))

        if dlg.exec() != QDialog.Accepted:
            return

        name = dlg.textValue().strip()
        if not name:
            return

        if name in self.folder_items:
            QMessageBox.warning(self, "Ошибка", "Папка с таким именем уже существует.")
            return

        # Добавляем папку в self.folders как словарь
        try:
            self.folders.append({"name": name, "color": "#FFFFFF", "pinned": False})
            self.save_folders()
        except Exception as e:
            self.logger.exception("Ошибка при создании папки")
            self.show_error("Ошибка", f"Не удалось сохранить папку.\nПричина: {e}")

        folder_item = QTreeWidgetItem([name])
        # Добавляем возможность перетаскивания папки
        folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled | Qt.ItemIsDragEnabled)
        folder_data = {'type': 'folder', 'name': name, 'pinned': False, 'color': '#FFFFFF'}
        folder_item.setData(0, Qt.UserRole, folder_data)
        folder_item.setBackground(0, QColor('#FFFFFF'))
        self.set_folder_icon(folder_item, expanded=False)
        self.folder_items[name] = folder_item
        self.chat_tree.addTopLevelItem(folder_item)
        dark = self.settings.get("theme", "light") == "dark"
        self.update_single_item_theme(folder_item, dark)
        self.save_tree_order()

    def create_chat_with_messages(self, title, messages, folder=''):
        """Создаёт чат с заданным названием, сообщениями и папкой, не запрашивая имя."""

        # Формируем безопасное имя файла
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        if not safe_title:
            safe_title = f"Чат {time.strftime('%Y%m%d_%H%M%S')}"
        filepath = os.path.join(self.chats_dir, f"{safe_title}.json")
        if os.path.exists(filepath):
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            filepath = os.path.join(self.chats_dir, f"{safe_title}_{timestamp}.json")

        new_chat = {
            'filepath': filepath,
            'title': title,
            'messages': messages,
            'folder': folder,
            'color': '#FFFFFF',
            'pinned': False,
            'background_color': None,
            'background_image': None,
            'model': self.settings.get('model', 'local-model'),
            'params': {},
            'system_prompt': '',
            'created_time': time.strftime('%d.%m.%Y %H:%M:%S'),
            'last_activity_time': time.strftime('%d.%m.%Y %H:%M:%S'),
            'total_tokens_sent': 0,
            'total_generation_time': 0.0,
            'background_image_original': None,
            'encrypted': False,
            'incognito': False
        }
        self.chats.append(new_chat)
        self.add_chat_to_tree(new_chat)
        self.save_chat_by_filepath(filepath)  # сохраняем на диск
        self.select_chat_by_filepath(filepath)
        if not os.path.exists(filepath):
            QMessageBox.warning(self, "Ошибка", "Не удалось сохранить импортированный чат.")

        # Прокручиваем импортированный чат вниз без вызова scroll_chat_to_bottom,
        # чтобы не сбивать флаги автоскроллинга
        QTimer.singleShot(100, lambda: self.chat_browser.verticalScrollBar().setValue(
            self.chat_browser.verticalScrollBar().maximum()
        ))
        self.chat_browser.verticalScrollBar().valueChanged.connect(self.on_scroll_changed)

        self._redraw_chat_with_history()

    def move_chat_to_folder(self, item, folder_name):
        filepath = item.data(0, Qt.UserRole)['filepath']
        idx = self.find_chat_index(filepath)
        if idx >= 0:
            self.chats[idx]['folder'] = folder_name
            try:
                self.save_chat_by_filepath(filepath)
            except Exception as e:
                self.logger.error(f"Ошибка сохранения чата при перемещении {filepath}: {e}")
                self.show_error("Ошибка", "Не удалось сохранить перемещение чата.")

        parent_item = None
        if folder_name:
            if folder_name not in self.folder_items:
                self.folders.append({"name": folder_name, "color": "#FFFFFF", "pinned": False})
                self.save_folders()
                folder_item = QTreeWidgetItem([folder_name])
                folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
                folder_item.setData(0, Qt.UserRole, {'type': 'folder', 'name': folder_name})
                self.folder_items[folder_name] = folder_item
                self.chat_tree.addTopLevelItem(folder_item)
            parent_item = self.folder_items[folder_name]

        current_parent = item.parent()
        if current_parent:
            current_parent.removeChild(item)
        else:
            self.chat_tree.takeTopLevelItem(self.chat_tree.indexOfTopLevelItem(item))

        if parent_item:
            parent_item.addChild(item)
            parent_item.setExpanded(True)
        else:
            self.chat_tree.addTopLevelItem(item)

        # Если исходная папка опустела, сворачиваем её и меняем иконку на закрытую
        if current_parent:
            parent_data = current_parent.data(0, Qt.UserRole)
            if parent_data and parent_data.get('type') == 'folder':
                if current_parent.childCount() == 0:
                    current_parent.setExpanded(False)
                    self.set_folder_icon(current_parent, expanded=False)

        # Обновляем иконки всех папок после изменения структуры
        self.update_all_folder_icons()
        self.save_tree_order()
        QTimer.singleShot(0, lambda: self.set_tree_items_dark(self.settings.get("theme", "light") == "dark"))

    def move_chat_to_folder_dialog(self, item):
        # Формируем список имён папок из self.folders (словари)
        folder_names = [f['name'] for f in self.folders if isinstance(f, dict) and f.get('name')]
        folder_name, ok = QInputDialog.getItem(
            self, "Переместить чат", "Выберите папку:",
            folder_names + ["Корень"], 0, False
        )
        if not ok:
            return
        if folder_name == "Корень":
            folder_name = ''
        self.move_chat_to_folder(item, folder_name)

    def delete_chat(self, item):
        filepath = item.data(0, Qt.UserRole)['filepath']
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        chat = self.chats[idx]

        # Подтверждение удаления
        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить чат «{chat['title']}»?\n"
            "Он будет перемещён в корзину.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # Добавляем чат в корзину
        trash_entry = {
            'type': 'chat',
            'filepath': filepath,
            'title': chat['title'],
            'messages': chat['messages'],
            'folder': chat.get('folder', ''),
            'background_color': chat.get('background_color'),
            'background_image': chat.get('background_image'),
            'background_image_original': chat.get('background_image_original'),
            'model': chat.get('model', self.settings.get('model', 'local-model')),
            'created_time': chat.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S')),
            'last_activity_time': chat.get('last_activity_time', time.strftime('%d.%m.%Y %H:%M:%S')),
            'total_tokens_sent': chat.get('total_tokens_sent', 0),
            'total_generation_time': chat.get('total_generation_time', 0.0),
            'deleted_time': time.strftime('%d.%m.%Y %H:%M:%S'),
            'params': chat.get('params', {}),
            'system_prompt': chat.get('system_prompt', ''),
            'color': chat.get('color', '#FFFFFF'),
            'pinned': chat.get('pinned', False),
            'tags': chat.get('tags', []),
            'encrypted': chat.get('encrypted', False),
            'password_hash': chat.get('password_hash'),
            'locked': chat.get('locked', False)
        }
        # Ищем существующую запись с тем же filepath
        existing = next((e for e in self.trash_items
                         if e.get('type') == 'chat' and e.get('filepath') == filepath), None)
        if existing:
            # Обновляем существующую запись
            idx_existing = self.trash_items.index(existing)
            self.trash_items[idx_existing] = trash_entry
        else:
            # Если не нашли – добавляем новую
            self.trash_items.append(trash_entry)

        try:
            self.save_trash()
        except Exception as e:
            self.logger.exception("Ошибка при удалении чата")
            self.show_error("Ошибка", f"Не удалось переместить чат в корзину.\nПричина: {e}")

        # Удаляем из списка чатов
        self.chats.pop(idx)

        # Удаляем из дерева
        parent = item.parent()
        if parent:
            parent.removeChild(item)
        else:
            self.chat_tree.takeTopLevelItem(self.chat_tree.indexOfTopLevelItem(item))
        self.chat_items.pop(filepath, None)

        # Сворачиваем пустую папку
        if parent:
            parent_data = parent.data(0, Qt.UserRole)
            if parent_data and parent_data.get('type') == 'folder' and parent.childCount() == 0:
                parent.setExpanded(False)
                self.set_folder_icon(parent, expanded=False)

        # Обновляем текущий индекс
        if self.current_chat_index == idx:
            self.current_chat_index = -1
            self.show_empty_chat_placeholder()
        elif self.current_chat_index > idx:
            self.current_chat_index -= 1

        # Перестраиваем дерево, чтобы убрать разделитель, если закреплённых не осталось
        self.reorder_tree_by_pinned()

    def delete_folder(self, item):
        folder_name = item.data(0, Qt.UserRole)['name']

        # Подтверждение удаления
        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить папку «{folder_name}»?\n"
            "Все чаты внутри будут перемещены в корзину вместе с папкой.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        # Собираем все чаты из элементов дерева внутри папки
        chats_data = []
        for i in range(item.childCount()):
            child = item.child(i)
            data = child.data(0, Qt.UserRole)
            if data and data.get('type') == 'chat':
                chats_data.append(data)

        # Удаляем старые записи этой папки из корзины
        self.trash_items = [
            e for e in self.trash_items
            if not (e.get('type') == 'folder' and e.get('name') == folder_name)
        ]

        folder_entry = {
            'type': 'folder',
            'name': folder_name,
            'chats': [],
            'deleted_time': time.strftime('%d.%m.%Y %H:%M:%S')
        }

        for chat_data in chats_data:
            filepath = chat_data['filepath']
            full_chat = None
            for c in self.chats:
                if c['filepath'] == filepath:
                    full_chat = c
                    break

            if full_chat:
                folder_entry['chats'].append({
                    'filepath': full_chat['filepath'],
                    'title': full_chat['title'],
                    'messages': full_chat['messages'],
                    'folder': folder_name,
                    'background_color': full_chat.get('background_color'),
                    'background_image': full_chat.get('background_image'),
                    'background_image_original': full_chat.get('background_image_original'),
                    'model': full_chat.get('model', self.settings.get('model', 'local-model')),
                    'params': full_chat.get('params', {}),
                    'system_prompt': full_chat.get('system_prompt', ''),
                    'color': full_chat.get('color', '#FFFFFF'),
                    'pinned': full_chat.get('pinned', False),
                    'tags': full_chat.get('tags', []),
                    'encrypted': full_chat.get('encrypted', False),
                    'password_hash': full_chat.get('password_hash'),
                    'locked': full_chat.get('locked', False),
                    'deleted_time': time.strftime('%d.%m.%Y %H:%M:%S'),
                    'created_time': full_chat.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S')),
                    'last_activity_time': full_chat.get('last_activity_time', time.strftime('%d.%m.%Y %H:%M:%S')),
                    'total_tokens_sent': full_chat.get('total_tokens_sent', 0),
                    'total_generation_time': full_chat.get('total_generation_time', 0.0)
                })
                self.chats.remove(full_chat)
            else:
                folder_entry['chats'].append({
                    'filepath': filepath,
                    'title': chat_data.get('title', ''),
                    'messages': [],
                    'folder': folder_name,
                    'background_color': None,
                    'background_image': None,
                    'model': self.settings.get('model', 'local-model'),
                    'deleted_time': time.strftime('%d.%m.%Y %H:%M:%S')
                })

        self.trash_items.append(folder_entry)
        try:
            self.save_trash()
        except Exception as e:
            self.logger.error(f"Ошибка сохранения корзины при удалении папки {folder_name}: {e}")
            self.show_error("Ошибка", "Не удалось сохранить корзину.")

        # Удаляем чаты из дерева
        for i in reversed(range(item.childCount())):
            child = item.child(i)
            data = child.data(0, Qt.UserRole)
            if data and data.get('type') == 'chat':
                item.removeChild(child)
                self.chat_items.pop(data['filepath'], None)

        # Удаляем папку из списка папок и дерева
        self.folders = [f for f in self.folders if not (isinstance(f, dict) and f.get('name') == folder_name)]
        self.save_folders()
        self.folder_items.pop(folder_name, None)
        self.chat_tree.takeTopLevelItem(self.chat_tree.indexOfTopLevelItem(item))

        # Перестраиваем дерево, чтобы убрать разделитель, если закреплённых не осталось
        self.reorder_tree_by_pinned()

    def archive_chat(self, filepath):
        """Перемещает чат в архив.
        Для зашифрованных locked-чатов запрашивает пароль, чтобы сохранить флаг archived в файле."""

        idx = self.find_chat_index(filepath)
        if idx < 0:
            return False

        chat = self.chats[idx]

        # === Если чат зашифрован и не разблокирован — спрашиваем пароль ===
        if chat.get('encrypted', False) and filepath not in self._session_unlocked:
            reply = QMessageBox.question(
                self, "Шифрованный чат",
                "Чат зашифрован. Чтобы сохранить архивацию в файле, нужно ввести пароль.\n\n"
                "Ввести пароль сейчас?\n\n"
                "(Если отказаться — архивация сохранится только в текущей сессии,\n"
                "и после перезапуска чат вернётся в обычный список.)",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                password, ok = QInputDialog.getText(
                    self, "Пароль",
                    f"Введите пароль для чата «{chat['title']}»:",
                    QLineEdit.Password
                )
                if ok and password:
                    try:
                        with open(filepath, 'rb') as f:
                            content = f.read()
                        data = json.loads(self._decrypt_with_password(content, password).decode('utf-8'))
                        self._current_chat_password = password

                        chat['title'] = data.get('title', chat['title'])
                        chat['messages'] = data.get('messages', [])
                        chat['folder'] = data.get('folder', '')
                        chat['background_color'] = data.get('background_color')
                        chat['background_image'] = data.get('background_image')
                        chat['background_image_original'] = data.get('background_image_original')
                        chat['model'] = data.get('model', self.settings.get('model', 'local-model'))
                        chat['created_time'] = data.get('created_time',
                                                        chat.get('created_time',
                                                                 time.strftime('%d.%m.%Y %H:%M:%S')))
                        chat['last_activity_time'] = data.get('last_activity_time',
                                                              chat.get('last_activity_time',
                                                                       time.strftime('%d.%m.%Y %H:%M:%S')))
                        chat['total_tokens_sent'] = data.get('total_tokens_sent', 0)
                        chat['total_generation_time'] = data.get('total_generation_time', 0.0)
                        chat['color'] = data.get('color', '#FFFFFF')
                        chat['params'] = data.get('params', {})
                        chat['system_prompt'] = data.get('system_prompt', '')
                        chat['pinned'] = data.get('pinned', False)
                        chat['tags'] = data.get('tags', [])
                        chat['password_hash'] = data.get('password_hash')
                        chat['locked'] = False
                        self._session_unlocked.add(filepath)
                    except Exception as e:
                        self.logger.warning(f"Ошибка расшифровки при архивации {filepath}: {e}")
                        QMessageBox.warning(
                            self, "Ошибка",
                            "Неверный пароль. Архивация сохранится только в памяти."
                        )

        # === Перемещаем в архив ===
        chat = self.chats.pop(idx)
        chat['archived'] = True
        self.archived_chats.append(chat)

        # Удаляем из дерева
        item = self.chat_items.pop(filepath, None)
        if item:
            parent = item.parent()
            if parent:
                parent.removeChild(item)
                if parent.childCount() == 0:
                    parent.setExpanded(False)
                    self.set_folder_icon(parent, expanded=False)
            else:
                self.chat_tree.takeTopLevelItem(self.chat_tree.indexOfTopLevelItem(item))

        try:
            # save_chat_by_filepath сам решит: перезаписывать файл или нет
            self.save_chat_by_filepath(filepath)
        except Exception as e:
            self.logger.error(f"Ошибка сохранения чата при архивации {filepath}: {e}")
            self.show_error("Ошибка", "Не удалось сохранить архивированный чат.")

        self.reorder_tree_by_pinned()
        self.save_tree_order()
        return True

    def unarchive_chat(self, filepath):
        """Восстанавливает чат из архива. Зашифрованный чат остаётся locked —
        пароль запросится при клике."""

        for i, chat in enumerate(self.archived_chats):
            if chat['filepath'] == filepath:

                # === Убираем из архива, добавляем в активные ===
                chat['archived'] = False
                chat['folder'] = ''
                chat['pinned'] = False

                # Если чат зашифрован — оставляем его locked, чтобы пароль
                # запрашивался при клике, а не сейчас
                if chat.get('encrypted', False):
                    chat['locked'] = True
                    chat['messages'] = []
                    self._session_unlocked.discard(filepath)

                self.archived_chats.pop(i)
                self.chats.append(chat)

                # Отключаем перерисовку
                self.chat_tree.setUpdatesEnabled(False)
                self.chat_tree.blockSignals(True)
                try:
                    self.add_chat_to_tree(chat)
                    self.reorder_tree_by_pinned()
                    self.save_tree_order()
                    self.set_tree_items_dark(self.settings.get("theme", "light") == "dark")
                except Exception as e:
                    self.logger.error(f"Ошибка восстановления чата {filepath}: {e}")
                    self.show_error("Ошибка", "Не удалось восстановить чат из архива.")
                    return False
                finally:
                    self.chat_tree.blockSignals(False)
                    self.chat_tree.setUpdatesEnabled(True)

                self.statusBar().showMessage(
                    f"Чат «{chat.get('title', '')}» восстановлен из архива"
                )

                # === НЕ открываем чат автоматически ===
                # Если зашифрован — пользователь кликнет и введёт пароль.
                # Если не зашифрован — можно сразу открыть.
                if not chat.get('encrypted', False):
                    self.select_chat_by_filepath(filepath)
                return True
        return False

    def show_archive_dialog(self):
        dialog = ArchiveDialog(self.archived_chats, self)
        dialog.exec()

    def show_trash_dialog(self):
        """Открывает диалог корзины."""

        dialog = TrashDialog(self.trash_items, self)
        if dialog.exec() == QDialog.Accepted:
            self.trash_items = dialog.get_updated_trash()
            self.save_trash()
            # Обновить дерево после восстановления
            self.refresh_tree_from_data()

    def restore_chat(self, filepath):
        for entry in self.trash_items:
            if entry['type'] == 'chat' and entry['filepath'] == filepath:
                chat_data = {
                    'filepath': entry['filepath'],
                    'title': entry['title'],
                    'messages': entry['messages'],
                    'folder': '',
                    'background_color': entry.get('background_color'),
                    'background_image': entry.get('background_image'),
                    'background_image_original': entry.get('background_image_original'),
                    'model': entry.get('model', self.settings.get('model', 'local-model')),
                    'params': entry.get('params', {}),
                    'system_prompt': entry.get('system_prompt', ''),
                    'color': entry.get('color', '#FFFFFF'),
                    'pinned': entry.get('pinned', False),
                    'tags': entry.get('tags', []),
                    'encrypted': entry.get('encrypted', False),
                    'password_hash': entry.get('password_hash'),
                    'locked': entry.get('locked', False),
                    'created_time': entry.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S')),
                    'last_activity_time': entry.get('last_activity_time', time.strftime('%d.%m.%Y %H:%M:%S')),
                    'total_tokens_sent': entry.get('total_tokens_sent', 0),
                    'total_generation_time': entry.get('total_generation_time', 0.0),
                }
                self.chats.append(chat_data)
                self.add_chat_to_tree(chat_data)

                # Если зашифрован, но пароль не введён — не трогаем файл
                if chat_data.get('encrypted', False):
                    pw = getattr(self, '_current_chat_password', None)
                    if not pw:
                        chat_data['locked'] = True
                        chat_data['messages'] = []
                        self.logger.info(
                            f"Восстановлен зашифрованный чат {entry['filepath']} — "
                            f"файл не перезаписан до ввода пароля."
                        )
                        self.trash_items.remove(entry)
                        break

                self.save_chat_by_filepath(entry['filepath'])
                self.trash_items.remove(entry)
                break
        self.save_trash()
        self.save_tree_order()

    def restore_folder(self, folder_name):
        """Восстанавливает папку и все чаты внутри неё из корзины."""

        # Находим все чаты, которые лежали в этой папке
        chats_in_folder = [entry for entry in self.trash_items if
                           entry['type'] == 'chat' and entry.get('folder') == folder_name]
        # Восстанавливаем папку, если её нет в дереве
        if folder_name not in self.folder_items:
            folder_item = QTreeWidgetItem([folder_name])
            folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
            folder_item.setData(0, Qt.UserRole, {'type': 'folder', 'name': folder_name})
            self.folder_items[folder_name] = folder_item
            self.chat_tree.addTopLevelItem(folder_item)
            self.folders.append({"name": folder_name, "color": "#FFFFFF", "pinned": False})
        # Восстанавливаем каждый чат
        for entry in chats_in_folder:
            self.restore_chat(entry['filepath'])
        # Удаляем папку из корзины, если она была сохранена как отдельная запись (мы не сохраняем папки в корзину, только чаты)
        self.save_folders()

    def permanently_delete(self, filepath):
        """Окончательно удаляет чат из корзины (и с диска)."""
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except Exception as e:
            print(f"Ошибка удаления файла {filepath}: {e}")
        self.trash_items = [entry for entry in self.trash_items if entry.get('filepath') != filepath]
        self.save_trash()

    def permanently_delete_entry(self, entry):
        """Удаляет конкретную запись из корзины безвозвратно."""

        try:
            if entry.get('type') == 'chat':
                filepath = entry.get('filepath')
                if filepath and os.path.exists(filepath):
                    os.remove(filepath)
            elif entry.get('type') == 'folder':
                # Удаляем файлы всех чатов внутри папки
                for chat_entry in entry.get('chats', []):
                    fp = chat_entry.get('filepath')
                    if fp and os.path.exists(fp):
                        os.remove(fp)
            # Удаляем саму запись из списка
            if entry in self.trash_items:
                self.trash_items.remove(entry)
            self.save_trash()
        except Exception as e:
            self.logger.error(f"Ошибка при окончательном удалении из корзины: {e}")
            self.show_error("Ошибка", "Не удалось окончательно удалить элемент.")

    def permanently_delete_archived(self, filepath):
        """Удаляет чат из архива безвозвратно (файл тоже удаляется)."""

        for i, chat in enumerate(self.archived_chats):
            if chat['filepath'] == filepath:
                try:
                    if os.path.exists(filepath):
                        os.remove(filepath)
                except Exception as e:
                    print(f"Ошибка удаления файла {filepath}: {e}")
                self.archived_chats.pop(i)
                return True
        return False

    def show_tree_context_menu(self, pos):
        # Определяем источник сигнала (viewport или сам tree) и координаты
        if self.sender() == self.chat_tree.viewport():
            item_pos = self.chat_tree.viewport().mapTo(self.chat_tree, pos)
            item = self.chat_tree.itemAt(item_pos)
            global_pos = self.chat_tree.viewport().mapToGlobal(pos)
        else:
            item = self.chat_tree.itemAt(pos)
            global_pos = self.chat_tree.mapToGlobal(pos)

        menu = QMenu(self)

        if item is None:
            # === Клик по пустому месту ===
            menu.addAction("Создать чат", lambda: self.create_new_chat(in_root=True))
            menu.addAction("Создать папку", self.create_new_folder)
            menu.exec(global_pos)
            return

        data = item.data(0, Qt.UserRole)
        if data and data.get('type') == 'chat':
            filepath = data['filepath']
            chat = None
            for c in self.chats:
                if c['filepath'] == filepath:
                    chat = c
                    break

            rename_action = menu.addAction("Переименовать")
            move_to_root = menu.addAction("Переместить в корень")
            move_to_folder = menu.addAction("Переместить в папку...")
            archive_action = menu.addAction("Архивировать")
            copy_dialog_action = menu.addAction("Копировать диалог")

            incognito_action = menu.addAction(
                "Отключить режим инкогнито" if chat.get('incognito', False) else "Включить режим инкогнито"
            )

            if chat.get('tags'):
                tags_action = menu.addAction("Изменить тег")
            else:
                tags_action = menu.addAction("Добавить тег")
            tags_action.triggered.connect(lambda: self.edit_chat_tags(filepath))

            if chat and chat.get('encrypted', False):
                if chat.get('locked', False):
                    # Чат зашифрован, но не разблокирован — расшифровывать нельзя
                    encrypt_action = menu.addAction("Не шифровать чат (требуется пароль)")
                    encrypt_action.setEnabled(False)
                    change_password_action = menu.addAction("Изменить пароль (требуется пароль)")
                    change_password_action.setEnabled(False)
                else:
                    encrypt_action = menu.addAction("Не шифровать чат")
                    change_password_action = menu.addAction("Изменить пароль")
            else:
                encrypt_action = menu.addAction("Шифровать чат")
                change_password_action = None

            delete_action = menu.addAction("Удалить чат")

            color_menu = menu.addMenu("Цвет")
            for color_name, color_hex in [("Белый", "#FFFFFF"), ("Жёлтый", "#FFF9C4"),
                                          ("Розовый", "#F8BBD0"), ("Фиолетовый", "#E1BEE7"),
                                          ("Синий", "#BBDEFB"), ("Зелёный", "#C8E6C9"),
                                          ("Серый", "#E0E0E0"), ("Без цвета", None)]:
                color_action = color_menu.addAction(color_name)
                color_action.triggered.connect(lambda checked, hex=color_hex, it=item: self.set_item_color(it, hex))

            pin_action = menu.addAction(
                self.icon_pin,
                "Открепить" if self.is_item_pinned(item) else "Закрепить"
            )
            pin_action.triggered.connect(lambda: self.toggle_pin(item))

            action = menu.exec(global_pos)
            if action == rename_action:
                self.rename_chat(item)
            elif action == move_to_root:
                self.move_chat_to_folder(item, '')
            elif action == move_to_folder:
                self.move_chat_to_folder_dialog(item)
            elif action == archive_action:
                self.archive_chat(filepath)
            elif action == copy_dialog_action:
                self.copy_full_dialog(chat)
            elif action == delete_action:
                self.delete_chat(item)
            elif action == encrypt_action:
                self.toggle_chat_encryption(filepath)
            elif action == incognito_action:
                self.toggle_incognito(filepath)
            elif change_password_action is not None and action == change_password_action:
                self.change_chat_password(filepath)

        elif data and data.get('type') == 'folder':
            rename_action = menu.addAction("Переименовать папку")
            delete_action = menu.addAction("Удалить папку")

            color_menu = menu.addMenu("Цвет")
            for color_name, color_hex in [("Белый", "#FFFFFF"), ("Жёлтый", "#FFF9C4"),
                                          ("Розовый", "#F8BBD0"), ("Фиолетовый", "#E1BEE7"),
                                          ("Синий", "#BBDEFB"), ("Зелёный", "#C8E6C9"),
                                          ("Серый", "#E0E0E0"), ("Без цвета", None)]:
                color_action = color_menu.addAction(color_name)
                color_action.triggered.connect(lambda checked, hex=color_hex, it=item: self.set_item_color(it, hex))

            pin_action = menu.addAction("Открепить" if self.is_item_pinned(item) else "Закрепить")
            pin_action.triggered.connect(lambda: self.toggle_pin(item))

            action = menu.exec(global_pos)
            if action == rename_action:
                self.rename_folder(item)
            elif action == delete_action:
                self.delete_folder(item)

    def get_folder_color(self, folder_name):
        for folder in self.folders:
            if isinstance(folder, dict) and folder.get('name') == folder_name:
                return folder.get('color', '#FFFFFF')
        return None

    def is_folder_pinned(self, folder_name):
        for folder in self.folders:
            if isinstance(folder, dict) and folder.get('name') == folder_name:
                return folder.get('pinned', False)
        return False

    def save_tree_order(self):
        """Сохраняет порядок элементов дерева в файл tree_order.json."""

        structure = self.get_tree_structure()
        try:
            with open('tree_order.json', 'w', encoding='utf-8') as f:
                json.dump(structure, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.logger.error(f"Ошибка сохранения порядка дерева: {e}")
            self.show_error("Ошибка сохранения", f"Не удалось сохранить порядок дерева.\nПричина: {e}")

    def load_tree_order(self):
        if os.path.exists('tree_order.json'):
            data = self._safe_json_load('tree_order.json', None)
            if isinstance(data, list):
                return data
            else:
                self.logger.warning("tree_order.json повреждён, порядок по умолчанию.")
                print("⚠️ tree_order.json повреждён, порядок по умолчанию.")
                return None
        return None

    def get_tree_structure(self):
        """Возвращает список, описывающий текущий порядок элементов в дереве."""

        def get_children(item):
            children = []
            for i in range(item.childCount()):
                child = item.child(i)
                child_data = child.data(0, Qt.UserRole)
                if child_data and child_data.get('type') == 'folder':
                    children.append({
                        'type': 'folder',
                        'name': child_data.get('name'),
                        'children': get_children(child)
                    })
                elif child_data and child_data.get('type') == 'chat':
                    children.append({
                        'type': 'chat',
                        'filepath': child_data.get('filepath')
                    })
            return children

        top = []
        for i in range(self.chat_tree.topLevelItemCount()):
            item = self.chat_tree.topLevelItem(i)
            data = item.data(0, Qt.UserRole)
            if data and data.get('type') == 'folder':
                top.append({
                    'type': 'folder',
                    'name': data.get('name'),
                    'children': get_children(item)
                })
            elif data and data.get('type') == 'chat':
                top.append({
                    'type': 'chat',
                    'filepath': data.get('filepath')
                })
        return top

    def apply_tree_order(self, order):
        """Строит дерево заново согласно сохранённому порядку."""

        # Очищаем дерево и словари
        self.chat_tree.clear()
        self.folder_items.clear()
        self.chat_items.clear()

        def add_folder(folder_name, parent_item=None):
            folder_params = next((f for f in self.folders if isinstance(f, dict) and f.get('name') == folder_name),
                                 None)
            pinned = folder_params.get('pinned', False) if folder_params else False
            color = folder_params.get('color', '#FFFFFF') if folder_params else '#FFFFFF'

            folder_item = QTreeWidgetItem([folder_name])
            folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
            folder_item.setData(0, Qt.UserRole,
                                {'type': 'folder', 'name': folder_name, 'pinned': pinned, 'color': color})
            if color:
                folder_item.setBackground(0, QColor(color))
            if pinned:
                folder_item.setIcon(1, self.icon_pin)
            self.set_folder_icon(folder_item, expanded=False)
            self.folder_items[folder_name] = folder_item

            if parent_item:
                parent_item.addChild(folder_item)
            else:
                self.chat_tree.addTopLevelItem(folder_item)
            return folder_item

        def add_chat(filepath, parent_item=None):
            for chat in self.chats:
                if chat['filepath'] == filepath:
                    pinned = chat.get('pinned', False)
                    color = chat.get('color', '#FFFFFF')
                    # Название чата с префиксом инкогнито
                    title = chat.get('title', '')
                    prefix = ""
                    if chat.get('incognito', False):
                        prefix += "🕶 "
                    chat_item = QTreeWidgetItem([prefix + title])
                    chat_item.setFlags(chat_item.flags() & ~Qt.ItemIsDropEnabled)
                    chat_item.setData(0, Qt.UserRole,
                                      {'type': 'chat', 'filepath': filepath, 'pinned': pinned, 'color': color})
                    if color:
                        chat_item.setBackground(0, QColor(color))
                    if pinned:
                        chat_item.setIcon(1, self.icon_pin)
                    self.set_chat_icon(chat_item)
                    self.chat_items[filepath] = chat_item

                    # Заполняем третью колонку тегами
                    tags = chat.get('tags', [])
                    if tags:
                        tags_str = ' '.join(['#' + tag for tag in tags])
                        chat_item.setText(2, tags_str)
                        chat_item.setForeground(2, QBrush(QColor("#1a73e8")))

                    if parent_item:
                        parent_item.addChild(chat_item)
                    else:
                        self.chat_tree.addTopLevelItem(chat_item)
                    break

        def process_items(items, parent_item=None):
            for item in items:
                if item['type'] == 'folder':
                    folder_item = add_folder(item['name'], parent_item)
                    process_items(item.get('children', []), folder_item)
                elif item['type'] == 'chat':
                    add_chat(item['filepath'], parent_item)

        process_items(order)

        self.save_folders()
        self.save_tree_order()

    def on_tree_structure_changed(self):
        """Обновляет folder у всех чатов и синхронизирует закрепления."""

        def update_item(item):
            data = item.data(0, Qt.UserRole)
            if data and data.get('type') == 'chat':
                parent = item.parent()
                new_folder = ''
                if parent:
                    parent_data = parent.data(0, Qt.UserRole)
                    if parent_data and parent_data.get('type') == 'folder':
                        new_folder = parent_data.get('name', '')
                filepath = data['filepath']
                for chat in self.chats:
                    if chat['filepath'] == filepath:
                        # Сохраняем только если папка реально изменилась
                        if chat.get('folder', '') != new_folder:
                            chat['folder'] = new_folder
                            self.save_chat_by_filepath(filepath)
                        break
            for i in range(item.childCount()):
                update_item(item.child(i))

        for i in range(self.chat_tree.topLevelItemCount()):
            update_item(self.chat_tree.topLevelItem(i))

        self.save_tree_order()
        self.sync_pinned_flags_from_positions()
        self.reorder_tree_by_pinned()

    def sync_pinned_flags_from_positions(self):
        """Устанавливает флаг pinned для верхнеуровневых элементов на основе их позиции относительно разделителя."""

        separator_idx = -1
        for i in range(self.chat_tree.topLevelItemCount()):
            item = self.chat_tree.topLevelItem(i)
            if item.data(0, Qt.UserRole).get('type') == 'separator':
                separator_idx = i
                break

        for i in range(self.chat_tree.topLevelItemCount()):
            item = self.chat_tree.topLevelItem(i)
            data = item.data(0, Qt.UserRole)
            if not data or data.get('type') == 'separator':
                continue

            new_pinned = (separator_idx != -1 and i < separator_idx)

            data['pinned'] = new_pinned
            item.setData(0, Qt.UserRole, data)

            if data['type'] == 'chat':
                filepath = data['filepath']
                for chat in self.chats:
                    if chat['filepath'] == filepath:
                        chat['pinned'] = new_pinned
                        self.save_chat_by_filepath(filepath)
                        break
            elif data['type'] == 'folder':
                folder_name = data['name']
                for folder in self.folders:
                    if isinstance(folder, dict) and folder.get('name') == folder_name:
                        folder['pinned'] = new_pinned
                        self.save_folders()
                        break

    def is_above_separator(self, item):
        """Проверяет, находится ли элемент выше разделителя на верхнем уровне."""

        separator_idx = None
        for i in range(self.chat_tree.topLevelItemCount()):
            top_item = self.chat_tree.topLevelItem(i)
            top_data = top_item.data(0, Qt.UserRole)
            if top_data and top_data.get('type') == 'separator':
                separator_idx = i
                break
        if separator_idx is None:
            return False
        for i in range(self.chat_tree.topLevelItemCount()):
            if self.chat_tree.topLevelItem(i) == item:
                return i < separator_idx
        return False

    def update_pinned_status_after_drop(self, item):
        """После перетаскивания определяет новый статус закрепления и обновляет данные."""

        data = item.data(0, Qt.UserRole)
        if not data:
            return

        # Определяем, находится ли элемент внутри папки
        parent = item.parent()
        is_inside_folder = parent is not None and parent.data(0, Qt.UserRole).get('type') == 'folder'

        if is_inside_folder:
            # Элемент внутри папки: его собственный pinned всегда False (закрепление наследуется от папки)
            self.set_item_pinned(item, False)
            # Эффективное закрепление будет определяться по родителю
        else:
            # Элемент на верхнем уровне: закреплён, если выше разделителя
            is_pinned_now = self.is_above_separator(item)
            self.set_item_pinned(item, is_pinned_now)

        # Перестраиваем дерево
        self.reorder_tree_by_pinned()

    def reset_children_pinned(self, folder_item):
        """Сбрасывает собственный флаг pinned у всех дочерних элементов папки."""

        for i in range(folder_item.childCount()):
            child = folder_item.child(i)
            data = child.data(0, Qt.UserRole)
            if not data:
                continue
            if data.get('type') == 'chat':
                data['pinned'] = False
                child.setData(0, Qt.UserRole, data)
                child.setIcon(1, QIcon())
                for chat in self.chats:
                    if chat['filepath'] == data['filepath']:
                        chat['pinned'] = False
                        self.save_chat_by_filepath(data['filepath'])
                        break
            elif data.get('type') == 'folder':
                # Вложенная папка: сбрасываем её собственный pinned и рекурсивно у её детей
                data['pinned'] = False
                child.setData(0, Qt.UserRole, data)
                child.setIcon(1, QIcon())
                for folder in self.folders:
                    if isinstance(folder, dict) and folder.get('name') == data.get('name'):
                        folder['pinned'] = False
                        self.save_folders()
                        break
                self.reset_children_pinned(child)

    def update_children_pin_icons(self, folder_item, parent_pinned):
        """Рекурсивно обновляет иконки Pin у дочерних элементов папки на основе наследования."""

        for i in range(folder_item.childCount()):
            child = folder_item.child(i)
            child_data = child.data(0, Qt.UserRole)
            if not child_data:
                continue
            if child_data.get('type') == 'chat':
                pinned = child_data.get('pinned', False) or parent_pinned
                if pinned:
                    child.setIcon(1, self.icon_pin)
                else:
                    child.setIcon(1, QIcon())
            elif child_data.get('type') == 'folder':
                # Вложенная папка: учитываем её собственный pinned и родительский
                pinned = child_data.get('pinned', False) or parent_pinned
                if pinned:
                    child.setIcon(1, self.icon_pin)
                else:
                    child.setIcon(1, QIcon())
                # Рекурсивно обновляем детей вложенной папки
                self.update_children_pin_icons(child, pinned)

    def check_auto_archive(self):
        """Архивирует чаты, неактивные более N дней."""

        days = self.settings.get("auto_archive_days", 0)
        if days <= 0:
            return
        now = time.time()
        threshold = days * 24 * 3600
        to_archive = []
        for chat in self.chats:
            if chat.get('pinned', False):
                continue  # закреплённые не архивируем
            date_str = chat.get('last_activity_time', '')
            try:
                t = time.mktime(time.strptime(date_str, '%d.%m.%Y %H:%M:%S'))
            except:
                continue
            if now - t > threshold:
                to_archive.append(chat['filepath'])
        for fp in to_archive:
            self.archive_chat(fp)

    # ======================================================================
    #  УПРАВЛЕНИЕ ТЕКУЩИМ ЧАТОМ И ОТОБРАЖЕНИЕ СООБЩЕНИЙ
    # ======================================================================

    def select_chat_by_filepath(self, filepath, keep_scroll=False, scroll_target=None):
        if self._selecting_chat:
            return True
        self._selecting_chat = True
        try:
            self.close_search_panel()
            self.hide_stop_button()

            # Запираем ВСЕ зашифрованные чаты, кроме того, куда идём
            for c in self.chats:
                if c.get('encrypted', False) and c['filepath'] != filepath:
                    self._session_unlocked.discard(c['filepath'])
                    c['locked'] = True
                    c['messages'] = []

            for i, chat in enumerate(self.chats):
                if chat['filepath'] == filepath:

                    # Если чат помечен зашифрованным и не разблокирован в текущей сессии
                    if chat.get('encrypted', False) and filepath not in self._session_unlocked:

                        # === ПРОВЕРКА: файл РЕАЛЬНО зашифрован? ===
                        file_is_encrypted = False
                        try:
                            with open(filepath, 'rb') as f:
                                head = f.read(4)
                            file_is_encrypted = head.startswith(b"ENC:")
                        except OSError:
                            file_is_encrypted = False

                        if not file_is_encrypted:
                            # Флаг encrypted=True, но файл открытый → снимаем флаг
                            self.logger.warning(
                                f"Чат {filepath} помечен зашифрованным, но файл не зашифрован. "
                                f"Снимаю флаг и загружаю как обычный JSON."
                            )
                            try:
                                with open(filepath, 'r', encoding='utf-8') as f:
                                    data = json.load(f)

                                chat['title'] = data.get('title', chat['title'])
                                chat['messages'] = data.get('messages', [])
                                chat['folder'] = data.get('folder', '')
                                chat['background_color'] = data.get('background_color')
                                chat['background_image'] = data.get('background_image')
                                chat['background_image_original'] = data.get('background_image_original')
                                chat['model'] = data.get('model', self.settings.get('model', 'local-model'))
                                chat['created_time'] = data.get(
                                    'created_time', chat.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S'))
                                )
                                chat['last_activity_time'] = data.get(
                                    'last_activity_time',
                                    chat.get('last_activity_time', time.strftime('%d.%m.%Y %H:%M:%S'))
                                )
                                chat['total_tokens_sent'] = data.get('total_tokens_sent', 0)
                                chat['total_generation_time'] = data.get('total_generation_time', 0.0)
                                chat['color'] = data.get('color', '#FFFFFF')
                                chat['params'] = data.get('params', {})
                                chat['system_prompt'] = data.get('system_prompt', '')
                                chat['pinned'] = data.get('pinned', False)
                                chat['tags'] = data.get('tags', [])
                                chat['password_hash'] = data.get('password_hash')
                            except Exception as e:
                                self.logger.warning(f"Не удалось прочитать открытый JSON {filepath}: {e}")
                                QMessageBox.warning(self, "Ошибка", f"Файл повреждён или пуст:\n{e}")
                                return False

                            # Снимаем флаг шифрования в памяти
                            chat['encrypted'] = False
                            chat['locked'] = False
                            self._session_unlocked.add(filepath)

                            # Пересохраняем файл как обычный JSON (чтобы флаг совпадал с содержимым)
                            self.save_chat_by_filepath(filepath)

                        else:
                            # === Файл действительно зашифрован — спрашиваем пароль ===
                            password, ok = QInputDialog.getText(
                                self, "Пароль",
                                f"Введите пароль для чата «{chat['title']}»:",
                                QLineEdit.Password
                            )
                            if not ok or not password:
                                return False
                            try:
                                with open(filepath, 'rb') as f:
                                    content = f.read()
                                data = json.loads(self._decrypt_with_password(content, password).decode('utf-8'))
                                self._current_chat_password = password

                                chat['title'] = data.get('title', chat['title'])
                                chat['messages'] = data.get('messages', [])
                                chat['folder'] = data.get('folder', '')
                                chat['background_color'] = data.get('background_color')
                                chat['background_image'] = data.get('background_image')
                                chat['background_image_original'] = data.get('background_image_original')
                                chat['model'] = data.get('model', self.settings.get('model', 'local-model'))
                                chat['created_time'] = data.get(
                                    'created_time', chat.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S'))
                                )
                                chat['last_activity_time'] = data.get(
                                    'last_activity_time',
                                    chat.get('last_activity_time', time.strftime('%d.%m.%Y %H:%M:%S'))
                                )
                                chat['total_tokens_sent'] = data.get('total_tokens_sent', 0)
                                chat['total_generation_time'] = data.get('total_generation_time', 0.0)
                                chat['color'] = data.get('color', '#FFFFFF')
                                chat['params'] = data.get('params', {})
                                chat['system_prompt'] = data.get('system_prompt', '')
                                chat['pinned'] = data.get('pinned', False)
                                chat['tags'] = data.get('tags', [])
                                chat['password_hash'] = data.get('password_hash')
                                chat['locked'] = False
                                self._session_unlocked.add(filepath)
                            except Exception as e:
                                self.logger.warning(f"Ошибка расшифровки {filepath}: {e}")
                                QMessageBox.warning(self, "Ошибка", "Неверный пароль.")
                                return False

                    # Устанавливаем текущий чат
                    self.current_chat_index = i
                    self._restore_scroll_value = None
                    self.pending_scroll_target = scroll_target

                    with self.busy("Открытие чата…"):
                        # 1. Убираем placeholder и старые сообщения
                        self.clear_chat_display()

                        # 2. Применяем обои на пустую область
                        self.apply_chat_background(chat)
                        self.chat_browser.setStyleSheet("")
                        self.apply_chat_background(chat)
                        self._current_search_idx = -1

                        # 3. Отложенное построение сообщений
                        self._redraw_chat_with_history(keep_scroll=keep_scroll)

                    item = self.chat_items.get(filepath)
                    if item and self.chat_tree.currentItem() is not item:
                        self.chat_tree.setCurrentItem(item)

                    self.update_chat_settings_button_visibility()
                    self.update_model_label()

                    if self.is_generating and filepath == self.generation_chat_filepath:
                        self.show_stop_button()
                    else:
                        self._remove_thinking_indicator()
                        self.hide_stop_button()

                    if filepath in self.unread_chats:
                        self.unread_chats.remove(filepath)
                        self.update_unread_indicators()
                        self.refresh_unread_ui()

                    return True

            self.current_chat_index = -1
            self.show_empty_chat_placeholder()
            self.update_chat_settings_button_visibility()
            self.update_model_label()
            return False
        finally:
            self._selecting_chat = False

    def _encrypt_with_password(self, data: bytes, password: str) -> bytes:
        import os as _os
        from cryptography.fernet import Fernet
        salt = _os.urandom(16).hex()
        iterations = self.settings.get("encryption_kdf_iterations", 200000)
        key = self._derive_key_from_password(password, salt, iterations)
        return f"ENC:{salt}:".encode() + Fernet(key).encrypt(data)

    def _decrypt_with_password(self, data: bytes, password: str) -> bytes:
        if not data.startswith(b"ENC:"):
            return data
        if len(data) > 37 and data[36:37] == b":":
            salt_hex = data[4:36].decode('ascii')
            int(salt_hex, 16)
            iterations = self.settings.get("encryption_kdf_iterations", 200000)
            key = self._derive_key_from_password(password, salt_hex, iterations)
            from cryptography.fernet import Fernet
            return Fernet(key).decrypt(data[37:])
        salt = self.settings.get("encryption_salt", "")
        iterations = self.settings.get("encryption_kdf_iterations", 200000)
        key = self._derive_key_from_password(password, salt, iterations)
        from cryptography.fernet import Fernet
        return Fernet(key).decrypt(data[4:])

    def find_chat_index(self, filepath):
        for i, chat in enumerate(self.chats):
            if chat['filepath'] == filepath:
                return i
        return -1

    def on_chat_selected(self):
        if self._selecting_chat:
            return
        item = self.chat_tree.currentItem()
        if item is None:
            return
        data = item.data(0, Qt.UserRole)
        if data and data.get('type') == 'chat':
            filepath = data['filepath']
            # Сохраняем текущий выделенный элемент и позицию скролла
            previous_item = self.chat_tree.currentItem()
            scroll_value = self.chat_tree.verticalScrollBar().value()

            success = self.select_chat_by_filepath(filepath)
            if not success:
                # Восстанавливаем предыдущее выделение и прокрутку
                self._selecting_chat = True
                if previous_item:
                    self.chat_tree.setCurrentItem(previous_item)
                self.chat_tree.verticalScrollBar().setValue(scroll_value)
                self._selecting_chat = False
        else:
            # Запираем все расшифрованные чаты при уходе из чата в папку
            for c in self.chats:
                if c.get('encrypted', False) and c['filepath'] in self._session_unlocked:
                    self._session_unlocked.discard(c['filepath'])
                    c['locked'] = True
                    c['messages'] = []
            self.current_chat_index = -1
            self.show_empty_chat_placeholder()

    def show_empty_chat_placeholder(self):
        """Показывает в области чата сообщение о необходимости выбора чата."""

        self.clear_chat_display()
        self.pinned_panel.setVisible(False)
        self.hide_stop_button()
        if self.thinking_cursor is not None:
            self._remove_thinking_indicator()

        # Добавляем заглушку в layout вместо QListWidget
        placeholder_label = QLabel("Выберите чат или создайте новый")
        placeholder_label.setAlignment(Qt.AlignCenter)
        placeholder_label.setStyleSheet("color: #888; font-size: 18px;")
        placeholder_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.chat_messages_layout.addWidget(placeholder_label)
        self.chat_messages_layout.setAlignment(Qt.AlignTop)

        self.close_search_panel()
        self.update_chat_settings_button_visibility()
        self.apply_chat_background(None)
        self.refresh_unread_ui()

    def clear_chat_display(self):
        while self.chat_messages_layout.count():
            item = self.chat_messages_layout.takeAt(0)
            if item.widget():
                w = item.widget()
                w.hide()
                w.setParent(None)
                w.deleteLater()
            elif item.layout():
                sub_layout = item.layout()
                while sub_layout.count():
                    sub_item = sub_layout.takeAt(0)
                    if sub_item.widget():
                        w = sub_item.widget()
                        w.hide()
                        w.setParent(None)
                        w.deleteLater()
        self.thinking_widget = None
        self.current_response_widget = None
        self.message_widgets.clear()
        self.load_more_widget = None

    def _redraw_chat_with_history(self, keep_scroll=False, reset_loaded_count=True):
        if self.current_chat_index < 0:
            return

        chat = self.chats[self.current_chat_index]
        filepath = chat['filepath']

        if reset_loaded_count:
            self.loaded_count[filepath] = 0

        # Токен — только он защищает от двойного построения
        self._redraw_token = getattr(self, '_redraw_token', 0) + 1
        token = self._redraw_token

        QTimer.singleShot(0, lambda: self._build_messages_async(chat, keep_scroll=keep_scroll, token=token))

    def _build_messages_async(self, chat, keep_scroll=False, token=None):
        if token is not None and token != getattr(self, '_redraw_token', 0):
            return
        if self.current_chat_index < 0:
            return
        if self.chats[self.current_chat_index]['filepath'] != chat['filepath']:
            return

        # Отключаем перерисовку на время построения
        self.chat_messages_widget.setUpdatesEnabled(False)

        try:
            self.clear_chat_display()
            self.preview_images.clear()
            self.preview_image_counter = 0
            self.message_positions = []

            filepath = chat['filepath']
            all_messages = chat['messages']
            total = len(all_messages)

            if filepath not in self.loaded_count:
                self.loaded_count[filepath] = 0

            # Если есть цель прокрутки и она в этом чате, загружаем все сообщения,
            # чтобы гарантировать наличие виджета для прокрутки.
            if self.pending_scroll_target and self.pending_scroll_target[0] == filepath:
                self.loaded_count[filepath] = total
                start_idx = 0
            else:
                if self.loaded_count[filepath] == 0:
                    start_idx = max(0, total - self.lazy_load_limit)
                    self.loaded_count[filepath] = total - start_idx
                else:
                    start_idx = total - self.loaded_count[filepath]

            # Добавляем сообщения
            for idx in range(start_idx, total):
                self._add_message_widget(all_messages[idx], idx)

            if self.is_generating and filepath == self.generation_chat_filepath:
                self._show_thinking_indicator()

            self.update_pinned_panel()

            # --- Восстановление прокрутки ---
            # 1. Если задана конкретная цель (сообщение) – прокручиваем к ней
            if self.pending_scroll_target:
                filepath, msg_idx = self.pending_scroll_target
                self.pending_scroll_target = None
                QTimer.singleShot(150, lambda fp=filepath, mi=msg_idx: self._scroll_to_message(fp, mi))
                return

            # 2. Если задано относительное положение (старый механизм) – используем его
            if self._restore_scroll_value is not None:
                saved_ratio = self._restore_scroll_value
                self._restore_scroll_value = None

                def restore_scroll():
                    vsb = self.chat_scroll_area.verticalScrollBar()
                    if vsb.maximum() > 0:
                        value = int(saved_ratio * vsb.maximum())
                        vsb.setValue(value)

                QTimer.singleShot(150, restore_scroll)
                return

            # 3. Если не требуется восстановление и keep_scroll=False – прокручиваем вниз
            if not keep_scroll:
                self.scroll_chat_to_bottom()
        finally:
            # Включаем перерисовку
            self.chat_messages_widget.setUpdatesEnabled(True)

            # Принудительно перерисовываем viewport, чтобы обои применились на всю область
            self.chat_scroll_area.viewport().update()
            self.chat_scroll_area.viewport().repaint()
            self.chat_messages_widget.updateGeometry()

    def _add_message_widget(self, msg, idx):
        role = msg.get('role')
        content = msg.get('content', '')
        if isinstance(content, list):
            text_parts = []
            for item in content:
                if isinstance(item, dict):
                    if 'text' in item:
                        text_parts.append(item['text'])
                    elif 'image_url' in item:
                        text_parts.append('[Изображение]')
                else:
                    text_parts.append(str(item))
            content = '\n'.join(text_parts)

        time_str = msg.get('time', '')
        reasoning = msg.get('reasoning', '') if (role == 'assistant' and self.show_reasoning) else ''
        pinned = msg.get('pinned', False)
        starred = msg.get('starred', False)

        if role == 'assistant':
            display_content = self._markdown_to_html(content)
        else:
            display_content = self._escape_html(content).replace('\n', '<br>')

        widget = MessageWidget(role, display_content, time_str, reasoning, pinned, idx, starred=starred,
                               parent=self.chat_messages_widget, comment=msg.get('comment', ''), reaction=msg.get('reaction'), quote=msg.get('quote'))
        widget.edit_clicked.connect(self._on_message_action_edit)
        widget.copy_clicked.connect(self._on_message_action_copy)
        widget.regenerate_clicked.connect(self._on_message_action_regenerate)
        widget.unpin_clicked.connect(self._on_message_action_unpin)
        widget.context_menu_requested.connect(lambda idx, gp, w: self.show_message_context_menu_at(idx, gp, w))
        widget.attachment_clicked.connect(self.show_attachment_preview)
        widget.decrypt_func = self._decrypt_bytes
        widget.thumbnails_dir = self.thumbnails_dir

        if 'attachments' in msg and msg['attachments']:
            widget.set_attachments(msg['attachments'])

        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(0)
        if role == 'user':
            row_layout.addStretch(1)
            row_layout.addWidget(widget)
        else:
            row_layout.addWidget(widget)
            row_layout.addStretch(1)

        self.chat_messages_layout.addLayout(row_layout)

        # === Кнопка "Показать схему" под сообщением, если есть схема или код ===
        def _contains_scheme(text):
            if '```dot' in text:
                return True
            has_arrow = re.search(r'(->|→|–>|—>|-->)', text)
            has_keyword = re.search(r'схем|блок-схем|диаграм|flowchart|diagram|граф|алгоритм', text, re.IGNORECASE)
            return has_arrow and has_keyword

        # Проверяем наличие Python-кода
        python_code = self.extract_python_code_from_message(content) if role == 'assistant' else ""
        has_scheme = _contains_scheme(content) or python_code

        if role == 'assistant' and has_scheme:
            btn_show_scheme = QPushButton("Показать схему")
            btn_show_scheme.setCursor(Qt.PointingHandCursor)
            btn_show_scheme.setIcon(QIcon("Image/Scheme.png"))
            btn_show_scheme.setIconSize(QSize(16, 16))
            btn_show_scheme.setMinimumHeight(32)
            btn_show_scheme.setMinimumWidth(140)
            btn_show_scheme.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
            btn_show_scheme.setStyleSheet("""
                QPushButton {
                    background-color: #e8f0fe;
                    border: 1px solid #1a73e8;
                    border-radius: 6px;
                    padding: 6px 12px;
                    color: #1a73e8;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background-color: #d2e3fc;
                }
            """)
            btn_show_scheme.clicked.connect(
                lambda checked, c=content, pc=python_code: self.show_diagram_from_text(c, python_code=pc)
            )

            scheme_layout = QHBoxLayout()
            scheme_layout.setContentsMargins(0, 6, 0, 0)
            scheme_layout.setSpacing(0)
            scheme_layout.setAlignment(Qt.AlignLeft)
            scheme_layout.addWidget(btn_show_scheme)
            self.chat_messages_layout.addLayout(scheme_layout)
        # ============================================================

        # Сохраняем ссылку для навигации
        self.message_widgets.append((self.chats[self.current_chat_index]['filepath'], idx, widget))

    def _add_message_widget_to_start(self, msg, idx):
        role = msg.get('role')
        content = msg.get('content', '')
        if isinstance(content, list):
            text_parts = []
            for item in content:
                if isinstance(item, dict):
                    if 'text' in item:
                        text_parts.append(item['text'])
                    elif 'image_url' in item:
                        text_parts.append('[Изображение]')
                else:
                    text_parts.append(str(item))
            content = '\n'.join(text_parts)

        time_str = msg.get('time', '')
        reasoning = msg.get('reasoning', '') if (role == 'assistant' and self.show_reasoning) else ''
        pinned = msg.get('pinned', False)
        starred = msg.get('starred', False)

        if role == 'assistant':
            display_content = self._markdown_to_html(content)
        else:
            display_content = self._escape_html(content).replace('\n', '<br>')

        widget = MessageWidget(role, display_content, time_str, reasoning, pinned, idx, starred=starred,
                               parent=self.chat_messages_widget, image_cache=self.image_cache, comment=msg.get('comment', ''), reaction=msg.get('reaction'), quote=msg.get('quote'))
        widget.edit_clicked.connect(self._on_message_action_edit)
        widget.copy_clicked.connect(self._on_message_action_copy)
        widget.regenerate_clicked.connect(self._on_message_action_regenerate)
        widget.unpin_clicked.connect(self._on_message_action_unpin)
        widget.context_menu_requested.connect(lambda idx, gp, w: self.show_message_context_menu_at(idx, gp, w))
        widget.decrypt_func = self._decrypt_bytes
        widget.attachment_clicked.connect(self.show_attachment_preview)
        if 'attachments' in msg and msg['attachments']:
            widget.set_attachments(msg['attachments'])

        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(0)
        if role == 'user':
            row_layout.addStretch(1)
            row_layout.addWidget(widget)
        else:
            row_layout.addWidget(widget)
            row_layout.addStretch(1)

        # Вставляем в начало layout
        self.chat_messages_layout.insertLayout(0, row_layout)

        # Сохраняем ссылку
        self.message_widgets.append((self.chats[self.current_chat_index]['filepath'], idx, widget))

    def _load_more_messages(self, filepath):
        if filepath not in self.loaded_count:
            return
        chat = None
        for c in self.chats:
            if c['filepath'] == filepath:
                chat = c
                break
        if not chat:
            return

        total = len(chat['messages'])
        current_loaded = self.loaded_count[filepath]
        if current_loaded >= total:
            return

        # Запоминаем текущее положение скролла
        scrollbar = self.chat_scroll_area.verticalScrollBar()
        old_value = scrollbar.value()
        old_max = scrollbar.maximum()

        # Определяем диапазон новых сообщений (старее уже загруженных)
        new_start = max(0, total - current_loaded - self.lazy_load_limit)
        new_end = total - current_loaded

        # Вставляем новые сообщения в начало layout
        for idx in range(new_start, new_end):
            msg = chat['messages'][idx]
            self._add_message_widget_to_start(msg, idx)

        # Обновляем счётчик
        self.loaded_count[filepath] += (new_end - new_start)

        # Корректируем прокрутку, чтобы остаться на месте
        def adjust_scroll():
            new_max = scrollbar.maximum()
            delta = new_max - old_max
            scrollbar.setValue(old_value + delta)

        QTimer.singleShot(0, adjust_scroll)
        QTimer.singleShot(100, adjust_scroll)

    def _add_load_more_indicator(self, chat):
        if self.load_more_widget is not None:
            try:
                self.load_more_widget.deleteLater()
            except RuntimeError:
                pass
            self.load_more_widget = None

        btn = QPushButton("Показать предыдущие сообщения")
        btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: 1px solid #ccc;
                border-radius: 10px;
                padding: 8px;
                color: #1a73e8;
            }
            QPushButton:hover {
                background: #f0f0f0;
            }
        """)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setFocusPolicy(Qt.NoFocus)
        btn.clicked.connect(lambda: self._load_more_messages(chat['filepath']))
        self.load_more_widget = btn
        self.chat_messages_layout.insertWidget(0, btn)

    def append_user_message(self, text: str, time_str: str = None):
        self._append_message("user", text, time_str)

    def append_assistant_message(self, text: str, time_str: str = None, reasoning: str = "", pinned: bool = False,
                                 idx: int = None):
        self._append_message("assistant", text, time_str, reasoning, pinned, idx)

    def _append_message(self, role: str, content: str, time_str: str = None, reasoning: str = "",
                        pinned: bool = False, idx: int = None):
        if not isinstance(content, str):
            if isinstance(content, list):
                text_parts = []
                for item in content:
                    if isinstance(item, dict):
                        if 'text' in item:
                            text_parts.append(str(item['text']))
                        elif 'image_url' in item:
                            text_parts.append('[Изображение]')
                        else:
                            text_parts.append(str(item))
                    else:
                        text_parts.append(str(item))
                content = ' '.join(text_parts)
            else:
                content = str(content)

        content = content.strip()
        reasoning = reasoning.strip()

        if role == "user":
            label = "👤 Пользователь"
            color = "#1a73e8"
            bg_color = "#e3f2fd"
        else:
            label = "🤖 Ассистент"
            color = "#188038"
            bg_color = "#f1f1f1"

        pin_marker = '📌 ' if pinned else ''

        unpin_link = ""
        if pinned and idx is not None:
            unpin_link = (
                f'<a href="unpin:{idx}" style="color:#d32f2f; text-decoration:none; '
                f'font-weight:bold; margin-left:8px;" title="Снять закрепление">✕</a>'
            )

        edit_link = ""
        if idx is not None:
            edit_link = f'<a href="edit:{idx}" style="color:#1a73e8; text-decoration:none; margin-left:8px;">✎</a>'

        copy_link = ""
        if idx is not None:
            copy_link = f'<a href="copy:{idx}" style="color:#1a73e8; text-decoration:none; margin-left:8px;" title="Копировать текст">📋</a>'

        regenerate_link = ""
        if role == "assistant" and idx is not None:
            regenerate_link = f'<a href="regenerate:{idx}" style="color:#1a73e8; text-decoration:none; margin-left:8px;" title="Перегенерировать">🔄</a>'

        if role == "assistant" and not content.strip() and reasoning.strip():
            content_display = self._markdown_to_html(reasoning)
        else:
            content_display = self._markdown_to_html(content) if role == "assistant" else self._escape_html(
                content).replace('\n', '<br>')

        time_html = f'<span style="color:#999; font-size:0.8em;">{time_str}</span>' if time_str else ""

        if role == "user":
            # Таблица прижата вправо, ширина по содержимому, max-width 80%
            html = f"""
            <table cellpadding="0" cellspacing="0" style="margin:8px 12px; max-width:80%;" align="right">
            <tr><td>
            <div style="background-color:{bg_color}; border-radius:12px; padding:12px 16px; color:#000000;">
            <span style="font-weight:600; color:{color};">{pin_marker}{label}</span> {time_html} {unpin_link} {edit_link} {copy_link}
            <div style="margin-top:6px; line-height:1.5; color:#000000;">{content_display}</div>
            </div>
            </td></tr></table>
            """
        else:
            # Таблица прижата влево, для мыслей и ответа отдельные блоки
            reasoning_html = ""
            if reasoning and self.show_reasoning and content.strip():
                reasoning_escaped = self._escape_html(reasoning).replace('\n', '<br>')
                reasoning_html = f"""
                <div style="background-color:#fff9c4; border-left:4px solid #fbc02d; padding:8px 12px; margin-top:6px; border-radius:8px; font-style:italic; color:#555;">
                    <span style="font-weight:bold;">🧠 Мысли:</span><br>{reasoning_escaped}
                </div>
                """
            html = f"""
            <table cellpadding="0" cellspacing="0" style="margin:8px 12px; max-width:80%;" align="left">
            <tr><td>
            <div style="background-color:{bg_color}; border-radius:12px; padding:10px 16px; color:#000000;">
            <span style="font-weight:600; color:{color};">{pin_marker}{label}</span> {time_html} {unpin_link} {edit_link} {copy_link} {regenerate_link}
            </div>{reasoning_html}<div style="background-color:{bg_color}; border-radius:12px; padding:12px 16px; margin-top:6px; color:#000000;">{content_display}</div>
            </td></tr></table>
            """
        self.chat_browser.append(html)

    def display_user_message_with_attachments(self, msg, idx=None):
        time_str = msg.get('time', '')
        content = msg.get('content', '')
        attachments = msg.get('attachments', [])

        time_html = f'<span style="color:#999; font-size:0.8em;">{time_str}</span>' if time_str else ""
        pin_marker = '📌 ' if msg.get('pinned', False) else ''

        unpin_link = ""
        if msg.get('pinned', False) and idx is not None:
            unpin_link = f'<a href="unpin:{idx}" style="color:#d32f2f; text-decoration:none; font-weight:bold; margin-left:8px;">✕</a>'

        edit_link = ""
        if idx is not None:
            edit_link = f'<a href="edit:{idx}" style="color:#1a73e8; text-decoration:none; margin-left:8px;">✎</a>'

        copy_link = ""
        if idx is not None:
            copy_link = f'<a href="copy:{idx}" style="color:#1a73e8; text-decoration:none; margin-left:8px;" title="Копировать текст">📋</a>'

        # Таблица с отступом от краёв и max-width, прижата вправо
        html = f"""
        <table cellpadding="0" cellspacing="0" style="margin:8px 12px; max-width:80%;" align="right">
        <tr><td>
        <div style="background-color:#e3f2fd; border-radius:12px; padding:12px 16px; color:#000000;">
        <span style="font-weight:600; color:#1a73e8;">{pin_marker}👤 Пользователь</span> {time_html} {unpin_link} {edit_link} {copy_link}
        <div style="margin-top:6px; line-height:1.5; color:#000000;">
        """
        if content:
            html += self._escape_html(content).replace('\n', '<br>') + '<br>'

        for att in attachments:
            att_type = att.get('type')
            if att_type == 'image':
                self.preview_image_counter += 1
                img_id = f"img_{self.preview_image_counter}"
                self.preview_images[img_id] = att["data"]
                html += f'<a href="image:{img_id}"><img src="{att["data"]}" style="max-width:400px; max-height:400px; margin:5px;"></a><br>'
            elif att_type == 'file_text':
                file_name = att.get('name', 'file')
                text_content = att.get('content', '')
                images = att.get('images', [])

                self.preview_counter += 1
                preview_id = f"preview_{self.preview_counter}"
                self.preview_texts[preview_id] = {
                    'text': text_content,
                    'images': images
                }
                html += f'<a href="{preview_id}" style="color:#1a73e8; text-decoration:underline;">📄 {file_name}</a><br>'
            elif att_type == 'file':
                path = att.get('path')
                if path and os.path.exists(path):
                    html += f'<a href="file:///{path}" style="color:#1a73e8;">{att.get("name", "Файл")}</a><br>'
                else:
                    html += f'<span>{att.get("name", "Файл")}</span><br>'
        html += '</div></div></td></tr></table>'
        self.chat_browser.append(html)

    def _markdown_to_html(self, text: str) -> str:
        if MARKDOWN_AVAILABLE:
            try:
                if PYGMENTS_AVAILABLE:
                    import re
                    pattern = re.compile(r'```(\w*)\n(.*?)```', re.DOTALL)
                    code_blocks = {}
                    counter = 0
                    pygments_formatter = HtmlFormatter(style='monokai', nowrap=True)

                    def replace_code_block(match):
                        nonlocal counter
                        lang = match.group(1).strip() or 'text'
                        code = match.group(2)
                        try:
                            lexer = get_lexer_by_name(lang, stripall=True)
                        except Exception:
                            lexer = TextLexer()
                        highlighted = highlight(code, lexer, pygments_formatter)
                        placeholder = f'@@CODEBLOCK{counter}@@'
                        code_blocks[placeholder] = (
                            f'<pre class="code-block"><code>{highlighted}</code></pre>'
                        )
                        counter += 1
                        return placeholder

                    text = pattern.sub(replace_code_block, text)
                    html_body = markdown.markdown(text, extensions=['fenced_code', 'tables', 'nl2br'])

                    for placeholder, code_html in code_blocks.items():
                        html_body = html_body.replace(placeholder, code_html)

                    # Получаем CSS для подсветки, применяем к .code-block code
                    pygments_css = pygments_formatter.get_style_defs('.code-block code')
                    styled_html = f"""
                    <div style="font-family: sans-serif; line-height: 1.4;">
                    <style>
                        * {{ user-select: text; }}
                        ::selection {{ background: #b3d4fc; }}
                        a {{ color: #1a73e8; text-decoration: underline; }}
                        a:visited {{ color: #1a73e8; }}
                        {pygments_css}
                        .code-block {{
                            background-color: #272822 !important;
                            padding: 8px !important;
                            border-radius: 6px !important;
                            margin: 4px 0 !important;
                            display: block !important;
                            width: 100% !important;
                            box-sizing: border-box !important;
                            line-height: 1 !important;          /* минимальный межстрочный интервал */
                            white-space: pre !important;        /* точное сохранение переносов */
                            overflow: hidden !important;        /* без прокрутки */
                            font-family: "Consolas", "Monaco", monospace !important;
                            font-size: 0.9em !important;
                        }}
                        .code-block code {{
                            display: inline !important;         /* спаны не создают блочных отступов */
                            margin: 0 !important;
                            padding: 0 !important;
                            line-height: inherit !important;
                            font-family: inherit !important;
                            font-size: inherit !important;
                            white-space: inherit !important;
                        }}
                        .code-block code span {{
                            display: inline !important;         /* и сами спаны строчные */
                            margin: 0 !important;
                            padding: 0 !important;
                            line-height: inherit !important;
                        }}
                        table {{ border-collapse: collapse; width: 100%; margin: 10px 0; }}
                        th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
                        th {{ background-color: #e0e0e0; font-weight: bold; }}
                        tr:nth-child(even) {{ background-color: #f9f9f9; }}
                    </style>
                    {html_body}
                    </div>
                    """
                    return styled_html
                else:
                    # если Pygments недоступен
                    html_body = markdown.markdown(text, extensions=['fenced_code', 'tables', 'nl2br'])
                    styled_html = f"""
                    <div style="font-family: sans-serif; line-height: 1.4;">
                    <style>
                        * {{ user-select: text; }}
                        ::selection {{ background: #b3d4fc; }}
                        a {{ color: #1a73e8; text-decoration: underline; }}
                        a:visited {{ color: #1a73e8; }}
                        table {{ border-collapse: collapse; width: 100%; margin: 10px 0; }}
                        th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
                        th {{ background-color: #e0e0e0; font-weight: bold; }}
                        tr:nth-child(even) {{ background-color: #f9f9f9; }}
                        pre {{ background-color: #f5f5f5; padding: 10px; border-radius: 4px; }}
                        code {{ background-color: #f0f0f0; padding: 2px 4px; border-radius: 3px; font-family: monospace; }}
                    </style>
                    {html_body}
                    </div>
                    """
                    return styled_html
            except Exception:
                pass
        return self._escape_html(text).replace('\n', '<br>')

    def read_file_content_structured(self, path):
        """Извлекает текст с сохранением структуры (заголовки, абзацы, списки).
        Для Markdown/HTML возвращает исходный текст (или обработанный HTML), для
        остальных — обычный текст с абзацами.
        """

        ext = os.path.splitext(path)[1].lower()
        try:
            if ext in ('.md', '.markdown'):
                # Markdown уже содержит разметку, возвращаем как есть
                with open(path, 'r', encoding='utf-8') as f:
                    return f.read()
            elif ext in ('.html', '.htm'):
                # Пытаемся извлечь текст с сохранением тегов заголовков и абзацев
                try:
                    from bs4 import BeautifulSoup
                    with open(path, 'r', encoding='utf-8') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                    # Удаляем скрипты и стили
                    for tag in soup(['script', 'style']):
                        tag.decompose()
                    # Собираем текст, вставляя переносы строк после блочных элементов
                    result = []
                    for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'li', 'div', 'br']):
                        if element.name == 'br':
                            result.append('\n')
                        else:
                            text = element.get_text(strip=True)
                            if text:
                                if element.name.startswith('h'):
                                    result.append('#' * int(element.name[1]) + ' ' + text)
                                else:
                                    result.append(text)
                    return '\n'.join(result)
                except ImportError:
                    # BeautifulSoup не установлен — fallback на простой текст
                    return self.read_file_content(path)
            elif ext == '.pptx':  # <-- ДОБАВЛЕНА ЭТА ВЕТКА
                return self.read_file_content(path)
            else:
                # Для PDF, DOCX и прочих используем обычный метод, но он уже даёт абзацы
                return self.read_file_content(path)
        except Exception as e:
            self.logger.warning(f"Ошибка извлечения структурированного текста из {path}: {e}")
            return self.read_file_content(path)

    def _escape_html(self, text: str) -> str:
        return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    def scroll_chat_to_bottom(self):
        self.btn_scroll_down.clearFocus()
        self.chat_scroll_area.setFocus()
        self._skip_scroll_restore = True
        QTimer.singleShot(0, self._do_scroll_to_bottom)

    def _do_scroll_to_bottom(self):
        # Откладываем установку, чтобы layout успел пересчитаться
        def set_scroll():
            vsb = self.chat_scroll_area.verticalScrollBar()
            vsb.setValue(vsb.maximum())
            self.btn_scroll_down.hide()
            self.update_floating_buttons_positions()

        # Даём время на обновление layout
        QTimer.singleShot(100, set_scroll)

    def on_chat_scrolled(self, value):
        if self._auto_scrolling:
            return
        scrollbar = self.chat_scroll_area.verticalScrollBar()
        max_scroll = scrollbar.maximum()
        threshold = 50

        # Пользователь у самого низа — «приклеиваем» скролл к низу
        if max_scroll - value <= threshold:
            self._stick_to_bottom = True
            if self.btn_scroll_down.isVisible():
                self.btn_scroll_down.hide()
                self.update_floating_buttons_positions()
        else:
            self._stick_to_bottom = False
            if not self.btn_scroll_down.isVisible():
                self.btn_scroll_down.show()
                self.update_floating_buttons_positions()

        # Автоподгрузка старых сообщений при прокрутке вверх
        if value < 50 and 0 <= self.current_chat_index < len(self.chats):
            chat = self.chats[self.current_chat_index]
            filepath = chat['filepath']
            if self.loaded_count.get(filepath, 0) < len(chat['messages']):
                self._load_more_messages(filepath)

    def update_pinned_panel(self):
        """Заполняет панель закреплённых сообщений текущими закреплениями."""

        # Очищаем старые элементы
        while self.pinned_layout.count():
            item = self.pinned_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if self.current_chat_index < 0:
            self.pinned_panel.setVisible(False)
            return

        chat = self.chats[self.current_chat_index]
        pinned_messages = [(idx, msg) for idx, msg in enumerate(chat['messages']) if msg.get('pinned', False)]

        if not pinned_messages:
            self.pinned_panel.setVisible(False)
            return

        self.pinned_panel.setVisible(True)

        for idx, msg in pinned_messages:
            # Контейнер для одного закреплённого сообщения
            item_widget = QWidget()
            item_layout = QHBoxLayout(item_widget)
            item_layout.setContentsMargins(4, 2, 4, 2)
            item_layout.setSpacing(4)

            # Текст сообщения (обрезаем до 50 символов)
            preview = msg.get('content', '')[:50]
            btn_label = QPushButton(preview)
            btn_label.setCursor(Qt.PointingHandCursor)
            btn_label.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    color: #333;
                    text-align: left;
                    padding: 0;
                }
                QPushButton:hover { color: #1a73e8; }
            """)
            btn_label.clicked.connect(
                lambda checked, i=idx: self.navigate_to_message(
                    self.chats[self.current_chat_index]['filepath'], i
                )
            )

            # Крестик для снятия закрепления
            unpin_btn = QPushButton("✕")
            unpin_btn.setFixedSize(20, 20)
            unpin_btn.setCursor(Qt.PointingHandCursor)
            unpin_btn.setStyleSheet("""
                QPushButton {
                    background: transparent;
                    border: none;
                    color: #d32f2f;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #f9eec1;
                }
            """)
            # Привязываем индекс к кнопке через lambda
            unpin_btn.clicked.connect(lambda checked, i=idx: self.toggle_pin_message(i))

            item_layout.addWidget(btn_label)
            item_layout.addWidget(unpin_btn)
            self.pinned_layout.addWidget(item_widget)

    def navigate_to_message(self, filepath, msg_idx):
        self.select_chat_by_filepath(filepath, keep_scroll=True, scroll_target=(filepath, msg_idx))

    def _scroll_to_message(self, filepath, msg_idx):
        target_widget = None
        for fp, idx, widget in self.message_widgets:
            if fp == filepath and idx == msg_idx:
                target_widget = widget
                break

        if target_widget:
            # Даём layout время обновиться
            QTimer.singleShot(50, lambda: self._perform_scroll_to_widget(target_widget))

    def _perform_scroll_to_widget(self, widget):
        self.chat_messages_layout.activate()
        self.chat_messages_widget.adjustSize()

        # Ручная прокрутка — предсказуемее, чем ensureWidgetVisible
        widget_pos = widget.mapTo(self.chat_messages_widget, QPoint(0, 0))
        target_y = widget_pos.y() - 100
        vsb = self.chat_scroll_area.verticalScrollBar()
        vsb.setValue(max(0, min(target_y, vsb.maximum())))

        # Принудительная перерисовка viewport (обои + содержимое)
        self.chat_scroll_area.viewport().repaint()
        self.chat_messages_widget.repaint()

        # Подсветка
        original_style = widget.styleSheet()
        widget.setStyleSheet("background-color: #ffe58f; border-radius: 18px;")

        def restore_style():
            try:
                widget.setStyleSheet(original_style)
            except RuntimeError:
                pass
            self.chat_scroll_area.viewport().repaint()

        QTimer.singleShot(2000, restore_style)

    def save_chat_by_filepath(self, filepath):
        chat = None
        for c in self.chats:
            if c['filepath'] == filepath:
                chat = c
                break
        if chat is None:
            for c in self.archived_chats:
                if c['filepath'] == filepath:
                    chat = c
                    break
        if chat is None:
            return

        if chat.get('incognito', False):
            return

        if chat.get('locked', False):
            return

        data = {
            "title": chat['title'],
            "messages": chat['messages'],
            "folder": chat.get('folder', ''),
            "background_color": chat.get('background_color'),
            "background_image": chat.get('background_image'),
            "model": chat.get('model', self.settings.get('model', 'local-model')),
            "created_time": chat.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S')),
            "last_activity_time": chat.get('last_activity_time', time.strftime('%d.%m.%Y %H:%M:%S')),
            "total_tokens_sent": chat.get('total_tokens_sent', 0),
            "total_generation_time": chat.get('total_generation_time', 0.0),
            "color": chat.get('color', '#FFFFFF'),
            "params": chat.get('params', {}),
            "system_prompt": chat.get('system_prompt', ''),
            "pinned": chat.get('pinned', False),
            "password_hash": chat.get('password_hash'),
            "background_image_original": chat.get('background_image_original'),
            "tags": chat.get('tags', []),
            "archived": chat.get('archived', False),
            "encrypted": chat.get('encrypted', False),
        }

        need_encrypt = chat.get('encrypted', False)

        # === ШАГ 1: готовим данные в памяти ===
        try:
            if need_encrypt:
                pw = getattr(self, '_current_chat_password', None)
                if not pw:
                    # Пароль потерян — файл НЕ перезаписываем
                    self.logger.warning(
                        f"Не сохранён {filepath}: пароль чата потерян. Файл не тронут."
                    )
                    self.show_error(
                        "Ошибка сохранения",
                        f"Не удалось сохранить чат «{chat.get('title', '')}».\n"
                        "Пароль не введён. Войдите в чат заново и повторите."
                    )
                    return
                encrypted_bytes = self._encrypt_with_password(
                    json.dumps(data, ensure_ascii=False).encode('utf-8'),
                    pw
                )
                json_text = None
            else:
                json_text = json.dumps(data, ensure_ascii=False, indent=2)
                encrypted_bytes = None
        except Exception as e:
            self.logger.error(f"Ошибка подготовки данных для {filepath}: {e}")
            self.show_error(
                "Ошибка сохранения",
                f"Не удалось подготовить чат «{chat.get('title', '')}» к сохранению.\nПричина: {e}"
            )
            return

        # === ШАГ 2: только теперь пишем на диск ===
        try:
            if encrypted_bytes is not None:
                with open(filepath, 'wb') as f:
                    f.write(encrypted_bytes)
            else:
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(json_text)
            self.logger.info(f"Сохранён чат: {filepath}")
        except Exception as e:
            self.logger.error(f"Ошибка записи {filepath}: {e}")
            self.show_error(
                "Ошибка сохранения",
                f"Не удалось сохранить чат «{chat.get('title', '')}».\nПричина: {e}"
            )

    def save_current_chat(self):
        """Сохраняет текущий чат."""

        if self.current_chat_index < 0 or self.current_chat_index >= len(self.chats):
            return
        chat = self.chats[self.current_chat_index]
        self.save_chat_by_filepath(chat['filepath'])

    def _force_scroll_to_bottom(self):
        """Принудительно прокручивает чат вниз, не затрагивая автоскроллинг."""

        scrollbar = self.chat_browser.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _invalidate_chat_cache(self, filepath):
        if filepath in self.chat_view_cache:
            del self.chat_view_cache[filepath]

    def _make_message_widget(self, msg, idx):
        role = msg.get('role')
        content = msg.get('content', '')
        if isinstance(content, list):
    # ... преобразование как в _add_message_widget ...
            time_str = msg.get('time', '')
            reasoning = msg.get('reasoning', '') if (role == 'assistant' and self.show_reasoning) else ''
            pinned = msg.get('pinned', False)
            starred = msg.get('starred', False)

        if role == 'assistant':
            display_content = self._markdown_to_html(content)
        else:
            display_content = self._escape_html(content).replace('\n', '<br>')

        widget = MessageWidget(role, display_content, time_str, reasoning, pinned, idx, starred=starred,
                               parent=self.chat_messages_widget, image_cache=self.image_cache)
        widget.edit_clicked.connect(self._on_message_action_edit)
        widget.copy_clicked.connect(self._on_message_action_copy)
        widget.regenerate_clicked.connect(self._on_message_action_regenerate)
        widget.unpin_clicked.connect(self._on_message_action_unpin)
        widget.context_menu_requested.connect(lambda idx, gp, w: self.show_message_context_menu_at(idx, gp, w))
        if 'attachments' in msg and msg['attachments']:
            widget.set_attachments(msg['attachments'])
        return widget

    def toggle_incognito(self, filepath):
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        chat = self.chats[idx]
        current = chat.get('incognito', False)
        if not current:
            # Включаем режим инкогнито
            chat['incognito'] = True
            # Удаляем существующий файл, если он есть
            if os.path.exists(filepath):
                try:
                    os.remove(filepath)
                except Exception as e:
                    print(f"Ошибка удаления файла {filepath}: {e}")
        else:
            # Выключаем режим инкогнито
            chat['incognito'] = False
            # Сохраняем текущую историю в файл
            self.save_chat_by_filepath(filepath)

        # Обновляем текст элемента с учётом инкогнито и непрочитанности
        self.update_unread_indicators()

    def add_or_edit_comment(self, idx):
        """Открывает диалог для добавления или редактирования заметки."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return
        msg = chat['messages'][idx]
        current_comment = msg.get('comment', '')
        text, ok = QInputDialog.getMultiLineText(
            self,
            "Заметка к сообщению",
            "Введите текст заметки:",
            current_comment
        )
        if ok:
            text = text.strip()
            if text:
                msg['comment'] = text
            else:
                msg.pop('comment', None)

            # Сохраняем путь к чату и индекс сообщения
            filepath = chat['filepath']

            self.save_current_chat()

            # Перерисовываем без автоскролла
            self._redraw_chat_with_history(keep_scroll=True)

            # После перестройки прокручиваем к этому сообщению
            QTimer.singleShot(200, lambda: self._scroll_to_message(filepath, idx))

            self.statusBar().showMessage("Заметка сохранена")

    def remove_comment(self, idx):
        """Удаляет заметку из сообщения."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return
        if 'comment' in chat['messages'][idx]:
            del chat['messages'][idx]['comment']

            filepath = chat['filepath']

            self.save_current_chat()
            self._redraw_chat_with_history(keep_scroll=True)

            # После перестройки прокручиваем к этому сообщению
            QTimer.singleShot(200, lambda: self._scroll_to_message(filepath, idx))

            self.statusBar().showMessage("Заметка удалена")

    def set_message_reaction(self, idx, emoji):
        """Устанавливает реакцию на сообщение. Если реакция совпадает, снимает её."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return
        msg = chat['messages'][idx]

        if msg.get('reaction') == emoji:
            # Повторный выбор — снимаем реакцию
            msg.pop('reaction', None)
        else:
            msg['reaction'] = emoji

        # Сохраняем позицию
        filepath = chat['filepath']
        self.save_current_chat()
        self._redraw_chat_with_history(keep_scroll=True)

        # Прокручиваем к этому сообщению после перерисовки
        QTimer.singleShot(200, lambda: self._scroll_to_message(filepath, idx))

    def remove_message_reaction(self, idx):
        """Удаляет реакцию с сообщения."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return
        if 'reaction' in chat['messages'][idx]:
            del chat['messages'][idx]['reaction']

        filepath = chat['filepath']
        self.save_current_chat()
        self._redraw_chat_with_history(keep_scroll=True)
        QTimer.singleShot(200, lambda: self._scroll_to_message(filepath, idx))

    def quote_message(self, idx):
        """Вставляет цитату из сообщения в поле ввода."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return
        msg = chat['messages'][idx]
        content = msg.get('content', '')
        if isinstance(content, list):
            # Для мультимодальных сообщений берём только текст
            content = ' '.join([part.get('text', '') for part in content if isinstance(part, dict)])
        if not content.strip():
            return

        # Формируем текст цитаты с префиксом '>' и переносами
        quote_text = "\n".join(["> " + line for line in content.splitlines()])
        current_text = self.input_edit.toPlainText()
        if current_text:
            self.input_edit.setPlainText(current_text + "\n\n" + quote_text + "\n\n")
        else:
            self.input_edit.setPlainText(quote_text + "\n\n")
        self.input_edit.setFocus()
        # Перемещаем курсор в конец
        cursor = self.input_edit.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.input_edit.setTextCursor(cursor)

    # ======================================================================
    #  СООБЩЕНИЯ И ДЕЙСТВИЯ С НИМИ
    # ======================================================================

    def _check_api_available(self, timeout=2) -> bool:
        """Быстрая проверка, отвечает ли API-сервер."""

        import requests
        base = self.settings.get("api_base", "").rstrip("/")
        if not base:
            return False
        try:
            resp = requests.get(f"{base}/models", timeout=timeout)
            return resp.status_code in (200, 401, 403)
        except Exception:
            return False

    def _check_lm_studio_on_startup(self):
        """Проверка доступности LM Studio при запуске программы."""

        if not self._check_api_available():
            api_base = self.settings.get('api_base', 'http://localhost:1234/v1')
            QMessageBox.information(
                self,
                "LM Studio не найден",
                f"Echos не видит LM Studio по адресу:\n{api_base}\n\n"
                "Что делать:\n"
                "  1. Установите LM Studio: https://lmstudio.ai\n"
                "  2. Загрузите модель (например, qwen2.5-7b-instruct)\n"
                "  3. Перейдите во вкладку Server → Start Server\n"
                "  4. Вернитесь в Echos и попробуйте снова"
            )

    def send_message(self):
        user_text = self.input_edit.toPlainText().strip()

        # Быстрая проверка доступности API
        if not self._check_api_available():
            reply = QMessageBox.question(
                self,
                "Сервер недоступен",
                "LM Studio не отвечает по адресу:\n"
                f"{self.settings.get('api_base', '')}\n\n"
                "Запустить LM Studio сейчас?",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.launch_lm_studio()
            return

        # Выделяем цитату, если текст начинается с '>'
        quote = None
        if user_text.lstrip().startswith('>'):
            lines = user_text.splitlines()
            quote_lines = []
            i = 0
            while i < len(lines) and lines[i].lstrip().startswith('>'):
                quote_lines.append(lines[i].lstrip()[1:].strip())
                i += 1
            quote = "\n".join(quote_lines)
            rest_lines = lines[i:]
            while rest_lines and rest_lines[0].strip() == '':
                rest_lines.pop(0)
            user_text = "\n".join(rest_lines).strip()

        if not user_text and not self.attachments:
            return
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "Занято", "Дождитесь завершения текущего запроса.")
            return
        if self.current_chat_index < 0:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите чат.")
            return

        # Обработка команд
        if user_text.startswith('/'):
            cmd = user_text.strip().lower()
            if cmd == '/clear':
                self.clear_chat_history()
                self.input_edit.clear()
                return
            elif cmd == '/retry':
                self.retry_last_request()
                self.input_edit.clear()
                return
            elif cmd == '/summary':
                self.summarize_dialog()
                self.input_edit.clear()
                return

        # === ВЫЗОВ ПЛАГИНОВ: обработка текста перед отправкой ===
        api = PluginAPI(self)
        for plugin in self.plugins:
            try:
                if hasattr(plugin, 'on_message_send'):
                    user_text = plugin.on_message_send(user_text, api)
            except Exception as e:
                self.logger.warning(f"Ошибка в плагине {plugin.__class__.__name__}: {e}")
                self.show_warning(
                    "Ошибка плагина",
                    f"Плагин {plugin.__class__.__name__} вызвал ошибку и был пропущен."
                )

        # === ЛОГИРОВАНИЕ ===
        self.logger.info(f"Отправка сообщения: {user_text[:100]}")

        # Сбор вложений
        attachments_data = []
        for path in self.attachments:
            # Проверка размера файла
            max_size_bytes = self.settings.get("max_attachment_size_mb", 20) * 1024 * 1024
            if os.path.getsize(path) > max_size_bytes:
                QMessageBox.warning(
                    self,
                    "Файл слишком большой",
                    f"Файл «{os.path.basename(path)}» превышает максимальный размер {self.settings.get('max_attachment_size_mb', 20)} МБ."
                )
                return
            ext = os.path.splitext(path)[1].lower()
            if ext in ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'):
                # === ИЗОБРАЖЕНИЯ: сохраняем в файл ===
                chat_id = os.path.splitext(os.path.basename(self.chats[self.current_chat_index]['filepath']))[0]
                attachments_dir = os.path.join("attachments", chat_id)
                os.makedirs(attachments_dir, exist_ok=True)

                timestamp = int(time.time() * 1000)
                img_filename = f"img_{timestamp}_{os.path.basename(path)}"
                img_path = os.path.join(attachments_dir, img_filename)

                # Определяем, нужно ли шифровать
                current_chat = self.chats[self.current_chat_index]
                need_encrypt = self.settings.get("encrypt_chats", False) or current_chat.get('encrypted', False)

                # Готовим байты в память
                img_bytes = None
                if PIL_AVAILABLE:
                    try:
                        img = Image.open(path)
                        img.thumbnail((512, 512))
                        buffer = io.BytesIO()
                        img.convert("RGB").save(buffer, "JPEG", quality=80)
                        img_bytes = buffer.getvalue()
                    except Exception as e:
                        print(f"Ошибка сжатия: {e}")
                if img_bytes is None:
                    with open(path, 'rb') as f:
                        img_bytes = f.read()

                # Сохраняем (шифруем или нет)
                if need_encrypt:
                    with open(img_path, 'wb') as f:
                        f.write(self._encrypt_bytes(img_bytes))
                    print(f"[ENCRYPT] Картинка зашифрована: {img_path}")
                else:
                    with open(img_path, 'wb') as f:
                        f.write(img_bytes)
                    print(f"[PLAIN] Картинка сохранена без шифрования: {img_path}")

                attachments_data.append({
                    "type": "image",
                    "path": img_path,
                    "name": os.path.basename(path)
                })
            else:
                text_content = self.read_file_content(path)
                images = self.extract_images_from_file(path)

                if not (text_content and text_content.strip()) and not images:
                    QMessageBox.warning(
                        self,
                        "Не удалось прочитать файл",
                        f"Файл «{os.path.basename(path)}» повреждён, пуст или имеет неподдерживаемый формат."
                    )
                    return

                if text_content and self.rag_enabled:
                    if self.rag_manager is None:
                        self.rag_manager = RAGManager()
                    self.rag_manager.add_document(text_content)

                if text_content:
                    file_text = f"[Содержимое файла {os.path.basename(path)}]:\n{text_content}"
                    attachments_data.append({
                        "type": "file_text",
                        "content": file_text,
                        "name": os.path.basename(path),
                        "images": images
                    })
                else:
                    if images:
                        attachments_data.append({
                            "type": "file_text",
                            "content": f"[Изображения из файла {os.path.basename(path)}]",
                            "name": os.path.basename(path),
                            "images": images
                        })
                    else:
                        attachments_dir = "attachments"
                        os.makedirs(attachments_dir, exist_ok=True)
                        dest_path = os.path.join(attachments_dir, os.path.basename(path))

                        current_chat = self.chats[self.current_chat_index]
                        need_encrypt = self.settings.get("encrypt_chats", False) or current_chat.get('encrypted', False)

                        try:
                            if need_encrypt:
                                with open(path, 'rb') as f:
                                    data = f.read()
                                with open(dest_path, 'wb') as f:
                                    f.write(self._encrypt_bytes(data))
                            else:
                                shutil.copy2(path, dest_path)
                        except Exception as e:
                            print(f"Не удалось скопировать файл: {e}")
                            dest_path = path
                        attachments_data.append({"type": "file", "path": dest_path, "name": os.path.basename(path)})
        # ПРОВЕРКА: количество изображений
        image_count = sum(1 for att in attachments_data if att.get('type') == 'image')
        model_for_limit = self.chats[self.current_chat_index].get('model', self.settings.get('model', 'local-model'))
        max_images = self.get_max_images_for_model(model_for_limit)

        if image_count > max_images:
            QMessageBox.warning(
                self,
                "Слишком много изображений",
                f"Максимально допустимое количество изображений в одном сообщении: {max_images}.\n"
                f"Вы пытаетесь отправить {image_count}. Пожалуйста, уменьшите количество вложений."
            )
            return

        # ПРОВЕРКА: длина промпта (включая текстовые вложения)
        max_prompt_chars = self.settings.get("max_prompt_chars", 8000)
        total_prompt_chars = len(user_text)
        for att in attachments_data:
            if att['type'] == 'file_text':
                total_prompt_chars += len(att.get('content', ''))
            elif att['type'] == 'file':
                total_prompt_chars += len(att.get('name', '')) + 10

        if total_prompt_chars > max_prompt_chars:
            QMessageBox.warning(
                self,
                "Слишком длинный запрос",
                f"Максимальная допустимая длина запроса: {max_prompt_chars} символов.\n"
                f"Текущая длина: {total_prompt_chars} символов.\n"
                "Пожалуйста, сократите текст или удалите часть вложений."
            )
            return

        # Предпросмотр только если есть вложения, не являющиеся изображениями
        has_non_image_attachments = any(att.get('type') != 'image' for att in attachments_data)
        if has_non_image_attachments:
            preview_dialog = MessagePreviewDialog(user_text, attachments_data, self)
            if preview_dialog.exec() != QDialog.Accepted:
                return

        # Создаём сообщение
        chat = self.chats[self.current_chat_index]
        msg = {
            "role": "user",
            "content": user_text,
            "attachments": attachments_data,
            "time": time.strftime("%d.%m.%Y %H:%M:%S")
        }
        if quote:
            msg["quote"] = quote
        chat['messages'].append(msg)

        chat['last_activity_time'] = time.strftime('%d.%m.%Y %H:%M:%S')
        user_tokens = len(user_text) / 4
        for att in attachments_data:
            if att['type'] == 'image':
                user_tokens += 1000
            elif att['type'] == 'file_text':
                user_tokens += len(att.get('content', '')) / 4
            elif att['type'] == 'file':
                user_tokens += 10
        chat['total_tokens_sent'] = chat.get('total_tokens_sent', 0) + int(user_tokens)

        self.save_current_chat()
        self._add_message_widget(msg, len(chat['messages']) - 1)
        self._stick_to_bottom = True
        self.input_edit.clear()
        self.attachments.clear()
        self.update_attachments_ui()
        self.save_draft("")

        stream = self.settings.get("stream", True)
        if stream:
            self.current_response_widget = MessageWidget(
                "assistant", "⏳ Думаю...", time_str="", reasoning="", idx=None,
                parent=self.chat_messages_widget,
                show_actions = False
            )
            row_layout = QHBoxLayout()
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(0)
            row_layout.addWidget(self.current_response_widget)
            row_layout.addStretch(1)
            self.chat_messages_layout.addLayout(row_layout)
            QTimer.singleShot(0, self._do_scroll_to_bottom)
        else:
            self.statusBar().showMessage("Ожидание ответа...")

        messages_to_send = self._prepare_messages_for_api()
        messages_to_send = self._make_json_safe(messages_to_send)

        if self.settings.get("rag_enabled", False):
            if self.rag_manager is None:
                self.rag_manager = RAGManager()
            if self.rag_manager.documents:
                relevant_chunks = self.rag_manager.search(user_text, top_k=3)
                if relevant_chunks:
                    context_text = "\n---\n".join(relevant_chunks)
                    system_context = {"role": "system", "content": f"Контекст из базы знаний:\n{context_text}"}
                    if messages_to_send and messages_to_send[0].get("role") == "system":
                        messages_to_send.insert(1, system_context)
                    else:
                        messages_to_send.insert(0, system_context)

        base_params = {
            "max_tokens": self.settings.get("max_tokens", 4096),
            "temperature": self.settings.get("temperature", 0.7),
            "top_p": self.settings.get("top_p", 0.9),
            "top_k": self.settings.get("top_k", 40),
            "repeat_penalty": self.settings.get("repeat_penalty", 1.1)
        }
        chat_params = chat.get('params', {})
        if chat_params:
            base_params.update(chat_params)
        params = base_params

        model = chat.get('model', self.settings.get('model', 'local-model'))

        self.worker = ChatWorker(
            api_base=self.settings["api_base"],
            model=model,
            messages=messages_to_send,
            params=params,
            stream=stream
        )
        self.worker.chunk_received.connect(self.on_chunk_received)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.error_occurred.connect(self.on_error)
        self.worker.stopped.connect(self.on_stop_generation)

        if stream:
            self.current_response_text = ""
        else:
            self.statusBar().showMessage("Ожидание ответа...")

        self.hide_stop_button()
        self.generation_start_time = time.time()
        self.is_generating = True
        self.generation_chat_index = self.current_chat_index
        self.generation_chat_filepath = chat['filepath']

        self.current_response_text = ""
        if hasattr(self, 'speed_timer'):
            self.speed_timer.start()
            self.speed_label.setText("Скорость: 0.0 ток/с")

        self.worker.start()
        self.show_stop_button()

    def _send_request_after_edit(self, cutoff_idx=None):
        """Отправляет текущую историю модели (опционально обрезанную до cutoff_idx включительно)."""

        if self.current_chat_index < 0:
            return
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "Занято", "Дождитесь завершения текущего запроса.")
            return

        chat = self.chats[self.current_chat_index]
        stream = self.settings.get("stream", True)

        # Выбираем сообщения для отправки
        if cutoff_idx is not None:
            # Обрезаем историю до указанного индекса включительно
            messages_for_api = chat['messages'][:cutoff_idx + 1]
        else:
            messages_for_api = chat['messages']

        # Формируем список сообщений для отправки
        messages_to_send = []
        for m in messages_for_api:
            role = m.get('role')
            content = m.get('content', '')
            if isinstance(content, list):
                text_parts = []
                for item in content:
                    if isinstance(item, dict):
                        if 'text' in item:
                            text_parts.append(item['text'])
                        elif 'image_url' in item:
                            text_parts.append("[Изображение]")
                        else:
                            text_parts.append(str(item))
                    else:
                        text_parts.append(str(item))
                content = ' '.join(text_parts)
            messages_to_send.append({"role": role, "content": content})

        # Проверяем, есть ли хотя бы одно сообщение пользователя
        if not any(msg.get("role") == "user" for msg in messages_to_send):
            QMessageBox.warning(self, "Ошибка", "Нет сообщений пользователя для отправки.")
            return

        # Добавляем системный промпт первым, если он задан
        system_prompt = (chat.get('system_prompt') or '').strip()
        if not system_prompt:
            system_prompt = (self.settings.get('system_prompt') or '').strip()
        if system_prompt:
            messages_to_send.insert(0, {"role": "system", "content": system_prompt})

        # Создаём виджет-заглушку для нового ответа
        if stream:
            self.current_response_widget = MessageWidget(
                "assistant", "⏳ Думаю...", time_str="", reasoning="",
                idx=None, parent=self.chat_messages_widget,
                show_actions=False
            )
            row_layout = QHBoxLayout()
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(0)
            row_layout.addWidget(self.current_response_widget)
            row_layout.addStretch(1)
            self.chat_messages_layout.addLayout(row_layout)
            self.current_response_text = ""
            QTimer.singleShot(0, self._do_scroll_to_bottom)
        else:
            self.statusBar().showMessage("Ожидание ответа...")

        # Параметры генерации
        base_params = {
            "max_tokens": self.settings.get("max_tokens", 4096),
            "temperature": self.settings.get("temperature", 0.7),
            "top_p": self.settings.get("top_p", 0.9),
            "top_k": self.settings.get("top_k", 40),
            "repeat_penalty": self.settings.get("repeat_penalty", 1.1)
        }
        chat_params = chat.get('params', {})
        if chat_params:
            base_params.update(chat_params)
        params = base_params

        model = chat.get('model', self.settings.get('model', 'local-model'))

        self.worker = ChatWorker(
            api_base=self.settings["api_base"],
            model=model,
            messages=messages_to_send,
            params=params,
            stream=stream
        )
        self.worker.chunk_received.connect(self.on_chunk_received)
        self.worker.finished_ok.connect(self.on_finished)
        self.worker.error_occurred.connect(self.on_error)
        self.worker.stopped.connect(self.on_stop_generation)

        self.is_generating = True
        self.generation_chat_index = self.current_chat_index
        self.generation_chat_filepath = chat['filepath']
        self.generation_start_time = time.time()

        if hasattr(self, 'speed_timer'):
            self.speed_timer.start()
            self.speed_label.setText("Скорость: 0.0 ток/с")

        self.hide_stop_button()
        self.worker.start()
        self.show_stop_button()

    def _prepare_messages_for_api(self) -> List[Dict[str, Any]]:
        if self.current_chat_index < 0:
            return []
        chat = self.chats[self.current_chat_index]
        messages = chat['messages']
        max_tokens = self.settings.get("max_context_tokens", 4096)
        max_chars = max_tokens * 4

        # Получаем системный промпт для чата (индивидуальный или глобальный)
        chat_system_prompt = chat.get('system_prompt', '').strip()
        if not chat_system_prompt:
            chat_system_prompt = self.settings.get('system_prompt', '').strip()

        # Игнорируем все системные сообщения из истории (они больше не используются)
        system_msgs = [{"role": "system", "content": chat_system_prompt}] if chat_system_prompt else []
        other_msgs = [m for m in messages if m["role"] != "system"]
        if not other_msgs:
            return system_msgs

        has_image = any(
            'attachments' in msg and any(att['type'] == 'image' for att in msg['attachments'])
            for msg in other_msgs
        )
        if has_image:
            max_chars = min(max_chars, 4000)

        def msg_effective_size(msg):
            content = msg.get("content", "")
            size = len(str(content))
            if 'attachments' in msg:
                for att in msg['attachments']:
                    if att['type'] == 'image':
                        # Грубая оценка изображения в символах (не зависит от размера файла)
                        size += 4000  # примерно 1000 токенов на изображение
                    elif att['type'] == 'file_text':
                        size += len(att.get('content', ''))
            return size

        selected = []
        total_chars = sum(len(str(m.get("content", ""))) for m in system_msgs)

        last_msg = other_msgs[-1]
        other_msgs = other_msgs[:-1]
        last_size = msg_effective_size(last_msg)
        total_chars += last_size
        selected.append(last_msg)

        for msg in reversed(other_msgs):
            msg_size = msg_effective_size(msg)
            if total_chars + msg_size <= max_chars:
                selected.append(msg)
                total_chars += msg_size
            else:
                break
        selected.reverse()

        api_messages = []
        for msg in system_msgs + selected:
            role = msg.get("role")
            content = msg.get("content", "")

            if role == "user" and 'attachments' in msg and msg['attachments']:
                images = [att for att in msg['attachments'] if att['type'] == 'image']
                print(f"[DEBUG API] attachments={msg['attachments']}")
                print(f"[DEBUG API] images={images}")
                print(f"[DEBUG API] multimodal={self.settings.get('multimodal', False)}")
                files_text = []
                for att in msg['attachments']:
                    if att['type'] == 'file_text':
                        files_text.append(att['content'])
                    elif att['type'] == 'file':
                        files_text.append(f"[Файл: {att['name']}]({att['path']})")

                combined_text = content
                if msg.get('quote'):
                    combined_text = f"Цитата: {msg['quote']}\n\n{combined_text}"
                if files_text:
                    combined_text = (combined_text + "\n" + "\n".join(files_text)).strip()

                if images and self.settings.get("multimodal", False):
                    content_parts = []
                    if combined_text:
                        content_parts.append({"type": "text", "text": combined_text})
                    for img in images:
                        img_data = img.get('data')  # может быть None (старый формат)
                        if not img_data and img.get('path'):
                            try:
                                with open(img['path'], 'rb') as f:
                                    img_bytes = f.read()
                                if img_bytes.startswith(b"ENC:"):
                                    pw = getattr(self, '_current_chat_password', None)
                                    if pw:
                                        img_bytes = self._decrypt_with_password(img_bytes, pw)
                                    else:
                                        img_bytes = self._decrypt_bytes(img_bytes)
                                mime = "image/jpeg" if img['path'].lower().endswith(('.jpg', '.jpeg')) else "image/png"
                                img_data = f"data:{mime};base64," + base64.b64encode(img_bytes).decode('utf-8')
                            except Exception as e:
                                print(f"Ошибка чтения изображения {img.get('path')}: {e}")
                                continue
                        if img_data:
                            content_parts.append({"type": "image_url", "image_url": {"url": img_data}})
                    if content_parts:
                        api_messages.append({"role": "user", "content": content_parts})
                    else:
                        # Если не удалось добавить изображения, отправляем текст
                        api_messages.append({"role": "user", "content": combined_text or " "})
                else:
                    # Мультимодальность выключена или изображений нет
                    if images:
                        # Добавляем имена файлов в текст, чтобы модель знала о картинках
                        img_names = " ".join([att.get('name', '') for att in images])
                        text_for_api = (combined_text + " " + img_names).strip()
                    else:
                        text_for_api = combined_text
                    api_messages.append({"role": "user", "content": text_for_api or " "})
            else:
                content_to_send = content
                if role == "user" and msg.get('quote'):
                    content_to_send = f"Цитата: {msg['quote']}\n\n{content}"
                api_messages.append({"role": role, "content": content_to_send if content_to_send and content_to_send.strip() else " "})
        return api_messages

    def quote_message(self, idx):
        """Вставляет цитату из сообщения в поле ввода."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return
        msg = chat['messages'][idx]
        content = msg.get('content', '')
        if isinstance(content, list):
            content = ' '.join([part.get('text', '') for part in content if isinstance(part, dict)])
        if not content.strip():
            return

        quote_text = "\n".join(["> " + line for line in content.splitlines()])
        current_text = self.input_edit.toPlainText()
        if current_text:
            self.input_edit.setPlainText(current_text + "\n\n" + quote_text + "\n\n")
        else:
            self.input_edit.setPlainText(quote_text + "\n\n")
        self.input_edit.setFocus()
        cursor = self.input_edit.textCursor()
        cursor.movePosition(QTextCursor.End)
        self.input_edit.setTextCursor(cursor)

    def add_message_to_current(self, role: str, content: str, reasoning: str = ""):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        msg = {"role": role, "content": content, "time": time.strftime("%d.%m.%Y %H:%M:%S")}
        if role == "assistant":
            msg["reasoning"] = reasoning
        chat['messages'].append(msg)
        chat['last_activity_time'] = time.strftime('%d.%m.%Y %H:%M:%S')
        self.save_current_chat()

    def delete_message(self, idx):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if idx < 0 or idx >= len(chat['messages']):
            return
        # Удаляем сообщение
        del chat['messages'][idx]
        # Если удалено сообщение пользователя, и следующим было сообщение ассистента, удаляем и его
        if idx < len(chat['messages']) and chat['messages'][idx].get('role') == 'assistant':
            del chat['messages'][idx]
        self.save_current_chat()
        self._redraw_chat_with_history()
        self.statusBar().showMessage("Сообщение удалено")

    def edit_message(self, idx):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return

        # Проверка занятости — до любых изменений
        if self.worker and self.worker.isRunning():
            QMessageBox.information(
                self, "Занято",
                "Дождитесь завершения текущего запроса или остановите генерацию."
            )
            return

        msg = chat['messages'][idx]
        if msg.get('role') != 'user':
            QMessageBox.information(self, "Информация", "Редактировать можно только сообщения пользователя.")
            return

        new_text, ok = QInputDialog.getMultiLineText(
            self, "Редактировать сообщение", "Измените текст:", msg.get('content', '')
        )
        if ok and new_text.strip():
            new_msg = {
                "role": "user",
                "content": new_text.strip(),
                "time": time.strftime("%d.%m.%Y %H:%M:%S")
            }
            if 'attachments' in msg:
                new_msg['attachments'] = msg['attachments']

            chat['messages'].append(new_msg)
            self.save_current_chat()
            self._redraw_chat_with_history()
            self.statusBar().showMessage("Создано новое сообщение с изменённым текстом")
            self._send_request_after_edit()

    def edit_assistant_message(self, idx):
        """Редактирование сообщения ассистента (текст ответа и, при наличии, reasoning)."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return

        # Проверка занятости — до открытия диалогов
        if self.worker and self.worker.isRunning():
            QMessageBox.information(
                self, "Занято",
                "Дождитесь завершения текущего запроса или остановите генерацию."
            )
            return

        msg = chat['messages'][idx]
        if msg.get('role') != 'assistant':
            QMessageBox.information(self, "Информация", "Редактировать можно только сообщения ассистента.")
            return

        current_content = msg.get('content', '')
        new_content, ok_content = QInputDialog.getMultiLineText(
            self, "Редактировать ответ ассистента", "Текст ответа:", current_content
        )
        if not ok_content:
            return

        new_reasoning = None
        if 'reasoning' in msg:
            current_reasoning = msg.get('reasoning', '')
            new_reasoning, ok_reasoning = QInputDialog.getMultiLineText(
                self, "Редактировать мысли ассистента",
                "Reasoning (можно оставить пустым):", current_reasoning
            )
            if not ok_reasoning:
                return

        msg['content'] = new_content.strip()
        if new_reasoning is not None:
            msg['reasoning'] = new_reasoning.strip()

        self.save_current_chat()
        self._redraw_chat_with_history()
        self.statusBar().showMessage("Сообщение ассистента обновлено")

    def toggle_pin_message(self, idx):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if 0 <= idx < len(chat['messages']):
            msg = chat['messages'][idx]
            msg['pinned'] = not msg.get('pinned', False)
            # Устанавливаем цель прокрутки на это сообщение
            self.pending_scroll_target = (chat['filepath'], idx)
            self.save_current_chat()
            # keep_scroll=True – не скроллить вниз, а восстановить позицию у сообщения
            self._redraw_chat_with_history(keep_scroll=True)

    def toggle_star_message(self, idx):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if 0 <= idx < len(chat['messages']):
            msg = chat['messages'][idx]
            msg['starred'] = not msg.get('starred', False)
            # Устанавливаем цель прокрутки на это сообщение
            self.pending_scroll_target = (chat['filepath'], idx)
            self.save_current_chat()
            # keep_scroll=True – не скроллить вниз, а восстановить позицию у сообщения
            self._redraw_chat_with_history(keep_scroll=True)

    def show_starred_messages(self):
        if self.current_chat_index < 0:
            QMessageBox.information(self, "Нет чата", "Откройте чат для просмотра закладок.")
            return
        chat = self.chats[self.current_chat_index]
        dialog = StarredMessagesDialog(chat, self)
        if dialog.exec() == QDialog.Accepted:
            idx = dialog.get_selected_index()
            if idx is not None:
                self.navigate_to_message(chat['filepath'], idx)

    def show_message_context_menu_at(self, idx, global_pos, widget=None):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        if not (0 <= idx < len(chat['messages'])):
            return
        msg = chat['messages'][idx]
        menu = QMenu(self)

        # Определяем выделенный текст
        selected_text = ""
        if widget and hasattr(widget, 'content_label'):
            selected_text = widget.content_label.selectedText().strip()
        if not selected_text and widget and hasattr(widget, 'reasoning_label'):
            selected_text = widget.reasoning_label.selectedText().strip()

        if selected_text:
            copy_action = menu.addAction("Копировать выделенный текст")
        else:
            copy_action = menu.addAction("Копировать текст")

        # Остальные пункты меню
        if msg.get('starred', False):
            star_action = menu.addAction("Удалить из избранного")
        else:
            star_action = menu.addAction("В избранное")

        if msg.get('pinned', False):
            pin_action = menu.addAction("Открепить сообщение")
        else:
            pin_action = menu.addAction("Закрепить сообщение")

        if msg.get('role') == 'user':
            edit_action = menu.addAction("Редактировать")
        else:
            edit_action = menu.addAction("Редактировать")
        delete_action = menu.addAction("Удалить")
        quote_action = menu.addAction("Цитировать")
        copy_all_dialog_action = menu.addAction("Копировать весь диалог")

        # Подменю реакций
        reaction_menu = menu.addMenu("Реакция")
        reaction_like = reaction_menu.addAction("👍 Хороший ответ")
        reaction_dislike = reaction_menu.addAction("👎 Плохой ответ")
        reaction_heart = reaction_menu.addAction("❤️ Очень полезно")
        if msg.get('reaction'):
            reaction_remove = reaction_menu.addAction("Убрать реакцию")
        else:
            reaction_remove = None

        # Заметки
        if msg.get('comment'):
            comment_action = menu.addAction("Изменить заметку")
            remove_comment_action = menu.addAction("Удалить заметку")
        else:
            comment_action = menu.addAction("Добавить заметку")
            remove_comment_action = None

        chosen = menu.exec(global_pos)

        # Обработка выбора (единая цепочка if/elif)
        if chosen == copy_action:
            if selected_text:
                QApplication.clipboard().setText(selected_text)
            else:
                content = msg.get('content', '')
                QApplication.clipboard().setText(content)
            self.statusBar().showMessage("Текст скопирован")
        elif chosen == star_action:
            self.toggle_star_message(idx)
        elif chosen == pin_action:
            self.toggle_pin_message(idx)
        elif chosen == edit_action:
            if msg.get('role') == 'user':
                self.edit_message(idx)
            else:
                self.edit_assistant_message(idx)
        elif chosen == delete_action:
            self.delete_message(idx)
        elif chosen == quote_action:
            self.quote_message(idx)
        elif chosen == copy_all_dialog_action:
            self.copy_full_dialog(chat)
        elif chosen == comment_action:
            self.add_or_edit_comment(idx)
        elif remove_comment_action is not None and chosen == remove_comment_action:
            self.remove_comment(idx)
        elif chosen == reaction_like:
            self.set_message_reaction(idx, "👍")
        elif chosen == reaction_dislike:
            self.set_message_reaction(idx, "👎")
        elif chosen == reaction_heart:
            self.set_message_reaction(idx, "❤️")
        elif reaction_remove is not None and chosen == reaction_remove:
            self.remove_message_reaction(idx)


    def show_message_context_menu(self, pos):
        if self.current_chat_index < 0:
            return
        cursor = self.chat_browser.cursorForPosition(pos)
        doc_pos = cursor.position()
        chat = self.chats[self.current_chat_index]
        for start, end, idx in self.message_positions:
            if start <= doc_pos <= end:
                msg = chat['messages'][idx]
                menu = QMenu(self)

                # Копирование
                selected_text = self.chat_browser.textCursor().selectedText().strip()
                if selected_text:
                    copy_action = menu.addAction("Копировать выделенный текст")
                else:
                    copy_action = menu.addAction("Копировать текст")

                # Пункт избранного
                if msg.get('starred', False):
                    star_action = menu.addAction("Удалить из избранного")
                else:
                    star_action = menu.addAction("В избранное")

                # Закрепление
                if msg.get('pinned', False):
                    pin_action = menu.addAction("Открепить сообщение")
                else:
                    pin_action = menu.addAction("Закрепить сообщение")

                # Редактирование и удаление
                if msg.get('role') == 'user':
                    edit_action = menu.addAction("Редактировать")
                    delete_action = menu.addAction("Удалить")
                else:
                    edit_action = menu.addAction("Редактировать")
                    delete_action = menu.addAction("Удалить")

                chosen = menu.exec(self.chat_browser.mapToGlobal(pos))

                if chosen == star_action:
                    self.toggle_star_message(idx)
                    return
                elif chosen == pin_action:
                    self.toggle_pin_message(idx)
                elif chosen == copy_action:
                    if selected_text:
                        QApplication.clipboard().setText(selected_text)
                    else:
                        self.copy_message_text(msg)
                elif edit_action is not None and chosen == edit_action:
                    if msg.get('role') == 'user':
                        self.edit_message(idx)
                    else:
                        self.edit_assistant_message(idx)
                elif chosen == delete_action:
                    self.delete_message(idx)
                return

    def copy_message_text(self, msg):
        content = msg.get('content', '')
        # Если контент представлен списком (мультимодальность), извлекаем текстовые части
        if isinstance(content, list):
            parts = []
            for item in content:
                if isinstance(item, dict) and 'text' in item:
                    parts.append(item['text'])
            content = '\n'.join(parts)
        QApplication.clipboard().setText(content)
        self.statusBar().showMessage("Текст скопирован")

    def _on_message_action_edit(self, idx):
        if self.current_chat_index >= 0:
            role = self.chats[self.current_chat_index]['messages'][idx].get('role')
            if role == 'user':
                self.edit_message(idx)
            else:
                self.edit_assistant_message(idx)

    def _on_message_action_copy(self, idx):
        if self.current_chat_index >= 0:
            msg = self.chats[self.current_chat_index]['messages'][idx]
            content = msg.get('content', '')
            if isinstance(content, list):
                content = '\n'.join([x.get('text', '') for x in content if isinstance(x, dict)])
            QApplication.clipboard().setText(content)
            self.statusBar().showMessage("Текст скопирован")

    def _on_message_action_regenerate(self, idx):
        if self.current_chat_index < 0:
            return

        # Запрещаем регенерацию во время активной генерации
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "Занято", "Дождитесь завершения текущей генерации.")
            return

        chat = self.chats[self.current_chat_index]
        if idx < 0 or idx >= len(chat['messages']):
            return

        # Убедимся, что нажали именно на сообщение ассистента
        msg = chat['messages'][idx]
        if msg.get('role') != 'assistant':
            return

        # Найдём последний пользовательский запрос перед этим ответом
        last_user_idx = -1
        for i in range(idx - 1, -1, -1):
            if chat['messages'][i].get('role') == 'user':
                last_user_idx = i
                break

        if last_user_idx == -1:
            QMessageBox.information(self, "Ошибка", "Не найден пользовательский запрос для регенерации.")
            return

        # Отправляем запрос с историей до последнего сообщения пользователя включительно
        # Сам чат не изменяется – старый ответ остаётся, новый добавится после генерации
        self._send_request_after_edit(cutoff_idx=last_user_idx)

    def _on_message_action_unpin(self, idx):
        self.toggle_pin_message(idx)

    # ======================================================================
    #  ОТПРАВКА ЗАПРОСОВ И ОБРАБОТКА ОТВЕТОВ
    # ======================================================================

    def on_chunk_received(self, chunk: str):
        # Игнорируем сигналы, если генерация уже остановлена или это не текущий поток
        if self.worker is None or self.sender() != self.worker:
            return

        if not hasattr(self, 'current_response_text'):
            self.current_response_text = ""

        # Если это первый чанк и виджет содержит заглушку, очищаем её
        if (self.current_response_text == "" and
                hasattr(self, 'current_response_widget') and
                self.current_response_widget is not None):
            try:
                self.current_response_widget.set_content_text("")
            except RuntimeError:
                self.current_response_widget = None
            except AttributeError:
                self.current_response_widget = None

        self.current_response_text += chunk
        self.update_generation_speed()

        if (self.current_chat_index == self.generation_chat_index and
                hasattr(self, 'current_response_widget') and
                self.current_response_widget is not None):
            try:
                self.current_response_widget.set_content_text(self.current_response_text)
            except RuntimeError:
                self.current_response_widget = None
            except AttributeError:
                self.current_response_widget = None
            QTimer.singleShot(0, self._do_scroll_to_bottom)

    def on_finished(self, full_text: str, reasoning_text: str = ""):
        self.hide_stop_button()

        if not full_text.strip() and not reasoning_text.strip():
            QMessageBox.warning(self, "Внимание", "Модель вернула пустой ответ.")
            self.statusBar().showMessage("Ошибка: пустой ответ")
            self.worker = None
            self.is_generating = False
            self.generation_chat_index = -1
            return

        # === ВЫЗОВ ПЛАГИНОВ: обработка ответа после получения ===
        api = PluginAPI(self)
        for plugin in self.plugins:
            try:
                if hasattr(plugin, 'on_response_received'):
                    full_text = plugin.on_response_received(full_text, api)
            except Exception as e:
                self.logger.warning(f"Ошибка в плагине {plugin.__class__.__name__}: {e}")
                self.show_warning(
                    "Ошибка плагина",
                    f"Плагин {plugin.__class__.__name__} вызвал ошибку и был пропущен."
                )

        # === ЛОГИРОВАНИЕ ===
        self.logger.info(f"Ответ получен: {full_text[:100]}")

        target_idx = self.generation_chat_index
        original_idx = self.current_chat_index

        if 0 <= target_idx < len(self.chats):
            chat = self.chats[target_idx]
            self.current_chat_index = target_idx
            self.add_message_to_current("assistant", full_text, reasoning_text)

            if self._pending_tail:
                chat['messages'].extend(self._pending_tail)
                self._pending_tail = None
                self.save_chat_by_filepath(chat['filepath'])

            if hasattr(self, 'generation_start_time'):
                duration = time.time() - self.generation_start_time
                chat['total_generation_time'] = chat.get('total_generation_time', 0.0) + duration
                chat['last_activity_time'] = time.strftime('%d.%m.%Y %H:%M:%S')
                assistant_tokens = (len(full_text) + len(reasoning_text)) / 4
                chat['total_tokens_sent'] = chat.get('total_tokens_sent', 0) + int(assistant_tokens)
                self.save_chat_by_filepath(chat['filepath'])
                del self.generation_start_time

            self.current_chat_index = original_idx

            if original_idx == target_idx:
                self._redraw_chat_with_history()
            else:
                self.notify_message_received(chat['filepath'])

        # Уведомления
        if self.settings.get("sound_notifications", True) or self.settings.get("toast_notifications", True):
            is_background = (target_idx != original_idx)
            window_inactive = not self.isActiveWindow() or self.isMinimized()
            if is_background or (target_idx == original_idx and window_inactive):
                chat_title = "?"
                if 0 <= target_idx < len(self.chats):
                    chat_title = self.chats[target_idx].get('title', '?')
                if self.settings.get("sound_notifications", True):
                    self.play_sound()
                if self.settings.get("toast_notifications", True):
                    if full_text.strip():
                        preview = full_text[:80] + ('...' if len(full_text) > 80 else '')
                    else:
                        preview = reasoning_text[:80] + ('...' if len(reasoning_text) > 80 else '')
                    self.show_toast(f"Ответ готов: {chat_title}", preview)

        self.speed_label.clear()
        self.statusBar().showMessage("Готов")
        self.worker = None
        self.is_generating = False
        self.generation_chat_index = -1
        self.speed_timer.stop()
        self.speed_label.clear()

    def _on_scroll_range_changed(self, min_val, max_val):
        """При росте содержимого, если пользователь у низа — прокручиваем вниз."""

        if self._stick_to_bottom:
            self.chat_scroll_area.verticalScrollBar().setValue(max_val)


    def on_error(self, error_msg: str):
        self.logger.error(f"Ошибка API: {error_msg}")
        self.hide_stop_button()
        self._remove_thinking_indicator()
        self._save_partial_assistant_response()

        msg_lower = error_msg.lower()
        base = self.settings.get("api_base", "")

        if "connection" in msg_lower or "соединения" in msg_lower or "refused" in msg_lower:
            friendly = (
                "Не удалось подключиться к LM Studio.\n\n"
                f"Адрес: {base}\n\n"
                "Что проверить:\n"
                "  1. LM Studio запущен?\n"
                "  2. Локальный сервер включён (Local Server → Start)?\n"
                "  3. Порт совпадает с указанным в настройках?\n"
                "  4. Модель загружена?"
            )
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Warning)
            box.setWindowTitle("Сервер недоступен")
            box.setText(friendly)
            btn_retry = box.addButton("Повторить", QMessageBox.AcceptRole)
            btn_launch = box.addButton("Запустить LM Studio", QMessageBox.ActionRole)
            box.addButton("Отмена", QMessageBox.RejectRole)
            box.exec()

            if box.clickedButton() == btn_retry:
                self.send_message()
                return
            elif box.clickedButton() == btn_launch:
                self.launch_lm_studio()
                return
        elif "timeout" in msg_lower or "timed out" in msg_lower:
            self.show_error(
                "Таймаут",
                "Сервер не ответил за отведённое время.\n\n"
                "Возможные причины:\n"
                "  • Модель ещё загружается — подожди и попробуй снова.\n"
                "  • Модель слишком большая для твоего железа.\n"
                "  • Сервер завис — перезапусти LM Studio."
            )
        elif "model not found" in msg_lower or "404" in msg_lower:
            self.show_error(
                "Модель не найдена",
                f"В LM Studio нет модели:\n{self.settings.get('model', '')}\n\n"
                "Проверь имя модели в настройках — оно должно совпадать с названием в LM Studio."
            )
        elif "401" in msg_lower or "403" in msg_lower or "unauthorized" in msg_lower:
            self.show_error(
                "Требуется авторизация",
                "Сервер требует API-ключ.\n\n"
                "Для локальных серверов это обычно означает неверную конфигурацию LM Studio."
            )
        elif "прервано на середине" in msg_lower:
            self.show_error(
                "Соединение прервано",
                "Связь с сервером оборвалась во время генерации.\n\n"
                "Проверь сеть или перезапусти LM Studio, затем попробуй снова."
            )
        else:
            self.show_error("Ошибка API", f"Произошла ошибка при генерации:\n\n{error_msg}")

        self.statusBar().showMessage("Ошибка")
        self.worker = None
        self.is_generating = False
        self.generation_chat_index = -1

        self.speed_timer.stop()
        self.speed_label.clear()

    def stop_generation(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()

    def on_stop_generation(self):
        self.hide_stop_button()
        self.statusBar().showMessage("Генерация остановлена")
        self.show_stop_notification()
        self._save_partial_assistant_response()  # сохранит ответ без перерисовки

        self.worker = None
        self.is_generating = False
        self.generation_chat_index = -1

        self.speed_timer.stop()
        self.speed_label.clear()

        # Прокручиваем вниз
        self.scroll_chat_to_bottom()

    def _save_partial_assistant_response(self):
        """Сохраняет частичный ответ ассистента, если он не пустой, БЕЗ перерисовки."""

        if not hasattr(self, 'current_response_text') or not self.current_response_text.strip():
            return
        target_idx = self.generation_chat_index
        if 0 <= target_idx < len(self.chats):
            chat = self.chats[target_idx]
            msg = {
                "role": "assistant",
                "content": self.current_response_text.strip(),
                "time": time.strftime("%d.%m.%Y %H:%M:%S")
            }
            chat['messages'].append(msg)
            if hasattr(self, 'generation_start_time'):
                duration = time.time() - self.generation_start_time
                chat['total_generation_time'] = chat.get('total_generation_time', 0.0) + duration
                chat['last_activity_time'] = time.strftime('%d.%m.%Y %H:%M:%S')
                assistant_tokens = len(self.current_response_text) / 4
                chat['total_tokens_sent'] = chat.get('total_tokens_sent', 0) + int(assistant_tokens)
                del self.generation_start_time
            self.save_chat_by_filepath(chat['filepath'])
            # НЕ вызываем _redraw_chat_with_history, чтобы сохранить скролл
            if self.current_chat_index != target_idx:
                self.notify_message_received(chat['filepath'])
        self.current_response_text = ""

    # ======================================================================
    #  ВЛОЖЕНИЯ И ФАЙЛЫ
    # ======================================================================

    def attach_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы",
            "",
            "Все поддерживаемые (*.txt *.md *.py *.json *.csv *.html *.log *.pdf *.docx *.pptx *.png *.jpg *.jpeg *.gif *.bmp *.webp);;Текстовые (*.txt *.md *.py *.json *.csv *.html *.log);;Документы (*.pdf *.docx *.pptx);;Изображения (*.png *.jpg *.jpeg *.gif *.bmp *.webp);;Все файлы (*)"
        )
        if not file_paths:
            return
        for path in file_paths:
            self.attachments.append(path)
        self.update_attachments_ui()

    def update_attachments_ui(self):
        # Очищаем layout
        while self.attachments_layout.count():
            item = self.attachments_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # Выравниваем по левому краю
        self.attachments_layout.setAlignment(Qt.AlignLeft)

        # Для каждого прикреплённого файла создаём компактный контейнер
        for path in self.attachments:
            container = QWidget()
            container.setStyleSheet("""
                QWidget {
                    background-color: #FCE4EC;  /* светло-розовый */
                    border-radius: 8px;
                }
            """)
            container_layout = QHBoxLayout(container)
            container_layout.setContentsMargins(4, 2, 4, 2)
            container_layout.setSpacing(2)

            # Обрезаем длинное имя файла
            name = os.path.basename(path)
            if len(name) > 20:
                name = name[:17] + "..."
            label = QLabel(name)
            label.setToolTip(path)
            label.setStyleSheet("background: transparent; border: none; font-size: 12px; color: #333;")

            # Маленькая круглая кнопка удаления
            btn_remove = QPushButton()
            btn_remove.setIcon(self.icon_close)
            btn_remove.setIconSize(QSize(12, 12))
            btn_remove.setFixedSize(16, 16)
            btn_remove.setStyleSheet("""
                QPushButton {
                    border-radius: 8px;
                    background-color: #D32F2F;
                    border: none;
                }
                QPushButton:hover {
                    background-color: #B71C1C;
                }
            """)
            btn_remove.clicked.connect(lambda checked, p=path: self.remove_attachment(p))

            container_layout.addWidget(label)
            container_layout.addWidget(btn_remove)

            # Фиксированная ширина контейнера
            container.setFixedWidth(150)

            self.attachments_layout.addWidget(container)

        # Растяжка, чтобы элементы прижимались влево
        self.attachments_layout.addStretch(1)

        # Показываем/скрываем область прокрутки
        self.attachments_scroll.setVisible(bool(self.attachments))

    def remove_attachment(self, path):
        """Удаляет прикреплённый файл из списка."""

        if path in self.attachments:
            self.attachments.remove(path)
            self.update_attachments_ui()

    def read_file_content(self, path):
        ext = os.path.splitext(path)[1].lower()
        text = ""
        try:
            if ext in ('.txt', '.md', '.py', '.json', '.csv', '.html', '.log'):
                with open(path, 'r', encoding='utf-8') as f:
                    text = f.read()
            elif ext == '.docx' and DOCX_AVAILABLE:
                doc = Document(path)
                text = '\n'.join([p.text for p in doc.paragraphs])
            elif ext == '.pdf' and PDF_AVAILABLE:
                reader = PdfReader(path)
                text = '\n'.join([page.extract_text() or '' for page in reader.pages])
            elif ext == '.pptx' and PPTX_AVAILABLE:
                prs = Presentation(path)
                slides_text = []
                for slide in prs.slides:
                    slide_text = []
                    for shape in slide.shapes:
                        if hasattr(shape, "text_frame") and shape.text_frame.text.strip():
                            slide_text.append(shape.text_frame.text.strip())
                        if shape.has_table:
                            for row in shape.table.rows:
                                cells = [cell.text.strip() for cell in row.cells]
                                slide_text.append(" | ".join(cells))
                    if slide_text:
                        slides_text.append("\n".join(slide_text))
                text = "\n\n".join(slides_text)
            else:
                # Попытка прочитать как текст
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    text = f.read()
            return text if text.strip() else None
        except Exception as e:
            print(f"Ошибка чтения файла {path}: {e}")
            return None

    def extract_images_from_file(self, path):
        """Извлекает изображения из PDF или DOCX и возвращает список data URI."""

        images = []
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == '.pdf' and PDF_IMAGES_AVAILABLE:
                pdf = pdfium.PdfDocument(path)
                try:
                    for page_index in range(len(pdf)):
                        page = pdf[page_index]
                        try:
                            objects = page.get_objects()
                        except Exception:
                            continue
                        for obj in objects:
                            if not isinstance(obj, pdfium.PdfImage):
                                continue
                            try:
                                bitmap = obj.get_bitmap(render=False)
                                pil_image = bitmap.to_pil()
                                buffer = io.BytesIO()
                                pil_image.save(buffer, format="PNG")
                                img_bytes = buffer.getvalue()
                                data_uri = "data:image/png;base64," + base64.b64encode(img_bytes).decode('utf-8')
                                images.append(data_uri)
                            except Exception as e:
                                self.logger.warning(f"Не удалось извлечь картинку из PDF: {e}")
                                continue
                finally:
                    pdf.close()
            elif ext == '.docx' and DOCX_AVAILABLE:
                doc = Document(path)
                for rel in doc.part.rels.values():
                    if "image" in rel.reltype:
                        image_part = rel.target_part
                        img_bytes = image_part.blob
                        data_uri = "data:image/png;base64," + base64.b64encode(img_bytes).decode('utf-8')
                        images.append(data_uri)
            elif ext == '.pptx' and PPTX_AVAILABLE:
                prs = Presentation(path)
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if shape.shape_type == 13:
                            image = shape.image
                            img_bytes = image.blob
                            data_uri = "data:image/png;base64," + base64.b64encode(img_bytes).decode('utf-8')
                            images.append(data_uri)
        except Exception as e:
            self.logger.warning(f"Ошибка извлечения изображений из {path}: {e}")
        return images

    def show_attachment_preview(self, attachment: dict):
        if attachment.get('type') == 'image':
            path = attachment.get('path')
            if not path or not os.path.exists(path):
                QMessageBox.warning(self, "Ошибка", "Файл не найден.")
                return
            try:
                with open(path, 'rb') as f:
                    raw = f.read()
                if raw.startswith(b"ENC:"):
                    raw = self._decrypt_bytes(raw)
                pixmap = QPixmap()
                pixmap.loadFromData(raw)
                if pixmap.isNull():
                    QMessageBox.warning(self, "Ошибка", "Не удалось загрузить изображение.")
                    return
                dialog = ImagePreviewDialog(pixmap, self)
                dialog.exec()
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось открыть изображение:\n{e}")
        else:
            dialog = MessagePreviewDialog("", [attachment], self, show_send_button=False)
            dialog.exec()

    # ======================================================================
    #  ИМПОРТ/ЭКСПОРТ
    # ======================================================================

    def import_deepseek_history(self):
        # Создаём диалог выбора файла с иконкой
        dlg = QFileDialog(self, "Импортировать историю DeepSeek", "", "JSON Files (*.json)")
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setWindowIcon(QIcon(self.resource_path("Image/Import_2.png")))
        dlg.setFileMode(QFileDialog.ExistingFile)
        dlg.setViewMode(QFileDialog.Detail)

        if not dlg.exec():
            return

        file_path = dlg.selectedFiles()[0]
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            imported_chats = []
            if isinstance(data, list):
                if all(isinstance(item, dict) and ('role' in item or 'content' in item) for item in data):
                    converted = []
                    for item in data:
                        role = item.get("role")
                        content = item.get("content")
                        if role in ("user", "assistant", "system"):
                            converted.append(
                                {"role": role, "content": str(content), "time": time.strftime("%d.%m.%Y %H:%M:%S")})
                    if converted:
                        imported_chats.append(("Импорт: " + os.path.basename(file_path), converted))
                else:
                    for chat_obj in data:
                        if not isinstance(chat_obj, dict):
                            continue
                        title = chat_obj.get("title") or "Без названия"
                        msgs = self.extract_messages_from_deepseek_chat(chat_obj)
                        if msgs:
                            imported_chats.append((title, msgs))
            elif isinstance(data, dict):
                if "mapping" in data:
                    title = data.get("title") or "Без названия"
                    msgs = self.extract_messages_from_deepseek_chat(data)
                    if msgs:
                        imported_chats.append((title, msgs))
                elif "messages" in data:
                    msgs = data["messages"]
                    if isinstance(msgs, list):
                        converted = []
                        for m in msgs:
                            if isinstance(m, dict) and m.get("role") in ("user", "assistant", "system"):
                                converted.append(m)
                        if converted:
                            imported_chats.append(("Импорт: " + os.path.basename(file_path), converted))
            if not imported_chats:
                raise ValueError("Не удалось извлечь сообщения из файла.")
            selected_chats = []
            if len(imported_chats) == 1:
                selected_chats = imported_chats
            else:
                dialog = DeepSeekImportDialog(imported_chats, self)
                if dialog.exec() == QDialog.Accepted:
                    selected_chats = dialog.get_selected()
                else:
                    return
            # Определяем папку, если нужно её создать
            folder_to_use = ''
            if len(selected_chats) > 1:
                reply = QMessageBox.question(
                    self, "Создать папку?",
                    "Создать общую папку для выбранных чатов?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    folder_to_use = self._prompt_for_folder_name()
            elif len(selected_chats) == 1:
                reply = QMessageBox.question(
                    self, "Создать папку?",
                    "Создать папку для этого чата?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    folder_to_use = self._prompt_for_folder_name()

            for title, msgs in selected_chats:
                clean_msgs = [m for m in msgs if m["role"] != "system"]
                self.create_chat_with_messages(title, clean_msgs, folder_to_use)
                self.statusBar().showMessage(f"Импортирован чат «{title}»: {len(clean_msgs)} сообщений")
        except Exception as e:
            self.logger.exception("Ошибка импорта DeepSeek")
            self.show_error("Ошибка импорта", f"Не удалось импортировать файл.\nПричина: {e}")



    def extract_messages_from_deepseek_chat(self, chat_obj: dict) -> List[Dict[str, str]]:
        """Рекурсивно ищет все фрагменты сообщений в объекте чата DeepSeek."""

        found = []

        def recursive_search(obj):
            if isinstance(obj, dict):
                # Ищем ключи, содержащие "fragments" или "message"
                if 'message' in obj and isinstance(obj['message'], dict):
                    fragments = obj['message'].get('fragments', [])
                    for frag in fragments:
                        if isinstance(frag, dict):
                            ftype = frag.get('type', '')
                            content = frag.get('content', '')
                            if ftype in ('REQUEST', 'RESPONSE'):
                                role = 'user' if ftype == 'REQUEST' else 'assistant'
                                found.append({"role": role, "content": content})
                # Обходим значения
                for value in obj.values():
                    recursive_search(value)
            elif isinstance(obj, list):
                for item in obj:
                    recursive_search(item)

        recursive_search(chat_obj)
        # Убираем дубликаты (по порядку)
        unique = []
        for msg in found:
            if msg not in unique:
                unique.append(msg)
        return unique

    def import_zip_archive(self):
        """Импортирует ZIP-архив с чатами и служебными файлами."""

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите ZIP-архив", "", "ZIP Archive (*.zip)"
        )
        if not file_path:
            return

        try:
            temp_dir = "import_temp"
            os.makedirs(temp_dir, exist_ok=True)

            with zipfile.ZipFile(file_path, 'r') as zipf:
                zipf.extractall(temp_dir)

            base = os.path.dirname(os.path.abspath(__file__))

            # Переносим папку chats
            chats_src = os.path.join(temp_dir, "chats")
            if os.path.isdir(chats_src):
                chats_dest = os.path.join(base, "chats")
                if os.path.exists(chats_dest):
                    shutil.rmtree(chats_dest)
                shutil.copytree(chats_src, chats_dest)

            # Переносим attachments
            attachments_src = os.path.join(temp_dir, "attachments")
            if os.path.isdir(attachments_src):
                attachments_dest = os.path.join(base, "attachments")
                if os.path.exists(attachments_dest):
                    shutil.rmtree(attachments_dest)
                shutil.copytree(attachments_src, attachments_dest)

            # Переносим служебные файлы
            for name in ["folders.json", "tree_order.json", "trash.json", "prompts.json", "last_session.json"]:
                src = os.path.join(temp_dir, name)
                if os.path.exists(src):
                    shutil.copy2(src, os.path.join(base, name))

            shutil.rmtree(temp_dir)

            # Перезагружаем данные
            self.chats.clear()
            self.archived_chats.clear()
            self.chat_tree.clear()
            self.chat_items.clear()
            self.folder_items.clear()
            self.folders = self.load_folders()
            self.trash_items = self.load_trash()
            self.prompts = self.load_prompts()
            self.load_chats_from_disk()
            self.save_tree_order()
            self.reorder_tree_by_pinned()
            self.show_empty_chat_placeholder()
            self.statusBar().showMessage("Импорт завершён")
        except Exception as e:
            self.logger.exception("Ошибка импорта ZIP")
            self.show_error("Ошибка импорта", f"Не удалось импортировать ZIP-архив.\nПричина: {e}")

    def export_current_chat_docx(self):
        """Экспортирует текущий чат в документ Word (DOCX)."""

        if self.current_chat_index < 0:
            QMessageBox.warning(self, "Ошибка", "Нет выбранного чата для экспорта.")
            return

        if not DOCX_AVAILABLE:
            QMessageBox.warning(
                self,
                "Библиотека не установлена",
                "Для экспорта в DOCX установите python-docx:\npip install python-docx"
            )
            return

        chat = self.chats[self.current_chat_index]

        # === Диалог сохранения с иконкой ===
        dlg = QFileDialog(self, "Экспорт чата в DOCX")
        dlg.setWindowIcon(QIcon(self.resource_path("Image/Export_2.png")))
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilters(["Word Document (*.docx)", "All Files (*)"])
        dlg.setDefaultSuffix("docx")
        if not dlg.exec():
            return
        selected = dlg.selectedFiles()
        if not selected:
            return
        file_path = selected[0]
        # ===================================

        try:
            doc = Document()

            # Заголовок документа
            doc.add_heading(f"Чат: {chat.get('title', 'Без названия')}", level=1)

            # Проходим по сообщениям
            for msg in chat['messages']:
                role = msg.get('role', '')
                content = msg.get('content', '')
                time_str = msg.get('time', '')
                reasoning = msg.get('reasoning', '')

                # Определяем отправителя
                if role == 'user':
                    sender = "Пользователь"
                elif role == 'assistant':
                    sender = "Ассистент"
                else:
                    continue  # пропускаем системные сообщения

                # Добавляем заголовок сообщения (отправитель и время)
                heading_text = sender
                if time_str:
                    heading_text += f" ({time_str})"
                doc.add_heading(heading_text, level=2)

                # Обрабатываем контент (может быть строкой или списком для мультимодальных)
                if isinstance(content, list):
                    # Для мультимодальных сообщений извлекаем текстовые части
                    text_parts = []
                    for item in content:
                        if isinstance(item, dict):
                            if 'text' in item:
                                text_parts.append(item['text'])
                            elif 'image_url' in item:
                                text_parts.append('[Изображение]')
                            else:
                                text_parts.append(str(item))
                        else:
                            text_parts.append(str(item))
                    content = '\n'.join(text_parts)

                if content and content.strip():
                    doc.add_paragraph(content)

                # Добавляем reasoning ассистента (если есть и нужно)
                if role == 'assistant' and reasoning and reasoning.strip():
                    p = doc.add_paragraph()
                    run = p.add_run("Мысли модели:\n" + reasoning)
                    run.italic = True
                    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

                # Обрабатываем вложения (если они есть)
                if 'attachments' in msg and msg['attachments']:
                    doc.add_paragraph("Вложения:", style='Caption')
                    for att in msg['attachments']:
                        att_type = att.get('type')
                        if att_type == 'image':
                            # Вставляем изображение, если путь существует
                            img_path = att.get('path')
                            if img_path and os.path.exists(img_path):
                                doc.add_picture(img_path, width=Inches(4.0))
                            else:
                                doc.add_paragraph(f"[Изображение: {att.get('name', 'без имени')}]")
                        elif att_type == 'file_text':
                            file_name = att.get('name', 'файл')
                            doc.add_paragraph(f"📄 {file_name} (текстовый файл)")
                        elif att_type == 'file':
                            file_name = att.get('name', 'файл')
                            doc.add_paragraph(f"📁 {file_name} (файл)")

                # Разделитель между сообщениями (пустая строка)
                doc.add_paragraph()

            # Сохраняем документ
            doc.save(file_path)
            self.statusBar().showMessage(f"Чат экспортирован в {file_path}")
        except Exception as e:
            self.logger.exception("Ошибка экспорта в DOCX")
            self.show_error("Ошибка экспорта", f"Не удалось создать DOCX.\nПричина: {e}")

    def export_current_chat_json(self):
        if self.current_chat_index < 0:
            QMessageBox.warning(self, "Ошибка", "Нет выбранного чата для экспорта.")
            return
        chat = self.chats[self.current_chat_index]

        # === Диалог сохранения с иконкой ===
        dlg = QFileDialog(self, "Экспорт чата в JSON")
        dlg.setWindowIcon(QIcon(self.resource_path("Image/Export_2.png")))
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilters(["JSON Files (*.json)", "All Files (*)"])
        dlg.setDefaultSuffix("json")
        if not dlg.exec():
            return
        selected = dlg.selectedFiles()
        if not selected:
            return
        file_path = selected[0]
        # ===================================

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(chat['messages'], f, ensure_ascii=False, indent=2)
            self.statusBar().showMessage(f"Чат экспортирован в {file_path}")
        except Exception as e:
            self.logger.exception("Ошибка экспорта в JSON")
            self.show_error("Ошибка экспорта", f"Не удалось сохранить JSON.\nПричина: {e}")

    def export_current_chat_pdf(self):
        if self.current_chat_index < 0:
            QMessageBox.warning(self, "Ошибка", "Нет выбранного чата для экспорта.")
            return

            # === Диалог сохранения с иконкой ===
        dlg = QFileDialog(self, "Экспорт чата в PDF")
        dlg.setWindowIcon(QIcon(self.resource_path("Image/Export_2.png")))
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilters(["PDF Files (*.pdf)", "All Files (*)"])
        dlg.setDefaultSuffix("pdf")
        if not dlg.exec():
            return
        selected = dlg.selectedFiles()
        if not selected:
            return
        file_path = selected[0]
        # ===================================

        chat = self.chats[self.current_chat_index]
        html_parts = []

        for msg in chat['messages']:
            role = msg.get('role', '')
            content = msg.get('content', '')
            time_str = msg.get('time', '')

            # Определяем отправителя и цвет
            if role == 'user':
                sender = "Пользователь"
                color = "#1a73e8"
            else:
                sender = "Ассистент"
                color = "#188038"

            # Подготавливаем текст сообщения (если есть)
            if isinstance(content, list):
                content_parts = []
                for item in content:
                    if isinstance(item, dict):
                        if 'text' in item:
                            content_parts.append(item['text'])
                        elif 'image_url' in item:
                            content_parts.append('[Изображение]')
                        else:
                            content_parts.append(str(item))
                    else:
                        content_parts.append(str(item))
                content = ' '.join(content_parts)

            # Блок отправителя и текста
            html_parts.append(
                f'<p style="color:{color}; font-weight:bold;">{sender} '
                f'<span style="color:#888; font-weight:normal;">({time_str})</span></p>'
            )
            if content.strip():
                html_parts.append(
                    f'<p style="white-space:pre-wrap;">{self._escape_html(content)}</p>'
                )

            # Обрабатываем вложения пользователя (если есть)
            if role == 'user' and 'attachments' in msg and msg['attachments']:
                for att in msg['attachments']:
                    att_type = att.get('type')
                    if att_type == 'image':
                        # Вставляем изображение
                        img_data = att.get('data', '')
                        if img_data:
                            html_parts.append(
                                f'<img src="{img_data}" style="max-width:500px; margin:5px;" />'
                            )
                    elif att_type == 'file_text':
                        # Для текстовых файлов показываем только название
                        file_name = att.get('name', 'файл')
                        html_parts.append(
                            f'<p style="color:#555;">📄 {self._escape_html(file_name)}</p>'
                        )
                    elif att_type == 'file':
                        file_name = att.get('name', 'файл')
                        html_parts.append(
                            f'<p style="color:#555;">📄 {self._escape_html(file_name)}</p>'
                        )

            # Разделитель между сообщениями
            html_parts.append("<hr>")

        full_html = "<html><body>" + "".join(html_parts) + "</body></html>"

        try:
            # Настройка принтера для PDF
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(file_path)

            # Печать документа
            doc = QTextDocument()
            doc.setHtml(full_html)
            doc.print_(printer)

            self.statusBar().showMessage(f"PDF сохранён в {file_path}")
        except Exception as e:
            self.logger.exception("Ошибка экспорта в PDF")
            self.show_error("Ошибка экспорта", f"Не удалось создать PDF.\nПричина: {e}")

        self.statusBar().showMessage(f"PDF сохранён в {file_path}")

    def export_all_chats_zip(self):
        # === Диалог сохранения с иконкой ===
        dlg = QFileDialog(self, "Экспорт всех чатов в ZIP")
        dlg.setWindowIcon(QIcon(self.resource_path("Image/Export_2.png")))
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilters(["ZIP Archive (*.zip)", "All Files (*)"])
        dlg.setDefaultSuffix("zip")
        if not dlg.exec():
            return
        selected = dlg.selectedFiles()
        if not selected:
            return
        file_path = selected[0]
        # ===================================

        try:
            with zipfile.ZipFile(file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Добавляем папку chats
                for root, dirs, files in os.walk(self.chats_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, os.path.dirname(self.chats_dir))
                        zipf.write(full_path, arcname)

                # Добавляем служебные файлы, если они есть
                for extra in ["folders.json", "tree_order.json", "trash.json", "prompts.json", "chat_config.json"]:
                    if os.path.exists(extra):
                        zipf.write(extra, extra)

            self.statusBar().showMessage(f"Архив сохранён в {file_path}")
        except Exception as e:
            self.logger.exception("Ошибка экспорта в ZIP")
            self.show_error("Ошибка экспорта", f"Не удалось создать ZIP-архив.\nПричина: {e}")

    def show_export_menu(self):
        menu = QMenu(self)
        dark = self.settings.get("theme", "light") == "dark"

        if dark:
            json_action = menu.addAction(self.icon_json_white, "Экспорт в JSON")
            pdf_action = menu.addAction(self.icon_pdf_white, "Экспорт в PDF")
            markdown_action = menu.addAction(self.icon_markdown_white, "Экспорт в Markdown")
            docx_action = menu.addAction(QIcon("Image_White/Docx_White.png"), "Экспорт в DOCX")
            html_action = menu.addAction(self.icon_html_white, "Экспорт в HTML")
            zip_action = menu.addAction(self.icon_zip_white, "Экспорт всех чатов в ZIP")
        else:
            json_action = menu.addAction(self.icon_json, "Экспорт в JSON")
            pdf_action = menu.addAction(self.icon_pdf, "Экспорт в PDF")
            markdown_action = menu.addAction(self.icon_markdown, "Экспорт в Markdown")
            docx_action = menu.addAction(QIcon("Image/Docx.png"), "Экспорт в DOCX")
            html_action = menu.addAction(self.icon_html, "Экспорт в HTML")
            zip_action = menu.addAction(self.icon_zip, "Экспорт всех чатов в ZIP")

        action = menu.exec(self.btn_export.mapToGlobal(QPoint(0, self.btn_export.height())))

        if action == json_action:
            self.export_current_chat_json()
        elif action == pdf_action:
            self.export_current_chat_pdf()
        elif action == zip_action:
            self.export_all_chats_zip()
        elif action == markdown_action:
            self.export_current_chat_markdown()
        elif action == html_action:
            self.export_current_chat_html()
        elif action == docx_action:
            self.export_current_chat_docx()

    def _render_file_preview(self, file_name, text_content):
        """Возвращает HTML-код для предпросмотра файла (кликабельная ссылка)."""

        # Временное сохранение текста для последующего открытия
        self.temp_file_text = text_content
        return f'<a href="#" style="color:#1a73e8; text-decoration:underline;" onclick="return false;">📄 {file_name}</a>'

    def export_current_chat_markdown(self):
        """Экспортирует текущий чат в Markdown-файл."""

        if self.current_chat_index < 0:
            QMessageBox.warning(self, "Ошибка", "Нет выбранного чата для экспорта.")
            return

        chat = self.chats[self.current_chat_index]

        # === Диалог сохранения с иконкой ===
        dlg = QFileDialog(self, "Экспорт чата в Markdown")
        dlg.setWindowIcon(QIcon(self.resource_path("Image/Export_2.png")))
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilters(["Markdown (*.md)", "All Files (*)"])
        dlg.setDefaultSuffix("md")
        if not dlg.exec():
            return
        selected = dlg.selectedFiles()
        if not selected:
            return
        file_path = selected[0]
        # ===================================

        try:
            lines = [f"# {chat.get('title', 'Чат')}\n"]

            for msg in chat.get('messages', []):
                role = msg.get('role', '')
                content = msg.get('content', '')
                time_str = msg.get('time', '')

                if isinstance(content, list):
                    text_parts = []
                    for item in content:
                        if isinstance(item, dict) and 'text' in item:
                            text_parts.append(item['text'])
                    content = '\n'.join(text_parts)

                if role == 'user':
                    header = "## 👤 Пользователь"
                elif role == 'assistant':
                    header = "## 🤖 Ассистент"
                else:
                    continue

                if time_str:
                    header += f"  <sub>{time_str}</sub>"

                lines.append(header)
                lines.append("")
                lines.append(content)
                lines.append("")
                lines.append("---")
                lines.append("")

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(lines))

            self.statusBar().showMessage(f"Чат экспортирован в {file_path}")
        except Exception as e:
            self.logger.exception("Ошибка экспорта в Markdown")
            self.show_error("Ошибка экспорта", f"Не удалось сохранить Markdown.\nПричина: {e}")

    def export_current_chat_html(self):
        """Экспортирует текущий чат в HTML-файл."""
        if self.current_chat_index < 0:
            QMessageBox.warning(self, "Ошибка", "Нет выбранного чата для экспорта.")
            return

        chat = self.chats[self.current_chat_index]

        # === Диалог сохранения с иконкой ===
        dlg = QFileDialog(self, "Экспорт чата в HTML")
        dlg.setWindowIcon(QIcon(self.resource_path("Image/Export_2.png")))
        dlg.setOption(QFileDialog.DontUseNativeDialog, True)
        dlg.setAcceptMode(QFileDialog.AcceptSave)
        dlg.setNameFilters(["HTML (*.html)", "All Files (*)"])
        dlg.setDefaultSuffix("html")
        if not dlg.exec():
            return
        selected = dlg.selectedFiles()
        if not selected:
            return
        file_path = selected[0]
        # ===================================

        try:
            title = chat.get('title', 'Чат')
            html_parts = [
                "<!DOCTYPE html>",
                "<html lang='ru'>",
                "<head>",
                "<meta charset='UTF-8'>",
                f"<title>{self._escape_html(title)}</title>",
                "<style>",
                "body { font-family: Arial, sans-serif; max-width: 900px; margin: 20px auto; padding: 0 20px; line-height: 1.6; color: #000; background: #fff; }",
                ".msg { margin: 16px 0; padding: 12px 16px; border-radius: 12px; }",
                ".user { background: #e3f2fd; margin-left: 15%; }",
                ".assistant { background: #f1f1f1; margin-right: 15%; }",
                ".role { font-weight: 600; margin-bottom: 6px; }",
                ".user .role { color: #1a73e8; }",
                ".assistant .role { color: #188038; }",
                ".time { color: #999; font-size: 0.85em; margin-left: 8px; }",
                ".reasoning { background: #fff9c4; border-left: 4px solid #fbc02d; padding: 8px 12px; margin-top: 6px; border-radius: 6px; font-style: italic; color: #555; }",
                ".code-block { background: #272822; color: #f8f8f2; padding: 10px; border-radius: 6px; overflow-x: auto; font-family: Consolas, Monaco, monospace; }",
                ".code-block code { white-space: pre; }",
                "table { border-collapse: collapse; width: 100%; margin: 10px 0; }",
                "th, td { border: 1px solid #ccc; padding: 8px; }",
                "th { background: #e0e0e0; }",
                "hr { border: none; border-top: 1px solid #ddd; margin: 20px 0; }",
                "</style>",
                "</head>",
                "<body>",
                f"<h1>{self._escape_html(title)}</h1>",
                "<hr>",
            ]

            for msg in chat.get('messages', []):
                role = msg.get('role', '')
                content = msg.get('content', '')
                time_str = msg.get('time', '')
                reasoning = msg.get('reasoning', '')

                if isinstance(content, list):
                    text_parts = []
                    for item in content:
                        if isinstance(item, dict) and 'text' in item:
                            text_parts.append(item['text'])
                    content = '\n'.join(text_parts)

                if role == 'user':
                    body_html = self._escape_html(content).replace('\n', '<br>')
                    html_parts.append(
                        f"<div class='msg user'>"
                        f"<div class='role'>👤 Пользователь"
                        f"<span class='time'>{self._escape_html(time_str)}</span></div>"
                        f"<div>{body_html}</div>"
                        f"</div>"
                    )
                elif role == 'assistant':
                    body_html = self._markdown_to_html(content)
                    reasoning_html = ""
                    if reasoning:
                        reasoning_html = (
                            f"<div class='reasoning'>"
                            f"<b>🧠 Мысли:</b><br>"
                            f"{self._escape_html(reasoning).replace(chr(10), '<br>')}"
                            f"</div>"
                        )
                    html_parts.append(
                        f"<div class='msg assistant'>"
                        f"<div class='role'>🤖 Ассистент"
                        f"<span class='time'>{self._escape_html(time_str)}</span></div>"
                        f"{reasoning_html}"
                        f"<div>{body_html}</div>"
                        f"</div>"
                    )

            html_parts.append("</body></html>")

            with open(file_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(html_parts))

            self.statusBar().showMessage(f"Чат экспортирован в {file_path}")
        except Exception as e:
            self.logger.exception("Ошибка экспорта в HTML")
            self.show_error("Ошибка экспорта", f"Не удалось сохранить HTML.\nПричина: {e}")

    # ======================================================================
    #  НАСТРОЙКИ И ПАРАМЕТРЫ
    # ======================================================================

    def open_settings_dialog(self):
        dialog = SettingsDialog(self.settings, self)
        if dialog.exec():
            new_settings = dialog.get_settings()
            self.settings.update(new_settings)
            self.save_settings()
            if self.settings.get("encrypt_chats", False) and not self.settings.get("encryption_salt"):
                self.setup_master_password()
            self.apply_system_prompt_to_all()
            self.apply_font_settings()
            self.rag_enabled = self.settings.get("rag_enabled", False)
            if self.rag_enabled:
                self.ensure_rag_folder()
                self.load_rag_documents_from_folder()
                self.start_rag_watcher()
            else:
                # Очищаем документы из памяти, чтобы они не использовались
                if self.rag_manager:
                    self.rag_manager.documents = []
                    self.rag_manager.index = None
                    self.rag_manager.embeddings = None
            self.statusBar().showMessage("Настройки обновлены")

    def open_chat_params_dialog(self):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        current_params = chat.get('params', {})
        current_system_prompt = chat.get('system_prompt', '')

        dialog = ChatParamsDialog(
            chat_params=current_params,
            global_settings=self.settings,
            current_system_prompt=current_system_prompt,
            parent=self
        )
        if dialog.exec() == QDialog.Accepted:
            result = dialog.get_params()
            chat['params'] = result['params']
            chat['system_prompt'] = result['system_prompt']
            self.save_current_chat()
            self.statusBar().showMessage("Параметры и системный промпт чата обновлены")

    def apply_system_prompt_to_current(self):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        sp = self.settings.get("system_prompt", "").strip()
        chat['messages'] = [m for m in chat['messages'] if m["role"] != "system"]
        if sp:
            chat['messages'].insert(0, {"role": "system", "content": sp, "time": time.strftime("%d.%m.%Y %H:%M:%S")})
        self.save_current_chat()

    def apply_system_prompt_to_all(self):
        for chat in self.chats:
            saved_idx = self.current_chat_index
            self.current_chat_index = self.chats.index(chat)
            self.apply_system_prompt_to_current()
            self.current_chat_index = saved_idx

    def get_max_images_for_model(self, model_name: str) -> int:
        """
        Возвращает максимальное допустимое количество изображений для конкретной модели.
        Если модель известна и имеет ограничение, возвращает это ограничение,
        иначе — значение из настроек (по умолчанию 4).
        """

        model_name_lower = model_name.lower()
        # Известные модели с ограничением в 1 изображение (мультимодальные)
        if any(kw in model_name_lower for kw in ["llava", "bakllava", "llama-3.2-vision", "vision"]):
            return 1
        # Можно добавить другие модели с известными лимитами
        # elif ...
        # Для неизвестных моделей используем настройку
        return self.settings.get("max_images_per_message", 4)

    def open_chat_appearance_dialog(self):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        dialog = ChatAppearanceDialog(
            self,
            current_bg_color=chat.get('background_color'),
            current_bg_image=chat.get('background_image'),
            current_bg_image_original=chat.get('background_image_original')
        )
        if dialog.exec() == QDialog.Accepted:
            chat['background_image'] = dialog.result_image_path
            chat['background_image_original'] = dialog.result_image_original
            chat['background_color'] = dialog.selected_color
            self.save_chat_by_filepath(chat['filepath'])
            self.apply_chat_background(chat)
            self._redraw_chat_with_history()

    def set_chat_background_color(self, color):
        self.chat_browser.setStyleSheet(f"QTextBrowser {{ background-color: {color}; }}")

    def set_chat_background_image(self, path):
        # Используем QPalette для установки фонового изображения
        palette = self.chat_browser.palette()
        brush = QBrush(QColor(255, 255, 255))
        brush.setTexture(QPixmap(path).scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        palette.setBrush(QPalette.Base, brush)
        self.chat_browser.setPalette(palette)
        self.chat_browser.setStyleSheet("")  # убираем цвет, если был

    def apply_chat_background(self, chat):

        if chat is None:
            # Фон для пустого места/папки
            if self.settings.get("theme", "light") == "dark":
                bg = "#2b2b2b"
            else:
                bg = "#ffffff"
            palette = self.chat_scroll_area.viewport().palette()
            palette.setBrush(QPalette.Base, QBrush(QColor(bg)))
            self.chat_scroll_area.viewport().setPalette(palette)
            self.chat_scroll_area.viewport().setAutoFillBackground(True)
            self.chat_scroll_area.setPalette(palette)
            self.chat_messages_widget.setStyleSheet("background: transparent;")
            return

        color = chat.get('background_color')
        image_path = chat.get('background_image')

        if image_path and os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            if not pixmap.isNull():
                # Масштабируем под размер viewport
                scaled = pixmap.scaled(
                    self.chat_scroll_area.viewport().size(),
                    Qt.KeepAspectRatioByExpanding,
                    Qt.SmoothTransformation
                )
                brush = QBrush(scaled)
                palette = self.chat_scroll_area.viewport().palette()
                palette.setBrush(QPalette.Base, brush)
                self.chat_scroll_area.viewport().setPalette(palette)
                self.chat_scroll_area.viewport().setAutoFillBackground(True)
                self.chat_scroll_area.setPalette(palette)
                self.chat_messages_widget.setStyleSheet("background: transparent;")
        elif color:
            palette = self.chat_scroll_area.viewport().palette()
            palette.setBrush(QPalette.Base, QBrush(QColor(color)))
            self.chat_scroll_area.viewport().setPalette(palette)
            self.chat_scroll_area.viewport().setAutoFillBackground(True)
            self.chat_scroll_area.setPalette(palette)
            self.chat_messages_widget.setStyleSheet("background: transparent;")
        else:
            if self.settings.get("theme", "light") == "dark":
                bg = "#2b2b2b"
            else:
                bg = "#ffffff"
            palette = self.chat_scroll_area.viewport().palette()
            palette.setBrush(QPalette.Base, QBrush(QColor(bg)))
            self.chat_scroll_area.viewport().setPalette(palette)
            self.chat_scroll_area.viewport().setAutoFillBackground(True)
            self.chat_scroll_area.setPalette(palette)
            self.chat_messages_widget.setStyleSheet("background: transparent;")

        # Принудительно обновляем всё после установки фона
        self.chat_scroll_area.viewport().update()
        self.chat_messages_widget.update()

    def open_logs_folder(self):
        """Открывает папку с логами в проводнике."""

        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(log_dir))

    def show_prompts_dialog(self):
        """Открывает диалог со списком сохранённых промптов."""

        dialog = PromptManagerDialog(self.prompts, self)
        if dialog.exec() == QDialog.Accepted:
            # После закрытия диалога обновляем список
            self.prompts = dialog.get_updated_prompts()
            self.save_prompts()

    def save_prompt(self, text):
        """Сохраняет текст как именованный промпт."""

        name, ok = QInputDialog.getText(self, "Сохранить промпт", "Введите название промпта:")
        if ok and name.strip():
            self.prompts.append({"name": name.strip(), "text": text})
            self.save_prompts()
            self.statusBar().showMessage("Промпт сохранён")

    def insert_prompt_text(self, text):
        """Вставляет текст промпта в поле ввода."""

        self.input_edit.insertPlainText(text)
        self.input_edit.setFocus()

    def show_statistics_dialog(self):
        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        dialog = StatisticsDialog(chat, self)
        dialog.exec()

    def show_input_context_menu(self, pos):
        """Контекстное меню для поля ввода: редактирование и сохранение промпта."""

        menu = QMenu(self)

        # Проверяем, есть ли выделенный текст
        has_selection = self.input_edit.textCursor().hasSelection()

        # Копировать
        copy_action = menu.addAction("Копировать")
        copy_action.setEnabled(has_selection)
        copy_action.triggered.connect(self.input_edit.copy)

        # Вырезать
        cut_action = menu.addAction("Вырезать")
        cut_action.setEnabled(has_selection)
        cut_action.triggered.connect(self.input_edit.cut)

        # Вставить
        paste_action = menu.addAction("Вставить")
        paste_action.triggered.connect(self.paste_plain_text)

        # Выделить всё
        select_all_action = menu.addAction("Выделить всё")
        select_all_action.triggered.connect(self.input_edit.selectAll)

        menu.addSeparator()

        # Сохранить промпт (только если есть текст)
        text = self.input_edit.toPlainText().strip()
        if text:
            save_action = menu.addAction("Сохранить промпт")
            save_action.triggered.connect(lambda: self.save_prompt(text))

        # Показываем меню
        menu.exec(self.input_edit.mapToGlobal(pos))

    def show_chat_settings_menu(self):
        menu = QMenu(self)
        dark = self.settings.get("theme", "light") == "dark"

        if dark:
            search_action = menu.addAction(self.icon_search_white, "Поиск")
            wallpaper_action = menu.addAction(self.icon_wallpaper_white, "Обои")
            stats_action = menu.addAction(self.icon_statistics_white, "Статистика")
            prompts_action = menu.addAction(self.icon_save_promt_white, "Промпты")
            clear_action = menu.addAction(self.icon_clear_history_white, "Очистить историю чата")
            retry_action = menu.addAction(self.icon_repeat_request_white, "Повторить последний запрос")
            summary_action = menu.addAction(self.icon_summary_dialogue_white, "Сделать резюме диалога")
            params_action = menu.addAction(self.icon_chat_parameters_white, "Параметры генерации")
            commands_menu = menu.addMenu(self.icon_commands_white, "Команды")
        else:
            search_action = menu.addAction(self.icon_search, "Поиск")
            wallpaper_action = menu.addAction(self.icon_wallpaper, "Обои")
            stats_action = menu.addAction(self.icon_statistics, "Статистика")
            prompts_action = menu.addAction(self.icon_promt, "Промпты")
            clear_action = menu.addAction(self.icon_clear_history, "Очистить историю чата")
            retry_action = menu.addAction(self.icon_repeat_request, "Повторить последний запрос")
            summary_action = menu.addAction(self.icon_summary_dialogue, "Сделать резюме диалога")
            params_action = menu.addAction(self.icon_chat_parameters, "Параметры генерации")
            commands_menu = menu.addMenu(self.icon_commands, "Команды")

        cmd_clear = commands_menu.addAction("/clear — очистить историю")
        cmd_retry = commands_menu.addAction("/retry — повторить последний запрос")
        cmd_summary = commands_menu.addAction("/summary — сделать резюме")

        action = menu.exec(self.btn_chat_settings.mapToGlobal(QPoint(0, self.btn_chat_settings.height())))

        if action == search_action:
            self.show_search_panel()
        elif action == wallpaper_action:
            self.open_chat_appearance_dialog()
        elif action == stats_action:
            self.show_statistics_dialog()
        elif action == clear_action:
            self.clear_chat_history()
        elif action == retry_action:
            self.retry_last_request()
        elif action == summary_action:
            self.summarize_dialog()
        elif action == cmd_clear:
            self.input_edit.insertPlainText("/clear")
            self.input_edit.setFocus()
        elif action == cmd_retry:
            self.input_edit.insertPlainText("/retry")
            self.input_edit.setFocus()
        elif action == cmd_summary:
            self.input_edit.insertPlainText("/summary")
            self.input_edit.setFocus()
        elif action == prompts_action:
            self.show_prompts_dialog()
        elif action == params_action:
            self.open_chat_params_dialog()

    def summarize_dialog(self):
        """Добавляет сообщение с просьбой сделать резюме и отправляет запрос."""

        if self.current_chat_index < 0:
            return
        user_text = "Пожалуйста, сделай краткое резюме нашего диалога."
        self.add_message_to_current("user", user_text)
        self.append_user_message(user_text, time.strftime("%d.%m.%Y %H:%M:%S"))
        self.input_edit.clear()
        self._send_request_after_edit()

    def retry_last_request(self):
        """Повторяет последний запрос, создавая новый ответ, не удаляя предыдущий."""

        if self.current_chat_index < 0:
            return

        # Проверка занятости — до добавления сообщения
        if self.worker and self.worker.isRunning():
            QMessageBox.information(
                self, "Занято",
                "Дождитесь завершения текущего запроса или остановите генерацию."
            )
            return

        chat = self.chats[self.current_chat_index]

        last_user_idx = -1
        for i in range(len(chat['messages']) - 1, -1, -1):
            if chat['messages'][i].get('role') == 'user':
                last_user_idx = i
                break

        if last_user_idx == -1:
            QMessageBox.information(self, "Информация", "Нет сообщений пользователя для повторения.")
            return

        last_user_msg = chat['messages'][last_user_idx]
        new_user_msg = {
            "role": "user",
            "content": last_user_msg.get("content", ""),
            "time": time.strftime("%d.%m.%Y %H:%M:%S"),
        }
        if "attachments" in last_user_msg:
            new_user_msg["attachments"] = last_user_msg["attachments"]

        chat['messages'].append(new_user_msg)
        self.save_current_chat()
        self._redraw_chat_with_history()
        self.statusBar().showMessage("Повторяю последний запрос...")

        QTimer.singleShot(0, self._send_request_after_edit)
        QTimer.singleShot(100, self._do_scroll_to_bottom)

    def clear_chat_history(self):
        """Полностью очищает историю текущего чата (оставляя системный промпт, если он есть)."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        sp = self.settings.get("system_prompt", "").strip()
        chat['messages'] = []
        if sp:
            chat['messages'].append({"role": "system", "content": sp, "time": time.strftime("%d.%m.%Y %H:%M:%S")})
        self.save_current_chat()
        self._redraw_chat_with_history()
        self.statusBar().showMessage("История чата очищена")

    # ======================================================================
    #  БЕЗОПАСНОСТЬ И ШИФРОВАНИЕ
    # ======================================================================

    def _derive_key_from_password(self, password: str, salt: str, iterations: int) -> bytes:
        """Генерирует Fernet key из мастер-пароля с помощью PBKDF2-HMAC-SHA256."""

        import base64
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt.encode('utf-8'),
            iterations=iterations,
        )
        key = base64.urlsafe_b64encode(kdf.derive(password.encode('utf-8')))
        return key

    def _get_fernet(self):
        """Возвращает Fernet-объект, используя мастер-ключ из памяти."""

        if not CRYPTO_AVAILABLE:
            raise RuntimeError("Библиотека cryptography не установлена.")
        if hasattr(self, '_fernet') and self._fernet is not None:
            return self._fernet

        # Ключ ещё не загружен — запросим мастер-пароль
        salt = self.settings.get("encryption_salt", "")
        if not salt:
            raise RuntimeError("Соль не задана. Сначала установите мастер-пароль.")
        iterations = self.settings.get("encryption_kdf_iterations", 200000)

        password, ok = QInputDialog.getText(
            self, "Мастер-пароль", "Введите мастер-пароль для шифрования:",
            QLineEdit.Password
        )
        if not ok or not password:
            raise RuntimeError("Мастер-пароль не введён.")

        key = self._derive_key_from_password(password, salt, iterations)
        self._fernet = Fernet(key)
        return self._fernet

    def _encrypt_data(self, data: dict, fernet=None) -> bytes:
        if fernet is None:
            fernet = self._get_fernet()
        json_str = json.dumps(data, ensure_ascii=False)
        encrypted = fernet.encrypt(json_str.encode('utf-8'))
        return b"ENC:" + encrypted

    def _encrypt_bytes(self, data: bytes, fernet=None) -> bytes:
        if fernet is None:
            fernet = self._get_fernet()
        return b"ENC:" + fernet.encrypt(data)

    def _decrypt_bytes(self, data: bytes, fernet=None) -> bytes:
        if not data.startswith(b"ENC:"):
            return data
        pw = getattr(self, '_current_chat_password', None)
        if pw:
            try:
                return self._decrypt_with_password(data, pw)
            except Exception:
                pass
        if fernet is None:
            fernet = self._get_fernet()
        return fernet.decrypt(data[4:])

    def _decrypt_data(self, data_bytes: bytes, fernet=None) -> dict:
        if fernet is None:
            fernet = self._get_fernet()
        decrypted = fernet.decrypt(data_bytes)
        return json.loads(decrypted.decode('utf-8'))

    def setup_master_password(self):
        """Устанавливает или меняет мастер-пароль, перешифровывая зашифрованные чаты при смене."""

        if not CRYPTO_AVAILABLE:
            QMessageBox.warning(self, "Ошибка", "Библиотека cryptography не установлена.")
            return

        current_salt = self.settings.get("encryption_salt", "")
        is_change = bool(current_salt)

        # Если мастер-пароль уже установлен, запрашиваем текущий пароль
        old_fernet = None
        if is_change:
            old_password, ok = QInputDialog.getText(
                self, "Текущий мастер-пароль",
                "Введите текущий мастер-пароль для подтверждения:",
                QLineEdit.Password
            )
            if not ok or not old_password:
                return

            # Проверяем старый пароль
            try:
                old_key = self._derive_key_from_password(old_password, current_salt,
                                                         self.settings.get("encryption_kdf_iterations", 200000))
                old_fernet = Fernet(old_key)
            except Exception:
                QMessageBox.warning(self, "Ошибка", "Неверный текущий пароль.")
                return

            # Дополнительная проверка: попытаемся расшифровать первый зашифрованный файл (если есть)
            if old_fernet:
                for filename in os.listdir(self.chats_dir):
                    if not filename.endswith('.json'):
                        continue
                    filepath = os.path.join(self.chats_dir, filename)
                    try:
                        with open(filepath, 'rb') as f:
                            content = f.read()
                        if content.startswith(b"ENC:"):
                            # Попытка расшифровать
                            self._decrypt_data(content[4:], fernet=old_fernet)
                            break  # успех
                    except Exception:
                        QMessageBox.warning(self, "Ошибка", "Неверный текущий пароль или повреждённые данные.")
                        return

        # Запрашиваем новый пароль
        password, ok = QInputDialog.getText(
            self, "Новый мастер-пароль" if is_change else "Установка мастер-пароля",
            "Введите новый мастер-пароль:",
            QLineEdit.Password
        )
        if not ok or not password:
            return

        confirm, ok = QInputDialog.getText(
            self, "Подтверждение", "Повторите новый пароль:",
            QLineEdit.Password
        )
        if not ok or password != confirm:
            QMessageBox.warning(self, "Ошибка", "Пароли не совпадают.")
            return

        # Генерируем новую соль и новый ключ
        import os as _os
        new_salt = _os.urandom(16).hex()
        new_key = self._derive_key_from_password(password, new_salt,
                                                 self.settings.get("encryption_kdf_iterations", 200000))
        new_fernet = Fernet(new_key)

        # Если это смена пароля, перешифровываем все зашифрованные чаты новым ключом
        if is_change and old_fernet:
            encrypted_files = []
            for filename in os.listdir(self.chats_dir):
                if not filename.endswith('.json'):
                    continue
                filepath = os.path.join(self.chats_dir, filename)
                try:
                    with open(filepath, 'rb') as f:
                        content = f.read()
                    if content.startswith(b"ENC:"):
                        encrypted_files.append(filepath)
                except Exception:
                    continue

            if encrypted_files:
                progress = QMessageBox(self)
                progress.setWindowTitle("Перешифрование")
                progress.setText(f"Перешифрование {len(encrypted_files)} файлов...")
                progress.setStandardButtons(QMessageBox.NoButton)
                progress.show()
                QApplication.processEvents()

                success = 0
                for filepath in encrypted_files:
                    try:
                        with open(filepath, 'rb') as f:
                            content = f.read()
                        # Расшифровываем старым ключом
                        data = self._decrypt_data(content[4:], fernet=old_fernet)
                        # Шифруем новым ключом
                        new_encrypted = self._encrypt_data(data, fernet=new_fernet)
                        with open(filepath, 'wb') as f:
                            f.write(new_encrypted)
                        success += 1
                    except Exception as e:
                        self.logger.error(f"Ошибка перешифрования {os.path.basename(filepath)}: {e}")
                        QMessageBox.warning(self, "Ошибка",
                                            f"Не удалось перешифровать {os.path.basename(filepath)}.\n"
                                            f"Файл остался зашифрован старым ключом.")
                        # В случае частичной неудачи лучше прервать и не сохранять новую соль
                        progress.close()
                        return

                progress.close()
                self.show_info("Готово", f"Успешно перешифровано файлов: {success} из {len(encrypted_files)}")
            else:
                self.show_info("Информация", "Зашифрованных чатов не найдено. Ключ обновлён.")

        # Сохраняем новую соль
        self.settings["encryption_salt"] = new_salt
        self.save_settings()
        self._fernet = new_fernet

        QMessageBox.information(self, "Готово", "Мастер-пароль обновлён.")

    def set_chat_password(self, filepath: str):
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return

        if self.chats[idx].get('password_hash'):
            QMessageBox.information(self, "Пароль уже установлен", "Используйте «Изменить пароль» для смены.")
            return

        password, ok = QInputDialog.getText(
            self, "Установить пароль", "Введите пароль:", QLineEdit.Password
        )
        if not ok or not password.strip():
            return

        confirm, ok2 = QInputDialog.getText(
            self, "Подтверждение", "Повторите пароль:", QLineEdit.Password
        )
        if not ok2 or password != confirm:
            QMessageBox.warning(self, "Ошибка", "Пароли не совпадают. Пароль не установлен.")
            return

        self.chats[idx]['password_hash'] = self._hash_password(password)
        self.save_chat_by_filepath(filepath)
        self.statusBar().showMessage("Пароль установлен")

    def change_chat_password(self, filepath: str):
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        chat = self.chats[idx]
        if not chat.get('encrypted'):
            QMessageBox.information(self, "Информация", "Чат не зашифрован.")
            return

        old_password, ok = QInputDialog.getText(
            self, "Текущий пароль", "Введите текущий пароль:", QLineEdit.Password)
        if not ok or not old_password:
            return
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
            json.loads(self._decrypt_with_password(content, old_password).decode('utf-8'))
        except Exception:
            QMessageBox.warning(self, "Ошибка", "Неверный пароль.")
            return

        new_password, ok = QInputDialog.getText(
            self, "Новый пароль", "Введите новый пароль:", QLineEdit.Password)
        if not ok or not new_password:
            return
        confirm, ok = QInputDialog.getText(
            self, "Подтверждение", "Повторите новый пароль:", QLineEdit.Password)
        if not ok or new_password != confirm:
            QMessageBox.warning(self, "Ошибка", "Пароли не совпадают.")
            return

        # Файл чата
        with open(filepath, 'rb') as f:
            content = f.read()
        decrypted = self._decrypt_with_password(content, old_password)
        with open(filepath, 'wb') as f:
            f.write(self._encrypt_with_password(decrypted, new_password))

        # Вложения этого чата
        chat_id = os.path.splitext(os.path.basename(filepath))[0]
        attachments_dir = os.path.join("attachments", chat_id)
        count = 0
        if os.path.isdir(attachments_dir):
            for name in os.listdir(attachments_dir):
                path = os.path.join(attachments_dir, name)
                if not os.path.isfile(path):
                    continue
                with open(path, 'rb') as f:
                    content = f.read()
                if not content.startswith(b"ENC:"):
                    continue
                try:
                    decrypted = self._decrypt_with_password(content, old_password)
                    with open(path, 'wb') as f:
                        f.write(self._encrypt_with_password(decrypted, new_password))
                    count += 1
                except Exception as e:
                    self.logger.warning(f"Не перешифрован {path}: {e}")

        # Запоминаем пароль для этой сессии
        self._current_chat_password = new_password
        QMessageBox.information(self, "Готово", f"Пароль изменён. Вложений: {count}")

    def reset_chat_password(self, filepath: str):
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return

        reply = QMessageBox.question(
            self, "Сброс пароля", "Вы уверены, что хотите удалить пароль с этого чата?",
            QMessageBox.Yes | QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        self.chats[idx]['password_hash'] = None
        self.save_chat_by_filepath(filepath)
        self.statusBar().showMessage("Пароль сброшен")

    def toggle_chat_encryption(self, filepath: str):
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        chat = self.chats[idx]
        chat_id = os.path.splitext(os.path.basename(filepath))[0]
        attachments_dir = os.path.join("attachments", chat_id)

        if not chat.get('encrypted', False):
            # === ШИФРУЕМ ЧАТ ===
            password, ok = QInputDialog.getText(
                self, "Пароль", "Введите пароль для шифрования чата:",
                QLineEdit.Password
            )
            if not ok or not password:
                return

            # Шифруем вложения
            count = 0
            if os.path.isdir(attachments_dir):
                for name in os.listdir(attachments_dir):
                    path = os.path.join(attachments_dir, name)
                    if not os.path.isfile(path):
                        continue
                    with open(path, 'rb') as f:
                        data = f.read()
                    if data.startswith(b"ENC:"):
                        continue
                    with open(path, 'wb') as f:
                        f.write(self._encrypt_with_password(data, password))
                    count += 1

            # Шифруем файл чата
            chat_data = {
                "title": chat.get('title', ''),
                "messages": chat.get('messages', []),
                "folder": chat.get('folder', ''),
                "background_color": chat.get('background_color'),
                "background_image": chat.get('background_image'),
                "background_image_original": chat.get('background_image_original'),
                "model": chat.get('model', self.settings.get('model', 'local-model')),
                "created_time": chat.get('created_time'),
                "last_activity_time": chat.get('last_activity_time'),
                "total_tokens_sent": chat.get('total_tokens_sent', 0),
                "total_generation_time": chat.get('total_generation_time', 0.0),
                "color": chat.get('color', '#FFFFFF'),
                "params": chat.get('params', {}),
                "system_prompt": chat.get('system_prompt', ''),
                "pinned": chat.get('pinned', False),
                "tags": chat.get('tags', []),
                "archived": chat.get('archived', False),
            }
            with open(filepath, 'wb') as f:
                f.write(self._encrypt_with_password(
                    json.dumps(chat_data, ensure_ascii=False).encode('utf-8'),
                    password
                ))

            chat['encrypted'] = True
            chat['locked'] = False
            self._current_chat_password = password
            self.statusBar().showMessage(f"Чат зашифрован, вложений: {count}")

        else:
            # === РАСШИФРОВЫВАЕМ ЧАТ ===
            password, ok = QInputDialog.getText(
                self, "Пароль", "Введите пароль для расшифровки чата:",
                QLineEdit.Password
            )
            if not ok or not password:
                return

            # Проверяем пароль на файле чата
            try:
                with open(filepath, 'rb') as f:
                    content = f.read()
                decrypted = self._decrypt_with_password(content, password)
                chat_data = json.loads(decrypted.decode('utf-8'))
            except Exception:
                QMessageBox.warning(self, "Ошибка", "Неверный пароль.")
                return

            # Расшифровываем вложения этим же паролем
            count = 0
            if os.path.isdir(attachments_dir):
                for name in os.listdir(attachments_dir):
                    path = os.path.join(attachments_dir, name)
                    if not os.path.isfile(path):
                        continue
                    with open(path, 'rb') as f:
                        data = f.read()
                    if not data.startswith(b"ENC:"):
                        continue
                    try:
                        dec = self._decrypt_with_password(data, password)
                        with open(path, 'wb') as f:
                            f.write(dec)
                        count += 1
                    except Exception as e:
                        self.logger.warning(f"Не расшифрован {path}: {e}")

            # Перезаписываем файл чата в открытом виде
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(chat_data, f, ensure_ascii=False, indent=2)

            chat['encrypted'] = False
            chat['locked'] = False
            self._current_chat_password = None
            self.statusBar().showMessage(f"Чат расшифрован, вложений: {count}")

    def decrypt_all_encrypted_chats(self):
        """Пересохраняет все зашифрованные чаты в незашифрованном виде."""

        if not CRYPTO_AVAILABLE:
            self.show_warning("Ошибка", "Библиотека cryptography не установлена.")
            return

        # Проверяем, есть ли зашифрованные файлы
        encrypted_files = []
        for filename in os.listdir(self.chats_dir):
            if not filename.endswith('.json'):
                continue
            filepath = os.path.join(self.chats_dir, filename)
            try:
                with open(filepath, 'rb') as f:
                    content = f.read()
                if content.startswith(b"ENC:"):
                    encrypted_files.append(filepath)
            except Exception:
                continue

        if not encrypted_files:
            self.show_info("Информация", "Зашифрованных чатов не найдено.")
            return

        # Запрашиваем мастер-пароль один раз
        try:
            fernet = self._get_fernet()
        except RuntimeError as e:
            self.show_warning("Ошибка", str(e))
            return

        # Расшифровываем и пересохраняем каждый файл
        success = 0
        for filepath in encrypted_files:
            try:
                with open(filepath, 'rb') as f:
                    content = f.read()
                data = self._decrypt_data(content[4:])  # используем уже полученный fernet
                # Сохраняем без шифрования
                with open(filepath, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                success += 1
            except Exception as e:
                self.logger.error(f"Ошибка расшифровки {os.path.basename(filepath)}: {e}")

        self.show_info("Готово", f"Расшифровано чатов: {success} из {len(encrypted_files)}")

        # Перезагружаем чаты
        # Полная очистка текущих данных
        self.chats.clear()
        self.archived_chats.clear()
        self.chat_tree.clear()
        self.chat_items.clear()
        self.folder_items.clear()
        self.loaded_count.clear()
        self.message_widgets.clear()
        self.current_chat_index = -1

        # Загружаем чаты заново из уже расшифрованных файлов
        self.load_chats_from_disk()

        # Принудительно перестраиваем дерево из текущего списка чатов
        self.refresh_tree_from_data()

        # Применяем закрепления и сохраняем порядок
        self.reorder_tree_by_pinned()
        self.save_tree_order()

        # Выбираем первый чат, если есть
        if self.chats:
            self.select_chat_by_filepath(self.chats[0]['filepath'])
        else:
            self.show_empty_chat_placeholder()

    # ======================================================================
    #  УВЕДОМЛЕНИЯ И ЗВУК
    # ======================================================================

    def play_sound(self):
        sound_file = self.settings.get("sound_file", "Sound_1.MP3")
        candidates = [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), sound_file),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "Sound", sound_file),
            sound_file
        ]
        path = next((p for p in candidates if os.path.exists(p)), None)
        if not path:
            print("Звук не найден")
            return
        try:
            from PySide6.QtMultimedia import QMediaPlayer, QAudioOutput
            self._media_player = QMediaPlayer()
            self._audio_output = QAudioOutput()
            self._media_player.setAudioOutput(self._audio_output)
            self._media_player.setSource(QUrl.fromLocalFile(path))
            self._audio_output.setVolume(0.8)
            self._media_player.play()
        except Exception as e:
            print(f"Не удалось воспроизвести звук: {e}")

    def show_toast(self, title, message):
        # Закрываем предыдущий toast и останавливаем его анимации
        if hasattr(self, '_toast') and self._toast is not None:
            try:
                if hasattr(self, '_fade_in_anim') and self._fade_in_anim is not None:
                    self._fade_in_anim.stop()
                if hasattr(self, '_fade_out_anim') and self._fade_out_anim is not None:
                    self._fade_out_anim.stop()
                self._toast.deleteLater()
            except RuntimeError:
                pass

        # Создаём дочерний виджет внутри области чата (не окно ОС)
        toast = QFrame(self.chat_scroll_area)
        toast.setStyleSheet("""
            QFrame {
                background-color: rgba(40, 40, 40, 220);
                border-radius: 10px;
            }
            QLabel {
                color: white;
                font-size: 12px;
                background: transparent;
            }
        """)

        layout = QVBoxLayout(toast)
        layout.setContentsMargins(10, 10, 10, 10)
        title_label = QLabel(title)
        title_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(title_label)
        msg_label = QLabel(message)
        msg_label.setWordWrap(True)
        layout.addWidget(msg_label)
        toast.adjustSize()

        # Позиционируем в правом нижнем углу области чата
        area_rect = self.chat_scroll_area.rect()
        x = area_rect.right() - toast.width() - 20
        y = area_rect.bottom() - toast.height() - 20
        toast.move(x, y)

        # Настройка прозрачности и анимации
        effect = QGraphicsOpacityEffect(toast)
        toast.setGraphicsEffect(effect)
        effect.setOpacity(0.0)

        self._fade_in_anim = QPropertyAnimation(effect, b"opacity")
        self._fade_in_anim.setDuration(300)
        self._fade_in_anim.setStartValue(0.0)
        self._fade_in_anim.setEndValue(1.0)
        self._fade_in_anim.start()

        toast.show()
        toast.raise_()
        self._toast = toast

        def start_fade_out():
            self._fade_out_anim = QPropertyAnimation(effect, b"opacity")
            self._fade_out_anim.setDuration(300)
            self._fade_out_anim.setStartValue(1.0)
            self._fade_out_anim.setEndValue(0.0)
            self._fade_out_anim.finished.connect(toast.deleteLater)
            self._fade_out_anim.start()

        QTimer.singleShot(5000, start_fade_out)

    def notify_message_received(self, filepath):
        if self.settings.get("notifications_enabled", True):
            self.unread_chats.add(filepath)
            self.update_unread_indicators()
            self.refresh_unread_ui()
            QApplication.alert(self, 3000)  # мигание окна в панели задач

    def update_unread_indicators(self):
        # Если уведомления отключены, очищаем все индикаторы и выходим
        if not self.settings.get("notifications_enabled", True):
            self.unread_chats.clear()
            for filepath, item in self.chat_items.items():
                chat = next((c for c in self.chats if c['filepath'] == filepath), None)
                if chat:
                    self.update_chat_item_text(chat)
            for folder_name, folder_item in self.folder_items.items():
                folder_item.setText(0, folder_name)
                self.refresh_unread_ui()
            return

        # Обновляем индикаторы у чатов
        for filepath, item in self.chat_items.items():
            chat = next((c for c in self.chats if c['filepath'] == filepath), None)
            if chat:
                self.update_chat_item_text(chat)

        # Обновляем индикаторы у папок
        for folder_name, folder_item in self.folder_items.items():
            has_unread = self._folder_has_unread_chats(folder_item)
            if has_unread:
                if not folder_item.text(0).startswith("🔴 "):
                    folder_item.setText(0, "🔴 " + folder_name)
            else:
                folder_item.setText(0, folder_name)

    # ======================================================================
    #  ПЛАГИНЫ
    # ======================================================================

    def load_plugins(self):
        plugins_dir = "plugins"
        if not os.path.isdir(plugins_dir):
            os.makedirs(plugins_dir, exist_ok=True)
        self.plugins = []
        for filename in os.listdir(plugins_dir):
            if not filename.endswith('.py'):
                continue
            path = os.path.join(plugins_dir, filename)
            try:
                spec = importlib.util.spec_from_file_location(f"plugin_{filename[:-3]}", path)
                module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(module)
                if hasattr(module, 'Plugin'):
                    plugin_instance = module.Plugin()
                    self.plugins.append(plugin_instance)
                    print(f"Плагин загружен: {filename}")
                    # Вызываем хук on_plugin_loaded, если он есть
                    try:
                        if hasattr(plugin_instance, 'on_plugin_loaded'):
                            from types import SimpleNamespace
                            api = PluginAPI(self)
                            plugin_instance.on_plugin_loaded(api)
                    except Exception as e:
                        self.logger.warning(f"Ошибка в on_plugin_loaded плагина {filename}: {e}")
            except Exception as e:
                self.logger.warning(f"Ошибка загрузки плагина {filename}: {e}")

    def reload_plugins(self):
        """Перезагружает плагины без перезапуска программы."""

        self.load_plugins()
        self.statusBar().showMessage(f"Плагины перезагружены. Загружено: {len(self.plugins)}")
        self.logger.info("Плагины перезагружены.")

    def open_plugins_folder(self):
        """Открывает папку с плагинами в проводнике."""

        plugins_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "plugins")
        os.makedirs(plugins_dir, exist_ok=True)
        QDesktopServices.openUrl(QUrl.fromLocalFile(plugins_dir))

    def show_plugins_manager(self):
        """Открывает диалог управления плагинами."""

        dialog = PluginsManagerDialog(self)
        dialog.exec()
        # После закрытия перезагружаем плагины (на случай удаления/изменения)
        self.load_plugins()

    # ======================================================================
    #  СЛУЖЕБНЫЕ МЕТОДЫ
    # ======================================================================

    def eventFilter(self, obj, event):
        # Обработка Drag&Drop для области чата
        if hasattr(self, 'chat_scroll_area') and hasattr(self, 'chat_messages_widget'):
            if obj in (self.chat_scroll_area, self.chat_scroll_area.viewport(), self.chat_messages_widget):
                if event.type() == QEvent.DragEnter:
                    if event.mimeData().hasUrls():
                        event.acceptProposedAction()
                        return True
                elif event.type() == QEvent.DragMove:
                    if event.mimeData().hasUrls():
                        event.acceptProposedAction()
                        return True
                elif event.type() == QEvent.Drop:
                    if event.mimeData().hasUrls():
                        for url in event.mimeData().urls():
                            if url.isLocalFile():
                                path = url.toLocalFile()
                                if os.path.exists(path):
                                    self.attachments.append(path)
                        self.update_attachments_ui()
                        event.acceptProposedAction()
                        return True

        # Обновление фона и ширины сообщений при изменении размера viewport ИЛИ контейнера сообщений
        if hasattr(self, 'chat_scroll_area') and hasattr(self, 'chat_messages_widget'):
            if obj in (self.chat_scroll_area.viewport(), self.chat_messages_widget) and event.type() == QEvent.Resize:
                if self.current_chat_index >= 0:
                    self.apply_chat_background(self.chats[self.current_chat_index])
                else:
                    self.apply_chat_background(None)
                self.update_floating_buttons_positions()

        # Глобальные горячие клавиши
        if event.type() == QEvent.Type.KeyPress:
            key = event.key()
            mods = event.modifiers()
            hotkeys = self.settings.get("hotkeys", {})

            if key == Qt.Key_F1:
                self.show_help_dialog()
                return True

            def matches(action):
                seq_str = hotkeys.get(action, "")
                if not seq_str:
                    return False
                parts = []
                if mods & Qt.ControlModifier:
                    parts.append("Ctrl")
                if mods & Qt.ShiftModifier:
                    parts.append("Shift")
                if mods & Qt.AltModifier:
                    parts.append("Alt")
                if mods & Qt.MetaModifier:
                    parts.append("Meta")
                key_text = QKeySequence(key).toString()
                if key_text:
                    parts.append(key_text)
                actual_str = "+".join(parts)
                return actual_str == seq_str

            if matches("new_folder"):
                self.create_new_folder()
                return True
            elif matches("new_chat"):
                self.create_new_chat()
                return True
            elif matches("close_chat"):
                self.close_current_chat()
                return True
            elif matches("search"):
                self.show_search_panel()
                return True
            elif matches("import_deepseek"):
                self.import_deepseek_history()
                return True
            elif matches("export"):
                self.show_export_menu()
                return True
            elif matches("settings"):
                self.open_settings_dialog()
                return True
            elif matches("send_message"):
                self.send_message()
                return True
            elif matches("toggle_theme"):
                self.toggle_theme()
                return True
            elif matches("toggle_sidebar"):
                self.toggle_sidebar()
                return True
            elif matches("clear_chat"):
                self.clear_chat_history()
                return True
            elif matches("starred_messages"):
                self.show_starred_messages()
                return True
            elif matches("statistics"):
                self.show_statistics_dialog()
                return True
            elif matches("prompts"):
                self.show_prompts_dialog()
                return True
            elif matches("focus_input"):
                self.input_edit.setFocus()
                return True
            elif matches("next_chat"):
                if self.chats:
                    new_idx = (self.current_chat_index + 1) % len(self.chats)
                    self.select_chat_by_filepath(self.chats[new_idx]['filepath'])
                return True
            elif matches("prev_chat"):
                if self.chats:
                    new_idx = (self.current_chat_index - 1) % len(self.chats)
                    self.select_chat_by_filepath(self.chats[new_idx]['filepath'])
                return True
            elif matches("chat_settings"):
                self.open_chat_params_dialog()
                return True

        return super().eventFilter(obj, event)

    def closeEvent(self, event):
        try:
            if self.worker and self.worker.isRunning():
                self.worker.stop()
                self.worker.wait(1000)
        except Exception:
            pass

        # Останавливаем звук, если он играет
        if hasattr(self, '_media_player') and self._media_player:
            try:
                self._media_player.stop()
            except Exception:
                pass

        try:
            # Сохраняем позицию скролла текущего чата
            if 0 <= self.current_chat_index < len(self.chats):
                filepath = self.chats[self.current_chat_index]['filepath']
                self.chat_scroll_positions[filepath] = self.chat_scroll_area.verticalScrollBar().value()
            self.save_scroll_positions()

            self.save_tree_order()
            self.save_last_chat_filepath()

            # Сохраняем обычные чаты, удаляем файлы инкогнито-чатов
            for chat in self.chats:
                if chat.get('incognito', False):
                    fp = chat['filepath']
                    if os.path.exists(fp):
                        try:
                            os.remove(fp)
                        except Exception:
                            pass
                else:
                    self.save_chat_by_filepath(chat['filepath'])

            # Архивные чаты тоже обрабатываем
            for chat in self.archived_chats:
                if chat.get('incognito', False):
                    fp = chat['filepath']
                    if os.path.exists(fp):
                        try:
                            os.remove(fp)
                        except Exception:
                            pass
                else:
                    self.save_chat_by_filepath(chat['filepath'])

            self.create_backup()
            self.save_draft(self.input_edit.toPlainText())
        except Exception as e:
            print(f"Ошибка при закрытии: {e}")
        self.cleanup_temp_files()
        event.accept()

    def resizeEvent(self, event):
        self._update_max_width()
        super().resizeEvent(event)
        if hasattr(self, 'chat_scroll_area') and self.current_chat_index >= 0:
            self.apply_chat_background(self.chats[self.current_chat_index])
        self.update_floating_buttons_positions()
        self.update_search_panel_position()

        # === ОБНОВЛЯЕМ ОВЕРЛЕЙ ПРИ ИЗМЕНЕНИИ РАЗМЕРА ===
        if hasattr(self, 'busy_overlay') and self.busy_overlay.isVisible():
            self.busy_overlay.setGeometry(self.rect())
        # ===============================================

        if hasattr(self, 'message_widgets'):
            viewport_width = self.chat_scroll_area.viewport().width()
            for fp, idx, widget in self.message_widgets:
                widget.setMaximumWidth(int(viewport_width * 0.8))

    def create_backup(self):
        """Создаёт резервную копию в фоновом потоке (если данные изменились)."""

        if not self.settings.get("auto_backup", True):
            return

        # Если копирование уже выполняется, не запускаем повторно
        if getattr(self, '_backup_in_progress', False):
            return

        self._backup_in_progress = True

        def do_backup():
            try:
                current_hash = self.compute_state_hash()
                if current_hash == self.last_backup_hash:
                    print("Нет изменений – резервная копия не требуется.")
                    return

                backup_root = "backups"
                os.makedirs(backup_root, exist_ok=True)

                timestamp = time.strftime("%Y%m%d_%H%M%S")
                backup_dir = os.path.join(backup_root, f"backup_{timestamp}")
                os.makedirs(backup_dir, exist_ok=True)

                items_to_copy = [
                    self.chats_dir,
                    self.config_file,
                    self.folders_file,
                    "tree_order.json",
                    "trash.json",
                    "prompts.json",
                    "last_session.json",
                    "attachments"
                ]

                for item in items_to_copy:
                    if os.path.isdir(item):
                        if os.path.exists(item):
                            shutil.copytree(item, os.path.join(backup_dir, os.path.basename(item)))
                    elif os.path.isfile(item):
                        if os.path.exists(item):
                            shutil.copy2(item, os.path.join(backup_dir, os.path.basename(item)))

                try:
                    with open(os.path.join(backup_dir, "backup_info.txt"), 'w', encoding='utf-8') as f:
                        f.write(f"Резервная копия создана: {time.strftime('%d.%m.%Y %H:%M:%S')}\n")
                        f.write(f"Количество чатов: {len(self.chats)}\n")
                except Exception:
                    pass

                max_backups = self.settings.get("max_backups", 5)
                all_backups = sorted(
                    [os.path.join(backup_root, d) for d in os.listdir(backup_root) if d.startswith("backup_")])
                if len(all_backups) > max_backups:
                    for old in all_backups[:len(all_backups) - max_backups]:
                        try:
                            shutil.rmtree(old)
                        except Exception:
                            pass

                self.last_backup_hash = current_hash
                print(f"Резервная копия сохранена в {backup_dir}")
            except Exception as e:
                self.logger.error(f"Ошибка резервного копирования: {e}")
            finally:
                self._backup_in_progress = False

        # Запускаем в отдельном потоке, чтобы не блокировать интерфейс
        threading.Thread(target=do_backup, daemon=True).start()

    def compute_state_hash(self):
        """Возвращает строку-хэш, отражающую текущее состояние чатов и настроек."""

        try:
            # Сериализуем основное состояние в JSON
            data = {
                "chats": self.chats,
                "settings": self.settings,
                "folders": self.folders,
                "trash_items": self.trash_items,
                "prompts": self.prompts,
            }
            json_str = json.dumps(data, ensure_ascii=False, sort_keys=True, default=str)
            return hashlib.sha256(json_str.encode('utf-8')).hexdigest()
        except Exception:
            # Если не удалось вычислить, возвращаем случайное значение, чтобы бэкап создался
            import random
            self.logger.warning("Не удалось вычислить хэш состояния, использую случайное значение")
            return str(random.random())

    def cleanup_temp_files(self):
        """Удаляет временные файлы, созданные программой."""

        # Удаляем временные обои
        temp_dir = tempfile.gettempdir()
        for f in os.listdir(temp_dir):
            if f.startswith("wallpaper_") and f.endswith(".png"):
                try:
                    os.remove(os.path.join(temp_dir, f))
                except Exception:
                    pass

        # Удаляем миниатюры изображений (если решите не хранить их)
        attachments_root = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachments")
        for root, dirs, files in os.walk(attachments_root):
            for file in files:
                if file.endswith("_thumb.jpg"):
                    try:
                        os.remove(os.path.join(root, file))
                    except Exception:
                        pass

    def launch_lm_studio(self):
        path = self.settings.get("lm_studio_path", "").strip()
        if not path or not os.path.exists(path):
            print("LM Studio путь не указан или файл не найден.")
            return
        try:
            subprocess.Popen([path])
            self.statusBar().showMessage("LM Studio запущен")
        except Exception as e:
            self.logger.warning(f"Не удалось запустить LM Studio: {e}")

    def update_token_count(self):
        """Обновляет метку примерного количества токенов в поле ввода."""

        text = self.input_edit.toPlainText()
        token_estimate = len(text) / 4
        self.token_count_label.setText(f"{token_estimate:.0f} токенов")

    def update_generation_speed(self):
        """Обновляет метку скорости и количества токенов."""

        if not hasattr(self, 'generation_start_time'):
            return
        elapsed = time.time() - self.generation_start_time
        if elapsed <= 0:
            return
        text = getattr(self, 'current_response_text', '')
        tokens_estimate = len(text) / 4
        speed = tokens_estimate / elapsed
        self.speed_label.setText(
            f"Скорость: {speed:.1f} т/с  |  {int(tokens_estimate)} токенов"
        )

    def update_model_label(self):
        if self.current_chat_index >= 0:
            chat = self.chats[self.current_chat_index]
            model = chat.get('model', self.settings.get('model', 'local-model'))
            self.model_label.setText(f"Модель: {model}")
        else:
            self.model_label.setText("")

    def show_error(self, title, message, log_message=None):
        """Показывает критическую ошибку и пишет в лог."""

        if log_message:
            self.logger.error(log_message)
        QMessageBox.critical(self, title, message)

    def show_warning(self, title, message, log_message=None):
        """Показывает предупреждение и пишет в лог."""

        if log_message:
            self.logger.warning(log_message)
        QMessageBox.warning(self, title, message)

    def show_info(self, title, message):
        """Показывает информационное сообщение."""

        QMessageBox.information(self, title, message)

    def check_for_updates(self, silent=True):
        """Проверяет наличие новой версии через GitHub Releases.
        silent=True — показывает диалог только если есть обновление.
        silent=False — всегда показывает результат (для кнопки в настройках)."""

        import requests

        try:
            resp = requests.get(
                GITHUB_API_URL,
                timeout=5,
                headers={"Accept": "application/vnd.github+json"}
            )
        except requests.exceptions.RequestException as e:
            self.logger.warning(f"Ошибка проверки обновлений: {e}")
            if not silent:
                QMessageBox.warning(
                    self, "Проверка обновлений",
                    "Не удалось подключиться к GitHub.\n"
                    "Проверьте интернет-соединение."
                )
            return

        if resp.status_code == 404:
            self.logger.warning("Репозиторий не найден")
            if not silent:
                QMessageBox.warning(
                    self, "Проверка обновлений",
                    "Репозиторий не найден на GitHub.\n"
                    "Проверьте настройку GITHUB_REPO в коде."
                )
            return

        if resp.status_code != 200:
            self.logger.warning(f"GitHub API вернул {resp.status_code}")
            if not silent:
                QMessageBox.warning(
                    self, "Проверка обновлений",
                    f"GitHub вернул ошибку {resp.status_code}."
                )
            return

        try:
            data = resp.json()
        except Exception as e:
            self.logger.warning(f"Не удалось разобрать ответ GitHub: {e}")
            if not silent:
                QMessageBox.warning(self, "Проверка обновлений", "Некорректный ответ от GitHub.")
            return

        latest_tag = data.get("tag_name", "").lstrip("v")
        release_url = data.get("html_url", "")
        release_notes = data.get("body", "").strip()
        release_name = data.get("name", "").strip() or f"Версия {latest_tag}"
        published_at = data.get("published_at", "")

        # Ищем .exe или .zip в ассетах релиза
        download_url = release_url  # по умолчанию — страница релиза
        for asset in data.get("assets", []):
            name = asset.get("name", "").lower()
            if name.endswith(".exe") or name.endswith(".zip"):
                download_url = asset.get("browser_download_url", release_url)
                break

        # Сравниваем версии
        if latest_tag and latest_tag != APP_VERSION:
            # Есть новая версия — показываем диалог
            self.show_update_dialog(
                current=APP_VERSION,
                latest=latest_tag,
                name=release_name,
                notes=release_notes,
                download_url=download_url,
                release_url=release_url,
                published_at=published_at
            )
        else:
            if not silent:
                QMessageBox.information(
                    self, "Проверка обновлений",
                    f"У вас последняя версия: {APP_VERSION}"
                )

    def show_update_dialog(self, current, latest, name, notes, download_url, release_url, published_at):
        """Кастомный диалог с информацией об обновлении."""

        dialog = QDialog(self)
        dialog.setWindowTitle("Доступно обновление")
        dialog.setModal(True)
        dialog.resize(650, 550)

        layout = QVBoxLayout(dialog)

        # Заголовок
        title = QLabel(f"<b>Доступна новая версия: {latest}</b>")
        title.setStyleSheet("font-size: 16px;")
        layout.addWidget(title)

        info = QLabel(f"Текущая версия: {current}")
        info.setStyleSheet("color: #888;")
        layout.addWidget(info)

        if published_at:
            date_label = QLabel(f"Опубликовано: {published_at[:10]}")
            date_label.setStyleSheet("color: #888;")
            layout.addWidget(date_label)

        layout.addSpacing(10)

        notes_label = QLabel("<b>Что нового:</b>")
        layout.addWidget(notes_label)

        notes_browser = QTextBrowser()
        notes_browser.setReadOnly(True)
        notes_browser.setOpenExternalLinks(True)

        # Цвета под тему
        is_dark = self.settings.get("theme", "light") == "dark"
        bg = "#2b2b2b" if is_dark else "#ffffff"
        text_color = "#dddddd" if is_dark else "#000000"
        link_color = "#8ab4f8" if is_dark else "#1a73e8"
        code_bg = "#1e1e1e" if is_dark else "#f5f5f5"
        border_color = "#444" if is_dark else "#ccc"

        # Палитра для ссылок и текста
        palette = notes_browser.palette()
        palette.setColor(QPalette.Base, QColor(bg))
        palette.setColor(QPalette.Text, QColor(text_color))
        palette.setColor(QPalette.Link, QColor(link_color))
        palette.setColor(QPalette.LinkVisited, QColor(link_color))
        notes_browser.setPalette(palette)

        notes_browser.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {bg};
                color: {text_color};
                border: 1px solid {border_color};
            }}
        """)

        # CSS для code/pre + ссылки через document
        notes_browser.document().setDefaultStyleSheet(
            f"body {{ color: {text_color}; }}"
            f"a {{ color: {link_color}; text-decoration: underline; }}"
            f"code {{ background-color: {code_bg}; padding: 2px 4px; border-radius: 3px; }}"
            f"pre {{ background-color: {code_bg}; padding: 8px; border-radius: 5px; }}"
        )

        if notes and MARKDOWN_AVAILABLE:
            try:
                body_html = markdown.markdown(notes, extensions=['fenced_code', 'tables', 'nl2br'])
                notes_browser.setHtml(body_html)
            except Exception:
                notes_browser.setPlainText(notes)
        elif notes:
            notes_browser.setPlainText(notes)
        else:
            notes_browser.setPlainText("Описание изменений отсутствует.")

        layout.addWidget(notes_browser, stretch=1)

        # Кнопки
        btn_box = QHBoxLayout()
        btn_download = QPushButton("Скачать обновление")
        btn_download.setStyleSheet("""
            QPushButton {
                background-color: #1a73e8;
                color: white;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #1557b0; }
        """)
        btn_download.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(download_url)))

        btn_release = QPushButton("Открыть страницу релиза")
        btn_release.clicked.connect(lambda: QDesktopServices.openUrl(QUrl(release_url)))

        btn_later = QPushButton("Позже")
        btn_later.clicked.connect(dialog.reject)

        btn_box.addWidget(btn_download)
        btn_box.addWidget(btn_release)
        btn_box.addStretch()
        btn_box.addWidget(btn_later)

        layout.addLayout(btn_box)
        dialog.exec()

    def show_about_dialog(self):
        dialog = AboutDialog(self)
        dialog.exec()

    def toggle_sidebar(self):
        left_panel = self.main_splitter.widget(0)

        if left_panel.isVisible():
            # Скрываем панель
            target_width = 0
            left_panel.hide()
            self.main_splitter.setSizes([0, self.main_splitter.width()])
        else:
            # Показываем панель
            left_panel.show()
            target_width = 250  # явная ширина
            # Если сохранён размер из настроек — используем его
            saved = self.settings.get("splitter_sizes")
            if isinstance(saved, list) and len(saved) == 2 and saved[0] > 100:
                target_width = saved[0]
            self.main_splitter.setSizes([target_width, self.main_splitter.width() - target_width])

        # Обновляем фон и кнопки
        QTimer.singleShot(0, self._after_sidebar_animation)

    def _after_sidebar_animation(self):
        if self.current_chat_index >= 0:
            self.apply_chat_background(self.chats[self.current_chat_index])
        else:
            self.apply_chat_background(None)

        self.update_floating_buttons_positions()

        if self.current_chat_index >= 0:
            self.btn_chat_settings.setVisible(True)
            # Показываем кнопку стоп только если генерация идёт в ТЕКУЩЕМ чате
            if (self.is_generating
                    and self.generation_chat_filepath == self.chats[self.current_chat_index]['filepath']):
                self.show_stop_button()
            else:
                self.hide_stop_button()
        else:
            self.hide_stop_button()

    def create_backup_async(self):
        threading.Thread(target=self.create_backup, daemon=True).start()

    def paste_plain_text(self):
        """Вставляет текст из буфера обмена без форматирования."""

        clipboard = QApplication.clipboard()
        mime = clipboard.mimeData()
        if mime.hasText():
            self.input_edit.insertPlainText(mime.text())

    def import_document_to_rag(self):
        if not RAG_AVAILABLE:
            QMessageBox.warning(self, "RAG недоступен",
                                "Установите библиотеки: pip install sentence-transformers faiss-cpu numpy")
            return
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "Выберите документы", "", "Документы (*.pdf *.docx *.txt *.md)"
        )
        if not file_paths:
            return
        for path in file_paths:
            text = self.read_file_content(path)
            if text:
                self.rag_manager.add_document(text)
        self.statusBar().showMessage(f"Добавлено документов в базу знаний: {len(self.rag_manager.documents)}")
        self.update_rag_status()

    def update_rag_status(self):
        count = len(self.rag_manager.documents)
        self.statusBar().showMessage(
            f"Документов в базе: {count}. RAG {'включён' if self.rag_enabled else 'выключен'}")

    def ensure_rag_folder(self):
        """Создаёт папку RAG_folder, если её не существует."""

        if not os.path.exists(self.rag_folder):
            os.makedirs(self.rag_folder, exist_ok=True)
            self.logger.info(f"Создана папка RAG_folder: {self.rag_folder}")

    def load_rag_documents_from_folder(self):
        """Сканирует RAG_folder, извлекает текст и загружает в менеджер RAG."""

        self.ensure_rag_folder()
        self.rag_manager.documents = []  # очищаем старые документы
        self.rag_manager.index = None
        self.rag_manager.embeddings = None

        supported_exts = ('.txt', '.md', '.py', '.json', '.csv', '.html', '.log', '.pdf', '.docx')
        files = [f for f in os.listdir(self.rag_folder)
                 if os.path.isfile(os.path.join(self.rag_folder, f))
                 and f.lower().endswith(supported_exts)]

        for filename in files:
            filepath = os.path.join(self.rag_folder, filename)
            text = self.read_file_content_structured(filepath)
            if text and text.strip():
                self.rag_manager.add_document_smart(text, max_chunk_size=1000)
                self.logger.info(f"Документ добавлен в RAG: {filename}")
            else:
                self.logger.warning(f"Не удалось извлечь текст из {filename}")

        # Пересчитываем индексы (вызывается в add_document, но на всякий случай)
        if self.rag_manager.model and self.rag_manager.documents:
            self.rag_manager._rebuild_index()

    def start_rag_watcher(self):
        """Запускает слежение за изменениями в папке RAG_folder."""

        self.ensure_rag_folder()
        if self.rag_watcher is not None:
            if self.rag_folder in self.rag_watcher.directories():
                return
            self.rag_watcher.addPath(self.rag_folder)
            return

        self.rag_watcher = QFileSystemWatcher([self.rag_folder], self)
        self.rag_watcher.directoryChanged.connect(self.on_rag_folder_changed)

    def on_rag_folder_changed(self, path):
        """Вызывается при изменении содержимого папки RAG_folder."""

        if not self.rag_enabled:
            return
        # Небольшая задержка, чтобы файловая система завершила операции
        QTimer.singleShot(500, self.load_rag_documents_from_folder)
        self.statusBar().showMessage("База знаний обновлена автоматически", 3000)

    def add_rag_documents_via_dialog(self):
        """Диалог для добавления документов в RAG_folder."""

        self.ensure_rag_folder()
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите документы для базы знаний",
            "",
            "Документы (*.pdf *.docx *.txt *.md *.py *.json *.csv *.html *.log);;Все файлы (*)"
        )
        if not file_paths:
            return

        for src in file_paths:
            dst = os.path.join(self.rag_folder, os.path.basename(src))
            if os.path.abspath(src) == os.path.abspath(dst):
                continue  # файл уже в папке
            if os.path.exists(dst):
                # Если файл с таким именем уже есть, добавляем суффикс
                base, ext = os.path.splitext(dst)
                counter = 1
                while os.path.exists(dst):
                    dst = f"{base}_{counter}{ext}"
                    counter += 1
            shutil.copy2(src, dst)
            self.logger.info(f"Документ скопирован в RAG_folder: {dst}")

        self.load_rag_documents_from_folder()
        self.update_rag_status()
        QMessageBox.information(self, "Готово", "Документы добавлены в базу знаний.")

    def remove_rag_document_via_dialog(self):
        """Диалог для удаления документа из RAG_folder."""

        self.ensure_rag_folder()
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите документ для удаления",
            self.rag_folder,
            "Документы (*.pdf *.docx *.txt *.md *.py *.json *.csv *.html *.log);;Все файлы (*)"
        )
        if not file_path:
            return
        # Проверяем, что файл действительно в RAG_folder
        if os.path.dirname(os.path.abspath(file_path)) != os.path.abspath(self.rag_folder):
            QMessageBox.warning(self, "Ошибка", "Можно удалять только файлы из папки RAG_folder.")
            return

        reply = QMessageBox.question(
            self,
            "Подтверждение удаления",
            f"Удалить файл «{os.path.basename(file_path)}» из базы знаний?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        try:
            os.remove(file_path)
            self.logger.info(f"Документ удалён из RAG_folder: {file_path}")
            self.load_rag_documents_from_folder()
            self.update_rag_status()
            QMessageBox.information(self, "Готово", "Файл удалён из базы знаний.")
        except Exception as e:
            self.logger.error(f"Ошибка удаления файла {file_path}: {e}")
            self.show_error("Ошибка", f"Не удалось удалить файл.\nПричина: {e}")

    def open_rag_folder(self):
        """Открывает папку RAG_folder в проводнике."""

        self.ensure_rag_folder()
        QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.abspath(self.rag_folder)))

    def on_scroll_changed(self, value):
        if 0 <= self.current_chat_index < len(self.chats):
            filepath = self.chats[self.current_chat_index]['filepath']
            self.chat_scroll_positions[filepath] = value

    def on_splitter_moved(self, pos, index):
        """Сохраняет текущие размеры панелей в настройки."""

        self.settings["splitter_sizes"] = self.main_splitter.sizes()
        self.save_settings()

    def _folder_has_unread_chats(self, folder_item):
        """Проверяет, есть ли в папке (включая вложенные) непрочитанные чаты."""

        for i in range(folder_item.childCount()):
            child = folder_item.child(i)
            data = child.data(0, Qt.UserRole)
            if not data:
                continue
            if data.get('type') == 'chat':
                if data.get('filepath') in self.unread_chats:
                    return True
            elif data.get('type') == 'folder':
                if self._folder_has_unread_chats(child):
                    return True
        return False

    def on_scroll_changed(self, value):
        if self._skip_scroll_restore:
            return
        if 0 <= self.current_chat_index < len(self.chats):
            filepath = self.chats[self.current_chat_index]['filepath']
            self.chat_cursor_positions[filepath] = self.chat_browser.textCursor().position()
            if hasattr(self, '_save_cursor_timer'):
                self._save_cursor_timer.stop()
            else:
                self._save_cursor_timer = QTimer(self)
                self._save_cursor_timer.setSingleShot(True)
                self._save_cursor_timer.timeout.connect(self.save_cursor_positions)
            self._save_cursor_timer.start(1000)

    def restore_folder_entry(self, folder_entry):
        folder_name = folder_entry['name']
        # Восстанавливаем папку, если её ещё нет
        if folder_name not in self.folder_items:
            folder_item = QTreeWidgetItem([folder_name])
            folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
            folder_item.setData(0, Qt.UserRole, {'type': 'folder', 'name': folder_name})
            self.folder_items[folder_name] = folder_item
            self.chat_tree.addTopLevelItem(folder_item)
            self.folders.append({"name": folder_name, "color": "#FFFFFF", "pinned": False})
            self.save_folders()
            self.set_folder_icon(folder_item, expanded=False)

        # Восстанавливаем каждый чат из вложенного списка
        for chat_data in folder_entry.get('chats', []):
            new_chat = {
                'filepath': chat_data['filepath'],
                'title': chat_data['title'],
                'messages': chat_data['messages'],
                'folder': folder_name,
                'background_color': chat_data.get('background_color'),
                'background_image': chat_data.get('background_image'),
                'background_image_original': chat_data.get('background_image_original'),
                'model': chat_data.get('model', self.settings.get('model', 'local-model')),
                'params': chat_data.get('params', {}),
                'system_prompt': chat_data.get('system_prompt', ''),
                'color': chat_data.get('color', '#FFFFFF'),
                'pinned': chat_data.get('pinned', False),
                'tags': chat_data.get('tags', []),
                'encrypted': chat_data.get('encrypted', False),
                'password_hash': chat_data.get('password_hash'),
                'locked': chat_data.get('locked', False),
                'created_time': chat_data.get('created_time', time.strftime('%d.%m.%Y %H:%M:%S')),
                'last_activity_time': chat_data.get('last_activity_time', time.strftime('%d.%m.%Y %H:%M:%S')),
                'total_tokens_sent': chat_data.get('total_tokens_sent', 0),
                'total_generation_time': chat_data.get('total_generation_time', 0.0)
            }
            self.chats.append(new_chat)
            self.add_chat_to_tree(new_chat)

            # Если чат зашифрован, но пароль не введён — не перезаписываем файл
            if new_chat.get('encrypted', False):
                pw = getattr(self, '_current_chat_password', None)
                if not pw:
                    new_chat['locked'] = True
                    new_chat['messages'] = []
                    self.logger.info(
                        f"Восстановлен зашифрованный чат {chat_data['filepath']} — "
                        f"файл не перезаписан до ввода пароля."
                    )
                    continue

            self.save_chat_by_filepath(chat_data['filepath'])

        # Удаляем запись папки из корзины
        self.trash_items = [e for e in self.trash_items if e != folder_entry]
        self.save_trash()
        self.save_tree_order()

    def permanently_delete_folder_entry(self, folder_entry):
        """Окончательно удаляет все чаты внутри папки и саму запись."""

        for chat_data in folder_entry.get('chats', []):
            filepath = chat_data['filepath']
            try:
                if os.path.exists(filepath):
                    os.remove(filepath)
            except Exception as e:
                print(f"Ошибка удаления файла {filepath}: {e}")
            # Удаляем отдельные записи чатов, если они есть
            self.trash_items = [e for e in self.trash_items if
                                not (e.get('type') == 'chat' and e.get('filepath') == filepath)]

        # Удаляем саму запись папки
        self.trash_items = [e for e in self.trash_items if e != folder_entry]
        self.save_trash()

    def _make_json_safe(self, obj):
        """Рекурсивно преобразует множества в списки для корректной сериализации JSON."""

        if isinstance(obj, set):
            return list(obj)
        elif isinstance(obj, dict):
            return {key: self._make_json_safe(value) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [self._make_json_safe(item) for item in obj]
        else:
            return obj


    def dropEvent(self, event):
        self._clear_highlight()
        dragged_item = self._dragged_item if self._dragged_item else self.currentItem()
        super().dropEvent(event)
        if dragged_item:
            self.main_window.on_tree_structure_changed()  # синхронизирует флаги и перестроит дерево
        self._dragged_item = None
        event.acceptProposedAction()

    def _reset_skip_scroll_flag(self):
        self._skip_scroll_restore = False

    def _reset_auto_scrolling(self):
        self._auto_scrolling = False
        self._skip_scroll_restore = False

    def show_sort_menu(self):
        menu = QMenu(self)
        menu.addAction("По дате создания", lambda: self.sort_chats_by('date'))
        menu.addAction("По названию", lambda: self.sort_chats_by('title'))
        menu.addAction("По количеству сообщений", lambda: self.sort_chats_by('messages'))
        menu.exec(self.btn_sort.mapToGlobal(QPoint(0, self.btn_sort.height())))

    def _chat_sort_key(self, chat, criterion):
        if criterion == 'date':
            date_str = chat.get('created_time', '')
            try:
                return -time.mktime(time.strptime(date_str, '%d.%m.%Y %H:%M:%S'))
            except:
                return 0
        elif criterion == 'title':
            return chat.get('title', '').lower()
        elif criterion == 'messages':
            return -len(chat.get('messages', []))
        else:
            return 0

    def toggle_reasoning(self):
        self.show_reasoning = self.btn_toggle_reasoning.isChecked()
        if self.current_chat_index >= 0:
            self._redraw_chat_with_history()

    def position_reasoning_button(self):
        if hasattr(self, 'btn_toggle_reasoning') and hasattr(self, 'btn_chat_settings'):
            x = self.chat_browser.width() - self.btn_chat_settings.width() - self.btn_toggle_reasoning.width() - 20
            y = 10
            self.btn_toggle_reasoning.move(x, y)

    def search_prev(self):
        query = self.search_edit.text().strip()
        if not query:
            return
        self._search(query, forward=False)

    def change_chat_model(self, item):
        filepath = item.data(0, Qt.UserRole)['filepath']
        idx = self.find_chat_index(filepath)
        if idx < 0:
            return
        current_model = self.chats[idx].get('model', self.settings['model'])
        # Список моделей
        models = ["local-model"]
        model, ok = QInputDialog.getItem(self, "Выбор модели", "Модель:", models, editable=True)
        if ok and model:
            self.chats[idx]['model'] = model
            self.save_chat_by_filepath(filepath)
            self.update_model_label()  # обновить отображение

    def _force_redraw_chat(self):
        if self.current_chat_index >= 0:
            self._redraw_chat_with_history()
            self.chat_browser.repaint()
            self.chat_browser.viewport().update()

    def _update_max_width(self):
        parent = self.parent()
        if parent and parent.width() > 0:
            self.setMaximumWidth(int(parent.width() * 0.8))

    def close_current_chat(self):
        """Закрывает (удаляет) текущий выбранный чат."""

        if self.current_chat_index < 0:
            return
        chat = self.chats[self.current_chat_index]
        filepath = chat['filepath']
        item = self.chat_items.get(filepath)
        if item:
            self.delete_chat(item)  # используем существующий метод удаления

    def _prompt_for_folder_name(self):
        """Запрашивает имя папки и создаёт её в дереве. Возвращает имя папки или ''."""

        folder_name, ok = QInputDialog.getText(self, "Новая папка", "Введите название папки:")
        if ok and folder_name.strip():
            folder_name = folder_name.strip()
            if folder_name not in self.folder_items:
                # Добавляем папку в список и дерево
                self.folders.append(folder_name)
                self.save_folders()
                folder_item = QTreeWidgetItem([folder_name])
                folder_item.setFlags(folder_item.flags() | Qt.ItemIsDropEnabled)
                folder_item.setData(0, Qt.UserRole, {'type': 'folder', 'name': folder_name})
                self.folder_items[folder_name] = folder_item
                self.chat_tree.addTopLevelItem(folder_item)
            return folder_name
        return ''

    def open_rag_manager_dialog(self):
        """Открывает окно управления базой знаний RAG."""

        dialog = RAGManagerDialog(self, self)
        dialog.exec()

    def show_help_dialog(self):
        """Открывает окно встроенной справки."""

        dialog = HelpDialog(self)
        dialog.exec()

    # ======================================================================
    #  ИНДИКАТОРЫ, КНОПКИ, УВЕДОМЛЕНИЯ
    # ======================================================================

    def _show_thinking_indicator(self):
        # Удаляем предыдущий индикатор
        if hasattr(self, 'thinking_widget') and self.thinking_widget is not None:
            try:
                self.thinking_widget.deleteLater()
            except RuntimeError:
                pass
            self.thinking_widget = None

        self.thinking_widget = MessageWidget(
            role="assistant",
            content="⏳ Думаю...",
            time_str="",
            reasoning="",
            pinned=False,
            idx=None,
            parent=self.chat_messages_widget,
            show_actions=False
        )

        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(0)
        row_layout.addWidget(self.thinking_widget)
        row_layout.addStretch(1)

        self.chat_messages_layout.addLayout(row_layout)
        QTimer.singleShot(0, self._do_scroll_to_bottom)

    def _remove_thinking_indicator(self):
        if hasattr(self, 'thinking_widget') and self.thinking_widget is not None:
            try:
                self.thinking_widget.deleteLater()
            except RuntimeError:
                pass
            self.thinking_widget = None

    def show_stop_notification(self):
        """Показывает полупрозрачное сообщение 'Процесс остановлен'."""

        # Если лейбл ещё не создан, создаём
        if not hasattr(self, 'stop_notification_label'):
            self.stop_notification_label = QLabel(self.chat_browser)
            self.stop_notification_label.setText("Процесс остановлен")
            self.stop_notification_label.setAlignment(Qt.AlignCenter)
            self.stop_notification_label.setStyleSheet("""
                QLabel {
                    background-color: rgba(0, 0, 0, 0.7);
                    color: white;
                    border-radius: 15px;
                    padding: 10px 20px;
                    font-size: 14px;
                }
            """)
            self.stop_notification_label.setVisible(False)
        # Позиционируем по центру
        self.position_stop_notification()
        self.stop_notification_label.setVisible(True)
        self.stop_notification_label.raise_()
        # Через 2 секунды скрываем
        QTimer.singleShot(2000, self.hide_stop_notification)

    def position_stop_notification(self):
        if hasattr(self, 'stop_notification_label'):
            vp = self.chat_browser.viewport()
            if vp:
                label = self.stop_notification_label
                x = (vp.width() - label.width()) // 2
                y = (vp.height() - label.height()) // 2
                label.move(x, y)

    def hide_stop_notification(self):
        if hasattr(self, 'stop_notification_label'):
            self.stop_notification_label.setVisible(False)

    def update_floating_buttons_positions(self):
        self.position_scroll_down_button()
        self.position_stop_button()
        self.position_chat_settings_button()

    def show_stop_button(self):
        self.update_floating_buttons_positions()
        self.btn_stop_generation.update_position()
        self.btn_stop_generation.setVisible(True)
        self.btn_stop_generation.raise_()

    def hide_stop_button(self):
        self.btn_stop_generation.setVisible(False)
        self.update_floating_buttons_positions()

    def position_stop_button(self):
        vp = self.chat_scroll_area.viewport()
        if vp:
            x = vp.width() - self.btn_stop_generation.width() - 10
            y = vp.height() - self.btn_stop_generation.height() - 10
            self.btn_stop_generation.move(x, y)

    def position_chat_settings_button(self):
        vp = self.chat_scroll_area.viewport()
        if vp:
            x = vp.width() - self.btn_chat_settings.width() - 15
            y = 10
            self.btn_chat_settings.move(x, y)

    def position_scroll_down_button(self):
        vp = self.chat_scroll_area.viewport()
        if vp:
            x = vp.width() - self.btn_scroll_down.width() - 10
            y = vp.height() - self.btn_scroll_down.height() - 10
            if hasattr(self, 'btn_stop_generation') and self.btn_stop_generation.isVisible():
                y -= self.btn_stop_generation.height() + 5
            self.btn_scroll_down.move(x, y)

    def update_chat_settings_button_visibility(self):
        has_chat = self.current_chat_index >= 0
        self.btn_chat_settings.setVisible(has_chat)
        if has_chat:
            self.position_chat_settings_button()
            self.btn_chat_settings.raise_()

    # ======================================================================
    #  ПОИСК И НАВИГАЦИЯ
    # ======================================================================

    def show_search_panel(self):
        if self.current_chat_index < 0:
            return
        self.search_panel.setVisible(True)
        self.search_edit.setFocus()
        self.search_edit.clear()
        self.search_counter_label.setText("0")
        self.update_search_panel_position()

    def _on_search_text_changed(self):
        """Сброс текущего индекса при изменении запроса."""

        self._current_search_idx = -1
        self._search_total = 0
        self.update_search_counter()

    def close_search_panel(self):
        self.search_panel.setVisible(False)
        self.search_counter_label.setText("")
        self._current_search_idx = -1
        self._search_total = 0
        for fp, idx, widget in self.message_widgets:
            widget.apply_search_highlight("")

    def update_search_counter(self):
        """Показывает 'текущее / всего' — например, '2 из 5'."""
        query = self.search_edit.text().strip()

        # Пустой запрос — очищаем
        if not query:
            self.search_counter_label.setText("")
            self._search_total = 0
            self._current_search_idx = -1
            return

        total = self._count_matches(query)
        self._search_total = total

        if total == 0:
            self.search_counter_label.setText("нет")
            return

        # Ещё не переходили ни к одному совпадению
        if not hasattr(self, '_current_search_idx') or self._current_search_idx < 0:
            self.search_counter_label.setText(f"— из {total}")
            return

        current = self._current_search_idx + 1
        self.search_counter_label.setText(f"{current} из {total}")

    def search_next(self):
        query = self.search_edit.text().strip()
        if not query:
            return
        self._search(query, forward=True)

    def search_prev(self):
        query = self.search_edit.text().strip()
        if not query:
            return
        self._search(query, forward=False)

    def _search(self, query, forward=True):
        """Ищет query среди сообщений текущего чата, переходит к найденному и подсвечивает."""

        if not query or not query.strip():
            return
        if self.current_chat_index < 0:
            return

        query = query.strip()
        filepath = self.chats[self.current_chat_index]['filepath']

        # Собираем индексы виджетов, где есть совпадения
        matches = []
        for fp, idx, widget in self.message_widgets:
            if fp != filepath:
                continue
            text = widget.raw_content or ""
            if self.check_regex.isChecked():
                try:
                    if re.search(query, text, re.IGNORECASE):
                        matches.append((idx, widget))
                except re.error:
                    pass
            else:
                if query.lower() in text.lower():
                    matches.append((idx, widget))

        if not matches:
            self._current_search_idx = -1
            self._search_total = 0
            self.update_search_counter()
            self.statusBar().showMessage("Совпадений не найдено")
            return

        # Сохраняем общее число совпадений
        self._search_total = len(matches)

        # Храним текущую позицию между вызовами
        if not hasattr(self, '_current_search_idx'):
            self._current_search_idx = -1

        if forward:
            self._current_search_idx = (self._current_search_idx + 1) % len(matches)
        else:
            self._current_search_idx = (self._current_search_idx - 1) % len(matches)

        target_idx, target_widget = matches[self._current_search_idx]

        # Прокручиваем к найденному сообщению
        self._scroll_to_message(filepath, target_idx)

        # Подсвечиваем совпадения во всех подходящих виджетах, снимаем в остальных
        self._highlight_all_matches(query, matches)
        self.update_search_counter()

    def _highlight_all_matches(self, query, matches):
        """Подсвечивает совпадения в подходящих виджетах, у остальных убирает подсветку."""

        # Индексы совпавших виджетов для быстрой проверки
        match_ids = {id(w) for _, w in matches}

        for fp, idx, widget in self.message_widgets:
            if id(widget) in match_ids:
                widget.apply_search_highlight(
                    query,
                    widget.role == 'assistant',
                    markdown_func=self._markdown_to_html,
                    escape_func=self._escape_html
                )
            else:
                # Убираем подсветку
                widget.apply_search_highlight("")

    def _count_matches(self, query):
        """Считает количество сообщений с совпадениями в текущем чате."""

        if not query or not query.strip():
            return 0
        if self.current_chat_index < 0:
            return 0

        query = query.strip()
        filepath = self.chats[self.current_chat_index]['filepath']
        count = 0
        for fp, idx, widget in self.message_widgets:
            if fp != filepath:
                continue
            text = widget.raw_content or ""
            if self.check_regex.isChecked():
                try:
                    if re.search(query, text, re.IGNORECASE):
                        count += 1
                except re.error:
                    pass
            else:
                if query.lower() in text.lower():
                    count += 1
        return count

    def update_search_panel_position(self):
        if self.search_panel:
            width = 400
            x = (self.chat_scroll_area.viewport().width() - width) // 2
            y = 5
            self.search_panel.move(x, y)
            self.search_panel.setFixedWidth(width)

    def open_global_search(self):
        """Открывает диалог глобального поиска по всем чатам."""

        dialog = GlobalSearchDialog(self.chats, self)
        if dialog.exec() == QDialog.Accepted:
            # Получаем выбранный результат
            filepath, msg_idx = dialog.get_selected_result()
            if filepath is not None and msg_idx is not None:
                self.navigate_to_message(filepath, msg_idx)


class DropImageLabel(QLabel):
    """Метка для приёма изображений перетаскиванием (используется в диалоге настройки обоев)."""

    def __init__(self, dialog, parent=None):
        super().__init__(parent)
        self.dialog = dialog
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                ext = os.path.splitext(urls[0].toLocalFile())[1].lower()
                if ext in ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp'):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
            if os.path.exists(path):
                self.dialog.set_image_from_path(path)
                event.acceptProposedAction()
                return
        event.ignore()


class TrashDialog(QDialog):
    """Диалог корзины.
    Показывает удалённые чаты и папки, позволяет восстановить или окончательно удалить их."""

    def __init__(self, trash_items, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Trash_2.png"))
        self.setWindowTitle("Корзина")
        self.setModal(True)
        self.resize(500, 400)
        self.trash_items = trash_items
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        for entry in self.trash_items:
            if entry['type'] == 'chat':
                label = f"Чат: {entry['title']} ({entry['deleted_time']})"
            elif entry['type'] == 'folder':
                chat_count = len(entry.get('chats', []))
                label = f"Папка: {entry['name']} ({chat_count} чат(ов)) ({entry['deleted_time']})"
            else:
                continue
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, entry)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        btn_restore = QPushButton("Восстановить")
        btn_restore.clicked.connect(self.restore_selected)
        btn_delete = QPushButton("Удалить навсегда")
        btn_delete.clicked.connect(self.permanent_delete_selected)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.reject)

        btn_box = QHBoxLayout()
        btn_box.addWidget(btn_restore)
        btn_box.addWidget(btn_delete)
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def restore_selected(self):
        selected = self.list_widget.currentItem()
        if selected:
            entry = selected.data(Qt.UserRole)
            if entry['type'] == 'chat':
                self.parent().restore_chat(entry['filepath'])
            elif entry['type'] == 'folder':
                self.parent().restore_folder_entry(entry)
            self.refresh_list()

    def permanent_delete_selected(self):
        selected = self.list_widget.currentItem()
        if selected:
            entry = selected.data(Qt.UserRole)
            self.parent().permanently_delete_entry(entry)
            self.refresh_list()

    def refresh_list(self):
        self.list_widget.clear()
        for entry in self.parent().trash_items:
            if entry['type'] == 'chat':
                label = f"Чат: {entry['title']} ({entry['deleted_time']})"
            elif entry['type'] == 'folder':
                chat_count = len(entry.get('chats', []))
                label = f"Папка: {entry['name']} ({chat_count} чат(ов)) ({entry['deleted_time']})"
            else:
                continue
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, entry)
            self.list_widget.addItem(item)

    def get_updated_trash(self):
        return self.parent().trash_items


class ArchiveDialog(QDialog):
    """Диалог архива чатов.
    Отображает заархивированные чаты, даёт возможность восстановить или удалить навсегда."""

    def __init__(self, archived_chats, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Archive_2.png"))
        self.setWindowTitle("Архив чатов")
        self.setModal(True)
        self.resize(500, 400)
        self.archived_chats = archived_chats
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        for chat in self.archived_chats:
            item = QListWidgetItem(chat['title'])
            item.setData(Qt.UserRole, chat['filepath'])
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        btn_restore = QPushButton("Восстановить")
        btn_restore.clicked.connect(self.restore_selected)
        btn_delete = QPushButton("Удалить навсегда")
        btn_delete.clicked.connect(self.permanent_delete_selected)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)

        btn_box = QHBoxLayout()
        btn_box.addWidget(btn_restore)
        btn_box.addWidget(btn_delete)
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def restore_selected(self):
        item = self.list_widget.currentItem()
        if item:
            filepath = item.data(Qt.UserRole)
            if self.parent.unarchive_chat(filepath):
                self.refresh_list()

    def permanent_delete_selected(self):
        item = self.list_widget.currentItem()
        if item:
            filepath = item.data(Qt.UserRole)
            if self.parent.permanently_delete_archived(filepath):
                self.refresh_list()

    def refresh_list(self):
        self.list_widget.clear()
        for chat in self.parent.archived_chats:
            item = QListWidgetItem(chat['title'])
            item.setData(Qt.UserRole, chat['filepath'])
            self.list_widget.addItem(item)


class PromptManagerDialog(QDialog):
    """Диалог управления сохранёнными промптами.
    Позволяет просматривать, вставлять и удалять промпты."""

    def __init__(self, prompts, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Promt_2.png"))
        self.setWindowTitle("Мои промпты")
        self.setModal(True)
        self.resize(400, 300)
        self.prompts = prompts
        self.parent = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.list_widget = QListWidget()
        for prompt in self.prompts:
            item = QListWidgetItem(prompt['name'])
            item.setData(Qt.UserRole, prompt)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        # Включаем контекстное меню для списка
        self.list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self.show_item_context_menu)

        btn_box = QHBoxLayout()
        btn_insert = QPushButton("Вставить")
        btn_insert.clicked.connect(self.insert_selected)
        btn_delete = QPushButton("Удалить")
        btn_delete.clicked.connect(self.delete_selected)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)

        btn_box.addWidget(btn_insert)
        btn_box.addWidget(btn_delete)
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def show_item_context_menu(self, pos):
        item = self.list_widget.itemAt(pos)
        if not item:
            return
        menu = QMenu(self)
        insert_action = menu.addAction("Выбрать (вставить)")
        delete_action = menu.addAction("Удалить")
        chosen = menu.exec(self.list_widget.mapToGlobal(pos))
        if chosen == insert_action:
            self.insert_prompt(item)
        elif chosen == delete_action:
            self.delete_prompt(item)

    def insert_selected(self):
        item = self.list_widget.currentItem()
        if item:
            self.insert_prompt(item)

    def delete_selected(self):
        item = self.list_widget.currentItem()
        if item:
            self.delete_prompt(item)

    def insert_prompt(self, item):
        prompt = item.data(Qt.UserRole)
        self.parent.insert_prompt_text(prompt['text'])
        QMessageBox.information(self, "Готово", "Промпт вставлен в поле ввода.")

    def delete_prompt(self, item):
        prompt = item.data(Qt.UserRole)
        self.prompts.remove(prompt)
        self.refresh_list()
        self.parent.save_prompts()

    def refresh_list(self):
        self.list_widget.clear()
        for prompt in self.prompts:
            item = QListWidgetItem(prompt['name'])
            item.setData(Qt.UserRole, prompt)
            self.list_widget.addItem(item)

    def get_updated_prompts(self):
        return self.prompts


class GlobalSearchDialog(QDialog):
    """Диалог глобального поиска по всем чатам.
    Поддерживает фильтр по дате, регулярные выражения, показывает результаты с подсветкой."""

    def __init__(self, chats, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Global_search.png"))
        self.setWindowTitle("Глобальный поиск по чатам")
        self.setModal(True)
        self.resize(600, 450)
        self.chats = chats
        self.selected_filepath = None
        self.selected_msg_idx = None
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Поле ввода запроса
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите текст для поиска...")
        self.search_edit.textChanged.connect(self.update_results)
        layout.addWidget(self.search_edit)

        # Чекбокс регулярных выражений
        self.check_regex = QCheckBox("Regex")
        self.check_regex.toggled.connect(self.update_results)

        # Фильтр по дате
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addMonths(-1))
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        self.date_from.dateChanged.connect(self.update_results)
        self.date_to.dateChanged.connect(self.update_results)

        # Layout для фильтров
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Даты:"))
        filter_layout.addWidget(self.date_from)
        filter_layout.addWidget(QLabel("–"))
        filter_layout.addWidget(self.date_to)
        filter_layout.addWidget(self.check_regex)
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        # Список результатов
        self.results_list = QListWidget()
        self.results_list.setAlternatingRowColors(True)
        self.results_list.itemDoubleClicked.connect(self.on_item_double_clicked)
        layout.addWidget(self.results_list, stretch=1)

        # Нижняя панель: счётчик и кнопки
        bottom_layout = QHBoxLayout()
        self.counter_label = QLabel("Найдено: 0")
        self.counter_label.setStyleSheet("color:#666;")
        btn_go = QPushButton("Перейти")
        btn_go.clicked.connect(self.on_go_clicked)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.reject)

        bottom_layout.addWidget(self.counter_label)
        bottom_layout.addStretch()
        bottom_layout.addWidget(btn_go)
        bottom_layout.addWidget(btn_close)
        layout.addLayout(bottom_layout)

        self.search_edit.setFocus()

    def update_results(self):
        query = self.search_edit.text().strip()
        self.results_list.clear()
        self.selected_filepath = None
        self.selected_msg_idx = None

        if not query:
            self.counter_label.setText("Найдено: 0")
            return

        use_regex = self.check_regex.isChecked()
        date_from = self.date_from.date().toPython()
        date_to = self.date_to.date().toPython()

        total_found = 0
        for chat in self.chats:
            chat_title = chat.get('title', 'Без названия')
            for idx, msg in enumerate(chat.get('messages', [])):
                content = msg.get('content', '')
                if not isinstance(content, str):
                    continue

                # Фильтр по дате (строгий)
                time_str = msg.get('time', '')
                if not time_str:
                    continue  # сообщения без даты пропускаем
                try:
                    dt = datetime.strptime(time_str, '%d.%m.%Y %H:%M:%S').date()
                    if not (date_from <= dt <= date_to):
                        continue
                except ValueError:
                    continue  # некорректная дата тоже пропускается

                # Поиск
                matched = False
                match_pos = -1
                if use_regex:
                    try:
                        m = re.search(query, content, re.IGNORECASE)
                        if m:
                            matched = True
                            match_pos = m.start()
                    except re.error:
                        continue
                else:
                    pos = content.lower().find(query.lower())
                    if pos != -1:
                        matched = True
                        match_pos = pos

                if not matched:
                    continue

                total_found += 1

                # Формирование предпросмотра
                start = max(0, match_pos - 30)
                end = min(len(content), match_pos + len(query) + 30)
                preview = content[start:end]
                if start > 0:
                    preview = "..." + preview
                if end < len(content):
                    preview = preview + "..."
                preview = preview.replace('\n', ' ')

                # Подсветка найденного фрагмента
                if use_regex:
                    highlighted = re.sub(query, lambda m: f"<b style='background-color: yellow;'>{m.group()}</b>",
                                         preview, flags=re.IGNORECASE)
                else:
                    lower_preview = preview.lower()
                    idx_pos = lower_preview.find(query.lower())
                    if idx_pos != -1:
                        original_fragment = preview[idx_pos:idx_pos + len(query)]
                        highlighted = preview[:idx_pos] + \
                                      f"<b style='background-color: yellow;'>{original_fragment}</b>" + \
                                      preview[idx_pos + len(query):]
                    else:
                        highlighted = preview

                # Создаём элемент списка
                item = QListWidgetItem()
                item.setText(f"{chat_title} → {highlighted}")
                item.setData(Qt.UserRole, (chat['filepath'], idx))
                self.results_list.addItem(item)

        self.counter_label.setText(f"Найдено: {total_found}")

    def on_item_double_clicked(self, item):
        data = item.data(Qt.UserRole)
        if data:
            self.selected_filepath, self.selected_msg_idx = data
            self.accept()

    def on_go_clicked(self):
        item = self.results_list.currentItem()
        if not item:
            return
        data = item.data(Qt.UserRole)
        if data:
            self.selected_filepath, self.selected_msg_idx = data
            self.accept()

    def get_selected_result(self):
        return self.selected_filepath, self.selected_msg_idx


class StarredMessagesDialog(QDialog):
    """Диалог закладок.
    Показывает избранные сообщения текущего чата и позволяет перейти к ним или снять отметку."""

    def __init__(self, chat, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Bookmarks_2.png"))
        self.setWindowTitle("Закладки")
        self.setModal(True)
        self.resize(450, 500)
        self.chat = chat
        self.parent = parent  # MainWindow
        self.init_ui()
        self.populate()

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self.on_go)
        layout.addWidget(self.list_widget)

        btn_box = QHBoxLayout()
        btn_go = QPushButton("Перейти")
        btn_go.clicked.connect(self.on_go)
        btn_remove = QPushButton("Удалить из избранного")
        btn_remove.clicked.connect(self.remove_selected)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.reject)

        btn_box.addWidget(btn_go)
        btn_box.addWidget(btn_remove)
        btn_box.addStretch()
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def populate(self):
        self.list_widget.clear()
        for idx, msg in enumerate(self.chat.get('messages', [])):
            if msg.get('starred', False):
                content = msg.get('content', '')
                preview = content[:80] + ('...' if len(content) > 80 else '')
                role = "Вы" if msg.get('role') == 'user' else "Ассистент"
                item = QListWidgetItem(f"{role}: {preview}")
                item.setData(Qt.UserRole, idx)
                self.list_widget.addItem(item)

    def get_selected_index(self):
        item = self.list_widget.currentItem()
        if item:
            return item.data(Qt.UserRole)
        return None

    def on_go(self):
        idx = self.get_selected_index()
        if idx is not None:
            self.accept()

    def remove_selected(self):
        idx = self.get_selected_index()
        if idx is not None:
            # Снимаем пометку
            self.parent.toggle_star_message(idx)
            self.populate()  # обновляем список


class DeepSeekImportDialog(QDialog):
    """Диалог выбора чатов для импорта из файла DeepSeek.
    Позволяет отметить нужные чаты галочками."""

    def __init__(self, chats: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Выбор чатов для импорта")
        self.setModal(True)
        self.resize(500, 400)
        self.chats = chats

        layout = QVBoxLayout(self)

        img_label = QLabel()
        pixmap = QPixmap("Image/Import_2.png")
        if not pixmap.isNull():
            img_label.setPixmap(pixmap.scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            img_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(img_label)

        label = QLabel("Отметьте галочками чаты, которые хотите импортировать:")
        layout.addWidget(label)

        self.list_widget = QListWidget()
        for title, _ in chats:
            item = QListWidgetItem(title)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        btn_box = QHBoxLayout()
        btn_ok = QPushButton("Импортировать отмеченные")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addStretch()
        btn_box.addWidget(btn_ok)
        btn_box.addWidget(btn_cancel)
        layout.addLayout(btn_box)

    def get_selected(self):
        selected = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                selected.append(self.chats[i])
        return selected


class StatisticsDialog(QDialog):
    """Диалог статистики чата.
    Отображает основные метрики (количество сообщений, токенов, время генерации) и графики (если доступен matplotlib)."""

    def __init__(self, chat_data, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Statistics_2.png"))
        self.setWindowTitle("Статистика чата")
        self.setModal(True)
        self.setFixedSize(500, 600)  # фиксированный размер окна
        self.chat_data = chat_data
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)

        # Заголовок
        title_label = QLabel(f"Чат: {self.chat_data.get('title', '')}")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        main_layout.addWidget(title_label)

        # Текстовая информация
        form_layout = QFormLayout()
        created = self.chat_data.get('created_time', 'Неизвестно')
        form_layout.addRow("Создан:", QLabel(created))
        last_activity = self.chat_data.get('last_activity_time', 'Неизвестно')
        form_layout.addRow("Последняя активность:", QLabel(last_activity))
        msg_count = len(self.chat_data.get('messages', []))
        form_layout.addRow("Сообщений:", QLabel(str(msg_count)))
        total_chars = sum(len(msg.get('content', '')) for msg in self.chat_data.get('messages', []))
        form_layout.addRow("Символов в истории:", QLabel(str(total_chars)))
        tokens_sent = self.chat_data.get('total_tokens_sent', 0)
        form_layout.addRow("Отправлено токенов (оценка):", QLabel(str(tokens_sent)))
        total_time = self.chat_data.get('total_generation_time', 0.0)
        form_layout.addRow("Время генерации (сек):", QLabel(f"{total_time:.2f}"))
        attachment_count = sum(len(msg.get('attachments', [])) for msg in self.chat_data.get('messages', []))
        form_layout.addRow("Вложений:", QLabel(str(attachment_count)))
        filepath = self.chat_data.get('filepath')
        if filepath and os.path.exists(filepath):
            size_kb = os.path.getsize(filepath) / 1024
            form_layout.addRow("Размер файла:", QLabel(f"{size_kb:.2f} КБ"))
        else:
            form_layout.addRow("Размер файла:", QLabel("Недоступен"))
        main_layout.addLayout(form_layout)

        # === Графики (если matplotlib доступен) ===
        if MATPLOTLIB_AVAILABLE:
            try:
                figure = Figure(figsize=(6, 6), dpi=80)

                messages = self.chat_data.get('messages', [])
                dates = []
                lengths = []
                roles = []
                for msg in messages:
                    time_str = msg.get('time', '')
                    if time_str:
                        date = time_str.split(' ')[0]
                        dates.append(date)
                    else:
                        dates.append('?')
                    content = msg.get('content', '')
                    lengths.append(len(content))
                    roles.append(msg.get('role', ''))

                # Подграфик 1: количество сообщений по дням
                ax1 = figure.add_subplot(311)
                if dates:
                    unique_dates = sorted(set(dates))
                    counts = [dates.count(d) for d in unique_dates]
                    ax1.bar(unique_dates, counts, color='#4CAF50')
                    ax1.set_title('Сообщений по дням', fontsize=9)
                    ax1.set_ylabel('Кол-во', fontsize=8)
                    ax1.tick_params(axis='x', rotation=0, labelsize=7)
                    ax1.tick_params(axis='y', labelsize=7)
                else:
                    ax1.text(0.5, 0.5, 'Нет данных', ha='center', va='center')

                # Подграфик 2: средняя длина ответа ассистента
                ax2 = figure.add_subplot(312)
                assistant_lengths = [lengths[i] for i, r in enumerate(roles) if r == 'assistant']
                if assistant_lengths:
                    avg_len = sum(assistant_lengths) / len(assistant_lengths)
                    ax2.bar(['Средняя длина'], [avg_len], color='#2196F3')
                    ax2.set_title('Средняя длина ответа ассистента (символов)', fontsize=9)
                    ax2.set_ylabel('Символов', fontsize=8)
                    ax2.tick_params(axis='both', labelsize=7)
                else:
                    ax2.text(0.5, 0.5, 'Нет ответов ассистента', ha='center', va='center')

                # Подграфик 3: расход токенов по дням (приблизительно)
                ax3 = figure.add_subplot(313)
                tokens_per_day = {}
                for i, msg in enumerate(messages):
                    if dates[i] != '?':
                        tokens = lengths[i] / 4
                        tokens_per_day[dates[i]] = tokens_per_day.get(dates[i], 0) + tokens
                if tokens_per_day:
                    sorted_dates = sorted(tokens_per_day.keys())
                    values = [tokens_per_day[d] for d in sorted_dates]
                    ax3.bar(sorted_dates, values, color='#FF9800')
                    ax3.set_title('Расход токенов по дням (оценка)', fontsize=9)
                    ax3.set_xlabel('Дата', fontsize=8)
                    ax3.set_ylabel('Токенов', fontsize=8)
                    ax3.tick_params(axis='x', rotation=0, labelsize=7)
                    ax3.tick_params(axis='y', labelsize=7)
                else:
                    ax3.text(0.5, 0.5, 'Нет данных', ha='center', va='center')

                figure.subplots_adjust(
                    left=0.15, right=0.9, top=0.92, bottom=0.12,
                    hspace=1.0, wspace=0.3
                )

                canvas = FigureCanvas(figure)
                canvas.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
                main_layout.addWidget(canvas)
            except Exception as e:
                err_label = QLabel(f"Ошибка построения графиков:\n{e}")
                err_label.setStyleSheet("color: #d32f2f; font-size: 11px;")
                err_label.setWordWrap(True)
                main_layout.addWidget(err_label)
        else:
            msg_label = QLabel(
                "Для отображения графиков установите библиотеку matplotlib:\n"
                "pip install matplotlib"
            )
            msg_label.setStyleSheet("color: #888; font-style: italic;")
            msg_label.setAlignment(Qt.AlignCenter)
            main_layout.addWidget(msg_label)

        # Кнопка закрытия
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_box.addWidget(btn_close)
        main_layout.addLayout(btn_box)


class SettingsDialog(QDialog):
    """Главный диалог настроек программы.
    Включает вкладки/группы для настройки подключения, параметров генерации, шрифтов, уведомлений, безопасности, горячих клавиш и др."""

    def __init__(self, settings: Dict[str, Any], parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Settings_2.png"))
        self.setWindowTitle("Настройки")
        self.settings = settings.copy()
        self.setModal(True)
        self.resize(750, 650)
        self.init_ui()

    def init_ui(self):
        # Основная вертикальная разметка окна
        main_layout = QVBoxLayout(self)

        # Область прокрутки
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)

        # Контейнер для содержимого прокрутки
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(10, 10, 10, 10)
        content_layout.setSpacing(10)

        # ---------- Подключение ----------
        group_conn = QGroupBox("Подключение")
        form_conn = QFormLayout()
        self.edit_api_base = QLineEdit(self.settings.get("api_base", "http://localhost:1234/v1"))
        self.edit_model = QLineEdit(self.settings.get("model", "local-model"))
        form_conn.addRow("API Base URL:", self.edit_api_base)
        form_conn.addRow("Model:", self.edit_model)
        group_conn.setLayout(form_conn)
        content_layout.addWidget(group_conn)

        # ---------- LM Studio ----------
        group_lm = QGroupBox("LM Studio")
        form_lm = QFormLayout()
        self.edit_lm_path = QLineEdit(self.settings.get("lm_studio_path", ""))
        self.btn_browse_lm = QPushButton("Обзор...")
        self.btn_browse_lm.clicked.connect(self.browse_lm_studio)
        path_layout = QHBoxLayout()
        path_layout.addWidget(self.edit_lm_path)
        path_layout.addWidget(self.btn_browse_lm)
        form_lm.addRow("Путь к LM Studio:", path_layout)
        self.check_auto_start = QCheckBox("Автоматически запускать LM Studio при старте программы")
        self.check_auto_start.setChecked(bool(self.settings.get("lm_studio_auto_start", False)))
        form_lm.addRow("", self.check_auto_start)
        self.btn_launch_lm = QPushButton("Запустить LM Studio")
        self.btn_launch_lm.clicked.connect(self.on_launch_lm_clicked)
        form_lm.addRow("", self.btn_launch_lm)
        group_lm.setLayout(form_lm)
        content_layout.addWidget(group_lm)

        # ---------- Параметры генерации ----------
        group_gen = QGroupBox("Параметры генерации")
        form_gen = QFormLayout()
        self.spin_max_tokens = QSpinBox()
        self.spin_max_tokens.setRange(1, 100000)
        self.spin_max_tokens.setValue(int(self.settings.get("max_tokens", 4096)))
        self.spin_temperature = QDoubleSpinBox()
        self.spin_temperature.setRange(0.0, 2.0)
        self.spin_temperature.setSingleStep(0.05)
        self.spin_temperature.setValue(float(self.settings.get("temperature", 0.7)))
        self.spin_top_p = QDoubleSpinBox()
        self.spin_top_p.setRange(0.0, 1.0)
        self.spin_top_p.setSingleStep(0.05)
        self.spin_top_p.setValue(float(self.settings.get("top_p", 0.9)))
        self.spin_top_k = QSpinBox()
        self.spin_top_k.setRange(0, 200)
        self.spin_top_k.setValue(int(self.settings.get("top_k", 40)))
        self.spin_repeat = QDoubleSpinBox()
        self.spin_repeat.setRange(0.0, 2.0)
        self.spin_repeat.setSingleStep(0.05)
        self.spin_repeat.setValue(float(self.settings.get("repeat_penalty", 1.1)))
        self.check_stream = QCheckBox("Потоковая передача (stream)")
        self.check_stream.setChecked(bool(self.settings.get("stream", True)))
        self.check_multimodal = QCheckBox("Мультимодальная модель (поддержка изображений)")
        self.check_multimodal.setChecked(bool(self.settings.get("multimodal", False)))

        form_gen.addRow("Max Tokens:", self.spin_max_tokens)
        form_gen.addRow("Temperature:", self.spin_temperature)
        form_gen.addRow("Top P:", self.spin_top_p)
        form_gen.addRow("Top K:", self.spin_top_k)
        form_gen.addRow("Repeat Penalty:", self.spin_repeat)
        form_gen.addRow("", self.check_stream)
        form_gen.addRow("", self.check_multimodal)

        self.spin_max_images = QSpinBox()
        self.spin_max_images.setRange(1, 20)
        self.spin_max_images.setValue(int(self.settings.get("max_images_per_message", 4)))
        form_gen.addRow("Макс. изображений в сообщении:", self.spin_max_images)

        # Ограничение на длину промпта
        self.spin_max_prompt_chars = QSpinBox()
        self.spin_max_prompt_chars.setRange(100, 1000000)
        self.spin_max_prompt_chars.setSingleStep(100)
        self.spin_max_prompt_chars.setValue(int(self.settings.get("max_prompt_chars", 8000)))
        form_gen.addRow("Макс. длина промпта (символов):", self.spin_max_prompt_chars)

        group_gen.setLayout(form_gen)
        content_layout.addWidget(group_gen)

        # ---------- Шрифты ----------
        group_fonts = QGroupBox("Шрифты")
        form_fonts = QFormLayout()
        self.combo_chat_font = QFontComboBox()
        self.combo_chat_font.setCurrentFont(QFont(self.settings.get("chat_font_family", "Arial")))
        self.spin_chat_font_size = QSpinBox()
        self.spin_chat_font_size.setRange(6, 72)
        self.spin_chat_font_size.setValue(int(self.settings.get("chat_font_size", 10)))
        self.combo_input_font = QFontComboBox()
        self.combo_input_font.setCurrentFont(QFont(self.settings.get("input_font_family", "Arial")))
        self.spin_input_font_size = QSpinBox()
        self.spin_input_font_size.setRange(6, 72)
        self.spin_input_font_size.setValue(int(self.settings.get("input_font_size", 10)))
        form_fonts.addRow("Шрифт чата:", self.combo_chat_font)
        form_fonts.addRow("Размер шрифта чата:", self.spin_chat_font_size)
        form_fonts.addRow("Шрифт редактора:", self.combo_input_font)
        form_fonts.addRow("Размер шрифта редактора:", self.spin_input_font_size)
        group_fonts.setLayout(form_fonts)
        content_layout.addWidget(group_fonts)

        # ---------- Контекст и системный промпт ----------
        group_ctx = QGroupBox("Контекст и системный промпт")
        form_ctx = QFormLayout()
        self.spin_max_ctx = QSpinBox()
        self.spin_max_ctx.setRange(1, 1000000)
        self.spin_max_ctx.setValue(int(self.settings.get("max_context_tokens", 4096)))
        self.edit_system = QTextEdit()
        if self.settings.get("theme", "light") == "dark":
            self.edit_system.setStyleSheet("""
                QTextEdit {
                    background-color: #1e1e1e;
                    color: #dddddd;
                    border: 1px solid #444;
                    border-radius: 5px;
                }
            """)
        else:
            self.edit_system.setStyleSheet("")
        self.edit_system.setPlainText(self.settings.get("system_prompt", ""))
        self.edit_system.setMaximumHeight(100)
        form_ctx.addRow("Max Context Tokens:", self.spin_max_ctx)
        form_ctx.addRow("System Prompt:", self.edit_system)
        group_ctx.setLayout(form_ctx)
        content_layout.addWidget(group_ctx)

        # ---------- Безопасность ----------
        group_security = QGroupBox("Безопасность")
        form_security = QFormLayout()
        self.check_encrypt = QCheckBox("Шифровать сохранённые чаты")
        self.check_encrypt.setChecked(bool(self.settings.get("encrypt_chats", False)))
        form_security.addRow("", self.check_encrypt)

        self.btn_master_password = QPushButton("Установить/изменить мастер-пароль")
        self.btn_master_password.clicked.connect(self.on_master_password_clicked)
        form_security.addRow("", self.btn_master_password)

        btn_decrypt_all = QPushButton("Расшифровать все зашифрованные чаты")
        btn_decrypt_all.clicked.connect(self.on_decrypt_all_clicked)
        form_security.addRow("", btn_decrypt_all)

        group_security.setLayout(form_security)
        content_layout.addWidget(group_security)

        # ---------- Обновления ----------
        group_updates = QGroupBox("Обновления")
        form_updates = QFormLayout()

        btn_check_updates = QPushButton("Проверить обновления")
        btn_check_updates.clicked.connect(self.on_check_updates_clicked)
        form_updates.addRow("", btn_check_updates)

        self.label_current_version = QLabel(f"Текущая версия: {APP_VERSION}")
        self.label_current_version.setStyleSheet("color: #888;")
        form_updates.addRow("", self.label_current_version)

        btn_whats_new = QPushButton("Что нового?")
        btn_whats_new.clicked.connect(self.on_whats_new_clicked)
        form_updates.addRow("", btn_whats_new)

        group_updates.setLayout(form_updates)
        content_layout.addWidget(group_updates)

        # ---------- База знаний (RAG) ----------
        group_rag = QGroupBox("База знаний (RAG)")
        form_rag = QFormLayout()

        self.check_rag = QCheckBox("Использовать базу знаний (RAG)")
        self.check_rag.setChecked(bool(self.settings.get("rag_enabled", False)))
        form_rag.addRow("", self.check_rag)

        btn_add_rag = QPushButton("Добавить документы в базу знаний...")
        btn_add_rag.clicked.connect(self.on_add_rag_documents)
        form_rag.addRow("", btn_add_rag)

        btn_remove_rag = QPushButton("Удалить документ из базы знаний...")
        btn_remove_rag.clicked.connect(self.on_remove_rag_document)
        form_rag.addRow("", btn_remove_rag)

        btn_open_rag_folder = QPushButton("Открыть папку RAG_folder")
        btn_open_rag_folder.clicked.connect(self.on_open_rag_folder)
        form_rag.addRow("", btn_open_rag_folder)

        btn_refresh_rag = QPushButton("Обновить базу знаний")
        btn_refresh_rag.clicked.connect(self.on_refresh_rag)
        form_rag.addRow("", btn_refresh_rag)

        btn_view_rag = QPushButton("Просмотр базы знаний...")
        btn_view_rag.clicked.connect(self.on_view_rag)
        form_rag.addRow("", btn_view_rag)

        group_rag.setLayout(form_rag)
        content_layout.addWidget(group_rag)

        # ---------- Резервное копирование ----------
        group_backup = QGroupBox("Резервное копирование")
        form_backup = QFormLayout()

        self.check_auto_backup = QCheckBox("Автосохранение резервной копии")
        self.check_auto_backup.setChecked(bool(self.settings.get("auto_backup", True)))
        form_backup.addRow("", self.check_auto_backup)

        self.spin_backup_interval = QSpinBox()
        self.spin_backup_interval.setRange(1, 60)
        self.spin_backup_interval.setValue(int(self.settings.get("auto_backup_interval_min", 5)))
        form_backup.addRow("Интервал бэкапа (мин):", self.spin_backup_interval)

        self.spin_max_backups = QSpinBox()
        self.spin_max_backups.setRange(1, 20)
        self.spin_max_backups.setValue(int(self.settings.get("max_backups", 5)))
        form_backup.addRow("Хранить копий:", self.spin_max_backups)

        self.spin_auto_archive = QSpinBox()
        self.spin_auto_archive.setRange(0, 365)
        self.spin_auto_archive.setValue(int(self.settings.get("auto_archive_days", 0)))
        form_backup.addRow("Автоархивация (дней неактивности):", self.spin_auto_archive)

        self.btn_import_zip = QPushButton("Импортировать ZIP-архив")
        self.btn_import_zip.clicked.connect(self.on_import_zip_clicked)
        form_backup.addRow("", self.btn_import_zip)

        group_backup.setLayout(form_backup)
        content_layout.addWidget(group_backup)

        # ---------- Интерфейс и уведомления ----------
        group_ui = QGroupBox("Интерфейс и уведомления")
        form_ui = QFormLayout()

        self.check_open_last = QCheckBox("Открывать последний чат при запуске")
        self.check_open_last.setChecked(bool(self.settings.get("open_last_chat_on_startup", False)))
        form_ui.addRow("", self.check_open_last)

        self.check_notifications = QCheckBox("Отключить уведомления о новых сообщениях")
        self.check_notifications.setChecked(not self.settings.get("notifications_enabled", True))
        form_ui.addRow("", self.check_notifications)

        self.check_sound_notifications = QCheckBox("Звуковые уведомления")
        self.check_sound_notifications.setChecked(bool(self.settings.get("sound_notifications", True)))
        form_ui.addRow("", self.check_sound_notifications)

        self.check_toast_notifications = QCheckBox("Всплывающие уведомления")
        self.check_toast_notifications.setChecked(bool(self.settings.get("toast_notifications", True)))
        form_ui.addRow("", self.check_toast_notifications)

        self.edit_sound_file = QLineEdit(self.settings.get("sound_file", "Sound_1.MP3"))
        btn_browse_sound = QPushButton("Обзор...")
        btn_browse_sound.clicked.connect(self.browse_sound_file)
        btn_default_sound = QPushButton("По умолчанию")
        btn_default_sound.clicked.connect(self.reset_sound_file)
        sound_layout = QHBoxLayout()
        sound_layout.addWidget(self.edit_sound_file)
        sound_layout.addWidget(btn_browse_sound)
        sound_layout.addWidget(btn_default_sound)
        form_ui.addRow("Звуковой файл:", sound_layout)

        group_ui.setLayout(form_ui)
        content_layout.addWidget(group_ui)

        # ---------- Файлы и лимиты ----------
        group_files = QGroupBox("Файлы и лимиты")
        form_files = QFormLayout()

        self.spin_max_attachment_size = QSpinBox()
        self.spin_max_attachment_size.setRange(1, 1000)
        self.spin_max_attachment_size.setValue(int(self.settings.get("max_attachment_size_mb", 20)))
        form_files.addRow("Макс. размер вложения (МБ):", self.spin_max_attachment_size)

        group_files.setLayout(form_files)
        content_layout.addWidget(group_files)

        # ---------- Прочее ----------
        group_misc = QGroupBox("Прочее")
        form_misc = QFormLayout()

        self.btn_hotkeys = QPushButton("Настроить горячие клавиши...")
        self.btn_hotkeys.clicked.connect(self.open_hotkeys_dialog)
        form_misc.addRow("", self.btn_hotkeys)

        btn_open_logs = QPushButton("Открыть папку с логами")
        btn_open_logs.clicked.connect(self.on_open_logs_clicked)
        form_misc.addRow("", btn_open_logs)

        group_misc.setLayout(form_misc)
        content_layout.addWidget(group_misc)

        # ---------- Плагины ----------
        group_plugins = QGroupBox("Плагины")
        form_plugins = QFormLayout()

        btn_open_plugins = QPushButton("Открыть папку плагинов...")
        btn_open_plugins.clicked.connect(self.open_plugins_folder)
        form_plugins.addRow("", btn_open_plugins)

        btn_manage_plugins = QPushButton("Управление плагинами...")
        btn_manage_plugins.clicked.connect(self.show_plugins_manager)
        form_plugins.addRow("", btn_manage_plugins)

        btn_reload_plugins = QPushButton("Перезагрузить плагины")
        btn_reload_plugins.clicked.connect(self.on_reload_plugins)
        form_plugins.addRow("", btn_reload_plugins)

        group_plugins.setLayout(form_plugins)
        content_layout.addWidget(group_plugins)

        # ---------- О программе ----------
        group_about = QGroupBox("О программе")
        form_about = QFormLayout()

        btn_about = QPushButton("О программе")
        btn_about.clicked.connect(self.show_about)
        form_about.addRow("", btn_about)

        btn_docs = QPushButton("Документация")
        btn_docs.clicked.connect(self.on_open_docs)
        form_about.addRow("", btn_docs)

        group_about.setLayout(form_about)
        content_layout.addWidget(group_about)

        # ---------- Кнопки (только одна пара, внизу) ----------
        btn_box = QHBoxLayout()
        btn_ok = QPushButton("Сохранить")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addStretch()
        btn_box.addWidget(btn_ok)
        btn_box.addWidget(btn_cancel)

        # Собираем прокрутку и добавляем всё в основную разметку
        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area)
        main_layout.addLayout(btn_box)

    def on_whats_new_clicked(self):
        """Показывает историю релизов с GitHub."""

        import requests
        try:
            resp = requests.get(
                f"https://api.github.com/repos/{GITHUB_REPO}/releases",
                timeout=5,
                headers={"Accept": "application/vnd.github+json"}
            )
            if resp.status_code != 200:
                QMessageBox.warning(self, "Ошибка", "Не удалось получить список релизов.")
                return
            releases = resp.json()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось подключиться к GitHub:\n{e}")
            return

        if not releases:
            QMessageBox.information(self, "Что нового", "Релизов пока нет.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Что нового в Echos")
        dialog.setModal(True)
        dialog.resize(700, 600)

        layout = QVBoxLayout(dialog)
        browser = QTextBrowser()
        browser.setReadOnly(True)
        browser.setOpenExternalLinks(True)

        # Цвета под тему
        is_dark = self.settings.get("theme", "light") == "dark"
        bg = "#2b2b2b" if is_dark else "#ffffff"
        text_color = "#dddddd" if is_dark else "#000000"
        link_color = "#8ab4f8" if is_dark else "#1a73e8"
        code_bg = "#1e1e1e" if is_dark else "#f5f5f5"
        muted_color = "#999999" if is_dark else "#888888"
        hr_color = "#555555" if is_dark else "#dddddd"

        # Палитра
        palette = browser.palette()
        palette.setColor(QPalette.Base, QColor(bg))
        palette.setColor(QPalette.Text, QColor(text_color))
        palette.setColor(QPalette.Link, QColor(link_color))
        palette.setColor(QPalette.LinkVisited, QColor(link_color))
        browser.setPalette(palette)

        browser.setStyleSheet(f"""
            QTextBrowser {{
                background-color: {bg};
                color: {text_color};
                border: 1px solid {hr_color};
            }}
        """)

        # CSS внутри документа
        browser.document().setDefaultStyleSheet(
            f"body {{ color: {text_color}; }}"
            f"a {{ color: {link_color}; text-decoration: underline; }}"
            f"code {{ background-color: {code_bg}; padding: 2px 4px; border-radius: 3px; }}"
            f"pre {{ background-color: {code_bg}; padding: 10px; border-radius: 5px; }}"
            f"hr {{ border: none; border-top: 1px solid {hr_color}; }}"
            f".muted {{ color: {muted_color}; font-size: 12px; }}"
        )

        html_parts = []
        for rel in releases[:10]:
            tag = rel.get("tag_name", "").lstrip("v")
            name = rel.get("name", f"Версия {tag}")
            published = rel.get("published_at", "")[:10]
            body = rel.get("body", "").strip()

            html_parts.append(f"<h2>{name} <span class='muted'>({published})</span></h2>")

            if body and MARKDOWN_AVAILABLE:
                try:
                    body_html = markdown.markdown(body, extensions=['fenced_code', 'tables', 'nl2br'])
                    html_parts.append(body_html)
                except Exception:
                    html_parts.append(f"<pre>{body}</pre>")
            else:
                html_parts.append(f"<pre>{body}</pre>")

            html_parts.append("<hr>")

        browser.setHtml("".join(html_parts))
        layout.addWidget(browser)

        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(dialog.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(btn_close)
        layout.addLayout(btn_row)

        dialog.exec()

    def on_check_updates_clicked(self):
        parent = self.parent()
        if parent and hasattr(parent, 'check_for_updates'):
            parent.check_for_updates(silent=False)
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось проверить обновления.")

    def on_open_docs(self):
        parent = self.parent()
        if parent and hasattr(parent, 'show_help_dialog'):
            parent.show_help_dialog()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к документации")

    def on_reload_plugins(self):
        parent = self.parent()
        if parent and hasattr(parent, 'reload_plugins'):
            parent.reload_plugins()
            QMessageBox.information(self, "Плагины", "Плагины успешно перезагружены.")
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции перезагрузки плагинов")

    def on_view_rag(self):
        parent = self.parent()
        if parent and hasattr(parent, 'open_rag_manager_dialog'):
            parent.open_rag_manager_dialog()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции просмотра базы знаний")

    def on_add_rag_documents(self):
        parent = self.parent()
        if parent and hasattr(parent, 'add_rag_documents_via_dialog'):
            parent.add_rag_documents_via_dialog()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции добавления документов")

    def on_remove_rag_document(self):
        parent = self.parent()
        if parent and hasattr(parent, 'remove_rag_document_via_dialog'):
            parent.remove_rag_document_via_dialog()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции удаления документов")

    def on_open_rag_folder(self):
        parent = self.parent()
        if parent and hasattr(parent, 'open_rag_folder'):
            parent.open_rag_folder()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции открытия папки")

    def on_refresh_rag(self):
        parent = self.parent()
        if parent and hasattr(parent, 'load_rag_documents_from_folder'):
            parent.load_rag_documents_from_folder()
            parent.update_rag_status()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции обновления базы знаний")

    def on_open_logs_clicked(self):
        parent = self.parent()
        if parent and hasattr(parent, 'open_logs_folder'):
            parent.open_logs_folder()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции открытия папки логов")

    def on_decrypt_all_clicked(self):
        parent = self.parent()
        if parent and hasattr(parent, 'decrypt_all_encrypted_chats'):
            parent.decrypt_all_encrypted_chats()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции расшифровки")

    def on_master_password_clicked(self):
        parent = self.parent()
        if parent and hasattr(parent, 'setup_master_password'):
            parent.setup_master_password()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции установки мастер-пароля")

    def show_about(self):
        parent = self.parent()
        if parent and hasattr(parent, 'show_about_dialog'):
            parent.show_about_dialog()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к окну «О программе»")

    def reset_sound_file(self):
        """Сбрасывает звуковой файл на стандартный Sound_1.MP3."""

        self.edit_sound_file.setText("Sound/Sound_1.MP3")


    def open_plugins_folder(self):
        parent = self.parent()
        if parent and hasattr(parent, 'open_plugins_folder'):
            parent.open_plugins_folder()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к папке плагинов")

    def show_plugins_manager(self):
        parent = self.parent()
        if parent and hasattr(parent, 'show_plugins_manager'):
            parent.show_plugins_manager()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к управлению плагинами")

    def on_import_zip_clicked(self):
        parent = self.parent()
        if parent and hasattr(parent, 'import_zip_archive'):
            self.done(QDialog.Accepted)  # закрываем настройки
            parent.import_zip_archive()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции импорта")

    def open_hotkeys_dialog(self):
        dialog = HotkeysDialog(self.settings.get("hotkeys", {}), self)
        if dialog.exec() == QDialog.Accepted:
            self.settings["hotkeys"] = dialog.get_hotkeys()

    def on_launch_lm_clicked(self):
        parent = self.parent()
        if parent and hasattr(parent, 'launch_lm_studio'):
            parent.launch_lm_studio()
        else:
            QMessageBox.warning(self, "Ошибка", "Нет доступа к функции запуска LM Studio")

    def browse_lm_studio(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите LM Studio", "", "Executable (*.exe);;All Files (*)")
        if file_path:
            self.edit_lm_path.setText(file_path)

    def browse_sound_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выберите звуковой файл", "",
                                                   "Звуковые файлы (*.mp3 *.wav *.ogg)")
        if file_path:
            self.edit_sound_file.setText(file_path)

    def get_settings(self) -> Dict[str, Any]:
        return {
            "api_base": self.edit_api_base.text().strip(),
            "model": self.edit_model.text().strip(),
            "max_tokens": self.spin_max_tokens.value(),
            "temperature": self.spin_temperature.value(),
            "top_p": self.spin_top_p.value(),
            "top_k": self.spin_top_k.value(),
            "repeat_penalty": self.spin_repeat.value(),
            "stream": self.check_stream.isChecked(),
            "max_context_tokens": self.spin_max_ctx.value(),
            "system_prompt": self.edit_system.toPlainText(),
            "multimodal": self.check_multimodal.isChecked(),
            "open_last_chat_on_startup": self.check_open_last.isChecked(),
            "auto_backup": self.check_auto_backup.isChecked(),
            "auto_backup_interval_min": self.spin_backup_interval.value(),
            "max_backups": self.spin_max_backups.value(),
            "backup_on_start": True,
            "lm_studio_path": self.edit_lm_path.text().strip(),
            "lm_studio_auto_start": self.check_auto_start.isChecked(),
            "encrypt_chats": self.check_encrypt.isChecked(),
            "max_images_per_message": self.spin_max_images.value(),
            "chat_font_family": self.combo_chat_font.currentFont().family(),
            "chat_font_size": self.spin_chat_font_size.value(),
            "input_font_family": self.combo_input_font.currentFont().family(),
            "input_font_size": self.spin_input_font_size.value(),
            "notifications_enabled": not self.check_notifications.isChecked(),
            "auto_archive_days": self.spin_auto_archive.value(),
            "rag_enabled": self.check_rag.isChecked(),
            "max_prompt_chars": self.spin_max_prompt_chars.value(),
            "sound_notifications": self.check_sound_notifications.isChecked(),
            "toast_notifications": self.check_toast_notifications.isChecked(),
            "max_attachment_size_mb": self.spin_max_attachment_size.value(),
            "sound_file": self.edit_sound_file.text().strip(),
            "hotkeys": self.settings.get("hotkeys", {})
        }


class ChatParamsDialog(QDialog):
    """Диалог настройки параметров генерации и системного промпта для конкретного чата."""

    def __init__(self, chat_params: dict, global_settings: dict, current_system_prompt: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметры генерации и системный промпт чата")
        self.setModal(True)
        self.chat_params = chat_params.copy()
        self.global_settings = global_settings
        self.current_system_prompt = current_system_prompt
        self._reset_clicked = False
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()

        # ---------- Параметры генерации ----------
        self.spin_max_tokens = QSpinBox()
        self.spin_max_tokens.setRange(1, 100000)
        self.spin_temperature = QDoubleSpinBox()
        self.spin_temperature.setRange(0.0, 2.0)
        self.spin_temperature.setSingleStep(0.05)
        self.spin_top_p = QDoubleSpinBox()
        self.spin_top_p.setRange(0.0, 1.0)
        self.spin_top_p.setSingleStep(0.05)
        self.spin_top_k = QSpinBox()
        self.spin_top_k.setRange(0, 200)
        self.spin_repeat = QDoubleSpinBox()
        self.spin_repeat.setRange(0.0, 2.0)
        self.spin_repeat.setSingleStep(0.05)

        self.spin_max_tokens.setValue(
            int(self.chat_params.get('max_tokens', self.global_settings.get('max_tokens', 4096))))
        self.spin_temperature.setValue(
            float(self.chat_params.get('temperature', self.global_settings.get('temperature', 0.7))))
        self.spin_top_p.setValue(float(self.chat_params.get('top_p', self.global_settings.get('top_p', 0.9))))
        self.spin_top_k.setValue(int(self.chat_params.get('top_k', self.global_settings.get('top_k', 40))))
        self.spin_repeat.setValue(
            float(self.chat_params.get('repeat_penalty', self.global_settings.get('repeat_penalty', 1.1))))

        form.addRow("Max Tokens:", self.spin_max_tokens)
        form.addRow("Temperature:", self.spin_temperature)
        form.addRow("Top P:", self.spin_top_p)
        form.addRow("Top K:", self.spin_top_k)
        form.addRow("Repeat Penalty:", self.spin_repeat)

        # ---------- Системный промпт ----------
        self.edit_system_prompt = QTextEdit()
        self.edit_system_prompt.setMaximumHeight(120)
        # Если у чата есть свой промпт — используем его, иначе глобальный
        if self.current_system_prompt:
            self.edit_system_prompt.setPlainText(self.current_system_prompt)
        else:
            self.edit_system_prompt.setPlainText(self.global_settings.get('system_prompt', ''))

        form.addRow("System Prompt:", self.edit_system_prompt)

        layout.addLayout(form)

        # Кнопка сброса к глобальным
        btn_reset = QPushButton("Сбросить к глобальным")
        btn_reset.clicked.connect(self.reset_to_global)
        layout.addWidget(btn_reset)

        # Кнопки ОК/Отмена
        btn_box = QHBoxLayout()
        btn_ok = QPushButton("Сохранить")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addStretch()
        btn_box.addWidget(btn_ok)
        btn_box.addWidget(btn_cancel)
        layout.addLayout(btn_box)

    def reset_to_global(self):
        self.spin_max_tokens.setValue(int(self.global_settings.get('max_tokens', 4096)))
        self.spin_temperature.setValue(float(self.global_settings.get('temperature', 0.7)))
        self.spin_top_p.setValue(float(self.global_settings.get('top_p', 0.9)))
        self.spin_top_k.setValue(int(self.global_settings.get('top_k', 40)))
        self.spin_repeat.setValue(float(self.global_settings.get('repeat_penalty', 1.1)))
        self.edit_system_prompt.setPlainText(self.global_settings.get('system_prompt', ''))
        self._reset_clicked = True

    def get_params(self) -> dict:
        if self._reset_clicked:
            return {'params': {}, 'system_prompt': ''}
        return {
            'params': {
                'max_tokens': self.spin_max_tokens.value(),
                'temperature': self.spin_temperature.value(),
                'top_p': self.spin_top_p.value(),
                'top_k': self.spin_top_k.value(),
                'repeat_penalty': self.spin_repeat.value(),
            },
            'system_prompt': self.edit_system_prompt.toPlainText().strip()
        }


class HotkeysDialog(QDialog):
    """Диалог настройки горячих клавиш.
    Позволяет изменить комбинации для основных действий."""

    ACTIONS = {
        "new_chat": "Новый чат",
        "new_folder": "Новая папка",
        "close_chat": "Закрыть чат",
        "search": "Поиск",
        "import_deepseek": "Импорт DeepSeek",
        "export": "Экспорт чата",
        "settings": "Открыть настройки",
        "send_message": "Отправить сообщение",
        "toggle_theme": "Переключить тему",
        "toggle_sidebar": "Показать/скрыть боковую панель",
        "clear_chat": "Очистить историю чата",
        "starred_messages": "Показать закладки",
        "statistics": "Показать статистику",
        "prompts": "Показать промпты",
        "focus_input": "Фокус на поле ввода",
        "next_chat": "Следующий чат",
        "prev_chat": "Предыдущий чат",
        "chat_settings": "Параметры чата",
    }

    DEFAULTS = {
        "new_chat": "Ctrl+N",
        "new_folder": "Ctrl+Shift+N",
        "close_chat": "Ctrl+W",
        "search": "Ctrl+F",
        "import_deepseek": "Ctrl+I",
        "export": "Ctrl+E",
        "settings": "Ctrl+P",
        "send_message": "Ctrl+Return",
        "toggle_theme": "Ctrl+T",
        "toggle_sidebar": "Ctrl+B",
        "clear_chat": "Ctrl+Shift+L",
        "starred_messages": "Ctrl+Shift+B",
        "statistics": "Ctrl+Shift+S",
        "prompts": "Ctrl+Alt+P",
        "focus_input": "Ctrl+L",
        "next_chat": "Ctrl+Tab",
        "prev_chat": "Ctrl+Shift+Tab",
        "chat_settings": "Ctrl+Shift+P",
    }

    def __init__(self, hotkeys, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Command.png"))
        self.setWindowTitle("Настройка горячих клавиш")
        self.setModal(True)
        self.hotkeys = hotkeys.copy()
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.seq_edits = {}
        for action, label in self.ACTIONS.items():
            cur_seq = self.hotkeys.get(action, self.DEFAULTS.get(action, ""))
            edit = QKeySequenceEdit(QKeySequence(cur_seq))
            self.seq_edits[action] = edit
            form.addRow(label, edit)
        layout.addLayout(form)

        btn_box = QHBoxLayout()
        btn_reset = QPushButton("Сбросить к стандартным")
        btn_reset.clicked.connect(self.reset_defaults)
        btn_save = QPushButton("Сохранить")
        btn_save.clicked.connect(self.accept)
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(btn_reset)
        btn_box.addStretch()
        btn_box.addWidget(btn_save)
        btn_box.addWidget(btn_cancel)
        layout.addLayout(btn_box)

    def reset_defaults(self):
        for action, edit in self.seq_edits.items():
            edit.setKeySequence(QKeySequence(self.DEFAULTS[action]))

    def get_hotkeys(self):
        return {action: edit.keySequence().toString() for action, edit in self.seq_edits.items()}


class MessagePreviewDialog(QDialog):
    """Диалог предпросмотра вложения или отправляемого сообщения.
    Отображает текст и изображения в едином HTML-представлении."""

    def __init__(self, text, attachments, parent=None, show_send_button=True):
        super().__init__(parent)
        self.setWindowTitle("Предпросмотр")
        self.setModal(True)
        self.resize(700, 700)
        self.show_send_button = show_send_button
        self.init_ui(text, attachments)

    def init_ui(self, text, attachments):
        # Определяем, используется ли тёмная тема
        dark_theme = False
        parent = self.parent()
        if parent and hasattr(parent, 'settings') and parent.settings.get("theme", "light") == "dark":
            dark_theme = True

        layout = QVBoxLayout(self)

        # Единая прокручиваемая область
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(False)
        self.browser.setReadOnly(True)
        layout.addWidget(self.browser)

        html_parts = []

        # Текст сообщения
        if text.strip():
            html_parts.append(f'<p style="white-space:pre-wrap;">{self._escape_html(text)}</p>')

        # Вложения
        for att in attachments:
            att_type = att.get('type')
            if att_type == 'image':
                img_html = self._image_to_html(att.get('path'))
                if img_html:
                    html_parts.append(img_html)
            elif att_type == 'file_text':
                name = att.get('name', 'Файл')
                content = att.get('content', '')
                images = att.get('images', [])

                # Стили для тёмной темы
                title_style = 'font-weight:bold;'
                content_style = 'background-color:#f9f9f9; padding:8px; border-radius:6px; white-space:pre-wrap;'
                if dark_theme:
                    title_style += ' color: white;'
                    content_style = 'background-color:#3c3c3c; padding:8px; border-radius:6px; white-space:pre-wrap; color: white;'

                if content:
                    html_parts.append(f'<p style="{title_style}">📄 {self._escape_html(name)}</p>')
                    html_parts.append(f'<div style="{content_style}">{self._escape_html(content)}</div>')
                else:
                    html_parts.append(f'<p style="{title_style}">📄 {self._escape_html(name)}</p>')

                for img_data in images:
                    img_html = self._data_uri_to_html(img_data)
                    if img_html:
                        html_parts.append(img_html)
            elif att_type == 'file':
                name = att.get('name', 'Файл')
                path = att.get('path', '')
                file_style = ''
                if dark_theme:
                    file_style = 'color: white;'
                html_parts.append(
                    f'<p style="{file_style}">📁 {self._escape_html(name)} <span style="color:gray;">({self._escape_html(path)})</span></p>')

        if not html_parts:
            html_parts.append('<p style="color:gray;">Нет данных</p>')

        full_html = '<html><body>' + ''.join(html_parts) + '</body></html>'
        self.browser.setHtml(full_html)

        # Кнопки
        btn_box = QHBoxLayout()
        if self.show_send_button:
            btn_send = QPushButton("Отправить")
            btn_send.clicked.connect(self.accept)
            btn_cancel = QPushButton("Отмена")
            btn_cancel.clicked.connect(self.reject)
            btn_box.addStretch()
            btn_box.addWidget(btn_send)
            btn_box.addWidget(btn_cancel)
        else:
            btn_close = QPushButton("Закрыть")
            btn_close.clicked.connect(self.accept)
            btn_box.addStretch()
            btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def _escape_html(self, text):
        return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

    def _image_to_html(self, path, max_width=400):
        if not path or not os.path.exists(path):
            return None
        try:
            import base64
            with open(path, 'rb') as f:
                img_bytes = f.read()
            mime = "image/jpeg" if path.lower().endswith(('.jpg', '.jpeg')) else "image/png"
            data_uri = f"data:{mime};base64," + base64.b64encode(img_bytes).decode('utf-8')
            return f'<img src="{data_uri}" style="max-width:{max_width}px; margin:5px; border-radius:5px;" />'
        except Exception as e:
            print(f"Ошибка загрузки изображения {path}: {e}")
            return None

    def _data_uri_to_html(self, data_uri, max_width=400):
        if not data_uri or not data_uri.startswith('data:image'):
            return None
        return f'<img src="{data_uri}" style="max-width:{max_width}px; margin:5px; border-radius:5px;" />'


class ImagePreviewDialog(QDialog):
    """Полноэкранный просмотр изображения (для вложений-картинок)."""

    def __init__(self, pixmap: QPixmap, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Просмотр изображения")
        self.setModal(True)
        self.resize(800, 800)

        layout = QVBoxLayout(self)
        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.label.setMinimumSize(1, 1)
        layout.addWidget(self.label)

        self.pixmap = pixmap
        if self.pixmap.isNull():
            self.label.setText("Ошибка загрузки изображения")
        else:
            self.update_pixmap()

        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)
        btn_box = QHBoxLayout()
        btn_box.addStretch()
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_pixmap()

    def update_pixmap(self):
        if hasattr(self, 'pixmap') and not self.pixmap.isNull():
            scaled = self.pixmap.scaled(
                self.label.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.label.setPixmap(scaled)


class RAGManagerDialog(QDialog):
    """Диалог просмотра и управления базой знаний RAG."""

    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.rag_folder = main_window.rag_folder
        self.setWindowTitle("База знаний (RAG)")
        self.setModal(True)
        self.resize(700, 500)
        self.init_ui()
        self.populate_list()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Поле поиска
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Поиск:"))
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Введите текст для поиска в документах...")
        self.search_edit.textChanged.connect(self.on_search_changed)
        search_layout.addWidget(self.search_edit)
        layout.addLayout(search_layout)

        # Список документов
        self.list_widget = QListWidget()
        self.list_widget.currentItemChanged.connect(self.on_item_selected)
        layout.addWidget(self.list_widget, stretch=2)

        # Предпросмотр
        self.preview_edit = QTextEdit()
        self.preview_edit.setReadOnly(True)
        self.preview_edit.setPlaceholderText("Выберите документ для предпросмотра")
        layout.addWidget(self.preview_edit, stretch=3)

        # Кнопки
        btn_box = QHBoxLayout()
        btn_open_folder = QPushButton("Открыть папку")
        btn_open_folder.clicked.connect(self.open_folder)
        btn_delete = QPushButton("Удалить файл")
        btn_delete.clicked.connect(self.delete_selected)
        btn_refresh = QPushButton("Обновить")
        btn_refresh.clicked.connect(self.populate_list)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)

        btn_box.addWidget(btn_open_folder)
        btn_box.addWidget(btn_delete)
        btn_box.addWidget(btn_refresh)
        btn_box.addStretch()
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def populate_list(self):
        """Сканирует папку RAG_folder и заполняет список документов."""

        self.list_widget.clear()
        self.preview_edit.clear()

        if not os.path.isdir(self.rag_folder):
            self.list_widget.addItem("Папка RAG_folder не существует")
            return

        files = [f for f in os.listdir(self.rag_folder)
                 if os.path.isfile(os.path.join(self.rag_folder, f))]
        if not files:
            self.list_widget.addItem("Нет документов в базе знаний")
            return

        for filename in sorted(files):
            filepath = os.path.join(self.rag_folder, filename)
            try:
                stat = os.stat(filepath)
                size_kb = stat.st_size / 1024
                mtime = datetime.fromtimestamp(stat.st_mtime).strftime('%d.%m.%Y %H:%M:%S')
                ext = os.path.splitext(filename)[1].lower()
                supported = ext in ('.txt', '.md', '.py', '.json', '.csv',
                                    '.html', '.log', '.pdf', '.docx')
                if not supported:
                    continue  # пропускаем неподдерживаемые

                # Читаем текст (используя метод главного окна)
                text = self.main_window.read_file_content(filepath) or ""

                # Вычисляем количество чанков (размер чанка 1000 символов)
                chunk_size = 1000
                chunks = (len(text) + chunk_size - 1) // chunk_size if text else 0

                item = QListWidgetItem(f"{filename}  |  {size_kb:.1f} КБ  |  {chunks} чанков  |  {mtime}")
                item.setData(Qt.UserRole, {
                    'path': filepath,
                    'text': text,
                    'chunks': chunks,
                    'size_kb': size_kb,
                    'mtime': mtime
                })
                self.list_widget.addItem(item)
            except Exception as e:
                self.main_window.logger.warning(f"Ошибка обработки {filename}: {e}")

    def on_item_selected(self, current, previous):
        """Показывает предпросмотр выбранного документа."""

        if current is None:
            self.preview_edit.clear()
            return
        data = current.data(Qt.UserRole)
        if data:
            text = data.get('text', '')
            # Показываем первые 2000 символов
            preview = text[:2000] + ("\n... (обрезано)" if len(text) > 2000 else "")
            self.preview_edit.setPlainText(preview)
        else:
            self.preview_edit.clear()

    def on_search_changed(self, query):
        """Фильтрует список документов по содержимому."""

        query = query.strip().lower()
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            data = item.data(Qt.UserRole)
            if data is None:
                continue  # заглушка "нет документов"
            text = data.get('text', '').lower()
            filename = os.path.basename(data.get('path', '')).lower()
            matched = query in text or query in filename
            item.setHidden(not matched)

    def open_folder(self):
        """Открывает папку RAG_folder в проводнике."""

        self.main_window.open_rag_folder()

    def delete_selected(self):
        """Удаляет выбранный файл из базы знаний."""

        item = self.list_widget.currentItem()
        if item is None:
            return
        data = item.data(Qt.UserRole)
        if not data:
            return

        filepath = data['path']
        filename = os.path.basename(filepath)
        reply = QMessageBox.question(
            self,
            "Удалить документ",
            f"Удалить файл «{filename}» из базы знаний?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return

        try:
            os.remove(filepath)
            self.main_window.logger.info(f"Документ удалён из RAG_folder: {filename}")
            # Перезагружаем базу RAG, если включена
            if self.main_window.rag_enabled:
                self.main_window.load_rag_documents_from_folder()
            # Обновляем список
            self.populate_list()
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Не удалось удалить файл.\nПричина: {e}")


class HelpDialog(QDialog):
    """Диалог встроенной справки: показывает список файлов из папки Documentation."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Info.png"))
        self.setWindowTitle("Справка Echos")
        self.setModal(True)
        self.resize(500, 400)
        self.docs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Documentation")
        self.init_ui()
        self.populate_list()

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.list_widget = QListWidget()
        self.list_widget.itemDoubleClicked.connect(self.open_selected)
        layout.addWidget(self.list_widget)

        btn_box = QHBoxLayout()
        btn_open = QPushButton("Открыть")
        btn_open.clicked.connect(self.open_selected)
        btn_open_folder = QPushButton("Открыть папку")
        btn_open_folder.clicked.connect(self.open_folder)
        btn_close = QPushButton("Закрыть")
        btn_close.clicked.connect(self.accept)

        btn_box.addWidget(btn_open)
        btn_box.addWidget(btn_open_folder)
        btn_box.addStretch()
        btn_box.addWidget(btn_close)
        layout.addLayout(btn_box)

    def populate_list(self):
        """Сканирует папку Documentation и заполняет список файлов."""

        self.list_widget.clear()
        if not os.path.isdir(self.docs_dir):
            item = QListWidgetItem("Папка документации не найдена")
            item.setFlags(Qt.NoItemFlags)
            self.list_widget.addItem(item)
            return

        files = [f for f in os.listdir(self.docs_dir)
                 if os.path.isfile(os.path.join(self.docs_dir, f))
                 and f.lower().endswith(('.md', '.txt', '.html', '.htm'))]
        if not files:
            item = QListWidgetItem("Нет доступных документов справки")
            item.setFlags(Qt.NoItemFlags)
            self.list_widget.addItem(item)
            return

        for filename in sorted(files):
            item = QListWidgetItem(filename)
            item.setData(Qt.UserRole, os.path.join(self.docs_dir, filename))
            self.list_widget.addItem(item)

    def open_selected(self):
        """Открывает выбранный файл справки в системном приложении."""

        item = self.list_widget.currentItem()
        if not item:
            return
        path = item.data(Qt.UserRole)
        if path and os.path.exists(path):
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def open_folder(self):
        """Открывает папку Documentation в проводнике."""

        if os.path.isdir(self.docs_dir):
            QDesktopServices.openUrl(QUrl.fromLocalFile(self.docs_dir))


class SchemeViewerDialog(QDialog):
    """Окно просмотра схемы с возможностью копирования."""

    def __init__(self, pixmap: QPixmap, parent=None):
        super().__init__(parent)
        self.setWindowIcon(QIcon("Image/Scheme.png"))
        self.setWindowTitle("Схема")
        self.setModal(True)
        self.resize(800, 600)

        self.pixmap = pixmap

        layout = QVBoxLayout(self)

        self.label = QLabel()
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setPixmap(pixmap)
        self.label.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)
        self.label.setScaledContents(False)
        self.label.setContextMenuPolicy(Qt.CustomContextMenu)
        self.label.customContextMenuRequested.connect(self.show_context_menu)
        layout.addWidget(self.label)

        # Горячая клавиша Ctrl+C для копирования
        copy_shortcut = QKeySequence("Ctrl+C")
        copy_action = QAction(self)
        copy_action.setShortcut(copy_shortcut)
        copy_action.triggered.connect(self.copy_to_clipboard)
        self.addAction(copy_action)

    def show_context_menu(self, pos):
        menu = QMenu(self)
        copy_action = menu.addAction("Копировать схему")
        copy_action.triggered.connect(self.copy_to_clipboard)

        save_action = menu.addAction("Сохранить как...")
        save_action.triggered.connect(self.save_as_file)

        menu.exec(self.label.mapToGlobal(pos))

    def copy_to_clipboard(self):
        """Копирует изображение в буфер обмена."""

        if not self.pixmap.isNull():
            QApplication.clipboard().setPixmap(self.pixmap)
            self.setWindowTitle("Схема — скопировано")
            QTimer.singleShot(1500, lambda: self.setWindowTitle("Схема"))

    def save_as_file(self):
        """Сохраняет схему как PNG-файл."""

        if self.pixmap.isNull():
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить схему", "scheme.png", "PNG (*.png);;JPG (*.jpg)"
        )
        if file_path:
            self.pixmap.save(file_path)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not self.pixmap.isNull():
            scaled = self.pixmap.scaled(
                self.label.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
            self.label.setPixmap(scaled)


def main():
    app = QApplication(sys.argv)

    # Установка AppUserModelID только для Windows
    if os.name == 'nt':
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myapp.1.0")
        except Exception:
            pass

    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Icon/Icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))
    else:
        print(f"Иконка не найдена: {icon_path}")

    # === Читает тему из настроек ДО создания сплэша ===
    is_dark = False
    config_file = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "chat_config.json"
    )
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            is_dark = (saved.get("theme", "light") == "dark")
        except Exception as e:
            print(f"Не удалось прочитать тему: {e}")

    # Показываем сплэш с нужной темой
    splash = LoadingSplash(dark=is_dark)
    splash.show()
    app.processEvents()

    # Создаём главное окно (загрузка идёт синхронно, сплэш обновляется)
    window = MainWindow(splash=splash)

    # Минимальное время показа сплэша, чтобы не мелькал
    QTimer.singleShot(300, lambda: (splash.close(), window.show()))

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
