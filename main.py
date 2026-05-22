#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import os
import ctypes
import markdown
try:
    from docx import Document
except Exception:
    Document = None
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QSplitter, QTextEdit,
    QFileDialog, QMessageBox, QMenuBar, QMenu,
    QToolBar, QStatusBar, QWidget, QVBoxLayout,
    QTabWidget, QDialog, QPushButton
)
from PyQt6.QtGui import QAction, QFont, QSyntaxHighlighter, QTextCharFormat, QColor, QIcon
from PyQt6.QtCore import Qt, QFileInfo


# ==================== Markdown 语法高亮器 ====================
class MarkdownHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.rules = []

        # 标题 (# ...)
        title_format = QTextCharFormat()
        title_format.setForeground(QColor(0, 120, 200))
        title_format.setFontWeight(QFont.Weight.Bold)
        self.rules.append((r"^#{1,6}\s+.*$", title_format))

        # 粗体 (**bold**)
        bold_format = QTextCharFormat()
        bold_format.setFontWeight(QFont.Weight.Bold)
        self.rules.append((r"\*\*[^*]+\*\*", bold_format))

        # 斜体 (*italic*)
        italic_format = QTextCharFormat()
        italic_format.setFontItalic(True)
        self.rules.append((r"\*[^*]+\*", italic_format))

        # 行内代码 (`code`)
        code_format = QTextCharFormat()
        code_format.setForeground(QColor(150, 100, 50))
        code_format.setFont(QFont("Courier New"))
        self.rules.append((r"`[^`]+`", code_format))

        # 链接 [text](url)
        link_format = QTextCharFormat()
        link_format.setForeground(QColor(0, 150, 0))
        link_format.setFontUnderline(True)
        self.rules.append((r"\[[^\]]+\]\([^\)]+\)", link_format))

    def highlightBlock(self, text):
        import re
        for pattern, fmt in self.rules:
            for match in re.finditer(pattern, text):
                start, end = match.span()
                self.setFormat(start, end - start, fmt)


# ==================== 自定义编辑器（带语法高亮和文件信息） ====================
class CodeEditor(QTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_path = None
        self.highlighter = MarkdownHighlighter(self.document())
        self.setFont(QFont("Consolas", 10))
        self.setTabStopDistance(4 * self.fontMetrics().horizontalAdvance(' '))

    def set_file_path(self, path):
        self.file_path = path

    def load_file(self, path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                self.setText(f.read())
            self.file_path = path
            self.document().setModified(False)
            return True
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法打开文件：{str(e)}")
            return False

    def save_to_path(self, path):
        try:
            with open(path, 'w', encoding='utf-8') as f:
                f.write(self.toPlainText())
            self.file_path = path
            self.document().setModified(False)
            return True
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法保存文件：{str(e)}")
            return False


# ==================== 可感知鼠标进入/离开的容器（用于悬停放大） ====================
class HoverWidget(QWidget):
    def __init__(self, widget, on_enter, on_leave, parent=None):
        super().__init__(parent)
        self.on_enter_callback = on_enter
        self.on_leave_callback = on_leave
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(widget)
        self.setLayout(layout)
        self.widget = widget
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)

    def enterEvent(self, event):
        self.on_enter_callback(self)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.on_leave_callback(self)
        super().leaveEvent(event)


# ==================== 主窗口 ====================
class MarkdownEditor(QMainWindow):
    def __init__(self, initial_paths=None):
        super().__init__()
        icon_path = get_icon_path()
        if icon_path:
            self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle("TML Markdown 编辑器")
        self.resize(700, 600)
        self.setAcceptDrops(True)

        self.split_enabled = False
        self.hovered_side = None
        self._syncing_scroll = False  # 防止循环同步

        # 中央分割器
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.setCentralWidget(self.splitter)

        # 左侧预览区
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setFont(QFont("Segoe UI", 10))
        self.left_container = HoverWidget(
            self.preview,
            self.on_hover_enter,
            self.on_hover_leave,
            self.splitter
        )

        # 右侧多标签页编辑区
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabsClosable(True)
        self.tab_widget.tabCloseRequested.connect(self.close_tab)
        self.tab_widget.currentChanged.connect(self.on_current_tab_changed)
        self.right_container = HoverWidget(
            self.tab_widget,
            self.on_hover_enter,
            self.on_hover_leave,
            self.splitter
        )

        self.splitter.addWidget(self.left_container)
        self.splitter.addWidget(self.right_container)

        # 初始单屏模式
        self.left_container.hide()
        self.splitter.setSizes([0, self.width()])

        self.create_menu_bar()

        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

        if not self.open_file_paths(initial_paths or []):
            self.new_tab()

    # ========== 滚动同步逻辑（百分比映射） ==========
    def sync_scroll(self, source_scrollbar, target_scrollbar):
        """根据源滚动条百分比设置目标滚动条相同百分比位置"""
        if self._syncing_scroll or not self.split_enabled:
            return
        self._syncing_scroll = True
        src_min = source_scrollbar.minimum()
        src_max = source_scrollbar.maximum()
        src_val = source_scrollbar.value()
        if src_max > src_min:
            percent = (src_val - src_min) / (src_max - src_min)
            tgt_min = target_scrollbar.minimum()
            tgt_max = target_scrollbar.maximum()
            tgt_val = int(tgt_min + percent * (tgt_max - tgt_min))
            target_scrollbar.setValue(tgt_val)
        self._syncing_scroll = False

    def connect_scroll_sync(self):
        """连接当前编辑区和预览区的滚动同步信号"""
        editor = self.current_editor()
        if editor and self.split_enabled:
            editor_vscroll = editor.verticalScrollBar()
            preview_vscroll = self.preview.verticalScrollBar()
            # 断开旧连接避免重复
            try:
                editor_vscroll.valueChanged.disconnect(self.on_editor_scroll)
            except TypeError:
                pass
            try:
                preview_vscroll.valueChanged.disconnect(self.on_preview_scroll)
            except TypeError:
                pass
            editor_vscroll.valueChanged.connect(self.on_editor_scroll)
            preview_vscroll.valueChanged.connect(self.on_preview_scroll)
            # 初始同步一次
            self.sync_scroll(editor_vscroll, preview_vscroll)

    def disconnect_scroll_sync(self):
        """断开滚动同步"""
        editor = self.current_editor()
        if editor:
            try:
                editor.verticalScrollBar().valueChanged.disconnect(self.on_editor_scroll)
            except TypeError:
                pass
        try:
            self.preview.verticalScrollBar().valueChanged.disconnect(self.on_preview_scroll)
        except TypeError:
            pass

    def on_editor_scroll(self, value):
        if not self.split_enabled:
            return
        editor = self.current_editor()
        if editor:
            self.sync_scroll(editor.verticalScrollBar(), self.preview.verticalScrollBar())

    def on_preview_scroll(self, value):
        if not self.split_enabled:
            return
        editor = self.current_editor()
        if editor:
            self.sync_scroll(self.preview.verticalScrollBar(), editor.verticalScrollBar())

    # ========== 悬停放大逻辑 ==========
    def on_hover_enter(self, hover_widget):
        if not self.split_enabled:
            return
        if hover_widget == self.left_container:
            self.hovered_side = "left"
        elif hover_widget == self.right_container:
            self.hovered_side = "right"
        self.adjust_splitter_sizes()

    def on_hover_leave(self, hover_widget):
        if not self.split_enabled:
            return
        self.hovered_side = None
        self.adjust_splitter_sizes()

    def adjust_splitter_sizes(self):
        if not self.split_enabled:
            return
        total = self.splitter.width()
        if total <= 0:
            return
        if self.hovered_side == "left":
            self.splitter.setSizes([int(total * 0.7), int(total * 0.3)])
        elif self.hovered_side == "right":
            self.splitter.setSizes([int(total * 0.3), int(total * 0.7)])
        else:
            self.splitter.setSizes([total // 2, total // 2])

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if self.split_enabled:
            self.adjust_splitter_sizes()

    # ========== 分屏模式切换 ==========
    def set_split_mode(self, enabled):
        self.split_enabled = enabled
        if enabled:
            self.left_container.show()
            total = self.splitter.width()
            if total > 0:
                self.splitter.setSizes([total // 2, total // 2])
            self.update_preview()
            self.connect_scroll_sync()
            self.status_bar.showMessage("分屏模式已开启（滚动百分比同步）")
        else:
            self.disconnect_scroll_sync()
            self.left_container.hide()
            self.splitter.setSizes([0, self.splitter.width()])
            self.status_bar.showMessage("分屏模式已关闭")
        self.split_action.setChecked(enabled)

    def toggle_split_mode(self):
        self.set_split_mode(not self.split_enabled)

    # ========== 多标签页管理 ==========
    def new_tab(self, file_path=None, content=""):
        editor = CodeEditor()
        if file_path:
            if not editor.load_file(file_path):
                return None
            tab_name = os.path.basename(file_path)
        else:
            editor.setPlainText(content)
            tab_name = "未命名"
        index = self.tab_widget.addTab(editor, tab_name)
        self.tab_widget.setCurrentIndex(index)
        editor.textChanged.connect(lambda: self.on_editor_text_changed(editor))
        editor.document().modificationChanged.connect(
            lambda modified: self.update_tab_title(editor, modified)
        )
        if self.split_enabled:
            self.connect_scroll_sync()
        return editor

    def open_file_paths(self, paths):
        valid_paths = [path for path in paths if path and os.path.isfile(path)]
        if not valid_paths:
            return False

        opened_any = False
        for path in valid_paths:
            editor = self.new_tab(file_path=path)
            if editor is None:
                continue
            opened_any = True
            self.status_bar.showMessage(f"已打开：{path}")

        if opened_any:
            current = self.current_editor()
            if current and current.file_path:
                ext = QFileInfo(current.file_path).suffix().lower()
                is_md = ext in ["md", "markdown"]
                if is_md != self.split_enabled:
                    self.set_split_mode(is_md)
                elif self.split_enabled:
                    self.update_preview()

        return opened_any

    def close_tab(self, index):
        editor = self.tab_widget.widget(index)
        if editor and editor.document().isModified():
            ret = QMessageBox.question(
                self, "未保存",
                f"文档“{self.tab_widget.tabText(index)}”已修改，是否保存？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )
            if ret == QMessageBox.StandardButton.Yes:
                if not self.save_current_editor(editor):
                    return
            elif ret == QMessageBox.StandardButton.Cancel:
                return
        self.tab_widget.removeTab(index)
        if self.tab_widget.count() == 0:
            self.new_tab()
        elif self.split_enabled:
            self.connect_scroll_sync()

    def update_tab_title(self, editor, modified):
        index = self.tab_widget.indexOf(editor)
        if index == -1:
            return
        base_name = os.path.basename(editor.file_path) if editor.file_path else "未命名"
        title = base_name + " *" if modified else base_name
        self.tab_widget.setTabText(index, title)

    def on_editor_text_changed(self, editor):
        if self.split_enabled and editor == self.current_editor():
            self.update_preview()

    def on_current_tab_changed(self, index):
        editor = self.current_editor()
        if editor:
            path = editor.file_path
            if path:
                self.setWindowTitle(f"TML Markdown 编辑器 - {os.path.basename(path)}")
                ext = QFileInfo(path).suffix().lower()
                is_md = ext in ["md", "markdown"]
                if is_md != self.split_enabled:
                    # 自动切换分屏模式
                    self.split_enabled = is_md
                    if is_md:
                        self.left_container.show()
                        total = self.splitter.width()
                        self.splitter.setSizes([total // 2, total // 2])
                        self.update_preview()
                        self.connect_scroll_sync()
                    else:
                        self.disconnect_scroll_sync()
                        self.left_container.hide()
                        self.splitter.setSizes([0, self.splitter.width()])
                    self.split_action.setChecked(is_md)
            else:
                self.setWindowTitle("TML Markdown 编辑器")
            if self.split_enabled:
                self.connect_scroll_sync()
            self.status_bar.showMessage(f"当前文件：{path if path else '未保存'}")

    def current_editor(self):
        return self.tab_widget.currentWidget()

    def save_current_editor(self, editor=None):
        if editor is None:
            editor = self.current_editor()
        if not editor:
            return False
        if editor.file_path:
            return editor.save_to_path(editor.file_path)
        else:
            return self.save_as_current_editor(editor)

    def save_as_current_editor(self, editor=None):
        if editor is None:
            editor = self.current_editor()
        if not editor:
            return False
        path, _ = QFileDialog.getSaveFileName(
            self, "保存文件", "",
            "Markdown 文件 (*.md);;文本文件 (*.txt);;LaTeX 文件 (*.tex);;所有文件 (*)"
        )
        if path:
            if editor.save_to_path(path):
                self.update_tab_title(editor, False)
                ext = QFileInfo(path).suffix().lower()
                if ext in ["md", "markdown"] and not self.split_enabled:
                    reply = QMessageBox.question(self, "建议", "是否开启分屏模式以预览 Markdown？",
                                                 QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                    if reply == QMessageBox.StandardButton.Yes:
                        self.set_split_mode(True)
                return True
        return False

    # ========== 文件操作 ==========
    def new_file(self):
        self.new_tab()

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "打开文件", "",
            "文本文件 (*.txt *.md *.markdown *.tex);;所有文件 (*)"
        )
        if path:
            self.open_file_paths([path])

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() and any(url.isLocalFile() for url in event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event):
        paths = [url.toLocalFile() for url in event.mimeData().urls() if url.isLocalFile()]
        if self.open_file_paths(paths):
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def save_file(self):
        editor = self.current_editor()
        if editor:
            if editor.file_path:
                editor.save_to_path(editor.file_path)
                self.update_tab_title(editor, False)
            else:
                self.save_as_current_editor(editor)

    def save_as_file(self):
        self.save_as_current_editor()

    # ========== 预览更新 ==========
    def update_preview(self):
        editor = self.current_editor()
        if not editor or not self.split_enabled:
            return
        md_text = editor.toPlainText()
        try:
            html = markdown.markdown(
                md_text,
                extensions=['extra', 'codehilite', 'tables', 'fenced_code']
            )
        except Exception as e:
            html = f"<p>渲染错误：{str(e)}</p>"
        css = """
        <style>
            body {
                font-family: 'Segoe UI', 'Roboto', sans-serif;
                font-size: 14px;
                line-height: 1.6;
                margin: 20px;
                background-color: #fafafa;
                color: #2c3e50;
            }
            h1 { color: #2980b9; border-bottom: 1px solid #ddd; }
            h2 { color: #3498db; }
            code {
                background-color: #f4f4f4;
                padding: 2px 4px;
                border-radius: 3px;
                font-family: 'Courier New', monospace;
                font-size: 0.9em;
            }
            pre {
                background-color: #f4f4f4;
                padding: 10px;
                border-radius: 5px;
                overflow-x: auto;
            }
            blockquote {
                border-left: 4px solid #3498db;
                margin-left: 0;
                padding-left: 15px;
                color: #7f8c8d;
            }
            table {
                border-collapse: collapse;
                width: 100%;
            }
            th, td {
                border: 1px solid #ddd;
                padding: 8px;
                text-align: left;
            }
            th {
                background-color: #ecf0f1;
            }
        </style>
        """
        full_html = f"<html><head>{css}</head><body>{html}</body></html>"
        # 保存当前滚动百分比
        editor_vscroll = editor.verticalScrollBar()
        src_min = editor_vscroll.minimum()
        src_max = editor_vscroll.maximum()
        src_val = editor_vscroll.value()
        percent = (src_val - src_min) / (src_max - src_min) if src_max > src_min else 0.0

        self.preview.setHtml(full_html)

        # 恢复预览区的滚动百分比
        preview_vscroll = self.preview.verticalScrollBar()
        tgt_min = preview_vscroll.minimum()
        tgt_max = preview_vscroll.maximum()
        tgt_val = int(tgt_min + percent * (tgt_max - tgt_min))
        self._syncing_scroll = True
        preview_vscroll.setValue(tgt_val)
        self._syncing_scroll = False

    # ========== 导出 Word ==========
    def export_to_word(self):
        editor = self.current_editor()
        if not editor:
            return
        if Document is None:
            QMessageBox.warning(self, "缺少依赖", "未安装 python-docx，无法导出 Word。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出为 Word", "", "Word 文档 (*.docx)"
        )
        if not path:
            return
        try:
            doc = Document()
            self.add_markdown_to_docx(doc, editor.toPlainText())
            doc.save(path)
            self.status_bar.showMessage(f"已导出：{path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"导出失败：{str(e)}")

    def add_markdown_to_docx(self, doc, text):
        import re
        lines = text.splitlines()
        in_code_block = False
        for line in lines:
            stripped = line.rstrip()
            if stripped.startswith("```"):
                in_code_block = not in_code_block
                continue
            if in_code_block:
                paragraph = doc.add_paragraph()
                run = paragraph.add_run(stripped)
                run.font.name = "Consolas"
                continue
            if not stripped:
                doc.add_paragraph("")
                continue
            if re.match(r"^#{1,6}\s+", stripped):
                level = len(stripped.split(" ", 1)[0])
                title = stripped[level + 1:]
                doc.add_heading(title, level=level)
                continue
            if re.match(r"^\d+\.\s+", stripped):
                content = re.sub(r"^\d+\.\s+", "", stripped)
                paragraph = doc.add_paragraph(style="List Number")
                self.add_inline_runs(paragraph, content)
                continue
            if stripped.startswith("- ") or stripped.startswith("* "):
                content = stripped[2:]
                paragraph = doc.add_paragraph(style="List Bullet")
                self.add_inline_runs(paragraph, content)
                continue
            if stripped.startswith("> "):
                paragraph = doc.add_paragraph(style="Intense Quote")
                self.add_inline_runs(paragraph, stripped[2:])
                continue
            paragraph = doc.add_paragraph()
            self.add_inline_runs(paragraph, stripped)

    def add_inline_runs(self, paragraph, text):
        import re
        token_re = re.compile(r"(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)")
        pos = 0
        for match in token_re.finditer(text):
            if match.start() > pos:
                paragraph.add_run(text[pos:match.start()])
            token = match.group(0)
            if token.startswith("**"):
                run = paragraph.add_run(token[2:-2])
                run.bold = True
            elif token.startswith("*"):
                run = paragraph.add_run(token[1:-1])
                run.italic = True
            elif token.startswith("`"):
                run = paragraph.add_run(token[1:-1])
                run.font.name = "Consolas"
            pos = match.end()
        if pos < len(text):
            paragraph.add_run(text[pos:])

    # ========== 帮助窗口 ==========
    def show_help(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("使用指导")
        dialog.resize(520, 420)

        layout = QVBoxLayout()
        guide = QTextEdit()
        guide.setReadOnly(True)
        guide.setPlainText(
            "使用指导\n"
            "\n"
            "1. 新建/打开\n"
            "- 文件菜单中可新建或打开文件。\n"
            "- 支持.txt, .md, .tex, .py文件。\n"
            "\n"
            "2. 分屏预览\n"
            "- 打开.md类型文件自动进入分屏模式。\n"
            "- 视图菜单可开启或关闭分屏模式。\n"
            "- 分屏时左侧预览，右侧编辑，滚动自动同步。\n"
            "- 鼠标悬停在某侧可自动放大该侧比例。\n"
            "\n"
            "3. 预览缩放\n"
            "- 视图菜单可放大或缩小预览字体。\n"
            "\n"
            "4. 导出 Word\n"
            "- 文件菜单可导出为 .docx。\n"
            "\n"  
            "本项目旨在提供一个方便、轻量的小窗口Markdown编辑体验\n"
            "欢迎反馈和建议！\n"
            "问题提交：https://github.com/Loryage\n"
        )
        layout.addWidget(guide)

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)

        dialog.setLayout(layout)
        dialog.exec()

    # ========== 界面构建 ==========
    def create_menu_bar(self):
        menubar = self.menuBar()
        file_menu = menubar.addMenu("文件(&F)")
        new_action = QAction("新建(&N)", self)
        new_action.triggered.connect(self.new_file)
        new_action.setShortcut("Ctrl+N")
        file_menu.addAction(new_action)

        open_action = QAction("打开(&O)...", self)
        open_action.triggered.connect(self.open_file)
        open_action.setShortcut("Ctrl+O")
        file_menu.addAction(open_action)

        save_action = QAction("保存(&S)", self)
        save_action.triggered.connect(self.save_file)
        save_action.setShortcut("Ctrl+S")
        file_menu.addAction(save_action)

        save_as_action = QAction("另存为(&A)...", self)
        save_as_action.triggered.connect(self.save_as_file)
        file_menu.addAction(save_as_action)

        export_action = QAction("导出为 Word(&W)...", self)
        export_action.triggered.connect(self.export_to_word)
        file_menu.addAction(export_action)

        file_menu.addSeparator()
        exit_action = QAction("退出(&X)", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        view_menu = menubar.addMenu("视图(&V)")
        zoom_in_action = QAction("放大预览(&+)", self)
        zoom_in_action.triggered.connect(lambda: self.preview_zoom_in())
        view_menu.addAction(zoom_in_action)

        zoom_out_action = QAction("缩小预览(&-)", self)
        zoom_out_action.triggered.connect(lambda: self.preview_zoom_out())
        view_menu.addAction(zoom_out_action)

        view_menu.addSeparator()
        self.split_action = QAction("分屏模式(&P)", self)
        self.split_action.setCheckable(True)
        self.split_action.setChecked(False)
        self.split_action.triggered.connect(self.toggle_split_mode)
        view_menu.addAction(self.split_action)

        help_menu = menubar.addMenu("帮助(&H)")
        help_action = QAction("使用指导(&G)", self)
        help_action.triggered.connect(self.show_help)
        help_menu.addAction(help_action)

    def preview_zoom_in(self):
        if self.split_enabled:
            font = self.preview.font()
            font.setPointSize(font.pointSize() + 1)
            self.preview.setFont(font)

    def preview_zoom_out(self):
        if self.split_enabled:
            font = self.preview.font()
            font.setPointSize(max(8, font.pointSize() - 1))
            self.preview.setFont(font)


def get_icon_path():
    base_dir = getattr(sys, "_MEIPASS", os.path.dirname(__file__))
    icon_path = os.path.join(base_dir, "tml.ico")
    return icon_path if os.path.exists(icon_path) else None


def set_windows_app_id(app_id):
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


if __name__ == "__main__":
    set_windows_app_id("TML.TMLEditor")
    app = QApplication(sys.argv)
    icon_path = get_icon_path()
    if icon_path:
        app.setWindowIcon(QIcon(icon_path))
    initial_paths = [path for path in sys.argv[1:] if os.path.isfile(path)]
    window = MarkdownEditor(initial_paths)
    window.show()
    sys.exit(app.exec())