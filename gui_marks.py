#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
唛头生成桌面工具(PySide6 GUI)
==============================

一个 PySide6 桌面工具「小小工具」:主窗口左侧是功能栏,点选其中一项即
在右侧区域显示对应功能页,不新开窗口,也不占用系统菜单栏。

当前已装功能「小小唛头-wichita」:选择 Excel 文件,列出所有 sheet,
指定其中一个为 packing list sheet、另一个为 gw sheet(两者不能相同),
然后一键生成南北唛头 Excel 和东西唛头 Excel。

增加新功能的方式:写一个功能页 Widget,在 MainWindow 里用 _add_feature
注册即可(左侧栏与右侧页面会同时加好)。

特点
- 不修改现有生成脚本:运行时动态加载并调用
  generate_marks_南北-0906.py / generate_marks_东西-0906.py 的 main(),
  完全复用它们读表、校验、算柜、生成的工作流;
- 可以随时重新选择文件,或单独清除已选的 sheet;
- sheet 列表默认只显示 Excel 标签栏里可见的 sheet。工作簿里的隐藏 sheet
  (openpyxl 会一并读出)不会混进来;需要时勾选「显示隐藏的 sheet」查看,
  此时隐藏项带 "(隐藏)" 后缀并以灰色显示;
- 输出文件写到所选 Excel 所在目录:
      <文件名>-南北唛头.xlsx / <文件名>-东西唛头.xlsx
- 生成在后台线程里跑,界面不会卡住。

用法(项目用 uv 管理,推荐通过 uv run 启动)
    uv run gui_marks.py                     # 打开图形界面
    uv run gui_marks.py --selftest <excel> --sheet <packing> --gw-sheet <gw>
                                            # 不开界面,直接跑两个生成器(自检)
macOS 也可以直接双击同目录的 "启动界面.command"。

编码与路径:
- 源码为 UTF-8;窗口内的标题、label 等文字由 Qt 以 Unicode 渲染,与系统区域无关,
  Windows 上不会出现乱码;
- 控制台输出在入口统一重设为 UTF-8(见 configure_stdio),避免 Windows 重定向
  输出时中文触发 UnicodeEncodeError;
- 路径一律走 os.path 处理(不硬编码分隔符),生成脚本也接受两种斜杠混用,
  并对 Windows 的 260 字符路径上限做了提前检查。

依赖: 见 pyproject.toml (PySide6 + openpyxl),执行 uv sync 安装。
"""

import argparse
import importlib.util
import io
import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QBrush, QColor, QFont, QFontDatabase
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QGroupBox, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QPlainTextEdit,
    QPushButton, QStackedWidget, QVBoxLayout, QWidget,
)


APP_NAME = "小小工具"
MARKS_FEATURE_NAME = "小小唛头-wichita"

NS_SCRIPT = "generate_marks_南北-0906.py"
EW_SCRIPT = "generate_marks_东西-0906.py"
NS_SUFFIX = "南北唛头"
EW_SUFFIX = "东西唛头"


def configure_stdio():
    """把控制台输出统一成 UTF-8。

    窗口里的文字走 Qt,内部本来就是 Unicode,不存在编码问题;这里只针对
    控制台输出(--selftest 的 print)。Windows 上如果 stdout 被重定向、而系统
    区域又不是中文,默认编码可能编不出中文并抛 UnicodeEncodeError,所以显式
    指定 UTF-8,并把 errors 放宽成 replace 作为兜底。
    打包成 windowed exe 时 sys.stdout / sys.stderr 为 None,直接跳过。
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (ValueError, OSError):
            pass


def output_names_for(excel_path):
    """两个输出文件名(南北 / 东西),都是纯文件名,写到所选 Excel 所在目录。"""
    stem = os.path.splitext(os.path.basename(excel_path))[0]
    return ["%s-%s.xlsx" % (stem, NS_SUFFIX), "%s-%s.xlsx" % (stem, EW_SUFFIX)]


def windows_path_limit_hit(folder, names):
    """Windows 上整条路径超过 260 字符就保存不了,提前查出超长的那一条。

    非 Windows 平台一律返回 None。放长路径支持(注册表 LongPathsEnabled)之外,
    这是 Windows 特有的失败点,提前提示比让 openpyxl 抛一个含糊的错更好。
    """
    if os.name != "nt":
        return None
    for name in names:
        full = os.path.abspath(os.path.join(folder or ".", name))
        if len(full) > 255:
            return full
    return None


def app_dir():
    """Folder that holds the generator scripts (works for PyInstaller too)."""
    if getattr(sys, "frozen", False):
        return getattr(sys, "_MEIPASS",
                       os.path.dirname(os.path.abspath(sys.executable)))
    return os.path.dirname(os.path.abspath(__file__))


def load_generator(filename, module_name):
    """Load a generator script as a module without modifying it."""
    path = os.path.join(app_dir(), filename)
    if not os.path.isfile(path):
        raise RuntimeError("找不到生成脚本: %s" % path)
    spec = importlib.util.spec_from_file_location(module_name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_generator(filename, module_name, excel_path, sheet, gw_sheet, out_name):
    """Run one generator's main() with CLI-style arguments; capture output.

    Returns (ok, message). The generator writes <out_name> into the folder
    of the selected Excel file (the scripts only accept a plain file name).
    """
    module = load_generator(filename, module_name)
    buf = io.StringIO()
    old_argv, old_stdout, old_cwd = sys.argv, sys.stdout, os.getcwd()
    sys.argv = [filename, excel_path, "--sheet", sheet,
                "--gw-sheet", gw_sheet, "--out", out_name]
    sys.stdout = buf
    os.chdir(os.path.dirname(excel_path) or ".")
    try:
        module.main()
        return True, buf.getvalue()
    except SystemExit as exc:
        text = str(exc.code) if exc.code not in (None, 0) else "已中止"
        return False, text + "\n" + buf.getvalue()
    except Exception:
        return False, traceback.format_exc() + "\n" + buf.getvalue()
    finally:
        sys.argv, sys.stdout = old_argv, old_stdout
        os.chdir(old_cwd)


def run_both(excel_path, sheet, gw_sheet, log=None):
    """Generate the N/S and E/W mark workbooks; returns list of results."""
    out_ns, out_ew = output_names_for(excel_path)
    results = []
    for label, script, module_name, out_name in (
            ("南北", NS_SCRIPT, "marks_ns", out_ns),
            ("东西", EW_SCRIPT, "marks_ew", out_ew)):
        if log:
            log("开始生成%s唛头 -> %s ..." % (label, out_name))
        ok, message = run_generator(script, module_name, excel_path,
                                    sheet, gw_sheet, out_name)
        if log:
            log(message.rstrip())
            log("%s唛头: %s" % (label, "成功" if ok else "失败"))
        results.append({"label": label, "ok": ok, "message": message,
                        "out_name": out_name})
    return results


class GenerateWorker(QThread):
    """Runs both generators off the GUI thread; emits log lines and results."""

    logged = Signal(str)
    finished_with = Signal(list)

    def __init__(self, excel_path, sheet, gw_sheet, parent=None):
        super().__init__(parent)
        self.excel_path = excel_path
        self.sheet = sheet
        self.gw_sheet = gw_sheet

    def run(self):
        try:
            results = run_both(self.excel_path, self.sheet, self.gw_sheet,
                               log=self.logged.emit)
        except Exception:
            results = [{"label": "-", "ok": False,
                        "message": traceback.format_exc(),
                        "out_name": ""}]
        self.finished_with.emit(results)


class MarkPage(QWidget):
    """唛头生成功能页: 选工作簿、指定两个 sheet、生成南北与东西唛头。"""

    status_changed = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.excel_path = None
        self.sheet_info = []          # [(sheet name, is_hidden), ...]
        self.ns_sheet = None
        self.gw_sheet = None
        self.worker = None
        self._build_ui()

    def is_busy(self):
        return self.worker is not None and self.worker.isRunning()

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(10)

        header = QLabel(MARKS_FEATURE_NAME, self)
        header_font = QFont()
        header_font.setPointSize(15)
        header_font.setBold(True)
        header.setFont(header_font)
        outer.addWidget(header)

        # file row
        file_box = QGroupBox("Excel 文件", self)
        file_row = QHBoxLayout(file_box)
        self.btn_choose = QPushButton("选择 Excel 文件…", file_box)
        self.btn_choose.clicked.connect(self.choose_file)
        file_row.addWidget(self.btn_choose)
        self.file_label = QLabel("(未选择文件)", file_box)
        self.file_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse)
        file_row.addWidget(self.file_label, 1)
        outer.addWidget(file_box)

        # sheets + selection
        mid = QHBoxLayout()
        mid.setSpacing(10)
        outer.addLayout(mid, 1)

        list_box = QGroupBox("Sheet 列表", self)
        list_layout = QVBoxLayout(list_box)
        self.sheet_list = QListWidget(list_box)
        self.sheet_list.setSelectionMode(
            QListWidget.SelectionMode.SingleSelection)
        self.sheet_list.itemDoubleClicked.connect(self._on_row_double_clicked)
        list_layout.addWidget(self.sheet_list)
        self.chk_hidden = QCheckBox("显示隐藏的 sheet", list_box)
        self.chk_hidden.toggled.connect(self._on_hidden_toggled)
        list_layout.addWidget(self.chk_hidden)
        mid.addWidget(list_box, 1)

        pick = QGroupBox("已选 sheet", self)
        pick_layout = QVBoxLayout(pick)

        ns_row = QHBoxLayout()
        ns_row.addWidget(QLabel("Packing List sheet:", pick))
        self.ns_label = QLabel("(未选择)", pick)
        self.ns_label.setStyleSheet("color: #0055aa;")
        ns_row.addWidget(self.ns_label, 1)
        self.btn_ns_clear = QPushButton("清除", pick)
        self.btn_ns_clear.setFixedWidth(64)
        self.btn_ns_clear.clicked.connect(lambda: self.clear_sheet("ns"))
        ns_row.addWidget(self.btn_ns_clear)
        pick_layout.addLayout(ns_row)

        gw_row = QHBoxLayout()
        gw_row.addWidget(QLabel("GW sheet:", pick))
        self.gw_label = QLabel("(未选择)", pick)
        self.gw_label.setStyleSheet("color: #0055aa;")
        gw_row.addWidget(self.gw_label, 1)
        self.btn_gw_clear = QPushButton("清除", pick)
        self.btn_gw_clear.setFixedWidth(64)
        self.btn_gw_clear.clicked.connect(lambda: self.clear_sheet("gw"))
        gw_row.addWidget(self.btn_gw_clear)
        pick_layout.addLayout(gw_row)

        pick_layout.addSpacing(12)
        self.btn_set_ns = QPushButton("把列表中选中的 sheet 设为 Packing List", pick)
        self.btn_set_ns.clicked.connect(lambda: self.set_sheet("ns"))
        pick_layout.addWidget(self.btn_set_ns)
        self.btn_set_gw = QPushButton("把列表中选中的 sheet 设为 GW sheet", pick)
        self.btn_set_gw.clicked.connect(lambda: self.set_sheet("gw"))
        pick_layout.addWidget(self.btn_set_gw)

        hint = QLabel("提示: 两个 sheet 不能相同;可随时重新选择文件", pick)
        hint.setStyleSheet("color: #666666;")
        hint.setWordWrap(True)
        pick_layout.addWidget(hint)
        pick_layout.addStretch(1)
        mid.addWidget(pick, 1)

        # generate + log
        self.btn_run = QPushButton("生成唛头(南北 + 东西)", self)
        self.btn_run.clicked.connect(self.on_generate)
        outer.addWidget(self.btn_run)

        self.log_text = QPlainTextEdit(self)
        self.log_text.setReadOnly(True)
        # 用各平台自己的等宽字体(Windows 上是 Consolas、macOS 上是 Menlo),
        # 而不是写死某个平台的字体名;缺字形时 Qt 会做字体回退,中文可正常显示。
        log_font = QFontDatabase.systemFont(
            QFontDatabase.SystemFont.FixedFont)
        log_font.setPointSize(12)
        self.log_text.setFont(log_font)
        self.log_text.setMinimumHeight(160)
        outer.addWidget(self.log_text, 1)

    # ------------------------------------------------------------- actions
    def _log(self, text):
        self.log_text.appendPlainText(text.rstrip())
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def _refresh_labels(self):
        self.ns_label.setText(self.ns_sheet or "(未选择)")
        self.gw_label.setText(self.gw_sheet or "(未选择)")

    def _set_busy(self, busy):
        for widget in (self.btn_choose, self.btn_run, self.btn_set_ns,
                       self.btn_set_gw, self.btn_ns_clear, self.btn_gw_clear):
            widget.setEnabled(not busy)
        self.status_changed.emit("生成中…" if busy else "就绪")

    def choose_file(self):
        if self.is_busy():
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel 文件", "",
            "Excel 文件 (*.xlsx *.xlsm);;所有文件 (*)")
        if not path:
            return
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            sheet_info = [(name, wb[name].sheet_state != "visible")
                          for name in wb.sheetnames]
            wb.close()
        except Exception as exc:
            QMessageBox.critical(self, "无法打开文件", "%s\n\n%s" % (path, exc))
            return
        self.excel_path = path
        self.sheet_info = sheet_info
        self.ns_sheet = None
        self.gw_sheet = None
        self.file_label.setText(path)
        self._populate_sheet_list()
        self._refresh_labels()
        visible = [n for n, hidden in sheet_info if not hidden]
        hiddens = [n for n, hidden in sheet_info if hidden]
        self._log("已选择文件: %s" % path)
        self._log("可见 sheet(%d): %s" % (len(visible), ", ".join(visible)))
        if hiddens:
            self._log("隐藏 sheet(%d, 未列出): %s"
                      % (len(hiddens), ", ".join(hiddens)))

    def _populate_sheet_list(self):
        """Fill the list from sheet_info, honouring the hidden-sheet filter."""
        previous = self._selected_sheet()
        self.sheet_list.clear()
        for name, hidden in self.sheet_info:
            if hidden and not self.chk_hidden.isChecked():
                continue
            item = QListWidgetItem("%s  (隐藏)" % name if hidden else name)
            item.setData(Qt.ItemDataRole.UserRole, name)
            if hidden:
                item.setForeground(QBrush(QColor("#999999")))
            self.sheet_list.addItem(item)
            if name == previous:
                item.setSelected(True)
                self.sheet_list.setCurrentItem(item)

    def _on_hidden_toggled(self, checked):
        self._populate_sheet_list()
        self._log("显示隐藏的 sheet: %s" % ("是" if checked else "否"))

    def _selected_sheet(self):
        item = self.sheet_list.currentItem()
        if item is None or not item.isSelected():
            return None
        return item.data(Qt.ItemDataRole.UserRole)

    def _on_row_double_clicked(self, _item):
        """Double-click a row to set whichever slot is still empty."""
        if self.ns_sheet is None:
            self.set_sheet("ns")
        elif self.gw_sheet is None:
            self.set_sheet("gw")

    def set_sheet(self, which):
        if self.is_busy():
            return
        if not self.excel_path:
            QMessageBox.warning(self, "提示", "请先选择 Excel 文件")
            return
        name = self._selected_sheet()
        if name is None:
            QMessageBox.warning(self, "提示", "请先在左侧列表中选择一个 sheet")
            return
        other = self.gw_sheet if which == "ns" else self.ns_sheet
        if other == name:
            QMessageBox.warning(self, "提示",
                                "Packing List sheet 与 GW sheet 不能相同")
            return
        if which == "ns":
            self.ns_sheet = name
        else:
            self.gw_sheet = name
        self._refresh_labels()
        self._log("已设置 %s = %s"
                  % ("Packing List sheet" if which == "ns" else "GW sheet", name))

    def clear_sheet(self, which):
        if self.is_busy():
            return
        if which == "ns":
            self.ns_sheet = None
        else:
            self.gw_sheet = None
        self._refresh_labels()
        self._log("已清除 %s"
                  % ("Packing List sheet" if which == "ns" else "GW sheet"))

    def on_generate(self):
        if self.is_busy():
            return
        if not self.excel_path:
            QMessageBox.warning(self, "提示", "请先选择 Excel 文件")
            return
        if not self.ns_sheet or not self.gw_sheet:
            QMessageBox.warning(self, "提示",
                                "请分别选择 Packing List sheet 和 GW sheet")
            return
        if self.ns_sheet == self.gw_sheet:
            QMessageBox.warning(self, "提示",
                                "Packing List sheet 与 GW sheet 不能相同")
            return
        too_long = windows_path_limit_hit(
            os.path.dirname(self.excel_path),
            output_names_for(self.excel_path))
        if too_long:
            QMessageBox.warning(
                self, "路径过长",
                "Windows 下完整路径超过 260 个字符就无法保存文件:\n\n%s\n\n"
                "当前 %d 个字符。请把 Excel 文件移到层级更浅的目录后重试。"
                % (too_long, len(too_long)))
            return
        self._set_busy(True)
        self._log("=" * 60)
        self._log("文件: %s" % self.excel_path)
        self._log("Packing List sheet: %s" % self.ns_sheet)
        self._log("GW sheet: %s" % self.gw_sheet)
        self.worker = GenerateWorker(self.excel_path, self.ns_sheet,
                                     self.gw_sheet, self)
        self.worker.logged.connect(self._log)
        self.worker.finished_with.connect(self._on_done)
        self.worker.start()

    def _on_done(self, results):
        self._set_busy(False)
        folder = os.path.dirname(self.excel_path)
        ok_all = all(r["ok"] for r in results)
        for r in results:
            if r["ok"]:
                self._log("输出: %s" % os.path.join(folder, r["out_name"]))
        if ok_all:
            files = "\n".join(os.path.join(folder, r["out_name"])
                              for r in results)
            QMessageBox.information(self, "完成", "已生成:\n%s" % files)
        else:
            bad = "\n".join("%s唛头: %s" % (r["label"], r["message"].strip())
                            for r in results if not r["ok"])
            QMessageBox.critical(self, "生成失败", bad)

class MainWindow(QMainWindow):
    """主窗口「小小工具」。

    功能都列在左侧栏,点选其中一项即在右侧区域显示对应功能页;
    不使用系统菜单栏,也不新开窗口。
    后续新增功能: 写一个功能页 Widget,在 __init__ 里用 _add_feature 注册即可。
    若功能页提供 is_busy(),状态栏会在它忙碌时自动显示"生成中…"。
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(900, 700)
        self.setMinimumSize(780, 560)

        self.mark_page = MarkPage(self)
        self.mark_page.status_changed.connect(self._on_status)

        self.stack = QStackedWidget(self)
        self.nav = QListWidget(self)

        self._build_ui()
        self._add_feature("首页", self._build_home())
        self._add_feature(MARKS_FEATURE_NAME, self.mark_page)
        self.nav.setCurrentRow(0)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        row = QHBoxLayout(central)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(0)
        row.addWidget(self._build_sidebar())
        row.addWidget(self.stack, 1)

    def _build_sidebar(self):
        """左侧栏: 顶部工具名,中间功能列表,底部退出按钮。"""
        side = QWidget(self)
        side.setObjectName("sidebar")
        side.setFixedWidth(200)
        side.setStyleSheet("#sidebar { background: #f4f4f6;"
                           " border-right: 1px solid #dcdce0; }")

        layout = QVBoxLayout(side)
        layout.setContentsMargins(14, 16, 14, 14)
        layout.setSpacing(12)

        title = QLabel(APP_NAME, side)
        title_font = QFont()
        title_font.setPointSize(17)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        self.nav.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.nav.setStyleSheet(
            "QListWidget { background: transparent; border: none;"
            " outline: none; }"
            "QListWidget::item { padding: 7px 10px; }"
            "QListWidget::item:hover { background: #e7e7ec; }"
            "QListWidget::item:selected { background: #d7e5f7;"
            " color: #12395f; }")
        self.nav.currentRowChanged.connect(self._on_nav_changed)
        layout.addWidget(self.nav, 1)

        btn_quit = QPushButton("退出", side)
        btn_quit.clicked.connect(self.close)
        layout.addWidget(btn_quit)
        return side

    def _build_home(self):
        """首页: 只做导航提示,功能要点了左侧栏才进入。"""
        page = QWidget(self)
        layout = QVBoxLayout(page)
        layout.addStretch(1)

        title = QLabel(APP_NAME, page)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_font = QFont()
        title_font.setPointSize(30)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)

        hint = QLabel("请从左侧选择要使用的功能", page)
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setStyleSheet("color: #666666;")
        layout.addWidget(hint)

        layout.addSpacing(20)
        features = QLabel("已安装功能: %s(生成南北 / 东西唛头)"
                          % MARKS_FEATURE_NAME, page)
        features.setAlignment(Qt.AlignmentFlag.AlignCenter)
        features.setStyleSheet("color: #888888;")
        layout.addWidget(features)

        layout.addStretch(1)
        return page

    def _add_feature(self, name, page):
        """注册一个功能: 左侧栏加一项、右侧容器加一页,两者索引保持一致。"""
        item = QListWidgetItem(name)
        item.setData(Qt.ItemDataRole.UserRole, page)
        self.nav.addItem(item)
        self.stack.addWidget(page)

    # ------------------------------------------------------------- actions
    def _on_nav_changed(self, row):
        if row < 0:
            return
        page = self.nav.item(row).data(Qt.ItemDataRole.UserRole)
        self.stack.setCurrentWidget(page)
        busy = hasattr(page, "is_busy") and page.is_busy()
        self._on_status("生成中…" if busy else "就绪")

    def _on_status(self, text):
        self.statusBar().showMessage(text)

    def closeEvent(self, event):
        if self.mark_page.is_busy():
            QMessageBox.warning(self, "提示", "正在生成,请等待完成后再关闭")
            event.ignore()
            return
        event.accept()


def run_selftest(excel_path, sheet, gw_sheet):
    """Headless check: run both generators exactly like the GUI does."""
    excel_path = os.path.abspath(excel_path)
    print("selftest file : %s" % excel_path)
    print("packing sheet : %s" % sheet)
    print("gw sheet      : %s" % gw_sheet)
    results = run_both(excel_path, sheet, gw_sheet, log=lambda t: print(t))
    ok = all(r["ok"] for r in results)
    print("selftest: %s" % ("OK" if ok else "FAILED"))
    return 0 if ok else 1


def main():
    # 先把控制台输出固定成 UTF-8,Windows 上重定向输出时才不会因中文编码失败
    configure_stdio()

    parser = argparse.ArgumentParser(description="唛头生成桌面工具")
    parser.add_argument("--selftest", metavar="EXCEL",
                        help="不开界面,直接跑两个生成器(自检)")
    parser.add_argument("--sheet", default="Block 5-8 Packing List",
                        help="packing list sheet name (selftest)")
    parser.add_argument("--gw-sheet", default="Batch_01_GW",
                        help="gw sheet name (selftest)")
    args = parser.parse_args()

    if args.selftest:
        sys.exit(run_selftest(args.selftest, args.sheet, args.gw_sheet))

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    window = MainWindow()
    window.show()
    window.raise_()
    window.activateWindow()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
