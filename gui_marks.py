#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
唛头生成桌面工具(PySide6 GUI)
==============================

一个 PySide6 界面小工具:菜单栏「唛头」→ 选择 Excel 文件,列出所有 sheet,
指定其中一个为 packing list sheet、另一个为 gw sheet(两者不能相同),
然后一键生成南北唛头 Excel 和东西唛头 Excel。

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

依赖: 见 pyproject.toml (PySide6 + openpyxl),执行 uv sync 安装。
"""

import argparse
import importlib.util
import io
import os
import sys
import traceback

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QBrush, QColor, QFont
from PySide6.QtWidgets import (
    QApplication, QCheckBox, QFileDialog, QGroupBox, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QPlainTextEdit,
    QPushButton, QVBoxLayout, QWidget,
)


NS_SCRIPT = "generate_marks_南北-0906.py"
EW_SCRIPT = "generate_marks_东西-0906.py"
NS_SUFFIX = "南北唛头"
EW_SUFFIX = "东西唛头"


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
    stem = os.path.splitext(os.path.basename(excel_path))[0]
    results = []
    for label, script, module_name, suffix in (
            ("南北", NS_SCRIPT, "marks_ns", NS_SUFFIX),
            ("东西", EW_SCRIPT, "marks_ew", EW_SUFFIX)):
        out_name = "%s-%s.xlsx" % (stem, suffix)
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


class MarkWindow(QMainWindow):
    """PySide6 UI: pick a workbook, pick two sheets, generate both marks."""

    def __init__(self):
        super().__init__()
        self.excel_path = None
        self.sheet_info = []          # [(sheet name, is_hidden), ...]
        self.ns_sheet = None
        self.gw_sheet = None
        self.worker = None

        self.setWindowTitle("唛头生成工具")
        self.resize(820, 620)
        self.setMinimumSize(720, 540)
        self._build_menu()
        self._build_ui()
        self.statusBar().showMessage("就绪")

    def _busy(self):
        return self.worker is not None and self.worker.isRunning()

    # ------------------------------------------------------------------ UI
    def _build_menu(self):
        menu = self.menuBar().addMenu("唛头")
        act_choose = QAction("选择 Excel 文件…", self)
        act_choose.triggered.connect(self.choose_file)
        menu.addAction(act_choose)
        act_rechoose = QAction("重新选择文件", self)
        act_rechoose.triggered.connect(self.choose_file)
        menu.addAction(act_rechoose)
        menu.addSeparator()
        act_clear_ns = QAction("清除 Packing List 选择", self)
        act_clear_ns.triggered.connect(lambda: self.clear_sheet("ns"))
        menu.addAction(act_clear_ns)
        act_clear_gw = QAction("清除 GW sheet 选择", self)
        act_clear_gw.triggered.connect(lambda: self.clear_sheet("gw"))
        menu.addAction(act_clear_gw)
        menu.addSeparator()
        act_quit = QAction("退出", self)
        act_quit.triggered.connect(self.close)
        menu.addAction(act_quit)

    def _build_ui(self):
        central = QWidget(self)
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(12, 12, 12, 12)
        outer.setSpacing(10)

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
        self.log_text.setFont(QFont("Menlo", 12))
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
        self.statusBar().showMessage("生成中…" if busy else "就绪")

    def choose_file(self):
        if self._busy():
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
        if self._busy():
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
        if self._busy():
            return
        if which == "ns":
            self.ns_sheet = None
        else:
            self.gw_sheet = None
        self._refresh_labels()
        self._log("已清除 %s"
                  % ("Packing List sheet" if which == "ns" else "GW sheet"))

    def on_generate(self):
        if self._busy():
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

    def closeEvent(self, event):
        if self._busy():
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
    app.setApplicationName("唛头生成工具")
    window = MarkWindow()
    window.show()
    window.raise_()
    window.activateWindow()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
