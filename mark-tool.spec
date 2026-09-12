# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置,目标是让 exe 体积尽量小。

用法:
    uv run pyinstaller --clean --noconfirm mark-tool.spec

精简手段(按实际效果排序):
1. Windows 上最省体积的一项:丢掉 PyInstaller 的 Qt hook 会无条件附带的
   "动态 OpenGL" 运行库(opengl32sw.dll 一项就约 20MB),见 DROP_BASENAMES。
   本工具全部用 Qt 光栅绘制,不碰 QOpenGLWidget / QQuick,用不到这些库。
2. 依赖 pyside6-essentials(见 pyproject.toml)而不是完整 pyside6:
   环境体积从 1.2G 降到 371M,同时保证 addons 里的 WebEngine / 3D / Charts
   没有机会被带进来。
3. excludes 列出的 Qt 模块属于兜底:PyInstaller 只收集被 import 到的模块,
   本工具只用 QtCore / QtGui / QtWidgets,这些模块正常不会被收集,
   列出来是为了防止以后误 import 或 hook 变化导致体积悄悄膨胀。
4. 去掉除简体中文外的 Qt 翻译文件。
5. 刻意不用 UPX:它压 Qt6 的 DLL 容易压出运行时崩溃,收益也不如上面几项。
"""

import os

EXE_NAME = "小小工具"
ENTRY_SCRIPT = "gui_marks.py"

# 这两个生成脚本没有在 gui_marks.py 里 import,而是运行时按文件名动态加载
# (见 gui_marks.py 的 load_generator),静态分析看不到,必须作为数据文件带上。
# 冻结后它们会落在 sys._MEIPASS,正好是 app_dir() 查找的位置。
GENERATOR_SCRIPTS = [
    "generate_marks_南北-0906.py",
    "generate_marks_东西-0906.py",
]

# 生成脚本自己的依赖,同样因为动态加载不会被自动发现,需要显式声明。
HIDDEN_IMPORTS = [
    "openpyxl",
    "openpyxl.styles",
    "openpyxl.workbook",
    "openpyxl.worksheet",
    "et_xmlfile",
]

EXCLUDES = [
    # --- Python 标准库里用不到的 ---
    "tkinter", "unittest", "pydoc", "doctest", "pdb", "test",
    "idlelib", "lib2to3", "ensurepip", "venv", "distutils",
    "setuptools", "pip", "wheel",
    # --- 没安装的第三方库,防止被间接收集进来 ---
    "numpy", "pandas", "scipy", "matplotlib", "PIL",
    "PyQt5", "PyQt6", "PySide2",
    # --- PySide6: 只保留 QtCore / QtGui / QtWidgets,其余全砍 ---
    "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtQuickWidgets",
    "PySide6.QtQuickControls2", "PySide6.QtNetwork", "PySide6.QtSql",
    "PySide6.QtTest", "PySide6.QtDesigner", "PySide6.QtHelp",
    "PySide6.QtUiTools", "PySide6.QtDBus", "PySide6.QtXml",
    "PySide6.QtConcurrent", "PySide6.QtStateMachine",
    "PySide6.QtSerialPort", "PySide6.QtWebSockets", "PySide6.QtWebChannel",
    "PySide6.QtOpenGL", "PySide6.QtOpenGLWidgets", "PySide6.QtSvg",
    "PySide6.QtSvgWidgets", "PySide6.QtPrintSupport",
    "PySide6.QtMultimedia", "PySide6.QtMultimediaWidgets",
    "PySide6.QtCharts", "PySide6.QtDataVisualization", "PySide6.QtGraphs",
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets",
    "PySide6.QtBluetooth", "PySide6.QtNfc", "PySide6.QtPositioning",
    "PySide6.QtLocation", "PySide6.QtSensors", "PySide6.QtSerialBus",
    "PySide6.QtRemoteObjects", "PySide6.QtScxml", "PySide6.QtTextToSpeech",
    "PySide6.QtPdf", "PySide6.QtPdfWidgets", "PySide6.QtNetworkAuth",
    "PySide6.QtHttpServer", "PySide6.QtSpatialAudio",
]

# Windows 上真正省体积的大头。
# PyInstaller 的 Qt hook 在 Windows 上会无条件附带这些"动态 OpenGL"运行库,
# 而本工具全部走 Qt 光栅绘制,不碰 QOpenGLWidget / QQuick,一个都用不到。
DROP_BASENAMES = {
    "opengl32sw.dll",     # 纯软件 OpenGL 回退,约 20MB,单项占比最大
    "libegl.dll",         # ANGLE(把 OpenGL 转到 D3D)相关
    "libglesv2.dll",
}
# d3dcompiler_XX.dll 带版本号(如 d3dcompiler_47.dll),按前缀匹配
DROP_BASENAME_PREFIXES = ("d3dcompiler_",)

# Qt 自带翻译只留简体中文,其余几十种语言包全部丢掉。
# 保留翻译是为了让 QMessageBox 这类标准对话框的按钮显示中文。
KEEP_TRANSLATION_SUFFIXES = ("_zh_cn.qm", "_en.qm")


def _keep(name):
    """判断某个收集到的文件是否保留。"""
    base = os.path.basename(name).lower()
    if base in DROP_BASENAMES:
        return False
    if base.startswith(DROP_BASENAME_PREFIXES):
        return False
    norm = name.replace(os.sep, "/")
    if "/translations/" in norm:
        return base.endswith(KEEP_TRANSLATION_SUFFIXES)
    if "/qml/" in norm:
        return False
    return True


a = Analysis(
    [ENTRY_SCRIPT],
    pathex=[],
    binaries=[],
    datas=[(p, ".") for p in GENERATOR_SCRIPTS],
    hiddenimports=HIDDEN_IMPORTS,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
)

# 在 EXE 之前过滤,打进包里的就只有留下来的部分
a.binaries = [b for b in a.binaries if _keep(b[0])]
a.datas = [d for d in a.datas if _keep(d[0])]

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=EXE_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,                      # GUI 程序,不弹控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
