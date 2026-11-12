# -*- coding: utf-8 -*-
"""
WP互相转换工具 — Word ⇄ PDF 双向转换器
======================================
转换后端：Microsoft Word COM（优先，保真度最高）/ LibreOffice（备用）
  · Word(.docx/.doc/.docm/.dotx/.rtf)  ->  PDF
  · PDF                               ->  Word(.docx)

结构：
  wp_converter.py      UI / 调度 / 环境自检 / 子进程生命周期
  convert_engine.ps1   一次转换（批量 JSONL + PROGRESS 协议 + 双后端）
  dismiss_dialogs.ps1  只负责处理 Office 弹窗
"""

import os
import re
import sys
import csv
import io
import json
import time
import shutil
import tempfile
import threading
import subprocess
import queue
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_NAME = "WP互相转换工具"
APP_SUB = "Word / Excel / PPT ⇄ PDF"
APP_VERSION = "2.2"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENGINE_PS1 = os.path.join(BASE_DIR, "convert_engine.ps1")
DISMISS_PS1 = os.path.join(BASE_DIR, "dismiss_dialogs.ps1")

WORD_EXTS = (".docx", ".doc", ".docm", ".dotx", ".rtf")
EXCEL_EXTS = (".xlsx", ".xlsm", ".xlsb", ".xls")
PPT_EXTS = (".pptx", ".pptm", ".ppsm", ".ppsx", ".ppt", ".pps")
PDF_EXTS = (".pdf",)

# 模式 -> 界面上显示的方向文字
MODE_ARROW = {
    "word2pdf": "Word → PDF",
    "excel2pdf": "Excel → PDF",
    "ppt2pdf": "PPT → PDF",
    "pdf2word": "PDF → Word",
    "pdf2ppt": "PDF → PPT",
    "pdf2excel": "PDF → Excel",
}

# 模式 -> 依赖的 Office 应用（LibreOffice 后端不需要这些）
# 注意：PDF -> PPT / Excel 需要先用 Word 把 PDF 重排成 DOCX，所以同时依赖 Word。
MODE_APPS = {
    "word2pdf": ("Word",),
    "excel2pdf": ("Excel",),
    "ppt2pdf": ("PowerPoint",),
    "pdf2word": ("Word",),
    "pdf2ppt": ("Word", "PowerPoint"),
    "pdf2excel": ("Word", "Excel"),
}

# 输入为 PDF 的模式（需要选择输出目标）
PDF_MODES = ("pdf2word", "pdf2ppt", "pdf2excel")
# 标签里直接写明质量定位：PDF → PPT/Excel 是“内容级重建”，避免用户误以为能保留版式
PDF_TARGET_LABELS = ("Word", "PPT（仅内容）", "Excel（仅内容）")
PDF_TARGET_MODES = {
    "Word": "pdf2word",
    "PPT（仅内容）": "pdf2ppt",
    "Excel（仅内容）": "pdf2excel",
}

OFFICE_APPS = ("Word", "Excel", "PowerPoint")
TIMEOUT = 180
STALL_GRACE = 60
STALL_NOTICE = 60
CHUNK_SIZE = 10
CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

LO_HOME_PAGE = "https://www.libreoffice.org/download/download-libreoffice/"
OFFICE_PAGE = "https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products"

BG = "#f4f6fb"
CARD = "#ffffff"
ACCENT = "#2f6fdb"
ACCENT_DARK = "#2459b0"
DANGER = "#c0392b"
DANGER_DARK = "#96281b"
OKC = "#1f9d55"
ERRC = "#d64545"
TEXT = "#22262e"
MUTED = "#6b7280"


# =========================== 基础工具 ===========================

def find_powershell():
    p = shutil.which("powershell")
    if p:
        return p
    cand = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"),
                        r"System32\WindowsPowerShell\v1.0\powershell.exe")
    return cand if os.path.isfile(cand) else "powershell"


POWERSHELL = find_powershell()


def decode_bytes(data):
    """PowerShell 5.1 重定向输出默认是系统代码页(中文=GBK)，故 mbcs 优先。"""
    if not data:
        return ""
    encs = (["mbcs"] if os.name == "nt" else []) + ["utf-8", "gbk"]
    for enc in encs:
        try:
            return data.decode(enc)
        except Exception:
            pass
    return data.decode("utf-8", "replace")


def office_version(app):
    """返回指定 Office 应用的主版本号（如 16），未安装返回 None。"""
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"%s.Application\CurVer" % app) as k:
            cur = str(winreg.QueryValueEx(k, "")[0])
        m = re.search(re.escape(app) + r"\.Application\.(\d+)", cur, re.IGNORECASE)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    try:
        import winreg
        for ver in ("16.0", "15.0", "14.0", "12.0"):
            try:
                winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                               r"Software\Microsoft\Office\%s\%s" % (ver, app))
                return int(ver.split(".")[0])
            except Exception:
                pass
    except Exception:
        pass
    return None


def office_versions():
    """返回 {应用名: 版本号或 None}。"""
    return dict((app, office_version(app)) for app in OFFICE_APPS)


def find_soffice():
    """定位 LibreOffice 的 soffice.exe。"""
    cands = []
    pf = os.environ.get("ProgramFiles")
    pf86 = os.environ.get("ProgramFiles(x86)")
    lad = os.environ.get("LOCALAPPDATA")
    if pf:
        cands.append(os.path.join(pf, "LibreOffice", "program", "soffice.exe"))
    if pf86:
        cands.append(os.path.join(pf86, "LibreOffice", "program", "soffice.exe"))
    if lad:
        cands.append(os.path.join(lad, "Programs", "LibreOffice", "program", "soffice.exe"))
    for c in cands:
        if os.path.isfile(c):
            return c
    return shutil.which("soffice")


_BACKEND = None


def refresh_backend():
    """探测可用转换后端：优先 Microsoft Office（Word/Excel/PPT 任一），其次 LibreOffice。"""
    global _BACKEND
    vers = office_versions()
    if any(v is not None for v in vers.values()):
        _BACKEND = ("office", vers)
    else:
        so = find_soffice()
        _BACKEND = ("libreoffice", so) if so else (None, None)
    return _BACKEND


def get_backend():
    if _BACKEND is None:
        refresh_backend()
    return _BACKEND


def describe_backend():
    kind, info = get_backend()
    if kind == "office":
        parts = []
        for app in OFFICE_APPS:
            v = info.get(app)
            if v is not None:
                parts.append("%s %d" % (app, v))
        missing = [app for app in OFFICE_APPS if info.get(app) is None]
        text = "Microsoft Office：" + "、".join(parts) if parts else "Microsoft Office（组件不全）"
        if missing:
            text += "（未安装 " + "、".join(missing) + "）"
        return text
    if kind == "libreoffice":
        return "LibreOffice（备用引擎，支持 Word/Excel/PPT → PDF）"
    return "未检测到可用引擎"


def missing_apps_for_modes(modes):
    """返回当前后端下这些模式所缺少的 Office 应用名集合。

    PDF -> PPT / Excel 因为要先经 Word 重排，所以同时要求 Word + 目标应用。
    LibreOffice 后端不支持 PDF -> PPT / Excel（会由调用方单独报错）。
    """
    kind, info = get_backend()
    need = set()
    for m in modes:
        for app in MODE_APPS.get(m, ()):
            need.add(app)
    if kind is None:
        return need
    if kind == "libreoffice":
        return set()
    return set(app for app in need if info.get(app) is None)


def _backend_cli_args():
    """生成传给引擎的后端参数。

    注意：Office 后端已经不只“Word”了 —— 引擎会根据每个任务的 mode
    自行选择 Word / Excel / PowerPoint 实例，所以这里叫 "office"。
    """
    kind, info = get_backend()
    if kind == "libreoffice" and info:
        return ["-Backend", "libreoffice", "-Soffice", info]
    return ["-Backend", "office"]


def mode_of(path, pdf_mode="pdf2word"):
    ext = os.path.splitext(path)[1].lower()
    if ext in PDF_EXTS:
        return pdf_mode if pdf_mode in PDF_MODES else "pdf2word"
    if ext in WORD_EXTS:
        return "word2pdf"
    if ext in EXCEL_EXTS:
        return "excel2pdf"
    if ext in PPT_EXTS:
        return "ppt2pdf"
    return None


def out_ext(mode):
    if mode == "pdf2word":
        return ".docx"
    if mode == "pdf2ppt":
        return ".pptx"
    if mode == "pdf2excel":
        return ".xlsx"
    return ".pdf"


def unique_path(folder, base, ext):
    cand = os.path.join(folder, base + ext)
    n = 1
    while os.path.exists(cand):
        cand = os.path.join(folder, "%s (%d)%s" % (base, n, ext))
        n += 1
    return cand


def looks_scanned_pdf(path):
    try:
        size = os.path.getsize(path)
        head_limit = 8 * 1024 * 1024
        with open(path, "rb") as f:
            head = f.read(head_limit)
            tail = b""
            if size > head_limit:
                f.seek(max(0, size - 1024 * 1024))
                tail = f.read(1024 * 1024)
        data = head + tail
        if (b"/Font" in data) or (b"/Type0" in data) or (b"/TrueType" in data) or (b"/CFF" in data):
            return False
        return ((b"/Image" in data) or (b"/DCTDecode" in data)
                or (b"/JPXDecode" in data) or (b"/CCITTFaxDecode" in data))
    except Exception:
        return False


# =========================== 环境自检 ===========================

def _port_open(host, port, timeout=3.0):
    try:
        import socket
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception:
        return False


def check_environment(with_network=True):
    """返回环境项列表：[{key,name,ok,detail,fixable}]"""
    items = []

    ps_ok = bool(shutil.which("powershell")) or os.path.isfile(POWERSHELL)
    items.append({
        "key": "powershell",
        "name": "Windows PowerShell",
        "ok": ps_ok,
        "detail": POWERSHELL if ps_ok else "未找到 powershell.exe（系统组件，请修复系统或重装）",
        "fixable": False,
    })

    eng_ok = os.path.isfile(ENGINE_PS1)
    items.append({
        "key": "engine",
        "name": "转换引擎脚本",
        "ok": eng_ok,
        "detail": "convert_engine.ps1 就位" if eng_ok else "缺少 convert_engine.ps1（请保持程序文件完整）",
        "fixable": False,
    })

    kind, info = refresh_backend()
    if kind == "office":
        items.append({"key": "backend", "name": "转换后端", "ok": True,
                      "detail": describe_backend() + "（直连 COM，保真度最高）", "fixable": False})
    elif kind == "libreoffice":
        items.append({"key": "backend", "name": "转换后端", "ok": True,
                      "detail": "LibreOffice（备用引擎，PDF→Word 排版质量较低）", "fixable": False})
    else:
        items.append({"key": "backend", "name": "转换后端", "ok": False,
                      "detail": "未检测到 Microsoft Office 或 LibreOffice，无法转换",
                      "fixable": True})

    try:
        import platform
        try:
            import tkinter as _tkmod
            tk_ver = str(_tkmod.TkVersion)
        except Exception:
            tk_ver = "缺失"
        items.append({
            "key": "python",
            "name": "Python 运行环境",
            "ok": sys.version_info >= (3, 8),
            "detail": "Python %s (%s)，tkinter/Tk %s" % (
                platform.python_version(), platform.architecture()[0], tk_ver),
            "fixable": False,
        })
    except Exception:
        items.append({"key": "python", "name": "Python 运行环境",
                      "ok": True, "detail": "Python " + sys.version.split()[0], "fixable": False})

    tmp_ok = True
    tmp_detail = tempfile.gettempdir()
    try:
        fd, probe = tempfile.mkstemp(prefix="wp_probe_")
        os.close(fd)
        os.remove(probe)
    except Exception as exc:
        tmp_ok = False
        tmp_detail = "临时目录不可写：%s" % exc
    items.append({"key": "tempdir", "name": "临时目录可写", "ok": tmp_ok,
                  "detail": tmp_detail, "fixable": False})

    if with_network:
        # 本工具的转换完全在本地完成，不需要网络。
        # 只有“没有任何可用后端、必须让用户自己去下载安装”时，网络才算必要条件；
        # 已有 Word / LibreOffice 时网络永远判通过 —— 否则没网的用户每次启动都会被误报。
        kind, _ = get_backend()
        net_ok = _port_open("download.documentfoundation.org", 443)
        if kind is None:
            items.append({"key": "network", "name": "网络（下载安装包需要）", "ok": net_ok,
                          "detail": "可连接下载源" if net_ok else "未联网，请手动下载安装 LibreOffice / Office",
                          "fixable": False})
        else:
            items.append({"key": "network", "name": "网络（可选，不影响转换）", "ok": True,
                          "detail": "可连接下载源" if net_ok else "未联网（转换不需要网络）",
                          "fixable": False})
    return items


# 刻意不提供“自动下载 + 提权静默安装第三方软件”的能力。
# 本工具的定位是“零网络依赖的本地转换”；而自动下载并以管理员权限执行安装包，
# 会把安全责任与信任成本转嫁给用户，收益（省一次下载）远小于风险。
# 缺失转换后端时，只提供官网链接，由用户自行下载安装。


# =========================== 子进程 / Word 清理 ===========================

OFFICE_IMAGES = ("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE")


def _office_snapshot():
    """拿到 Word/Excel/PowerPoint 三类进程的 {pid: 窗口标题}（无窗口时为空字符串）。

    用 tasklist /V 一次拿到 PID 与窗口标题：相比启动 PowerShell，tasklist 冷启动快得多。
    tasklist 的 /FI 不支持 OR 条件，所以对每种进程各查一次。
    """
    snap = {}
    for image in OFFICE_IMAGES:
        try:
            proc = subprocess.run(
                ["tasklist", "/V", "/FI", "IMAGENAME eq " + image, "/FO", "CSV", "/NH"],
                capture_output=True, stdin=subprocess.DEVNULL,
                creationflags=CREATE_NO_WINDOW, timeout=15)
            for row in csv.reader(io.StringIO(decode_bytes(proc.stdout))):
                if len(row) < 9:
                    continue
                if row[0].strip().lower() != image.lower():
                    continue
                try:
                    pid = int(row[1].strip())
                except Exception:
                    continue
                title = row[-1].strip()
                if title.upper() == "N/A":
                    title = ""
                snap[pid] = title
        except Exception:
            continue
    return snap


def _list_office_pids():
    return list(_office_snapshot().keys())


def _kill_pids(pids):
    count = 0
    for pid in pids:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           stdin=subprocess.DEVNULL,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           creationflags=CREATE_NO_WINDOW, timeout=10)
            count += 1
        except Exception:
            pass
    return count


def _kill_tree(pid):
    try:
        subprocess.run(["taskkill", "/T", "/F", "/PID", str(pid)],
                       stdin=subprocess.DEVNULL,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                       creationflags=CREATE_NO_WINDOW, timeout=15)
    except Exception:
        pass


def _read_session_lines(path):
    try:
        if not path or not os.path.isfile(path):
            return []
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            return [ln.strip() for ln in f if ln.strip()]
    except Exception:
        return []


def _reap_owned_office(session_path, before_pids=None):
    """清理本引擎创建的 Office 进程（Word/Excel/PowerPoint），且绝不伤害用户文档。

    两道闸门：
      1) 引擎发现实例里有非本工具的文档时写 KEEP —— 一律不杀（引擎已把窗口切前台）。
      2) 即使上报了 PID，也只杀“无主窗口”的实例；带窗口的实例可能正承载用户文档
         （Office 会把用户双击的文档合并进已存在的实例），一律保留。
    返回 (killed, skipped, used_fallback, notice)
    """
    lines = _read_session_lines(session_path)
    if any(x.upper() == "KEEP" for x in lines):
        return [], [], False, ("检测到 Office 实例中存在用户文档，已保留该实例未关闭"
                               "（已切到前台，请手动保存/关闭）")
    reported = [int(x) for x in lines if x.isdigit()]
    if reported:
        snap = _office_snapshot()
        windowless = set(p for p, t in snap.items() if not t)
        target = [p for p in reported if p in snap and p in windowless]
        skipped = [p for p in reported if p in snap and p not in windowless]
        if target:
            _kill_pids(target)
        notice = ""
        if skipped:
            notice = "Office 进程 %s 带有窗口（可能承载用户文档），已保留未清理" % ",".join(str(p) for p in skipped)
        return target, skipped, False, notice
    return _reap_unowned_office(before_pids), [], True, ""


def _reap_unowned_office(before_pids):
    """兜底路径：只清理“本次新增 + 没有主窗口”的自动化 Office 进程。

    残余风险（已知并接受）：用户点“开始转换”后立刻双击打开 Office 文档，
    若在主窗口出现前（约 0.5~1 秒窗口期）被扫到，理论上存在被误杀的可能。
    该路径仅在引擎没有上报 session（崩溃 / 尚未创建实例）时才会走到，
    概率极低，且已通过“无主窗口”条件把有文档的实例排除在外。
    """
    if before_pids is None:
        return []
    snap = _office_snapshot()
    new_pids = [p for p in snap if p not in before_pids]
    if not new_pids:
        return []
    target = [p for p in new_pids if not snap.get(p)]
    if target:
        _kill_pids(target)
    return target


# =========================== 资源注册表 ===========================

class ResourceRegistry(object):
    """统一登记所有子进程资源，提供唯一的 shutdown 路径。"""

    def __init__(self):
        self._lock = threading.RLock()
        self._procs = set()
        self._dismissers = set()
        self._word_baseline = None

    def register_proc(self, proc):
        with self._lock:
            self._procs.add(proc)

    def unregister_proc(self, proc):
        with self._lock:
            self._procs.discard(proc)

    def register_dismisser(self, proc):
        with self._lock:
            self._dismissers.add(proc)

    def unregister_dismisser(self, proc):
        with self._lock:
            self._dismissers.discard(proc)

    def set_word_baseline(self):
        pids = _list_office_pids()
        with self._lock:
            self._word_baseline = pids
        return pids

    def take_word_baseline(self):
        with self._lock:
            baseline = self._word_baseline
            self._word_baseline = None
            return baseline

    def snapshot(self):
        with self._lock:
            return {"procs": len(self._procs), "dismissers": len(self._dismissers),
                    "baseline": list(self._word_baseline) if self._word_baseline else None}

    def shutdown_all(self):
        with self._lock:
            procs = list(self._procs)
            self._procs.clear()
            dismissers = list(self._dismissers)
            self._dismissers.clear()
        for p in procs:
            try:
                _kill_tree(p.pid)
            except Exception:
                pass
        for d in dismissers:
            try:
                d.kill()
            except Exception:
                pass


REGISTRY = ResourceRegistry()


def _start_dismisser(duration_seconds):
    if not os.path.isfile(DISMISS_PS1):
        return None
    try:
        proc = subprocess.Popen(
            [POWERSHELL, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", DISMISS_PS1, "-DurationSeconds", str(int(duration_seconds)),
             "-ParentPid", str(os.getpid())],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW)
    except Exception:
        return None
    REGISTRY.register_dismisser(proc)
    return proc


def _stop_dismisser(proc):
    if proc is None:
        return
    try:
        proc.kill()
    except Exception:
        pass
    REGISTRY.unregister_dismisser(proc)


def set_word_baseline():
    return REGISTRY.set_word_baseline()


def reap_baseline_word():
    baseline = REGISTRY.take_word_baseline()
    if baseline is None:
        return []
    return _reap_unowned_office(baseline)


# =========================== 单文件转换 ===========================

def convert_file(src, dst, mode, timeout=TIMEOUT, manage_dismisser=True):
    """调用引擎转换单个文件。返回 (ok, timed_out, log_text)。"""
    tmpdir = tempfile.mkdtemp(prefix="wpconv1_")
    session_path = os.path.join(tmpdir, "session.json")
    cmd = [POWERSHELL, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
           "-File", ENGINE_PS1, "-InputFile", src, "-OutputFile", dst, "-Mode", mode,
           "-SessionJson", session_path] + _backend_cli_args()

    kind, _ = get_backend()
    office_before = _list_office_pids() if kind == "office" else None
    dismisser = _start_dismisser(timeout + 90) if (manage_dismisser and mode == "pdf2word" and kind == "office") else None
    proc = None
    try:
        try:
            proc = subprocess.Popen(cmd, stdin=subprocess.DEVNULL,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    creationflags=CREATE_NO_WINDOW)
        except Exception as exc:
            return False, False, "无法启动转换引擎: %s" % exc
        REGISTRY.register_proc(proc)

        timed_out = False
        try:
            out_b, err_b = proc.communicate(timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out = True
            _kill_tree(proc.pid)
            try:
                out_b, err_b = proc.communicate(timeout=15)
            except Exception:
                out_b, err_b = b"", b""

        log = (decode_bytes(out_b) + decode_bytes(err_b)).strip()
        ok = (not timed_out) and proc.returncode == 0 \
            and os.path.isfile(dst) and os.path.getsize(dst) > 0

        if not ok:
            reaped, skipped, used_fallback, notice = _reap_owned_office(session_path, office_before)
            if reaped:
                log += "\n[已清理残留 Office 进程%s: %s]" % (
                    "（兜底判断）" if used_fallback else "", ",".join(str(p) for p in reaped))
            if notice:
                log += "\n[注意] " + notice
            if not log:
                log = "转换失败（输出文件未生成）"
        return ok, timed_out, log
    finally:
        if proc is not None:
            REGISTRY.unregister_proc(proc)
        _stop_dismisser(dismisser)
        shutil.rmtree(tmpdir, ignore_errors=True)


# =========================== 批量转换 ===========================

def convert_batch(jobs, timeout_per_file=TIMEOUT, on_progress=None,
                  should_cancel=None, on_notice=None):
    """批量转换。返回 (results, timed_out, cancelled, engine_log)。"""
    if not jobs:
        return [], False, False, ""

    tmpdir = tempfile.mkdtemp(prefix="wpconv_")
    jobs_path = os.path.join(tmpdir, "jobs.json")
    res_path = os.path.join(tmpdir, "results.jsonl")
    session_path = os.path.join(tmpdir, "session.json")
    kind, _ = get_backend()
    office_before = _list_office_pids() if kind == "office" else None
    has_pdf2word = any(j.get("mode") in PDF_MODES for j in jobs)
    dismisser = _start_dismisser(max(300, timeout_per_file * len(jobs) + 180)) \
        if (has_pdf2word and kind == "office") else None

    timed_out = False
    cancelled = False
    log_lines = []
    last_beat = [time.time()]
    stall_at = [time.time() + timeout_per_file]
    state = {"done": 0, "total": len(jobs)}
    proc = None
    try:
        with open(jobs_path, "w", encoding="utf-8") as f:
            json.dump(jobs, f, ensure_ascii=False)

        cmd = [POWERSHELL, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
               "-File", ENGINE_PS1, "-JobsJson", jobs_path, "-ResultJson", res_path,
               "-SessionJson", session_path] + _backend_cli_args()
        try:
            proc = subprocess.Popen(cmd, stdin=subprocess.DEVNULL,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    creationflags=CREATE_NO_WINDOW)
        except Exception as exc:
            return ([{"ok": False, "error": "无法启动转换引擎: %s" % exc} for _ in jobs],
                    False, False, "")
        REGISTRY.register_proc(proc)

        def _read_stdout():
            try:
                for raw in proc.stdout:
                    line = decode_bytes(raw).strip()
                    if not line:
                        continue
                    if line.startswith("PROGRESS "):
                        parts = line.split()
                        if len(parts) >= 3:
                            try:
                                idx, tot = int(parts[1]), int(parts[2])
                            except ValueError:
                                continue
                            ok_flag = (len(parts) >= 4 and parts[3] == "1")
                            state["done"] = idx
                            state["total"] = tot or len(jobs)
                            last_beat[0] = time.time()
                            stall_at[0] = time.time() + timeout_per_file
                            if on_progress is not None:
                                try:
                                    on_progress(idx, state["total"], ok_flag)
                                except Exception:
                                    pass
                    elif line.startswith("NOTICE "):
                        msg = line[len("NOTICE "):]
                        if on_notice is not None:
                            try:
                                on_notice("  ⚠ " + msg)
                            except Exception:
                                pass
                        else:
                            log_lines.append("  ⚠ " + msg)
                    else:
                        log_lines.append(line)
            except Exception:
                pass

        def _read_stderr():
            try:
                for raw in proc.stderr:
                    line = decode_bytes(raw).strip()
                    if line:
                        log_lines.append(line)
            except Exception:
                pass

        t_out = threading.Thread(target=_read_stdout, daemon=True)
        t_err = threading.Thread(target=_read_stderr, daemon=True)
        t_out.start()
        t_err.start()

        while True:
            if should_cancel is not None:
                try:
                    if should_cancel():
                        cancelled = True
                        break
                except Exception:
                    pass
            if proc.poll() is not None:
                break
            now = time.time()
            if now >= stall_at[0]:
                waited = int(now - last_beat[0])
                msg = "    ⏳ 已等待 %d 秒仍无进展，Word 可能被对话框阻塞…" % waited
                if now - last_beat[0] >= timeout_per_file:
                    if on_notice is not None:
                        try:
                            on_notice(msg)
                        except Exception:
                            pass
                    else:
                        log_lines.append(msg)
                stall_at[0] = now + STALL_NOTICE
            if now - last_beat[0] > (timeout_per_file + STALL_GRACE):
                timed_out = True
                break
            time.sleep(0.2)

        if cancelled or timed_out:
            _kill_tree(proc.pid)
        try:
            proc.wait(timeout=30)
        except Exception:
            _kill_tree(proc.pid)
        t_out.join(timeout=6)
        t_err.join(timeout=6)

        results = [None] * len(jobs)
        if os.path.isfile(res_path):
            try:
                with open(res_path, "r", encoding="utf-8-sig", errors="replace") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            obj = json.loads(line)
                        except Exception:
                            continue
                        idx = obj.get("index")
                        if not isinstance(idx, int) or not (0 <= idx < len(results)):
                            continue
                        ok = bool(obj.get("ok"))
                        err = str(obj.get("error") or "")
                        if ok:
                            outp = jobs[idx].get("output")
                            if not (outp and os.path.isfile(outp) and os.path.getsize(outp) > 0):
                                ok, err = False, (err or "输出文件未生成")
                        results[idx] = {"ok": ok, "error": err}
            except Exception:
                pass

        for i in range(len(results)):
            if results[i] is None:
                if cancelled:
                    results[i] = {"ok": False, "error": "已取消"}
                elif timed_out:
                    results[i] = {"ok": False, "error": "超时，未完成"}
                else:
                    results[i] = {"ok": False, "error": "未执行"}

        engine_log = "\n".join(log_lines).strip()
        if office_before is not None and (cancelled or timed_out or (proc.returncode or 0) != 0):
            reaped, skipped, used_fallback, notice = _reap_owned_office(session_path, office_before)
            if reaped:
                engine_log += "\n[已清理残留 Office 进程%s: %s]" % (
                    "（兜底判断）" if used_fallback else "", ",".join(str(p) for p in reaped))
            if notice:
                engine_log += "\n[注意] " + notice
        return results, timed_out, cancelled, engine_log
    finally:
        if proc is not None:
            REGISTRY.unregister_proc(proc)
            for stream in (proc.stdout, proc.stderr):
                try:
                    if stream is not None:
                        stream.close()
                except Exception:
                    pass
        _stop_dismisser(dismisser)
        shutil.rmtree(tmpdir, ignore_errors=True)


# =========================== CMYK 进度条 ===========================

class CmykProgress(tk.Canvas):
    """打印机 CMYK 四色方框进度条：C / M / Y / K 四个色块依次填充。"""

    SPEC = (("C", "#00A6E0"), ("M", "#E4007F"), ("Y", "#F5C400"), ("K", "#2B2B2B"))

    def __init__(self, master, width=380, height=24, bg=BG):
        super(CmykProgress, self).__init__(master, width=width, height=height,
                                           bg=bg, highlightthickness=0, bd=0)
        # 注意：不能叫 self._w / self._h —— 那是 tkinter 内部的控件路径名，会被覆盖导致 TclError
        self._cw = width
        self._ch = height
        self._value = 0.0
        self._draw()

    def set(self, pct):
        try:
            v = max(0.0, min(100.0, float(pct)))
        except Exception:
            return
        if abs(v - self._value) < 0.05:
            return
        self._value = v
        self._draw()

    def get(self):
        return self._value

    def reset(self):
        self.set(0)

    def _draw(self):
        self.delete("all")
        gap = 5
        bw = (self._cw - gap * 3) / 4.0
        for i, (label, color) in enumerate(self.SPEC):
            x0 = i * (bw + gap)
            x1 = x0 + bw
            self.create_rectangle(x0, 1, x1, self._ch - 1,
                                  outline="#c9d1de", width=1, fill="#ffffff")
            local = max(0.0, min(1.0, (self._value - i * 25.0) / 25.0))
            if local > 0:
                inner = (x1 - x0 - 2) * local
                self.create_rectangle(x0 + 1, 2, x0 + 1 + inner, self._ch - 2,
                                      fill=color, outline="")
            fg = "#4a5262"
            if local > 0.5:
                fg = "#4a3b00" if label == "Y" else "#ffffff"
            self.create_text((x0 + x1) / 2.0, self._ch / 2.0, text=label,
                             fill=fg, font=("Microsoft YaHei UI", 8, "bold"))


# =========================== 环境检测对话框 ===========================

class EnvDialog(tk.Toplevel):
    def __init__(self, master, items, on_retry):
        super(EnvDialog, self).__init__(master)
        self.title("环境检测")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        self._master = master
        self._on_retry = on_retry

        tk.Label(self, text="运行环境检测", bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 14, "bold")).pack(anchor="w", padx=18, pady=(14, 2))
        tk.Label(self, text="转换后端：" + describe_backend(), bg=BG, fg=MUTED,
                 font=("Microsoft YaHei UI", 9)).pack(anchor="w", padx=18, pady=(0, 10))

        body = tk.Frame(self, bg=CARD, highlightbackground="#dfe3ea", highlightthickness=1)
        body.pack(fill="both", expand=True, padx=18)
        for it in items:
            row = tk.Frame(body, bg=CARD)
            row.pack(fill="x", padx=10, pady=5)
            tk.Label(row, text=("✔" if it["ok"] else "✘"), bg=CARD,
                     fg=(OKC if it["ok"] else ERRC),
                     font=("Microsoft YaHei UI", 11, "bold"), width=2).pack(side="left")
            tk.Label(row, text=it["name"], bg=CARD, fg=TEXT, width=18, anchor="w",
                     font=("Microsoft YaHei UI", 9, "bold")).pack(side="left")
            tk.Label(row, text=it["detail"], bg=CARD, fg=MUTED, anchor="w",
                     wraplength=380, justify="left",
                     font=("Microsoft YaHei UI", 9)).pack(side="left", fill="x", expand=True)

        if get_backend()[0] is None:
            tk.Label(self, text="缺少转换引擎。本工具不会自动下载安装软件，请用下面的链接自行安装：",
                     bg=BG, fg=ERRC, font=("Microsoft YaHei UI", 9)).pack(anchor="w", padx=18, pady=(0, 6))
        btns = tk.Frame(self, bg=BG)
        btns.pack(fill="x", padx=18, pady=14)
        ttk.Button(btns, text="打开 LibreOffice 官网（免费）",
                   command=lambda: webbrowser.open(LO_HOME_PAGE)).pack(side="left")
        ttk.Button(btns, text="打开 Office 安装页",
                   command=lambda: webbrowser.open(OFFICE_PAGE)).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="重新检测", command=self._retry).pack(side="left", padx=(8, 0))
        ttk.Button(btns, text="关闭", command=self.destroy).pack(side="right")
        tip = ("提示：如果在别的电脑上双击没反应，请改用「启动WP互相转换工具.bat」启动 ——\n"
               "它会先检测 Python，缺失时可自动下载安装（本程序由 Python 编写，必须先有 Python 才能运行）。")
        tk.Label(self, text=tip, bg=BG, fg=MUTED, justify="left",
                 font=("Microsoft YaHei UI", 8)).pack(anchor="w", padx=18, pady=(0, 12))
        self.geometry("+%d+%d" % (master.winfo_rootx() + 80, master.winfo_rooty() + 60))

    def _retry(self):
        self.destroy()
        self._on_retry()


# =========================== 主窗口 ===========================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.configure(bg=BG)
        self.geometry("940x720")
        self.minsize(860, 640)
        self.rows = []
        self.q = queue.Queue()
        self.worker = None
        self.cancel_flag = False
        self._build()
        self._center()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(120, self._poll)
        self.after(200, lambda: self._env_check(auto=True))

    # ---------------- UI ----------------
    def _build(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TButton", font=("Microsoft YaHei UI", 9), padding=5)
        style.configure("TEntry", font=("Microsoft YaHei UI", 9))
        style.configure("Treeview", font=("Microsoft YaHei UI", 9), rowheight=26)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"))

        head = tk.Frame(self, bg=BG)
        head.pack(fill="x", padx=18, pady=(14, 2))
        tk.Label(head, text=APP_NAME, bg=BG, fg=TEXT,
                 font=("Microsoft YaHei UI", 17, "bold")).pack(side="left")
        tk.Label(head, text="   " + APP_SUB, bg=BG, fg=MUTED,
                 font=("Microsoft YaHei UI", 9)).pack(side="left", pady=(6, 0))
        tk.Label(head, text="v" + APP_VERSION, bg=BG, fg="#a3aab8",
                 font=("Consolas", 8)).pack(side="right", pady=(6, 0))

        tk.Label(self, bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 9),
                 text="支持 Word / Excel / PPT → PDF，以及 PDF → Word / PPT / Excel；可批量转换"
                 ).pack(fill="x", padx=18, pady=(0, 1))
        tk.Label(self, bg=BG, fg="#9a6a00", font=("Microsoft YaHei UI", 9),
                 text="提示：扫描版 PDF（纯图片）转 Word 只能还原为图片、无法提取文字"
                 ).pack(fill="x", padx=18, pady=(0, 8))

        bar = tk.Frame(self, bg=BG)
        bar.pack(fill="x", padx=18, pady=(0, 8))
        ttk.Button(bar, text="＋ 添加文件", command=self.add_files_dialog).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="添加文件夹", command=self.add_folder_dialog).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="移除选中", command=self.remove_selected).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="清空列表", command=self.clear_all).pack(side="left", padx=(0, 6))
        ttk.Button(bar, text="环境检测", command=lambda: self._env_check(auto=False)).pack(side="right")
        # PDF 可转三种目标，这里让用户选（对列表里所有 PDF 生效）
        self.pdf_target = ttk.Combobox(bar, values=list(PDF_TARGET_LABELS), width=13, state="readonly")
        self.pdf_target.set("Word")
        self.pdf_target.pack(side="right", padx=(4, 12))
        self.pdf_target.bind("<<ComboboxSelected>>", self._on_pdf_target_changed)
        tk.Label(bar, text="PDF 转换为：", bg=BG, fg=MUTED,
                 font=("Microsoft YaHei UI", 9)).pack(side="right")

        card = tk.Frame(self, bg=CARD, highlightbackground="#dfe3ea", highlightthickness=1)
        card.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        self.tree = ttk.Treeview(card, columns=("name", "mode", "status"),
                                 show="headings", selectmode="extended")
        self.tree.heading("name", text="文件名")
        self.tree.heading("mode", text="转换方向")
        self.tree.heading("status", text="状态")
        self.tree.column("name", width=560, anchor="w")
        self.tree.column("mode", width=130, anchor="center")
        self.tree.column("status", width=150, anchor="center")
        self.tree.tag_configure("ok", foreground=OKC)
        self.tree.tag_configure("err", foreground=ERRC)
        vsb = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=(2, 0), pady=2)
        vsb.pack(side="right", fill="y", pady=2)

        out = tk.Frame(self, bg=BG)
        out.pack(fill="x", padx=18, pady=(0, 6))
        self.out_mode = tk.IntVar(value=0)
        tk.Radiobutton(out, text="输出到源文件所在目录", variable=self.out_mode, value=0,
                       bg=BG, fg=TEXT, activebackground=BG, selectcolor=CARD,
                       font=("Microsoft YaHei UI", 9), command=self._toggle_out).pack(side="left")
        tk.Radiobutton(out, text="指定输出目录：", variable=self.out_mode, value=1,
                       bg=BG, fg=TEXT, activebackground=BG, selectcolor=CARD,
                       font=("Microsoft YaHei UI", 9), command=self._toggle_out
                       ).pack(side="left", padx=(16, 4))
        self.out_entry = ttk.Entry(out, state="disabled")
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.out_btn = ttk.Button(out, text="浏览…", command=self.choose_out, state="disabled")
        self.out_btn.pack(side="left")

        act = tk.Frame(self, bg=BG)
        act.pack(fill="x", padx=18, pady=(2, 6))
        self.btn_start = tk.Button(act, text="▶  开始转换", command=self.start,
                                   bg=ACCENT, fg="white", activebackground=ACCENT_DARK,
                                   activeforeground="white", relief="flat", bd=0,
                                   font=("Microsoft YaHei UI", 10, "bold"),
                                   padx=20, pady=7, cursor="hand2")
        self.btn_start.pack(side="left")
        self.progress = CmykProgress(act, width=400, height=24)
        self.progress.pack(side="left", padx=12)
        self.progress_label = tk.Label(act, text="0 / 0", bg=BG, fg=MUTED,
                                       font=("Microsoft YaHei UI", 9), width=10)
        self.progress_label.pack(side="left")

        logcard = tk.Frame(self, bg=CARD, highlightbackground="#dfe3ea", highlightthickness=1)
        logcard.pack(fill="both", expand=False, padx=18, pady=(0, 10))
        self.log = tk.Text(logcard, height=9, bg="#0f1420", fg="#d7e2f5", relief="flat",
                           font=("Consolas", 9), wrap="word", padx=8, pady=6, state="disabled")
        lsb = ttk.Scrollbar(logcard, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=lsb.set)
        self.log.pack(side="left", fill="both", expand=True, padx=(1, 0), pady=1)
        lsb.pack(side="right", fill="y", pady=1)

        self.status = tk.Label(self, text="就绪", anchor="w", bg="#e9edf5", fg=MUTED,
                               font=("Microsoft YaHei UI", 9), padx=10, pady=4)
        self.status.pack(fill="x", side="bottom")

    def _center(self):
        self.update_idletasks()
        w, h = 940, 720
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2), max(0, (sh - h) // 3)))

    # ---------------- helpers ----------------
    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_status(self, text):
        self.status.configure(text=text)

    def _toggle_out(self):
        on = (self.out_mode.get() == 1)
        self.out_entry.configure(state="normal" if on else "disabled")
        self.out_btn.configure(state="normal" if on else "disabled")

    # ---------------- 环境检测 ----------------
    def _env_check(self, auto=True):
        """环境检测一律在后台线程执行：网络探测最长 3 秒，绝不能阻塞 UI。"""
        self._log("环境检测中…（转换过程不需要网络；仅首次缺少 Python 时引导安装需要联网）")

        def work():
            try:
                # 启动自检不测网络：转换完全本地进行，没必要让没网的用户多等或被误报
                items = check_environment(with_network=not auto)
            except Exception as exc:
                items = [{"key": "selfcheck", "name": "环境检测", "ok": False,
                          "detail": "检测过程出错：%s" % exc, "fixable": False}]
            self.q.put(("env_items", (items, auto)))

        threading.Thread(target=work, daemon=True).start()

    def _show_env_result(self, items, auto):
        bad = [it for it in items if not it["ok"]]
        self._log("环境检测：转换后端 = %s" % describe_backend())
        if auto:
            if not bad:
                return
            self._log("  ⚠ 有 %d 项未通过，请点右上角「环境检测」查看详情" % len(bad))
            EnvDialog(self, items, on_retry=lambda: self._env_check(auto=False))
        else:
            EnvDialog(self, items, on_retry=lambda: self._env_check(auto=False))

    # ---------------- file list ----------------
    def _add_paths(self, paths):
        try:
            pdf_mode = PDF_TARGET_MODES.get(self.pdf_target.get(), "pdf2word")
        except Exception:
            pdf_mode = "pdf2word"
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if not os.path.isfile(p):
                continue
            mode = mode_of(p, pdf_mode)
            if mode is None:
                continue
            if any(r["path"].lower() == p.lower() for r in self.rows):
                continue
            iid = "r%d" % (len(self.rows) + 1 + int(time.time() * 1000) % 100000)
            arrow = MODE_ARROW.get(mode, mode)
            self.tree.insert("", "end", iid=iid, values=(os.path.basename(p), arrow, "待转换"))
            self.rows.append({"path": p, "mode": mode, "iid": iid, "arrow": arrow})
            added += 1
        if added:
            self._log("已添加 %d 个文件（当前共 %d 个）。" % (added, len(self.rows)))
        self._set_status("已选择 %d 个文件" % len(self.rows))

    def add_files_dialog(self):
        paths = filedialog.askopenfilenames(
            title="选择要转换的文件",
            filetypes=[
                ("Office / PDF 文件",
                 "*.docx *.doc *.docm *.dotx *.rtf *.xlsx *.xlsm *.xlsb *.xls "
                 "*.pptx *.pptm *.ppsm *.ppsx *.ppt *.pps *.pdf"),
                ("Word 文档", "*.docx *.doc *.docm *.dotx *.rtf"),
                ("Excel 表格", "*.xlsx *.xlsm *.xlsb *.xls"),
                ("PowerPoint 演示文稿", "*.pptx *.pptm *.ppsm *.ppsx *.ppt *.pps"),
                ("PDF 文件", "*.pdf"),
                ("所有文件", "*.*")])
        if paths:
            self._add_paths(list(paths))

    def add_folder_dialog(self):
        folder = filedialog.askdirectory(title="选择包含文件的文件夹")
        if not folder:
            return
        found = []
        for root, _dirs, files in os.walk(folder):
            for f in files:
                if mode_of(f):
                    found.append(os.path.join(root, f))
        if found:
            self._add_paths(found)
        else:
            messagebox.showinfo("提示", "该文件夹中未找到 Word 或 PDF 文件。")

    def _on_pdf_target_changed(self, _evt=None):
        """切换 PDF 输出目标后，把列表里已有的 PDF 行一起更新。"""
        new_mode = PDF_TARGET_MODES.get(self.pdf_target.get(), "pdf2word")
        changed = 0
        for r in self.rows:
            if os.path.splitext(r["path"])[1].lower() in PDF_EXTS:
                r["mode"] = new_mode
                r["arrow"] = MODE_ARROW.get(new_mode, new_mode)
                if self.tree.exists(r["iid"]):
                    status = self.tree.set(r["iid"], "status")
                    self.tree.item(r["iid"],
                                   values=(os.path.basename(r["path"]), r["arrow"], status))
                changed += 1
        if changed:
            label = self.pdf_target.get()
            self._log("PDF 输出目标切换为「%s」，已更新 %d 个 PDF 文件。" % (label, changed))
            if new_mode in ("pdf2ppt", "pdf2excel"):
                self._log("  说明：PDF → %s 是「内容级重建」——文字/表格/图片会保留，"
                          "但原有排版、字体、配色不会完整还原。"
                          % ("PPT" if new_mode == "pdf2ppt" else "Excel"))
            self._set_status("PDF 输出目标：%s" % label)

    def remove_selected(self):
        if self.worker and self.worker.is_alive():
            return
        sel = list(self.tree.selection())
        for iid in sel:
            self.rows = [r for r in self.rows if r["iid"] != iid]
            self.tree.delete(iid)
        if sel:
            self._set_status("已选择 %d 个文件" % len(self.rows))

    def clear_all(self):
        if self.worker and self.worker.is_alive():
            return
        self.rows = []
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.progress.reset()
        self.progress_label.configure(text="0 / 0")
        self._set_status("就绪")

    def choose_out(self):
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, folder)

    # ---------------- 窗口关闭 ----------------
    def _on_close(self):
        self.cancel_flag = True
        try:
            REGISTRY.shutdown_all()
        except Exception:
            pass
        try:
            self.withdraw()
        except Exception:
            pass
        try:
            reap_baseline_word()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    # ---------------- conversion ----------------
    def _set_busy(self, busy):
        if busy:
            self.btn_start.configure(state="normal", text="⏹ 取消转换",
                                     bg=DANGER, activebackground=DANGER_DARK,
                                     command=self.cancel)
        else:
            self.btn_start.configure(state="normal", text="▶  开始转换",
                                     bg=ACCENT, activebackground=ACCENT_DARK,
                                     command=self.start)

    def cancel(self):
        if self.worker and self.worker.is_alive():
            self.cancel_flag = True
            self._set_status("正在取消，已中断当前批次…")
            self._log(">>> 收到取消请求，正在中断当前批次并保留已完成结果…")

    def _reset_rows_status(self):
        for r in self.rows:
            if self.tree.exists(r["iid"]):
                self.tree.item(r["iid"], values=(os.path.basename(r["path"]), r["arrow"], "待转换"),
                               tags=())

    def start(self):
        if self.worker and self.worker.is_alive():
            return
        if not self.rows:
            messagebox.showinfo("提示", "请先添加要转换的文件。")
            return
        kind, info = refresh_backend()
        if kind is None:
            messagebox.showerror("无法转换",
                                 "未检测到可用的转换引擎。\n\n"
                                 "本工具不会自动下载安装软件，请：\n"
                                 "  1) 安装 Microsoft Word，或\n"
                                 "  2) 点「环境检测」取得 LibreOffice 官网下载链接，自行安装后重试。")
            self._env_check(auto=False)
            return
        # 按实际要做的模式，检查所需 Office 组件是否齐全
        modes = set(r["mode"] for r in self.rows)
        if kind == "libreoffice" and (modes & set(["pdf2ppt", "pdf2excel"])):
            messagebox.showerror(
                "当前后端不支持该方向",
                "PDF → PPT / PDF → Excel 需要 Microsoft Office。\n\n"
                "Excel 与 PowerPoint 的 COM 接口无法直接打开 PDF，\n"
                "LibreOffice 备用后端也不能完成这类转换。")
            return
        missing = missing_apps_for_modes(modes)
        if missing:
            messagebox.showerror(
                "缺少所需的 Office 组件",
                "当前任务需要以下组件，但本机未检测到：\n\n    " + "、".join(sorted(missing)) +
                "\n\n请安装对应的 Microsoft Office 组件，或安装 LibreOffice 作为备用引擎。")
            self._env_check(auto=False)
            return
        if kind == "office" and (modes & set(PDF_MODES)):
            word_ver = info.get("Word")
            if word_ver is not None and word_ver < 15:
                messagebox.showerror("无法转换 PDF",
                                     "检测到 Word %d.x。\n"
                                     "PDF 读取/重排需要 Word 2013 及以上，请升级 Office 后再试。" % word_ver)
                return
        out_dir = ""
        if self.out_mode.get() == 1:
            out_dir = self.out_entry.get().strip()
            if not out_dir:
                messagebox.showwarning("提示", "请选择输出目录。")
                return
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as exc:
                messagebox.showerror("错误", "无法创建输出目录：%s" % exc)
                return
        self.cancel_flag = False
        self.progress.reset()
        self.progress_label.configure(text="0 / 0")
        self._reset_rows_status()
        self._set_busy(True)
        self.worker = threading.Thread(target=self._work, args=(list(self.rows), out_dir), daemon=True)
        self.worker.start()

    def _work(self, rows, out_dir):
        jobs = []
        for row in rows:
            src, mode = row["path"], row["mode"]
            base = os.path.splitext(os.path.basename(src))[0]
            folder = out_dir or os.path.dirname(src)
            dst = unique_path(folder, base, out_ext(mode))
            jobs.append({"row": row, "input": src, "output": dst, "mode": mode})

        total = len(jobs)
        set_word_baseline()
        self.q.put(("log", "=" * 52))
        self.q.put(("log", "共 %d 个文件，后端：%s" % (total, describe_backend())))

        for j in jobs:
            if j["mode"] in PDF_MODES and looks_scanned_pdf(j["input"]):
                self.q.put(("log", "  ℹ %s 可能是扫描版 PDF：只能还原为图片，文字无法提取。"
                            % os.path.basename(j["input"])))

        ok_count = 0
        fail = []
        done = 0
        t0 = time.time()
        i = 0
        while i < total:
            if self.cancel_flag:
                self.q.put(("log", "已取消：剩余 %d 个文件未处理。" % (total - i)))
                break
            chunk = jobs[i:i + CHUNK_SIZE]
            offset = i
            for j in chunk:
                self.q.put(("row", (j["row"]["iid"], "排队中…", "", j["row"]["arrow"])))
            self.q.put(("log", "[%d-%d / %d] 批量转换中…" % (offset + 1, offset + len(chunk), total)))

            def _on_prog(done_idx, batch_total, ok_flag, offset=offset, chunk=chunk):
                if 1 <= done_idx <= len(chunk):
                    j = chunk[done_idx - 1]
                    if ok_flag:
                        self.q.put(("row", (j["row"]["iid"], "✓ 完成", "ok", j["row"]["arrow"])))
                    else:
                        self.q.put(("row", (j["row"]["iid"], "✗ 失败", "err", j["row"]["arrow"])))
                g = offset + done_idx
                self.q.put(("progress", (g * 100.0 / total, "%d / %d" % (g, total))))

            def _on_notice(msg):
                self.q.put(("log", msg))

            payload = [{"input": j["input"], "output": j["output"], "mode": j["mode"]} for j in chunk]
            results, timed_out, cancelled, engine_log = convert_batch(
                payload, timeout_per_file=TIMEOUT,
                on_progress=_on_prog, should_cancel=lambda: self.cancel_flag,
                on_notice=_on_notice)

            if engine_log:
                self.q.put(("log", engine_log))

            if cancelled:
                for j, res in zip(chunk, results):
                    if (res or {}).get("ok"):
                        ok_count += 1
                        self.q.put(("row", (j["row"]["iid"], "✓ 完成", "ok", j["row"]["arrow"])))
                    else:
                        self.q.put(("row", (j["row"]["iid"], "已取消", "", j["row"]["arrow"])))
                    done += 1
                self.q.put(("log", "已按请求中断当前批次，已完成的结果已保留。"))
                break

            if (timed_out or all(not (r or {}).get("ok") for r in results)) and len(chunk) > 1:
                self.q.put(("log", "    ⚠ 批次异常，改为逐文件转换以定位问题…"))
                base_done = done

                def _fb_prog(n, n_total, idx, base_done=base_done):
                    g = base_done + n
                    self.q.put(("progress", (g * 100.0 / total, "%d / %d" % (g, total))))

                results = self._fallback_per_file(chunk, results, progress_cb=_fb_prog)

            for j, res in zip(chunk, results):
                ok = bool((res or {}).get("ok"))
                if ok:
                    ok_count += 1
                    self.q.put(("row", (j["row"]["iid"], "✓ 完成", "ok", j["row"]["arrow"])))
                    self.q.put(("log", "      ✓ %s → %s" % (os.path.basename(j["input"]),
                                                            os.path.basename(j["output"]))))
                else:
                    fail.append(os.path.basename(j["input"]))
                    self.q.put(("row", (j["row"]["iid"], "✗ 失败", "err", j["row"]["arrow"])))
                    msg = str((res or {}).get("error") or "转换失败").splitlines()
                    self.q.put(("log", "      ✗ %s：%s" % (os.path.basename(j["input"]),
                                                          (msg[0] if msg else "转换失败")[:220])))
                done += 1
                self.q.put(("progress", (done * 100.0 / total, "%d / %d" % (done, total))))
            i += len(chunk)

        self.q.put(("done", (ok_count, total, fail, time.time() - t0)))

    def _fallback_per_file(self, chunk, partial, progress_cb=None):
        merged = list(partial)
        todo = [idx for idx, r in enumerate(merged) if not (r or {}).get("ok")]
        if not todo:
            return merged
        kind, _ = get_backend()
        dismisser = _start_dismisser(len(todo) * TIMEOUT + 120) if kind == "office" else None
        try:
            for n, idx in enumerate(todo, 1):
                if self.cancel_flag:
                    break
                j = chunk[idx]
                ok, timed_out, log = convert_file(j["input"], j["output"], j["mode"],
                                                  manage_dismisser=False)
                if ok:
                    merged[idx] = {"ok": True, "error": ""}
                    if log:
                        # 成功时也可能带 NOTICE/清理提示，不能丢
                        self.q.put(("log", "  " + log.splitlines()[0][:200]))
                else:
                    err = "超时（Word 可能被对话框阻塞）" if timed_out else (log.splitlines()[0] if log else "转换失败")
                    merged[idx] = {"ok": False, "error": err}
                if progress_cb is not None:
                    try:
                        progress_cb(n, len(todo), idx)
                    except Exception:
                        pass
        finally:
            _stop_dismisser(dismisser)
        return merged

    # ---------------- queue pump ----------------
    def _handle(self, kind, payload):
        if kind == "log":
            self._log(payload)
        elif kind == "row":
            iid, text, tag, arrow = payload
            if self.tree.exists(iid):
                name = self.tree.set(iid, "name")
                self.tree.item(iid, values=(name, arrow, text), tags=(tag,) if tag else ())
        elif kind == "progress":
            self.progress.set(payload[0])
            self.progress_label.configure(text=payload[1])
        elif kind == "env_items":
            items, auto = payload
            self._show_env_result(items, auto)
        elif kind == "done":
            self._on_done(payload)

    def _poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                try:
                    self._handle(kind, payload)
                except Exception:
                    pass
        except queue.Empty:
            pass
        except Exception:
            pass
        finally:
            try:
                self.after(100, self._poll)
            except Exception:
                pass

    def _on_done(self, payload):
        ok, total, fail, elapsed = payload
        self._set_busy(False)
        # 列表已处理完毕：进度条推到 100%，标签显示真实成功数，避免"停在中间像没做完"
        self.progress.set(100)
        self.progress_label.configure(text="完成 %d / %d" % (ok, total))
        skipped = max(0, total - ok - len(fail))
        self._log("=" * 52)
        self._log("转换结束：成功 %d / 共 %d，失败 %d，跳过 %d，用时 %.1f 秒"
                  % (ok, total, len(fail), skipped, elapsed))
        if self.cancel_flag:
            self._set_status("已取消：成功 %d 个" % ok)
            messagebox.showinfo("已取消", "转换已取消。\n\n已完成 %d 个文件。" % ok)
            return
        if not fail:
            self._set_status("全部完成：%d 个文件" % ok)
            messagebox.showinfo("完成", "全部转换成功！\n\n成功 %d 个文件，用时 %.1f 秒。" % (ok, elapsed))
        else:
            self._set_status("完成 %d 个，失败 %d 个" % (ok, len(fail)))
            preview = "\n".join(fail[:10])
            messagebox.showwarning("部分失败",
                                   "成功 %d 个，失败 %d 个。\n\n失败文件：\n%s" % (ok, len(fail), preview))


def main():
    App().mainloop()


if __name__ == "__main__":
    if "--env-check" in sys.argv:
        # 说明：GUI 的启动自检不含网络探测（转换本来就不需要网络），
        # 而 --env-check 是完整诊断，默认包含网络；用 --no-network 可关闭。
        with_net = "--no-network" not in sys.argv
        print("== 环境诊断（%s网络探测）==" % ("含" if with_net else "不含"))
        for it in check_environment(with_network=with_net):
            print(("  [OK]  " if it["ok"] else "  [FAIL]") + " " + it["name"] + " : " + it["detail"])
        print("backend =", describe_backend())
        sys.exit(0)
    if "--convert" in sys.argv:
        idx = sys.argv.index("--convert")
        try:
            src, dst = sys.argv[idx + 1], sys.argv[idx + 2]
        except IndexError:
            print("usage: wp_converter.py --convert <input> <output>")
            print("       支持 .docx/.doc/.rtf→PDF、.xlsx/.xls→PDF、.pptx/.ppt→PDF、.pdf→.docx")
            sys.exit(2)
        m = sys.argv[idx + 3] if len(sys.argv) > idx + 3 else (mode_of(src) or "word2pdf")
        if m not in MODE_ARROW:
            print("不认识的模式:", m)
            print("可选: " + ", ".join(sorted(MODE_ARROW)))
            sys.exit(2)
        ok, to, lg = convert_file(src, dst, m)
        print("RESULT:", "OK" if ok else "FAIL", "| timeout:", to)
        if lg:
            print(lg)
        sys.exit(0 if ok else 1)
    main()
