import os
import sys
import threading
import queue
import locale
import glob
import shutil
import tempfile
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from concurrent.futures import ThreadPoolExecutor, as_completed

# ------------------ 高 DPI 与资源路径 ------------------

def enable_high_dpi():
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
            return
        except Exception:
            pass
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            return
        except Exception:
            pass
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    except Exception:
        pass


def resource_path(relative_path: str) -> str:
    """
    获取资源文件路径（兼容 PyInstaller 冻结后的临时目录）。
    """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


# ------------------ 可选拖拽支持 ------------------

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

# ------------------ 引用后端逻辑 ------------------

from main import DigitalSignatureTool, FileFormats, SignatureStatus, SigningConfig


def detect_lang():
    """
    语言检测（优先系统显示语言）。
    Windows：
      1) GetUserPreferredUILanguages（zh-CN/zh-TW/en-US 等）
      2) GetUserDefaultLocaleName
    跨平台回退：
      3) locale.getlocale()
      4) 环境变量（LC_ALL/LANG/LANGUAGE）
    """
    code = ""

    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes

            MUI_LANGUAGE_NAME = 0x8
            GetUserPreferredUILanguages = ctypes.windll.kernel32.GetUserPreferredUILanguages
            GetUserPreferredUILanguages.argtypes = [
                wintypes.DWORD,
                ctypes.POINTER(wintypes.ULONG),
                wintypes.LPWSTR,
                ctypes.POINTER(wintypes.ULONG),
            ]
            GetUserPreferredUILanguages.restype = wintypes.BOOL

            num = wintypes.ULONG(0)
            buflen = wintypes.ULONG(0)
            if GetUserPreferredUILanguages(MUI_LANGUAGE_NAME, ctypes.byref(num), None, ctypes.byref(buflen)):
                buf = ctypes.create_unicode_buffer(buflen.value)
                if GetUserPreferredUILanguages(MUI_LANGUAGE_NAME, ctypes.byref(num), buf, ctypes.byref(buflen)):
                    parts = buf[:buflen.value].split("\x00")
                    first = next((p for p in parts if p), "")
                    if first:
                        code = first  # zh-CN/zh-TW/en-US
        except Exception:
            pass

        if not code:
            try:
                import ctypes
                buf = ctypes.create_unicode_buffer(85)
                if ctypes.windll.kernel32.GetUserDefaultLocaleName(buf, 85):
                    code = buf.value  # zh-CN
            except Exception:
                pass

    if not code:
        try:
            try:
                locale.setlocale(locale.LC_CTYPE, "")
            except Exception:
                pass
            loc = locale.getlocale()
            if loc and loc[0]:
                code = loc[0]
        except Exception:
            pass

    if not code:
        for env in ("LC_ALL", "LANG", "LANGUAGE"):
            v = os.environ.get(env, "")
            if v:
                code = v
                break

    code = (code or "").lower()
    if code.startswith("zh"):
        if "tw" in code or "hk" in code or "mo" in code or "hant" in code:
            return "zh_TW"
        return "zh_CN"
    return "en"


I18N = {
    "zh_CN": {
        "app_title": "数字签名生成/签名程序(非认证) v0.0.1.0",
        "pending_files": "待处理文件",
        "add_files_btn": "添加文件",
        "remove_selected_btn": "移除选中",
        "clear_list_btn": "清空列表",
        "cert_ts": "证书与时间戳",
        "pfx_file": "PFX 文件:",
        "browse": "浏览…",
        "password": "密码:",
        "timestamp_server": "时间戳服务器:",
        "create_self_signed_btn": "创建自签名 PFX…",
        "create_cer_btn": "仅创建安全证书 (.cer 文件)",
        "verify_btn": "验证签名",
        "sign_btn": "签名并加时间戳",
        "sign_no_ts_btn": "签名（不加时间戳）",
        "timestamp_btn": "仅添加时间戳",
        "log_title": "日志",
        "select_files_title": "选择文件",
        "supported_formats": "支持的格式",
        "all_files": "所有文件",
        "added_files": "已添加 {n} 个文件。",
        "removed_selected": "已移除选中项。",
        "list_cleared": "列表已清空。",
        "no_files": "请先添加至少一个文件。",
        "need_valid_pfx": "请先选择有效的 .pfx 文件。",
        "start_verify": "开始验证 {n} 个文件…",
        "verifying": "[{i}/{n}] 验证：{name}",
        "result": "  结果：{status}",
        "stats": "验证完成。统计：",
        "trusted_friendly": "受信任的签名",
        "self_signed_friendly": "自签名证书（未经认证）",
        "unsigned_friendly": "未签名（程序不存在证书）",
        "invalid_friendly": "签名无效或证书错误",
        "unknown_friendly": "未知状态",
        "signer": "签名者",
        "issuer": "颁发者",
        "timestamp": "时间戳",
        "signing": "[{i}/{n}] 签名：{name}",
        "signing_no_index": "签名：{name}",
        "done": "  ✓ 完成",
        "sign_all_done": "全部签名完成。",
        "start_timestamp": "开始为 {n} 个文件添加时间戳…",
        "timestamp_item": "[{i}/{n}] 时间戳：{name}",
        "timestamp_done": "时间戳添加完成。",
        "create_self_signed": "开始创建自签名证书…",
        "self_signed_done": "自签名 PFX 创建完成。桌面已保存副本：Key.pfx",
        "self_signed_note": "注意：这是自签名证书，未经权威机构认证。",
        "create_pfx_failed": "创建 PFX 失败。",
        "create_cert_failed": "创建证书失败：{err}",
        "drag_tip": "可将文件拖拽到列表中添加。",
        "drag_not_available": "拖放功能不可用：未安装 tkinterdnd2。可通过 pip install tkinterdnd2 启用。",
        "concurrency": "并发数:",
        "stats_item": "  {label}：{n} 个",
        "concurrency_prompt": "检测到选择了 {n} 个文件，将使用多线程（并发数：{workers}）进行签名（不加时间戳）。是否继续？",
        "seq_info_ts": "为避免时间戳服务器限流，本操作将按顺序处理。",
        "cer_done": "CER 证书已创建并复制到桌面：{name}",
        "cer_not_found": "未找到生成的 CER 文件。",
        "enter_pwd": "请输入 PFX 密码（{name}）：",
    },
    "zh_TW": {
        "app_title": "數位簽章產生/簽署程式（非認證） v0.0.1.0",
        "pending_files": "待處理檔案",
        "add_files_btn": "新增檔案",
        "remove_selected_btn": "移除選取",
        "clear_list_btn": "清空列表",
        "cert_ts": "憑證與時間戳",
        "pfx_file": "PFX 檔案:",
        "browse": "瀏覽…",
        "password": "密碼:",
        "timestamp_server": "時間戳伺服器:",
        "create_self_signed_btn": "建立自簽名 PFX…",
        "create_cer_btn": "僅建立安全憑證 (.cer 檔案)",
        "verify_btn": "驗證簽名",
        "sign_btn": "簽名並加時間戳",
        "sign_no_ts_btn": "簽名（不加時間戳）",
        "timestamp_btn": "僅新增時間戳",
        "log_title": "日誌",
        "select_files_title": "選擇檔案",
        "supported_formats": "支援的格式",
        "all_files": "所有檔案",
        "added_files": "已新增 {n} 個檔案。",
        "removed_selected": "已移除選取項。",
        "list_cleared": "列表已清空。",
        "no_files": "請先新增至少一個檔案。",
        "need_valid_pfx": "請先選擇有效的 .pfx 檔案。",
        "start_verify": "開始驗證 {n} 個檔案…",
        "verifying": "[{i}/{n}] 驗證：{name}",
        "result": "  結果：{status}",
        "stats": "驗證完成。統計：",
        "trusted_friendly": "受信任的簽名",
        "self_signed_friendly": "自簽名憑證（未經認證）",
        "unsigned_friendly": "未簽名（程式不存在憑證）",
        "invalid_friendly": "簽名無效或憑證錯誤",
        "unknown_friendly": "未知狀態",
        "signer": "簽名者",
        "issuer": "簽發者",
        "timestamp": "時間戳",
        "signing": "[{i}/{n}] 簽名：{name}",
        "signing_no_index": "簽名：{name}",
        "done": "  ✓ 完成",
        "sign_all_done": "全部簽名完成。",
        "start_timestamp": "開始為 {n} 個檔案新增時間戳…",
        "timestamp_item": "[{i}/{n}] 時間戳：{name}",
        "timestamp_done": "時間戳新增完成。",
        "create_self_signed": "開始建立自簽名憑證…",
        "self_signed_done": "自簽名 PFX 建立完成。桌面已保存副本：Key.pfx",
        "self_signed_note": "注意：這是自簽名憑證，未經權威機構認證。",
        "create_pfx_failed": "建立 PFX 失敗。",
        "create_cert_failed": "建立憑證失敗：{err}",
        "drag_tip": "可將檔案拖曳到列表中新增。",
        "drag_not_available": "拖放功能不可用：未安裝 tkinterdnd2。可透過 pip install tkinterdnd2 啟用。",
        "concurrency": "並發數:",
        "stats_item": "  {label}：{n} 個",
        "concurrency_prompt": "偵測到選取 {n} 個檔案，將以多執行緒（並發數：{workers}）進行簽名（不加時間戳）。是否繼續？",
        "seq_info_ts": "為避免時間戳伺服器限流，本操作將按順序處理。",
        "cer_done": "CER 憑證已建立並複製到桌面：{name}",
        "cer_not_found": "找不到產生的 CER 檔案。",
        "enter_pwd": "請輸入 PFX 密碼（{name}）：",
    },
    "en": {
        "app_title": "Digital Signature Generator/Signer (Non-certified) v0.0.1.0",
        "pending_files": "Pending Files",
        "add_files_btn": "Add Files",
        "remove_selected_btn": "Remove Selected",
        "clear_list_btn": "Clear List",
        "cert_ts": "Certificate & Timestamp",
        "pfx_file": "PFX File:",
        "browse": "Browse…",
        "password": "Password:",
        "timestamp_server": "Timestamp Server:",
        "create_self_signed_btn": "Create Self-signed PFX…",
        "create_cer_btn": "Create CER certificate only",
        "verify_btn": "Verify Signatures",
        "sign_btn": "Sign + Timestamp",
        "sign_no_ts_btn": "Sign (no timestamp)",
        "timestamp_btn": "Timestamp Only",
        "log_title": "Log",
        "select_files_title": "Select Files",
        "supported_formats": "Supported formats",
        "all_files": "All files",
        "added_files": "{n} file(s) added.",
        "removed_selected": "Selected items removed.",
        "list_cleared": "List cleared.",
        "no_files": "Please add at least one file.",
        "need_valid_pfx": "Please select a valid .pfx file first.",
        "start_verify": "Verifying {n} file(s)…",
        "verifying": "[{i}/{n}] Verify: {name}",
        "result": "  Result: {status}",
        "stats": "Verification summary:",
        "trusted_friendly": "Trusted signature",
        "self_signed_friendly": "Self-signed certificate (not CA-issued)",
        "unsigned_friendly": "Unsigned (no certificate present)",
        "invalid_friendly": "Invalid signature or certificate error",
        "unknown_friendly": "Unknown status",
        "signer": "Signer",
        "issuer": "Issuer",
        "timestamp": "Timestamp",
        "signing": "[{i}/{n}] Sign: {name}",
        "signing_no_index": "Sign: {name}",
        "done": "  ✓ Done",
        "sign_all_done": "All signatures completed.",
        "start_timestamp": "Timestamping {n} file(s)…",
        "timestamp_item": "[{i}/{n}] Timestamp: {name}",
        "timestamp_done": "Timestamping completed.",
        "create_self_signed": "Creating self-signed certificate…",
        "self_signed_done": "Self-signed PFX created. A copy has been saved to Desktop: Key.pfx",
        "self_signed_note": "Note: This is a self-signed certificate (not CA-issued).",
        "create_pfx_failed": "Failed to create PFX.",
        "create_cert_failed": "Failed to create certificate: {err}",
        "drag_tip": "Drag and drop files into the list.",
        "drag_not_available": "Drag-and-drop disabled: tkinterdnd2 not installed. Install with: pip install tkinterdnd2",
        "concurrency": "Concurrency:",
        "stats_item": "  {label}: {n}",
        "concurrency_prompt": "You selected {n} files. Signing without timestamp will run concurrently (workers: {workers}). Continue?",
        "seq_info_ts": "To avoid TSA rate limits, this operation runs sequentially.",
        "cer_done": "CER certificate created and copied to Desktop: {name}",
        "cer_not_found": "Generated CER file not found.",
        "enter_pwd": "Enter PFX password ({name}):",
    },
}


class App((TkinterDnD.Tk if DND_AVAILABLE else tk.Tk)):
    def __init__(self):
        super().__init__()

        # 语言
        self.lang = detect_lang()
        self.t = lambda k, **kw: (I18N.get(self.lang, I18N["en"]).get(k, I18N["en"].get(k, k))).format(**kw)

        # 标题与图标（随语言）
        self.title(self.t("app_title"))
        self._set_icon()

        self.geometry("1000x680")
        self.minsize(880, 560)

        # 颜色标签
        self.TAG_COLORS = {
            "ok": "#2e7d32",
            "warn": "#b26a00",
            "error": "#c62828",
            "info": "#2e3b4e",
        }

        # 后端实例与工具检查
        self.tool = DigitalSignatureTool()
        if not self.tool._check_tools():
            messagebox.showerror("Error", f"Missing tools folder:\n{self.tool.tools_path}")
            self.destroy()
            return

        # 状态
        self.selected_files = []
        self.pfx_path_var = tk.StringVar(value="")
        self.pfx_pwd_var = tk.StringVar(value="")
        self.ts_server_var = tk.StringVar(value=self.tool.current_timestamp_url)
        self.workers_var = tk.IntVar(value=min(4, (os.cpu_count() or 4)))

        # 额外：缓存每个 PFX 的密码（用户输入一次后复用）
        self._pfx_pwd_cache = {}

        # 消息队列
        self.msg_queue = queue.Queue()

        # 构建 UI
        self._build_ui()

        # 拖拽提示
        if DND_AVAILABLE:
            self._tip_text = self.t("drag_tip")
            self._tip_tag = "info"
        else:
            self._tip_text = self.t("drag_not_available")
            self._tip_tag = "warn"
        self._log(self._tip_text, tag=self._tip_tag)

        # 定时处理队列
        self.after(100, self._process_queue)

    # ------------------ UI ------------------

    def _build_ui(self):
        # 顶部：文件列表
        file_frame = ttk.LabelFrame(self, text=self.t("pending_files"))
        file_frame.pack(side=tk.TOP, fill=tk.BOTH, padx=10, pady=8)

        self.file_listbox = tk.Listbox(file_frame, height=8, selectmode=tk.EXTENDED)
        self.file_listbox.grid(row=0, column=0, rowspan=4, sticky="nsew", padx=(8, 0), pady=8)

        file_scroll = ttk.Scrollbar(file_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        file_scroll.grid(row=0, column=1, rowspan=4, sticky="ns", pady=8, padx=(0, 8))
        self.file_listbox.config(yscrollcommand=file_scroll.set)

        # 支持拖放
        if DND_AVAILABLE:
            self.file_listbox.drop_target_register(DND_FILES)
            self.file_listbox.dnd_bind("<<Drop>>", self._on_drop_files)

        # 右侧按钮
        btn_add = ttk.Button(file_frame, text=self.t("add_files_btn"), command=self._on_add_files)
        btn_del = ttk.Button(file_frame, text=self.t("remove_selected_btn"), command=self._on_remove_selected)
        btn_clear = ttk.Button(file_frame, text=self.t("clear_list_btn"), command=self._on_clear_files)
        btn_add.grid(row=0, column=2, padx=8, pady=(12, 4), sticky="ew")
        btn_del.grid(row=1, column=2, padx=8, pady=4, sticky="ew")
        btn_clear.grid(row=2, column=2, padx=8, pady=(4, 12), sticky="ew")

        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(3, weight=1)

        # 中部：证书与时间戳
        cfg_frame = ttk.LabelFrame(self, text=self.t("cert_ts"))
        cfg_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=6)

        ttk.Label(cfg_frame, text=self.t("pfx_file")).grid(row=0, column=0, padx=(8, 4), pady=8, sticky="e")
        pfx_entry = ttk.Entry(cfg_frame, textvariable=self.pfx_path_var)
        pfx_entry.grid(row=0, column=1, padx=4, pady=8, sticky="ew")
        ttk.Button(cfg_frame, text=self.t("browse"), command=self._on_browse_pfx).grid(row=0, column=2, padx=8, pady=8)

        ttk.Label(cfg_frame, text=self.t("password")).grid(row=0, column=3, padx=(18, 4), pady=8, sticky="e")
        pwd_entry = ttk.Entry(cfg_frame, textvariable=self.pfx_pwd_var, show="•")
        pwd_entry.grid(row=0, column=4, padx=4, pady=8, sticky="ew")

        ttk.Label(cfg_frame, text=self.t("timestamp_server")).grid(row=1, column=0, padx=(8, 4), pady=8, sticky="e")
        self.ts_combo = ttk.Combobox(cfg_frame, values=self.tool.TIMESTAMP_URLS, textvariable=self.ts_server_var, state="readonly")
        self.ts_combo.grid(row=1, column=1, columnspan=2, padx=4, pady=8, sticky="ew")

        btn_create_pfx = ttk.Button(cfg_frame, text=self.t("create_self_signed_btn"), command=self._on_create_self_signed)
        btn_create_pfx.grid(row=1, column=3, padx=8, pady=8, sticky="e")

        # 保留你的对齐修改：sticky="ew"
        btn_create_cer = ttk.Button(cfg_frame, text=self.t("create_cer_btn"), command=self._on_create_cer_only)
        btn_create_cer.grid(row=1, column=4, padx=8, pady=8, sticky="ew")

        cfg_frame.columnconfigure(1, weight=2)
        cfg_frame.columnconfigure(4, weight=1)  # 使第 4 列可伸展，配合 sticky="ew"

        # 操作区
        op_frame = ttk.Frame(self)
        op_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=6)

        self.btn_verify = ttk.Button(op_frame, text=self.t("verify_btn"), command=self._on_verify_files)
        self.btn_sign = ttk.Button(op_frame, text=self.t("sign_btn"), command=self._on_sign_files_seq)  # 顺序
        self.btn_sign_no_ts = ttk.Button(op_frame, text=self.t("sign_no_ts_btn"), command=self._on_sign_files_no_ts)  # 并发
        self.btn_timestamp = ttk.Button(op_frame, text=self.t("timestamp_btn"), command=self._on_timestamp_files_seq)  # 顺序

        self.btn_verify.pack(side=tk.LEFT, padx=(0, 8))
        self.btn_sign.pack(side=tk.LEFT, padx=8)
        self.btn_sign_no_ts.pack(side=tk.LEFT, padx=8)
        self.btn_timestamp.pack(side=tk.LEFT, padx=8)

        # 并发数选择（仅用于 验证 / 不加时间戳签名）
        ttk.Label(op_frame, text=self.t("concurrency")).pack(side=tk.LEFT, padx=(16, 4))
        workers_spin = ttk.Spinbox(op_frame, from_=1, to=max(1, (os.cpu_count() or 1)), textvariable=self.workers_var, width=4)
        workers_spin.pack(side=tk.LEFT, padx=(0, 8))

        # 进度条
        self.progress = ttk.Progressbar(op_frame, mode="determinate")
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(8, 0))

        # 日志
        log_frame = ttk.LabelFrame(self, text=self.t("log_title"))
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(6, 10))

        self.log_text = tk.Text(log_frame, height=16, wrap="word")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0), pady=8)
        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 8), pady=8)
        self.log_text.config(yscrollcommand=log_scroll.set)

        # 配置日志颜色标签
        for tag, color in self.TAG_COLORS.items():
            self.log_text.tag_config(tag, foreground=color)

    # ------------------ 文件管理 ------------------

    def _exts(self):
        return FileFormats.get_all_extensions()

    def _accept_file(self, path: str) -> bool:
        p = path.lower()
        return any(p.endswith(e) for e in self._exts())

    def _add_files(self, paths):
        added = 0
        for p in paths:
            p = os.path.abspath(p)
            if os.path.isfile(p) and self._accept_file(p):
                if p not in self.selected_files:
                    self.selected_files.append(p)
                    self.file_listbox.insert(tk.END, p)
                    added += 1
        if added:
            self._log(self.t("added_files", n=added), tag="info")

    def _on_add_files(self):
        exts = self._exts()
        pattern_all = " ".join(f"*{e}" for e in exts)
        filetypes = [
            (f"{self.t('supported_formats')} ({', '.join(exts)})", pattern_all),
            (self.t("all_files"), "*.*"),
        ]
        paths = filedialog.askopenfilenames(title=self.t("select_files_title"), filetypes=filetypes)
        if not paths:
            return
        self._add_files(paths)

    def _on_drop_files(self, event):
        try:
            parts = self.tk.splitlist(event.data)
        except Exception:
            parts = [event.data]
        self._add_files(parts)

    def _on_remove_selected(self):
        sel = list(self.file_listbox.curselection())
        if not sel:
            return
        sel.reverse()
        for idx in sel:
            path = self.file_listbox.get(idx)
            if path in self.selected_files:
                self.selected_files.remove(path)
            self.file_listbox.delete(idx)
        self._log(self.t("removed_selected"), tag="info")

    def _on_clear_files(self):
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self._log(self.t("list_cleared"), tag="info")

    def _on_browse_pfx(self):
        p = filedialog.askopenfilename(
            title=self.t("pfx_file"),
            filetypes=[("PFX", "*.pfx"), (self.t("all_files"), "*.*")]
        )
        if p:
            self.pfx_path_var.set(p)

    # ------------------ 子进程（signtool；Unicode 安全 + 隐藏窗口） ------------------

    def _find_signtool(self) -> str:
        tools_root = getattr(self.tool, "tools_path", "")
        if tools_root and os.path.isdir(tools_root):
            c1 = os.path.join(tools_root, "signtool.exe")
            if os.path.exists(c1):
                return os.path.abspath(c1)
            for base, dirs, files in os.walk(tools_root):
                if "signtool.exe" in files:
                    return os.path.abspath(os.path.join(base, "signtool.exe"))
        which = shutil.which("signtool.exe")
        if which:
            return which
        return "signtool.exe"

    def _run_signtool(self, args, check=True) -> str:
        """
        以 shell=False + 参数列表方式执行 signtool，避免代码页问题；
        同时隐藏子进程控制台窗口（CREATE_NO_WINDOW + STARTF_USESHOWWINDOW）。
        """
        exe = self._find_signtool()
        cmd = [exe] + list(args)
        enc = locale.getpreferredencoding(False) or "utf-8"

        startupinfo = None
        creationflags = 0
        if sys.platform == "win32":
            startupinfo = subprocess.STARTUPINFO()
            # 隐藏窗口
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0  # SW_HIDE
            # 不创建控制台窗口
            creationflags |= subprocess.CREATE_NO_WINDOW

        try:
            cp = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding=enc,
                errors="replace",
                shell=False,
                startupinfo=startupinfo,
                creationflags=creationflags,
            )
        except FileNotFoundError:
            raise RuntimeError("signtool.exe not found")
        out = (cp.stdout or "") + (cp.stderr or "")
        if check and cp.returncode != 0:
            raise RuntimeError(out.strip() or f"signtool failed (exit {cp.returncode})")
        return out

    # ------------------ 密码提示（GUI 线程同步） ------------------

    def _ask_password_sync(self, title: str) -> str | None:
        """
        在主线程弹出密码输入框，并同步等待用户输入（后台线程可调用）。
        """
        result = {"value": None}
        ev = threading.Event()
        self.msg_queue.put(("ask_pwd", (title, result, ev)))
        ev.wait()
        return result["value"]

    # ------------------ signtool 操作 ------------------

    def _verify_file(self, path: str) -> str:
        return self._run_signtool(["verify", "/pa", "/v", path], check=False)

    @staticmethod
    def _msg_indicates_password(err: str) -> bool:
        s = (err or "").lower()
        keywords = ["password", "pfx", "pass", "密碼", "密码"]
        return any(k in s for k in keywords)

    @staticmethod
    def _msg_wrong_password(err: str) -> bool:
        s = (err or "").lower()
        keywords = ["wrong password", "password is incorrect", "密碼不正確", "密码不正确"]
        return any(k in s for k in keywords)

    def _sign_one(self, path: str, pfx_path: str, password: str | None, add_timestamp: bool, ts_url: str | None):
        """
        对单个文件执行 signtool sign，必要时在 GUI 提示密码并重试一次。
        """
        base_args = ["sign", "/f", pfx_path, "/fd", "sha256", "/v"]

        def try_sign(pwd: str | None):
            args = list(base_args)
            if pwd is not None:
                args += ["/p", pwd]
            if add_timestamp and ts_url:
                # 先 RFC3161
                try:
                    self._run_signtool(args + ["/tr", ts_url, "/td", "sha256", path], check=True)
                    return
                except RuntimeError:
                    # 回退 /t
                    self._run_signtool(args + ["/t", ts_url, path], check=True)
                    return
            # 无时间戳
            self._run_signtool(args + [path], check=True)

        pwd_to_use = password if (password is not None and password != "") else None
        try:
            if pwd_to_use is None:
                try_sign("")
            else:
                try_sign(pwd_to_use)
            return
        except RuntimeError as e:
            msg = str(e)
            if self._msg_indicates_password(msg):
                cached = self._pfx_pwd_cache.get(pfx_path)
                if cached is not None and cached != pwd_to_use:
                    try:
                        try_sign(cached)
                        return
                    except RuntimeError as e2:
                        if not self._msg_wrong_password(str(e2)) and not self._msg_indicates_password(str(e2)):
                            raise
                ask_title = self.t("enter_pwd", name=os.path.basename(pfx_path))
                new_pwd = self._ask_password_sync(ask_title)
                if new_pwd is None:
                    raise RuntimeError("Signing cancelled by user (password prompt).")
                self._pfx_pwd_cache[pfx_path] = new_pwd
                try_sign(new_pwd)
                return
            raise

    def _timestamp_one(self, path: str, ts_url: str):
        try:
            self._run_signtool(["timestamp", "/tr", ts_url, "/td", "sha256", path], check=True)
        except RuntimeError:
            self._run_signtool(["timestamp", "/t", ts_url, path], check=True)

    # ------------------ 操作按钮 ------------------

    def _on_verify_files(self):
        files = self._get_files_or_warn()
        if not files:
            return
        self._clear_log(preserve_tip=True)
        self._run_bg(self._task_verify_parallel, files)

    def _on_sign_files_seq(self):
        files = self._get_files_or_warn()
        if not files:
            return
        pfx = self.pfx_path_var.get().strip('" ')
        if not (pfx and os.path.exists(pfx) and pfx.lower().endswith(".pfx")):
            messagebox.showwarning(title=self.t("app_title"), message=self.t("need_valid_pfx"))
            return
        self.tool.current_timestamp_url = self.ts_server_var.get()
        self._log(self.t("seq_info_ts"), tag="info")
        self._run_bg(self._task_sign_sequential_with_ts, files, pfx, self.pfx_pwd_var.get())

    def _on_sign_files_no_ts(self):
        files = self._get_files_or_warn()
        if not files:
            return
        pfx = self.pfx_path_var.get().strip('" ')
        if not (pfx and os.path.exists(pfx) and pfx.lower().endswith(".pfx")):
            messagebox.showwarning(title=self.t("app_title"), message=self.t("need_valid_pfx"))
            return
        workers = self._get_workers()
        if len(files) > 1:
            if not messagebox.askyesno(
                title=self.t("app_title"),
                message=self.t("concurrency_prompt", n=len(files), workers=workers)
            ):
                return
        self._run_bg(self._task_sign_parallel_no_ts, files, pfx, self.pfx_pwd_var.get(), workers)

    def _on_timestamp_files_seq(self):
        files = self._get_files_or_warn()
        if not files:
            return
        self.tool.current_timestamp_url = self.ts_server_var.get()
        self._log(self.t("seq_info_ts"), tag="info")
        self._run_bg(self._task_timestamp_sequential, files)

    def _on_create_self_signed(self):
        name = simpledialog.askstring(self.t("create_self_signed_btn"), "CN:", parent=self)
        if name is None or not name.strip():
            return
        email = simpledialog.askstring(self.t("create_self_signed_btn"), "E-mail (optional):", parent=self)
        pwd = simpledialog.askstring(self.t("password"), "PFX Password (optional):", parent=self, show="•")

        def _create():
            try:
                self._qlog(self.t("create_self_signed"), tag="info")
                cfg = SigningConfig(name=name.strip(), email=(email.strip() if email else None))
                self.tool._create_certificate(cfg)
                ok, final_pwd = self.tool._create_pfx(password=pwd if pwd else None)
                if ok:
                    self.tool._copy_to_desktop("Key.pfx")
                    self.tool._cleanup_temp_files()
                    pfx_full = os.path.join(self.tool.tools_path, "Key.pfx")
                    self.pfx_path_var.set(pfx_full if os.path.exists(pfx_full) else "")
                    self.pfx_pwd_var.set(final_pwd or "")
                    self._qlog(self.t("self_signed_done"), tag="warn")
                    self._qlog(self.t("self_signed_note"), tag="warn")
                else:
                    self._qlog(self.t("create_pfx_failed"), tag="error")
            except Exception as e:
                self._qlog(self.t("create_cert_failed", err=str(e)), tag="error")

        self._run_bg(_create)

    def _on_create_cer_only(self):
        name = simpledialog.askstring(self.t("create_cer_btn"), "CN:", parent=self)
        if name is None or not name.strip():
            return
        email = simpledialog.askstring(self.t("create_cer_btn"), "E-mail (optional):", parent=self)

        def _create_cer():
            try:
                cfg = SigningConfig(name=name.strip(), email=(email.strip() if email else None))
                self.tool._create_certificate(cfg)
                # 尝试定位生成的 .cer 文件
                cer_path = None
                preferred_names = [
                    "Key.cer",
                    f"{cfg.name}.cer",
                    "name.cer",
                    "certificate.cer",
                ]
                search_dirs = [os.getcwd(), self.tool.tools_path]
                for base in search_dirs:
                    for fname in preferred_names:
                        p = os.path.join(base, fname)
                        if os.path.exists(p):
                            cer_path = p
                            break
                    if cer_path:
                        break
                if not cer_path:
                    candidates = []
                    for base in search_dirs:
                        candidates.extend(glob.glob(os.path.join(base, "*.cer")))
                    if candidates:
                        cer_path = max(candidates, key=lambda x: os.path.getmtime(x))

                if not cer_path or not os.path.exists(cer_path):
                    self._qlog(self.t("cer_not_found"), tag="error")
                    return

                key_cer_path = os.path.join(self.tool.tools_path, "Key.cer")
                try:
                    shutil.copy2(cer_path, key_cer_path)
                except Exception:
                    key_cer_path = os.path.join(os.getcwd(), "Key.cer")
                    shutil.copy2(cer_path, key_cer_path)

                try:
                    prev_dir = os.getcwd()
                    os.chdir(os.path.dirname(key_cer_path))
                    self.tool._copy_to_desktop("Key.cer")
                    os.chdir(prev_dir)
                except Exception:
                    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                    shutil.copy2(key_cer_path, os.path.join(desktop, "Key.cer"))

                self.tool._cleanup_temp_files()
                self._qlog(self.t("cer_done", name="Key.cer"), tag="ok")
                self._qlog(self.t("self_signed_note"), tag="warn")
            except Exception as e:
                self._qlog(self.t("create_cert_failed", err=str(e)), tag="error")

        self._run_bg(_create_cer)

    # ------------------ 后台任务 ------------------

    def _status_label_and_tag(self, status: SignatureStatus):
        if status == SignatureStatus.TRUSTED:
            return self.t("trusted_friendly"), "ok"
        if status == SignatureStatus.SELF_SIGNED:
            return self.t("self_signed_friendly"), "warn"
        if status == SignatureStatus.UNSIGNED:
            return self.t("unsigned_friendly"), "error"
        if status == SignatureStatus.INVALID:
            return self.t("invalid_friendly"), "error"
        return self.t("unknown_friendly"), "info"

    def _task_verify_parallel(self, files):
        n = len(files)
        self._qlog(self.t("start_verify", n=n), tag="info")
        self._qset_progress(0, n)

        stats = {
            SignatureStatus.TRUSTED: 0,
            SignatureStatus.SELF_SIGNED: 0,
            SignatureStatus.UNSIGNED: 0,
            SignatureStatus.INVALID: 0,
            SignatureStatus.UNKNOWN: 0,
        }

        def verify_one(path: str):
            raw = self._verify_file(path)
            info = self.tool._parse_signature_info(raw)
            return os.path.basename(path), info

        workers = self._get_workers()
        completed = 0
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futures = [ex.submit(verify_one, f) for f in files]
            for fut in as_completed(futures):
                try:
                    name, info = fut.result()
                    completed += 1
                    self._qlog(self.t("verifying", i=completed, n=n, name=name), tag="info")
                    label, tag = self._status_label_and_tag(info.status)
                    self._qlog(self.t("result", status=label), tag=tag)
                    details = []
                    if info.signer_name:
                        details.append(f"{self.t('signer')}: {info.signer_name}")
                    if info.issuer:
                        details.append(f"{self.t('issuer')}: {info.issuer}")
                    if info.timestamp:
                        details.append(f"{self.t('timestamp')}: {info.timestamp}")
                    if details:
                        self._qlog("  " + " | ".join(details), tag="info")
                    stats[info.status] += 1
                except Exception as e:
                    completed += 1
                    self._qlog(self.t("verifying", i=completed, n=n, name="(error)"), tag="info")
                    self._qlog(f"  ✗ {str(e)}", tag="error")
                    stats[SignatureStatus.INVALID] += 1
                finally:
                    self._qstep()

        self._qlog(self.t("stats"), tag="info")
        for st in [SignatureStatus.TRUSTED, SignatureStatus.SELF_SIGNED, SignatureStatus.UNSIGNED,
                   SignatureStatus.INVALID, SignatureStatus.UNKNOWN]:
            c = stats[st]
            if c > 0:
                label, tag = self._status_label_and_tag(st)
                self._qlog(self.t("stats_item", label=label, n=c), tag=tag)

    def _task_sign_sequential_with_ts(self, files, pfx_path, pwd):
        n = len(files)
        self._qset_progress(0, n)
        ts_url = self.tool.current_timestamp_url
        for i, f in enumerate(files, 1):
            name = os.path.basename(f)
            self._qlog(self.t("signing", i=i, n=n, name=name), tag="info")
            try:
                self._sign_one(f, pfx_path, pwd, add_timestamp=True, ts_url=ts_url)
                self._qlog(self.t("done"), tag="ok")
            except Exception as e:
                self._qlog(f"  ✗ {str(e)}", tag="error")
            self._qstep()
        self._qlog(self.t("sign_all_done"), tag="ok")

    def _task_sign_parallel_no_ts(self, files, pfx_path, pwd, workers):
        n = len(files)
        self._qset_progress(0, n)

        def sign_one(path: str):
            try:
                self._sign_one(path, pfx_path, pwd, add_timestamp=False, ts_url=None)
                return os.path.basename(path), True, ""
            except Exception as e:
                return os.path.basename(path), False, str(e)

        completed = 0
        with ThreadPoolExecutor(max_workers=max(1, workers)) as ex:
            futures = [ex.submit(sign_one, f) for f in files]
            for fut in as_completed(futures):
                name, ok, err = fut.result()
                completed += 1
                self._qlog(self.t("signing", i=completed, n=n, name=name), tag="info")
                if ok:
                    self._qlog(self.t("done"), tag="ok")
                else:
                    self._qlog(f"  ✗ {err}", tag="error")
                self._qstep()
        self._qlog(self.t("sign_all_done"), tag="ok")

    def _task_timestamp_sequential(self, files):
        n = len(files)
        self._qlog(self.t("start_timestamp", n=n), tag="info")
        self._qset_progress(0, n)
        ts_url = self.tool.current_timestamp_url
        for i, f in enumerate(files, 1):
            name = os.path.basename(f)
            self._qlog(self.t("timestamp_item", i=i, n=n, name=name), tag="info")
            try:
                self._timestamp_one(f, ts_url)
                self._qlog(self.t("done"), tag="ok")
            except Exception as e:
                self._qlog(f"  ✗ {str(e)}", tag="error")
            self._qstep()
        self._qlog(self.t("timestamp_done"), tag="ok")

    # ------------------ 工具 ------------------

    def _set_icon(self):
        try:
            candidates = [
                resource_path("icon.ico"),
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico"),
                "icon.ico",
            ]
            ico_path = next((os.path.abspath(p) for p in candidates if os.path.exists(p)), None)
            if not ico_path or sys.platform != "win32":
                return
            tmp_ico = os.path.join(tempfile.gettempdir(), "app_icon.ico")
            try:
                if (not os.path.exists(tmp_ico)) or (os.path.getmtime(tmp_ico) < os.path.getmtime(ico_path)):
                    shutil.copy2(ico_path, tmp_ico)
            except Exception:
                tmp_ico = ico_path
            self.iconbitmap(tmp_ico)
            try:
                self.iconbitmap(default=tmp_ico)
            except Exception:
                pass
        except Exception:
            pass

    def _get_workers(self, cap: int | None = None) -> int:
        try:
            n = int(self.workers_var.get())
        except Exception:
            n = 1
        n = max(1, n)
        if cap is not None:
            n = min(n, cap)
        return n

    def _get_files_or_warn(self):
        if not self.selected_files:
            messagebox.showinfo(title=self.t("app_title"), message=self.t("no_files"))
            return None
        return list(self.selected_files)

    def _run_bg(self, target, *args, **kwargs):
        for b in (self.btn_verify, self.btn_sign, self.btn_sign_no_ts, self.btn_timestamp):
            b.config(state=tk.DISABLED)
        self.progress.config(value=0)
        t = threading.Thread(target=self._bg_wrapper, args=(target, args, kwargs), daemon=True)
        t.start()

    def _bg_wrapper(self, target, args, kwargs):
        try:
            target(*args, **kwargs)
        finally:
            self.msg_queue.put(("enable_buttons", None))

    def _qlog(self, msg: str, tag: str = None):
        self.msg_queue.put(("log", (msg, tag)))

    def _qset_progress(self, value, maximum):
        self.msg_queue.put(("progress_set", (value, maximum)))

    def _qstep(self):
        self.msg_queue.put(("progress_step", 1))

    def _process_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "log":
                    msg, tag = payload
                    self._log(msg, tag=tag)
                elif kind == "progress_set":
                    value, maximum = payload
                    self.progress.config(maximum=maximum, value=value)
                elif kind == "progress_step":
                    self.progress.step(payload)
                elif kind == "enable_buttons":
                    for b in (self.btn_verify, self.btn_sign, self.btn_sign_no_ts, self.btn_timestamp):
                        b.config(state=tk.NORMAL)
                elif kind == "ask_pwd":
                    title, result, ev = payload
                    try:
                        pwd = simpledialog.askstring(self.t("password"), title, parent=self, show="•")
                        result["value"] = pwd
                    finally:
                        ev.set()
        except queue.Empty:
            pass
        self.after(100, self._process_queue)

    def _clear_log(self, preserve_tip=False):
        self.log_text.delete("1.0", tk.END)
        if preserve_tip:
            self._log(self._tip_text, tag=self._tip_tag)

    def _log(self, msg: str, tag: str = None):
        if tag:
            self.log_text.insert(tk.END, msg + "\n", (tag,))
        else:
            self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)


if __name__ == "__main__":
    enable_high_dpi()
    App().mainloop()