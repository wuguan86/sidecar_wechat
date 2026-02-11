from __future__ import annotations

import base64
import ctypes
import dataclasses
import datetime as _dt
import hashlib
import io
import json
import logging
import logging.handlers
import os
import queue
import random
import re
import subprocess
import threading
import time
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests

try:
    import pyperclip
except ImportError:
    pyperclip = None

try:
    import uiautomation as auto
except Exception:  # noqa: BLE001
    auto = None

try:
    import pythoncom  # type: ignore
except Exception:  # noqa: BLE001
    pythoncom = None

try:
    from PIL import ImageGrab
except Exception:  # noqa: BLE001
    ImageGrab = None

import ctypes
try:
    # 强制开启 DPI 感知，防止截图错位
    ctypes.windll.shcore.SetProcessDpiAwareness(1) 
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass    


@dataclasses.dataclass(frozen=True)
class BridgeConfig:
    java_receive_url: str = "http://localhost:8080/api/wechat/receive"
    java_timeout_seconds: float = 5.0
    java_retry_max: int = 3
    java_retry_backoff_base_seconds: float = 0.6

    window_class_name: str = "mmui::MainWindow"
    window_name: str = "微信"

    scan_interval_seconds: float = 0.6
    scan_jitter_seconds: float = 0.3
    unread_max_per_round: int = 5
    message_scan_limit: int = 10

    send_delay_min_seconds: float = 0.5
    send_delay_max_seconds: float = 2.0
    click_move_min_seconds: float = 0.18
    click_move_max_seconds: float = 0.55

    # 新增：未读消息扫描间隔（模拟人类偶尔查看列表的行为，而不是每秒都看）
    # 调整为“精力集中”模式：检查频率加快 (1.5~4s)
    unread_scan_interval_min_seconds: float = 1.5
    unread_scan_interval_max_seconds: float = 4.0

    server_host: str = "127.0.0.1"
    server_port: int = 51234

    log_file: str = "wechat_bridge.log"
    log_level: str = "INFO"
    log_max_bytes: int = 10 * 1024 * 1024
    log_backup_count: int = 3


def _strip_yaml_comment(line: str) -> str:
    in_single = False
    in_double = False
    for i, ch in enumerate(line):
        if ch == "'" and not in_double:
            in_single = not in_single
        elif ch == '"' and not in_single:
            in_double = not in_double
        elif ch == "#" and not in_single and not in_double:
            return line[:i]
    return line


def _parse_yaml_scalar(value: str) -> Any:
    value = value.strip()
    if value == "":
        return ""
    if value.lower() in {"true", "false"}:
        return value.lower() == "true"
    if value.lower() in {"null", "none", "~"}:
        return None
    if (value.startswith('"') and value.endswith('"')) or (value.startswith("'") and value.endswith("'")):
        return value[1:-1]
    if re.fullmatch(r"-?\d+", value):
        try:
            return int(value)
        except Exception:  # noqa: BLE001
            return value
    if re.fullmatch(r"-?\d+\.\d+", value):
        try:
            return float(value)
        except Exception:  # noqa: BLE001
            return value
    return value


def _parse_simple_yaml(text: str) -> Dict[str, Any]:
    root: Dict[str, Any] = {}
    stack: List[Tuple[int, Dict[str, Any]]] = [(0, root)]

    for raw in text.splitlines():
        line = _strip_yaml_comment(raw).rstrip("\r\n")
        if not line.strip():
            continue

        indent = len(line) - len(line.lstrip(" "))
        line = line.strip()
        if ":" not in line:
            continue

        key, rest = line.split(":", 1)
        key = key.strip()
        rest = rest.strip()

        while stack and indent < stack[-1][0]:
            stack.pop()
        if not stack:
            stack = [(0, root)]

        current = stack[-1][1]
        if rest == "":
            new_obj: Dict[str, Any] = {}
            current[key] = new_obj
            stack.append((indent + 2, new_obj))
        else:
            current[key] = _parse_yaml_scalar(rest)

    return root


def _deep_get(data: Dict[str, Any], path: List[str], default: Any) -> Any:
    cur: Any = data
    for key in path:
        if not isinstance(cur, dict):
            return default
        if key not in cur:
            return default
        cur = cur[key]
    return cur


def load_config(config_path: str) -> BridgeConfig:
    config_dir = os.path.dirname(os.path.abspath(config_path))
    with open(config_path, "r", encoding="utf-8") as f:
        raw = f.read()

    parsed: Dict[str, Any]
    yaml_module = None
    try:
        import yaml as yaml_module  # type: ignore
    except Exception:  # noqa: BLE001
        yaml_module = None

    if yaml_module is not None:
        parsed = yaml_module.safe_load(raw) or {}
    else:
        parsed = _parse_simple_yaml(raw)

    cfg = BridgeConfig(
        java_receive_url=str(_deep_get(parsed, ["java", "receive_url"], BridgeConfig.java_receive_url)),
        java_timeout_seconds=float(_deep_get(parsed, ["java", "timeout_seconds"], BridgeConfig.java_timeout_seconds)),
        java_retry_max=int(_deep_get(parsed, ["java", "retry_max"], BridgeConfig.java_retry_max)),
        java_retry_backoff_base_seconds=float(
            _deep_get(parsed, ["java", "retry_backoff_base_seconds"], BridgeConfig.java_retry_backoff_base_seconds)
        ),
        window_class_name=str(_deep_get(parsed, ["window", "class_name"], BridgeConfig.window_class_name)),
        window_name=str(_deep_get(parsed, ["window", "name"], BridgeConfig.window_name)),
        scan_interval_seconds=float(_deep_get(parsed, ["listener", "scan_interval_seconds"], BridgeConfig.scan_interval_seconds)),
        scan_jitter_seconds=float(_deep_get(parsed, ["listener", "scan_jitter_seconds"], BridgeConfig.scan_jitter_seconds)),
        unread_max_per_round=int(_deep_get(parsed, ["listener", "unread_max_per_round"], BridgeConfig.unread_max_per_round)),
        message_scan_limit=int(_deep_get(parsed, ["listener", "message_scan_limit"], BridgeConfig.message_scan_limit)),
        send_delay_min_seconds=float(_deep_get(parsed, ["executor", "send_delay_min_seconds"], BridgeConfig.send_delay_min_seconds)),
        send_delay_max_seconds=float(_deep_get(parsed, ["executor", "send_delay_max_seconds"], BridgeConfig.send_delay_max_seconds)),
        click_move_min_seconds=float(_deep_get(parsed, ["executor", "click_move_min_seconds"], BridgeConfig.click_move_min_seconds)),
        click_move_max_seconds=float(_deep_get(parsed, ["executor", "click_move_max_seconds"], BridgeConfig.click_move_max_seconds)),
        unread_scan_interval_min_seconds=float(
            _deep_get(parsed, ["listener", "unread_scan_interval_min_seconds"], BridgeConfig.unread_scan_interval_min_seconds)
        ),
        unread_scan_interval_max_seconds=float(
            _deep_get(parsed, ["listener", "unread_scan_interval_max_seconds"], BridgeConfig.unread_scan_interval_max_seconds)
        ),
        server_host=str(_deep_get(parsed, ["server", "host"], BridgeConfig.server_host)),
        server_port=int(_deep_get(parsed, ["server", "port"], BridgeConfig.server_port)),
        log_file=str(_deep_get(parsed, ["logging", "file"], BridgeConfig.log_file)),
        log_level=str(_deep_get(parsed, ["logging", "level"], BridgeConfig.log_level)),
        log_max_bytes=int(_deep_get(parsed, ["logging", "max_bytes"], BridgeConfig.log_max_bytes)),
        log_backup_count=int(_deep_get(parsed, ["logging", "backup_count"], BridgeConfig.log_backup_count)),
    )

    env_receive = os.getenv("JAVA_RECEIVE_URL")
    env_host = os.getenv("SIDECAR_SERVER_HOST")
    env_port = os.getenv("SIDECAR_SERVER_PORT")
    if env_receive:
        cfg = dataclasses.replace(cfg, java_receive_url=env_receive)
    if env_host:
        cfg = dataclasses.replace(cfg, server_host=env_host)
    if env_port:
        try:
            cfg = dataclasses.replace(cfg, server_port=int(env_port))
        except Exception:  # noqa: BLE001
            pass

    if not os.path.isabs(cfg.log_file):
        cfg = dataclasses.replace(cfg, log_file=os.path.join(config_dir, cfg.log_file))
    return cfg


def setup_logging(cfg: BridgeConfig) -> logging.Logger:
    logger = logging.getLogger("wechat_bridge")
    # 设置日志级别，默认为 INFO
    logger.setLevel(getattr(logging, cfg.log_level.upper(), logging.INFO))

    if logger.handlers:
        return logger

    # 格式：时间 [级别] [线程] 消息
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s")

    os.makedirs(os.path.dirname(cfg.log_file), exist_ok=True)
    
    # 使用 RotatingFileHandler 防止文件过大
    file_handler = logging.handlers.RotatingFileHandler(
        cfg.log_file, maxBytes=cfg.log_max_bytes, backupCount=cfg.log_backup_count, encoding="utf-8"
    )
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # 添加控制台输出
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


def _now_iso() -> str:
    return _dt.datetime.now().isoformat(timespec="seconds")


def _sleep_with_jitter(base_seconds: float, jitter_seconds: float) -> None:
    if base_seconds < 0:
        base_seconds = 0
    if jitter_seconds < 0:
        jitter_seconds = 0
    time.sleep(base_seconds + random.random() * jitter_seconds)


def _rect_to_bbox(rect: Any) -> Optional[Tuple[int, int, int, int]]:
    if rect is None:
        return None
    try:
        left = int(rect.left)
        top = int(rect.top)
        right = int(rect.right)
        bottom = int(rect.bottom)
    except Exception:  # noqa: BLE001
        try:
            if isinstance(rect, (tuple, list)) and len(rect) == 4:
                left, top, right, bottom = [int(v) for v in rect]
            else:
                return None
        except Exception:  # noqa: BLE001
            return None
    if right <= left or bottom <= top:
        return None
    return (left, top, right, bottom)


def _image_to_base64_png(img: Any) -> str:
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("ascii")


def _sha1_text(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8", errors="ignore")).hexdigest()


class JavaClient:
    def __init__(self, cfg: BridgeConfig, logger: logging.Logger) -> None:
        self._cfg = cfg
        self._logger = logger
        self._session = requests.Session()

    def post_message(self, payload: Dict[str, Any]) -> None:
        if not self._cfg.java_receive_url:
            return

        for attempt in range(self._cfg.java_retry_max + 1):
            try:
                res = self._session.post(
                    self._cfg.java_receive_url,
                    json=payload,
                    timeout=(self._cfg.java_timeout_seconds, self._cfg.java_timeout_seconds),
                )
                if 200 <= res.status_code < 300:
                    return
                self._logger.warning("Java 接口返回非 2xx: status=%s body=%s", res.status_code, res.text[:500])
            except Exception as e:  # noqa: BLE001
                self._logger.warning("上报 Java 失败: %s", e)

            if attempt < self._cfg.java_retry_max:
                backoff = self._cfg.java_retry_backoff_base_seconds * (2**attempt) * (0.7 + random.random() * 0.6)
                time.sleep(backoff)


def _ensure_deps() -> None:
    missing = []
    if auto is None:
        missing.append("uiautomation")
    if ImageGrab is None:
        missing.append("Pillow(ImageGrab)")
    if missing:
        raise RuntimeError("缺少依赖: " + ", ".join(missing))


def _ensure_com_initialized(logger: logging.Logger) -> None:
    if pythoncom is None:
        logger.warning("pythoncom 未安装，无法显式初始化 COM")
        return
    try:
        pythoncom.CoInitialize()
    except Exception as e:  # noqa: BLE001
        logger.warning("COM 初始化失败: %s", e)


def _warmup_uia(logger: logging.Logger) -> None:
    if auto is None:
        return
    try:
        auto.GetRootControl()
    except Exception as e:  # noqa: BLE001
        logger.warning("UIA 核心组件预加载失败: %s", e)


def _control_type_name(ctrl: Any) -> str:
    name = getattr(ctrl, "ControlTypeName", None)
    if isinstance(name, str) and name:
        return name
    try:
        ct = getattr(ctrl, "ControlType", None)
        if ct is None:
            return ""
        if hasattr(auto, "ControlType") and hasattr(auto.ControlType, "GetControlTypeName"):
            return str(auto.ControlType.GetControlTypeName(ct))
        return str(ct)
    except Exception:  # noqa: BLE001
        return ""


def _safe_attr(ctrl: Any, attr: str) -> str:
    try:
        val = getattr(ctrl, attr, "")
        if val is None:
            return ""
        return str(val)
    except Exception:  # noqa: BLE001
        return ""


def _rect_text(ctrl: Any) -> str:
    try:
        rect = getattr(ctrl, "BoundingRectangle", None)
        if rect is None:
            return ""
        return str(rect)
    except Exception:  # noqa: BLE001
        return ""


def _inspect_window_tree(window: Any, logger: logging.Logger) -> None:
    keywords = ["未读", "消息", "输入", "发送"]
    stack: List[Tuple[Any, int]] = [(window, 0)]
    while stack:
        node, depth = stack.pop()
        control_type = _control_type_name(node)
        name = _safe_attr(node, "Name")
        automation_id = _safe_attr(node, "AutomationId")
        rect_text = _rect_text(node)
        keyword_hit = any(k in name or k in automation_id for k in keywords)
        prefix = "HIGHLIGHT" if keyword_hit else "NODE"
        indent = "  " * depth
        logger.info(
            "%s %s%s | Name=%s | AutomationId=%s | Rect=%s",
            prefix,
            indent,
            control_type,
            name,
            automation_id,
            rect_text,
        )
        try:
            children = node.GetChildren() or []
        except Exception:  # noqa: BLE001
            children = []
        for child in reversed(children):
            stack.append((child, depth + 1))


def _iter_descendants(root: Any, max_depth: int) -> Iterable[Any]:
    if max_depth <= 0:
        return
    q: List[Tuple[Any, int]] = [(root, 0)]
    while q:
        node, depth = q.pop(0)
        if depth >= max_depth:
            continue
        children = []
        try:
            children = node.GetChildren() or []
        except Exception:  # noqa: BLE001
            children = []
        for child in children:
            yield child
            q.append((child, depth + 1))


class WeChatUI:
    def __init__(self, cfg: BridgeConfig, logger: logging.Logger) -> None:
        self._cfg = cfg
        self._logger = logger
        self._uia_lock = threading.RLock()
        self._cached_main = None
        self._tree_logged_handle = None

    def _normalize_contact_name(self, name: str) -> str:
        if not name:
            return ""
        name = name.strip()
        if not name:
            return ""
        parts = [part.strip() for part in name.splitlines() if part.strip()]
        if parts:
            name = parts[0]
        name = re.sub(r"\s+", " ", name).strip()
        name = re.sub(r"(?:\d+\s*条新消息|未读)$", "", name).strip()
        return name

    def _get_wechat_pids(self) -> List[int]:
        pids = []
        for name in ["WeChat.exe", "Weixin.exe"]:
            try:
                # 使用 tasklist 查找，/NH 不显示表头，/FO CSV CSV格式
                cmd = f'tasklist /FI "IMAGENAME eq {name}" /FO CSV /NH'
                # Windows 中文环境通常是 GBK
                output = subprocess.check_output(cmd, shell=True).decode("gbk", errors="ignore")
                for line in output.splitlines():
                    if not line.strip():
                        continue
                    parts = line.split(',')
                    if len(parts) >= 2:
                        # "Image Name","PID",...
                        pid_str = parts[1].strip('"')
                        if pid_str.isdigit():
                            pids.append(int(pid_str))
            except Exception as e:
                self._logger.warning(f"获取进程PID失败 {name}: {e}")
        return list(set(pids))

    def _find_main_window(self) -> Any:
        if auto is None:
            raise RuntimeError("uiautomation 未加载")
        
        self._logger.info("Debug: _find_main_window start (Optimized)")
        
        if hasattr(auto, "SetGlobalSearchTimeout"):
            try:
                auto.SetGlobalSearchTimeout(0.5) # Reduced to 0.5s
            except Exception as e:
                self._logger.warning("SetGlobalSearchTimeout failed: %s", e)

        if hasattr(auto, "SetTransactionTimeout"):
            try:
                auto.SetTransactionTimeout(500) # 500ms
            except Exception as e:
                self._logger.warning("SetTransactionTimeout failed: %s", e)
        
        # 策略0：使用 Win32 API 快速查找 (避免 UIA 遍历挂死)
        try:
            hwnd = ctypes.windll.user32.FindWindowW(self._cfg.window_class_name, self._cfg.window_name)
            if hwnd and hwnd != 0:
                self._logger.info(f"FindWindowW found handle: {hwnd}")
                # 检查窗口是否挂起
                if ctypes.windll.user32.IsHungAppWindow(hwnd):
                    self._logger.warning(f"Handle {hwnd} is HUNG (Not Responding). Skipping.")
                else:
                    window = auto.ControlFromHandle(hwnd)
                    if window.Exists(0, 0):
                        self._logger.info("Window from handle validated via UIA")
                        try:
                            self._logger.info("Window PID: %s", window.ProcessId)
                        except:
                            pass
                        return window
                    else:
                        self._logger.warning("Window from handle exists check failed")
        except Exception as e:
            self._logger.warning(f"Win32 FindWindow optimization failed: {e}")

        # 策略1：全局直接查找（最常用，通常最快且不易卡死在特定僵尸进程）
        try:
            window = auto.WindowControl(
                searchDepth=1,
                ClassName=self._cfg.window_class_name,
                Name=self._cfg.window_name,
            )
            self._logger.info("Checking Global Window Exists...")
            if window.Exists(0, 0):
                self._logger.info("Found window via Global Search")
                # 顺便记录一下PID，确认是哪个进程
                try:
                    self._logger.info("Window PID: %s", window.ProcessId)
                except:
                    pass
                return window
        except Exception as e:
            self._logger.warning("Global search failed: %s", e)
        
        # 策略2：通过 PID 查找（备选）
        pids = self._get_wechat_pids()
        self._logger.info("Found WeChat PIDs: %s", pids)
        
        for pid in pids:
            self._logger.info(f"Checking PID: {pid}")
            # 尝试通过 PID + Name/ClassName 查找
            try:
                self._logger.info(f"Creating WindowControl for PID {pid}")
                window = auto.WindowControl(
                    searchDepth=1,
                    ProcessId=pid,
                    ClassName=self._cfg.window_class_name,
                    Name=self._cfg.window_name,
                )
                self._logger.info(f"Checking Exists for PID {pid}")
                exists = window.Exists(0, 0)
                self._logger.info(f"PID {pid} Exists result: {exists}")
                if exists:
                    self._logger.info("Found window in PID %s", pid)
                    return window
                
                # 备选：只用 Name 查找（防止 ClassName 变动）
                self._logger.info(f"Creating WindowControl (loose) for PID {pid}")
                window_loose = auto.WindowControl(
                    searchDepth=1,
                    ProcessId=pid,
                    Name=self._cfg.window_name,
                )
                self._logger.info(f"Checking Exists (loose) for PID {pid}")
                exists_loose = window_loose.Exists(0, 0)
                self._logger.info(f"PID {pid} Exists (loose) result: {exists_loose}")
                if exists_loose:
                     self._logger.info("Found window (loose match) in PID %s", pid)
                     return window_loose
            except Exception as e:
                self._logger.warning("Check window for PID %s failed: %s", pid, e)
                
        self._logger.warning("未找到微信主窗口 (scanned PIDs: %s)", pids)
        return None


    def get_main_window(self) -> Any:
        with self._uia_lock:
            if self._cached_main is not None:
                try:
                    if self._cached_main.Exists(0, 0):
                        return self._cached_main
                except Exception:  # noqa: BLE001
                    self._cached_main = None
            win = self._find_main_window()
            self._cached_main = win
            if win is not None:
                self._logger.info("Main window found, calling _log_window_tree")
                self._log_window_tree(win)
            return win

    def _log_window_tree(self, window: Any) -> None:
        handle = getattr(window, "NativeWindowHandle", None)
        if handle is None:
            handle = id(window)
        
        self._logger.info("Preparing to log tree for handle: %s", handle)
        
        if self._tree_logged_handle == handle:
            self._logger.info("Tree already logged for handle %s", handle)
            return
        self._tree_logged_handle = handle
        try:
            children = window.GetChildren() or []
            self._logger.info("Window has %d children", len(children))
        except Exception as e:  # noqa: BLE001
            self._logger.warning("GetChildren failed: %s", e)
            children = []
        for child in children:
            try:
                name = getattr(child, "Name", "") or ""
            except Exception:  # noqa: BLE001
                name = ""
            self._logger.info("WindowChild L1: %s | %s", _control_type_name(child), name)
            try:
                grand_children = child.GetChildren() or []
            except Exception:  # noqa: BLE001
                grand_children = []
            for grand in grand_children:
                try:
                    gname = getattr(grand, "Name", "") or ""
                except Exception:  # noqa: BLE001
                    gname = ""
                self._logger.info("WindowChild L2: %s | %s", _control_type_name(grand), gname)

    def guard_popups(self) -> None:
        if auto is None:
            return
        with self._uia_lock:
            root = auto.GetRootControl()
            for child in root.GetChildren() or []:
                if _control_type_name(child) != "WindowControl":
                    continue
                try:
                    class_name = getattr(child, "ClassName", "")
                    if class_name == self._cfg.window_class_name:
                        continue
                    title = getattr(child, "Name", "") or ""
                except Exception:  # noqa: BLE001
                    continue

                if not title:
                    continue

                # 弹窗防御说明（中文）：
                # - 微信运行过程中常出现更新/通话等干扰弹窗，可能遮挡 UI 或抢占焦点导致流程卡死。
                # - 这里用标题关键词匹配（Regex）先粗筛，再在弹窗内部寻找“取消/稍后/关闭”等按钮点击关闭。
                if not re.search(r"(语音|通话|更新|版本|升级|安装|提示|确认)", title):
                    continue

                self._logger.info("检测到干扰弹窗: %s", title)

                button_names = ["取消", "稍后", "关闭", "知道了", "忽略", "否", "不升级"]
                for btn_name in button_names:
                    try:
                        # 选择器说明（中文）：
                        # - ButtonControl + Name=按钮文案：在不同弹窗里 AutomationId 不稳定，但按钮文案在中文系统下更稳定。
                        btn = auto.ButtonControl(searchFromControl=child, searchDepth=6, Name=btn_name)
                        if btn and btn.Exists(0, 0):
                            btn.Click(simulateMove=True, waitTime=random.uniform(0.1, 0.3))
                            self._logger.info("已点击弹窗按钮: %s", btn_name)
                            break
                    except Exception:  # noqa: BLE001
                        continue

    def _locate_session_list(self, window: Any) -> Optional[Any]:
        if auto is None:
            return None
            
        # 策略0：精确查找 Name='会话' 的 ListControl (新版微信特征)
        try:
            # 搜索深度适当大一点，因为可能在 XSplitterView -> XView -> ... 下面
            target = auto.ListControl(searchFromControl=window, searchDepth=12, Name="会话")
            if target and target.Exists(0, 0):
                self._logger.debug("_locate_session_list 通过 Name='会话' 找到控件")
                return target
        except Exception:
            pass
            
        candidates = []
        # 增加搜索深度到 25，因为控件树显示 ListControl 在第 10+ 层
        for ctrl in _iter_descendants(window, max_depth=25):
            ct = _control_type_name(ctrl)
            if ct not in {"ListControl", "PaneControl"}:
                continue
            rect = getattr(ctrl, "BoundingRectangle", None)
            bbox = _rect_to_bbox(rect) if rect is not None else None
            if bbox is None:
                continue
            left, top, right, bottom = bbox
            width = right - left
            height = bottom - top
            if width < 160 or height < 200:
                continue
            # 选择器说明（中文）：
            # - 左侧会话列表位于窗口左侧区域；用 BoundingRectangle.left 做空间过滤，避免把右侧消息列表误判为会话列表。
            if left > 280: # 稍微放宽一点限制，原为 200，实际为 80，安全
                continue
            
            # 增加检查：是否有 ListItem 子节点
            try:
                items = auto.ListItemControl(searchFromControl=ctrl, searchDepth=2)
                if not items.Exists(0, 0):
                     # 如果没有 ListItem，可能不是我们要的列表
                     continue
            except Exception:
                pass
                
            candidates.append((width * height, ctrl))
            
        if not candidates:
            self._logger.debug("_locate_session_list 未找到任何符合条件的列表控件")
            return None
            
        candidates.sort(key=lambda x: x[0], reverse=True)
        best = candidates[0][1]
        # self._logger.info(f"DEBUG: _locate_session_list 找到列表: {getattr(best, 'Name', '')} Rect={getattr(best, 'BoundingRectangle', '')}")
        return best

    def _locate_message_list(self, window: Any) -> Optional[Any]:
        if auto is None:
            return None
            
        # 策略0：精确查找 Name='消息' 的 ListControl (新版微信特征)
        try:
            target = auto.ListControl(searchFromControl=window, searchDepth=15, Name="消息")
            if target and target.Exists(0, 0):
                self._logger.debug("_locate_message_list 通过 Name='消息' 找到控件")
                return target
        except Exception:
            pass

        candidates = []
        
        # 尝试从内容根节点和窗口根节点搜索 (双重保障)
        roots = []
        content_root = self._get_content_root(window)
        roots.append(content_root)
        if content_root != window:
            roots.append(window)
            
        checked_ids = set()

        for root in roots:
            # self._logger.info(f"DEBUG: Searching message list in root: {_control_type_name(root)}")
            for ctrl in _iter_descendants(root, max_depth=25):
                # 避免重复检查同一个控件
                try:
                    h = id(ctrl)
                    if h in checked_ids:
                        continue
                    checked_ids.add(h)
                except:
                    pass

                ct = _control_type_name(ctrl)
                # 增加 PaneControl 支持，因为某些版本消息列表可能是 Pane
                if ct not in ("ListControl", "PaneControl"):
                    continue
                
                rect = getattr(ctrl, "BoundingRectangle", None)
                bbox = _rect_to_bbox(rect) if rect is not None else None
                if bbox is None:
                    continue
                left, top, right, bottom = bbox
                width = right - left
                height = bottom - top
                
                # self._logger.info(f"DEBUG: Found {ct} candidate: {width}x{height} at ({left}, {top}) Name={getattr(ctrl, 'Name', '')}")

                # 放宽尺寸限制 (原: 300x260 -> 200x200)
                if width < 200 or height < 200:
                    continue
                
                # 放宽左侧位置限制 (原: 220 -> 180)
                if left < 180:
                    continue
                    
                candidates.append((width * height, ctrl))
            
            if candidates:
                break

        if not candidates:
            self._logger.debug("_locate_message_list 失败: 未找到候选控件 (ListControl/PaneControl > 200x200 @ Right).")
            # 尝试打印一下所有找到的 ListControl/PaneControl 以便调试
            try:
                debug_list = []
                for root in roots:
                    for ctrl in _iter_descendants(root, max_depth=10):
                        ct = _control_type_name(ctrl)
                        if ct in ("ListControl", "PaneControl"):
                            r = getattr(ctrl, "BoundingRectangle", None)
                            b = _rect_to_bbox(r) if r else (0,0,0,0)
                            w, h = b[2]-b[0], b[3]-b[1]
                            debug_list.append(f"{ct}:{w}x{h}@{b[0]},{b[1]}")
                self._logger.info(f"DEBUG: All List/Pane candidates in top 10 layers: {debug_list[:10]}...")
            except:
                pass
            return None
            
        candidates.sort(key=lambda x: x[0], reverse=True)
        best = candidates[0][1]
        # self._logger.info(f"DEBUG: _locate_message_list selected: {_control_type_name(best)} {getattr(best, 'BoundingRectangle', '')}")
        return best

    def find_unread_sessions(self) -> List[Any]:
        if auto is None:
            return []
        with self._uia_lock:
            window = self.get_main_window()
            if window is None:
                return []

            session_list = self._locate_session_list(window)
            if session_list is None:
                return []

            unread_items: List[Any] = []
            # 获取所有子项，限制扫描前 20 个，防止卡顿
            children = session_list.GetChildren() or []
            for item in children[:20]:
                if len(unread_items) >= self._cfg.unread_max_per_round:
                    break
                try:
                    if self._is_session_item_unread(item):
                        unread_items.append(item)
                except Exception:
                    continue
            return unread_items

    def _is_session_item_unread(self, item: Any) -> bool:
        # 未读判断说明（中文）：
        # - 不使用 OCR，只依赖 UIA 文本属性。
        # - 微信未读红点/数字徽标通常以 TextControl 形式存在（Name 为纯数字），或在条目 Name 中包含“x条新消息/未读”。
        item_name = getattr(item, "Name", "") or ""
        if re.search(r"(\d+条新消息|未读)", item_name):
            # self._logger.info(f"DEBUG: Item Name hit unread: {item_name}")
            return True

        # 降低深度到 5，防止遍历耗时过长导致卡顿
        for ctrl in _iter_descendants(item, max_depth=5):
            ct = _control_type_name(ctrl)
            # 有些版本红点可能是 GroupControl 或其他，只要 Name 是数字就可能是
            text = getattr(ctrl, "Name", "") or ""
            if not text:
                continue
                
            if re.fullmatch(r"\d+", text.strip()):
                # self._logger.info(f"DEBUG: Found numeric indicator: {text} in {ct}")
                return True
            if "条新消息" in text or "未读" in text:
                # self._logger.info(f"DEBUG: Found text indicator: {text} in {ct}")
                return True
        return False

    def click_session_item(self, item: Any) -> Optional[str]:
        if auto is None:
            return None
        with self._uia_lock:
            name = self._normalize_contact_name((getattr(item, "Name", "") or ""))
            try:
                # 拟人化说明（中文）：
                # - simulateMove=True 让库内部模拟鼠标移动到控件再点击，避免“瞬移点击”的机械行为特征。
                # - waitTime 使用随机范围，降低行为一致性，减少风控风险。
                item.Click(simulateMove=True, waitTime=random.uniform(self._cfg.click_move_min_seconds, self._cfg.click_move_max_seconds))
                time.sleep(random.uniform(0.2, 0.5))
            except Exception:  # noqa: BLE001
                try:
                    item.Click()
                except Exception:  # noqa: BLE001
                    return name or None
            window = self.get_main_window()
            current = self.get_current_chat_title(window) if window is not None else None
            current_name = self._normalize_contact_name(current or "")
            return current_name or name or None

    def ensure_chat_target(self, target: str) -> bool:
        if auto is None:
            return False
        with self._uia_lock:
            window = self.get_main_window()
            if window is None:
                return False

            # 选择器说明（中文）：
            # - 先读取当前聊天标题（顶部区域的 TextControl），避免每次都扫描会话列表并点击，降低操作频率。
            target_name = self._normalize_contact_name(target)
            current = self.get_current_chat_title(window)
            current_name = self._normalize_contact_name(current or "")
            print(f"DEBUG: 目标联系人 [{target}]，当前窗口标题识别为 [{current}]")
            
            if current_name and target_name and target_name in current_name:
                print("DEBUG: 当前已在目标窗口，无需切换")
                return True

            print("DEBUG: 尝试在左侧会话列表中寻找目标...")
            session_list = self._locate_session_list(window)
            if session_list is None:
                print("DEBUG: 找不到会话列表控件 (session_list is None)")
                return False

            best = None
            # 遍历会话列表
            children = session_list.GetChildren() or []
            print(f"DEBUG: 会话列表共有 {len(children)} 个子项")
            
            for item in children:
                raw_name = getattr(item, "Name", "") or ""
                # 处理换行符，只取第一行通常是名字
                name_lines = raw_name.split('\n')
                name = name_lines[0].strip() if name_lines else ""
                name = self._normalize_contact_name(name)
                
                if not name:
                    continue
                
                # 调试打印（只打印前几个或匹配的）
                # print(f"DEBUG: 列表项: raw=[{raw_name}] parsed=[{name}]")

                # 选择器说明（中文）：
                # - 会话条目 Name 通常包含联系人名（可能还包含预览/未读提示），因此用“全等或包含”进行宽松匹配。
                if target_name == name or (target_name and target_name in name):
                    best = item
                    print(f"DEBUG: 在列表中找到匹配项: [{name}] (Raw: {raw_name[:20]}...)")
                    break
            
            if best is None:
                print(f"DEBUG: 在左侧列表中未找到 [{target}] (遍历了 {len(children)} 个项)")
                return False
                
            self.click_session_item(best)
            time.sleep(random.uniform(0.5, 1.0)) # 增加等待时间让UI刷新
            
            # 二次确认
            current = self.get_current_chat_title(window)
            current_name = self._normalize_contact_name(current or "")
            print(f"DEBUG: 切换后再次确认窗口标题: [{current}]")
            return bool(current_name and target_name and target_name in current_name)

    def get_current_chat_title(self, window: Any) -> Optional[str]:
        if auto is None:
            return None
        rect = getattr(window, "BoundingRectangle", None)
        bbox = _rect_to_bbox(rect) if rect is not None else None
        if bbox is None:
            return None
        left, top, right, bottom = bbox
        
        # 调整区域：避开左侧列表 (通常宽度<300)，从 280 开始找 (放宽)
        header_top = top
        header_bottom = min(bottom, top + 100) # 缩小高度范围，只看顶部
        header_left = left + 280
        header_right = right

        best = None
        # 收集所有候选标题，增加搜索深度到 25
        candidates = []
        for ctrl in _iter_descendants(window, max_depth=25):
            # 标题通常是 TextControl，但也可能是 ButtonControl (如某些群名)
            # 放宽类型限制，增加 GroupControl/CustomControl 以防万一
            ct = _control_type_name(ctrl)
            if ct not in ("TextControl", "ButtonControl", "PaneControl", "GroupControl", "CustomControl"):
                continue

            text = (getattr(ctrl, "Name", "") or "").strip()
            if not text:
                continue
                
            r = getattr(ctrl, "BoundingRectangle", None)
            b = _rect_to_bbox(r) if r is not None else None
            if b is None:
                continue
                
            # 必须完全在 Header 区域内
            cx = (b[0] + b[2]) // 2
            cy = (b[1] + b[3]) // 2
            
            # self._logger.info(f"DEBUG: Checking title candidate: '{text}' at ({cx}, {cy}) HeaderRegion: X[{header_left}-{header_right}] Y[{header_top}-{header_bottom}]")

            if not (header_left <= cx <= header_right and header_top <= cy <= header_bottom):
                continue
                
            if len(text) > 40:
                continue
            
            # 排除一些常见的顶部干扰词
            if text in ("微信", "通讯录", "发现", "我", "朋友圈", "小程序", "视频号", "搜一搜", "看一看", "文件传输助手", "置顶", "最小化", "最大化", "关闭", "还原"):
                continue

            # 记录 (中心点Y坐标, 文本)
            candidates.append((cy, text))

        if candidates:
            # 按 Y 坐标排序，最靠上的通常是标题
            candidates.sort(key=lambda x: x[0])
            
            # 调试：打印前3个候选
            debug_candidates = [f"{t}@{y}" for y, t in candidates[:3]]
            self._logger.debug(f"标题候选: {debug_candidates}")
            
            best = candidates[0][1]
            self._logger.debug(f"识别到当前聊天标题: {best}")
            
        return best

    def _get_content_root(self, window: Any) -> Any:
        root = window
        try:
            for child in window.GetChildren() or []:
                # 兼容 ClassName 或 Name 为 MMUIRenderSubWindowHW 的情况
                name = getattr(child, "Name", "")
                cls = getattr(child, "ClassName", "")
                if cls == "MMUIRenderSubWindowHW" or name == "MMUIRenderSubWindowHW":
                    return child
        except Exception:
            pass
        return root

    def find_input_box(self, window: Any) -> Optional[Any]:
        if auto is None:
            return None
        
        # 准备搜索根节点列表：优先尝试 ContentRoot，失败则兜底尝试 Window
        roots = []
        content_root = self._get_content_root(window)
        roots.append(content_root)
        if content_root != window:
            roots.append(window)

        # 策略0：精确查找 ClassName='mmui::ChatInputField' (新版微信特征)
        for root in roots:
            try:
                edit = auto.EditControl(searchFromControl=root, searchDepth=25, ClassName="mmui::ChatInputField")
                if edit and edit.Exists(0, 0):
                    return edit
            except Exception:
                pass

        # 策略1：优先按 AutomationId 匹配新版输入框
        for root in roots:
            try:
                # 增加 searchDepth 到 25 以防层级过深
                edit = auto.EditControl(searchFromControl=root, searchDepth=25, AutomationId="chat_input_field")
                if edit and edit.Exists(0, 0):
                    return edit
            except Exception:
                pass

        # 策略2：尝试旧版 Name="输入" (兼容旧版微信)
        for root in roots:
            try:
                edit = auto.EditControl(searchFromControl=root, searchDepth=15, Name="输入")
                if edit and edit.Exists(0, 0):
                    return edit
            except Exception:
                pass

        # 策略3：按当前聊天标题名称匹配（某些版本 Edit.Name=当前会话名）
        try:
            title = self.get_current_chat_title(window)
            if title:
                for root in roots:
                    try:
                        edit = auto.EditControl(searchFromControl=root, searchDepth=25, Name=title)
                        if edit and edit.Exists(0, 0):
                            return edit
                    except Exception:
                        pass
        except Exception:
            pass

        # 策略4：空间位置推断法（最靠下的 Edit）
        try:
            candidates = self._collect_edit_controls(window)
            if candidates:
                candidates.sort(key=lambda x: x.BoundingRectangle.top)
                best_edit = candidates[-1]
                win_rect = window.BoundingRectangle
                if best_edit.BoundingRectangle.top > (win_rect.top + (win_rect.height() * 0.3)):
                    return best_edit
        except Exception as e:
            self._logger.warning(f"查找输入框兜底策略异常: {e}")
            
        return None

    def _collect_edit_controls(self, window: Any) -> List[Any]:
        edits: List[Any] = []
        # 这里也尝试从 window 搜索，防止 _get_content_root 找错
        roots = []
        content_root = self._get_content_root(window)
        roots.append(content_root)
        if content_root != window:
            roots.append(window)
            
        for root in roots:
            current_edits = []
            for ctrl in _iter_descendants(root, max_depth=25):
                if _control_type_name(ctrl) != "EditControl":
                    continue
                rect = getattr(ctrl, "BoundingRectangle", None)
                bbox = _rect_to_bbox(rect)
                if bbox is None:
                    continue
                width = bbox[2] - bbox[0]
                height = bbox[3] - bbox[1]
                if width < 10 or height < 10:
                    continue
                current_edits.append(ctrl)
            
            if current_edits:
                edits.extend(current_edits)
                # 如果从 content_root 找到了，就不需要再从 window 找了（避免重复和性能浪费）
                break
                
        return edits

    def find_send_button(self, window: Any) -> Optional[Any]:
        if auto is None:
            return None
            
        # 准备搜索根节点列表
        roots = []
        content_root = self._get_content_root(window)
        roots.append(content_root)
        if content_root != window:
            roots.append(window)

        try:
            # 策略1：精确查找
            for root in roots:
                # 优先精确匹配按钮文案
                for name in ("发送", "发送(S)"):
                    btn = auto.ButtonControl(searchFromControl=root, searchDepth=25, Name=name)
                    if btn and btn.Exists(0, 0):
                        return btn

            # 策略2：兜底遍历控件树，选择最靠下且名称包含“发送”的按钮
            for root in roots:
                candidates = []
                for ctrl in _iter_descendants(root, max_depth=25):
                    if _control_type_name(ctrl) != "ButtonControl":
                        continue
                    text = (getattr(ctrl, "Name", "") or "")
                    if "发送" not in text:
                        continue
                    rect = getattr(ctrl, "BoundingRectangle", None)
                    b = _rect_to_bbox(rect)
                    if b is None:
                        continue
                    candidates.append((b[3], ctrl))
                if candidates:
                    candidates.sort(key=lambda x: x[0])
                    return candidates[-1][1]
                    
        except Exception:  # noqa: BLE001
            pass
        return None

    def _analyze_item_alignment(self, item: Any, list_bbox: Tuple[int, int, int, int]) -> str:
        from PIL import ImageStat, ImageGrab
        l, _, r, _ = list_bbox
        list_width = r - l
        
        try:
            rect = item.BoundingRectangle
            # 增加高度补偿：上下各扩展 5 像素，防止条目太窄截不到头像中心
            bbox = (rect.left, rect.top - 5, rect.right, rect.bottom + 5)
            
            screenshot = ImageGrab.grab(bbox=bbox)
            img_w, img_h = screenshot.size
            
            # 头像区域（左右各取 70 像素）
            zone_w = 70
            l_zone = screenshot.crop((0, 0, zone_w, img_h))
            r_zone = screenshot.crop((img_w - zone_w, 0, img_w, img_h))
            
            def get_score(region):
                stat = ImageStat.Stat(region)
                return sum(stat.stddev) / 3.0

            l_score = get_score(l_zone)
            r_score = get_score(r_zone)
            
            # 如果两边都很纯净（得分都小于 2），可能是系统消息，默认为对方
            if l_score < 2 and r_score < 2:
                return "other"

            # 判定：哪边颜色杂乱（得分高），哪边就有头像
            return "self" if r_score > l_score else "other"
                
        except Exception as e:
            return "other"

    def extract_latest_messages(self, contact_hint: str) -> List[Dict[str, Any]]:
        if auto is None: return []
        with self._uia_lock:
            window = self.get_main_window()
            if window is None: return []
            normalized_contact = self._normalize_contact_name(contact_hint)
            if not normalized_contact:
                normalized_contact = self._normalize_contact_name(self.get_current_chat_title(window) or "")
            msg_list = self._locate_message_list(window)
            if msg_list is None: return []
            
            list_bbox = _rect_to_bbox(getattr(msg_list, "BoundingRectangle", None))
            if not list_bbox: return []

            items = msg_list.GetChildren() or []
            if not items: return []

            collected_messages = []
            
            scan_limit = max(5, int(self._cfg.message_scan_limit))
            for item in items[-scan_limit:]:
                msg = self._extract_message_from_item(normalized_contact or contact_hint, item)
                if not msg: continue 

                # 判定消息方向
                direction = self._analyze_item_alignment(item, list_bbox)
                
                msg['is_self'] = (direction == "self")
                collected_messages.append(msg)
            
            # 标记需要回复的消息：
            # 只有当最新的一条消息是“对方”发的，才触发回复。
            # 如果最后一条是自己发的，说明已经回复过了。
            if collected_messages:
                last_msg = collected_messages[-1]
                if not last_msg.get('is_self', False):
                    last_msg['trigger_reply'] = True
            
            return collected_messages

    def _extract_message_from_item(self, contact_hint: str, item: Any) -> Optional[Dict[str, Any]]:
        try:
            ui_id = item.GetRuntimeId()
        except:
            ui_id = None

        raw_name = (getattr(item, "Name", "") or "").strip()
        cls_name = _safe_attr(item, "ClassName")
        
        # 排除时间戳（如 "08:18"）
        if re.fullmatch(r"(\d{1,2}:\d{2}|昨天.*|星期.*|202\d年.*)", raw_name):
            return None
            
        # 针对 mmui 版微信，TextItemView 的 Name 就是消息文字
        if "ChatTextItemView" in cls_name and raw_name:
            return {"contact": contact_hint, "type": "text", "content": raw_name, "timestamp": _now_iso(), "ui_id": ui_id}
        
        # 处理非文本消息（如表情、图片、文件），遍历子级
        for ctrl in _iter_descendants(item, max_depth=5):
            name = (getattr(ctrl, "Name", "") or "").strip()
            if name and not re.fullmatch(r"(\d{1,2}:\d{2})", name):
                return {"contact": contact_hint, "type": "text", "content": name, "timestamp": _now_iso(), "ui_id": ui_id}
        
        return None

    def set_text_and_send(self, target: str, text: str) -> bool:
        if auto is None:
            print("DEBUG: auto 库未加载")
            return False
        
        # 模拟真人反应延迟：收到指令后不会立即动作，而是有一个自然的反应时间 (0.2 - 0.6秒，熟练客服)
        time.sleep(random.uniform(0.2, 0.6))

        with self._uia_lock:
            window = self.get_main_window()
            if window is None:
                print("DEBUG: 找不到微信主窗口")
                return False
            
            try:
                window.SetActive()
            except:
                pass

            # 切换到目标聊天窗口
            if not self.ensure_chat_target(target):
                print(f"DEBUG: 无法切换到 [{target}]，可能是名字不匹配")
                return False

            # 寻找输入框
            edit = self.find_input_box(window)
            if edit is None:
                print("DEBUG: 致命错误 - 找不到输入框！")
                # 尝试打印一下窗口里所有的 Edit 控件，看看有没有活着的
                edits = self._collect_edit_controls(window)
                print(f"DEBUG: 当前窗口共找到 {len(edits)} 个 Edit 控件")
                return False

            print(f"DEBUG: 找到输入框 (Rect={edit.BoundingRectangle})，正在粘贴...")
            
            # 粘贴文本
            ok = self._set_edit_value(edit, text)
            if not ok:
                print("DEBUG: 粘贴动作失败")
                return False

            # 随机延迟模拟真人
            time.sleep(random.uniform(0.3, 0.8))

            # 寻找发送按钮
            send_btn = self.find_send_button(window)
            if send_btn is not None:
                print("DEBUG: 找到发送按钮，点击中...")
                try:
                    send_btn.Click(simulateMove=True)
                    return True
                except Exception as e:
                    print(f"DEBUG: 点击发送按钮出错: {e}")

            # 如果找不到按钮，回车发送
            print("DEBUG: 未找到发送按钮，尝试回车发送...")
            try:
                auto.SendKeys("{Enter}")
                return True
            except Exception as e:
                print(f"DEBUG: 回车发送出错: {e}")
                return False

    def _set_edit_value(self, edit: Any, text: str) -> bool:
        """
        使用剪贴板粘贴的方式输入文本，支持中文、Emoji和长文本。
        步骤：聚焦 -> 全选删除(防残留) -> 复制 -> 粘贴
        """
        if pyperclip is None:
            self._logger.error("缺少 pyperclip 依赖，无法使用剪贴板发送文本。请运行: pip install pyperclip")
            return False

        try:
            # 1. 聚焦输入框
            # simulateMove=True 模拟鼠标移动，规避部分风控
            edit.Click(simulateMove=True, waitTime=0.2)
            
            # 2. 清空现有内容 (防止之前遗留文字)
            # 发送 Ctrl+A
            edit.SendKeys("{Ctrl}a", waitTime=0.1) 
            # 发送 Delete
            edit.SendKeys("{Delete}", waitTime=0.1)

            # 3. 将文本存入系统剪贴板
            pyperclip.copy(text)
            
            # 4. 粘贴内容 (Ctrl+V)
            # 注意：这里使用 auto.SendKeys 配合 waitTime 确保粘贴动作完成
            edit.SendKeys("{Ctrl}v", waitTime=0.2)

            # 5. 再次短暂等待，防止后续立即回车导致粘贴未上屏
            time.sleep(0.1)
            
            return True
        except Exception as e:
            self._logger.error(f"输入文本异常: {e}")
            return False


class Reporter(threading.Thread):
    def __init__(self, java_client: JavaClient, logger: logging.Logger) -> None:
        super().__init__(name="Reporter", daemon=True)
        self._java_client = java_client
        self._logger = logger
        self._q: "queue.Queue[Dict[str, Any]]" = queue.Queue()
        self._stopping = threading.Event()

    def submit(self, payload: Dict[str, Any]) -> None:
        self._q.put(payload)

    def stop(self) -> None:
        self._stopping.set()
        self._q.put({})

    def run(self) -> None:
        while not self._stopping.is_set():
            payload = self._q.get()
            if not payload:
                continue
            try:
                self._java_client.post_message(payload)
            except Exception as e:  # noqa: BLE001
                self._logger.warning("Reporter 发送失败: %s", e)


class Listener:
    def __init__(self, cfg: BridgeConfig, ui: WeChatUI, reporter: Reporter, logger: logging.Logger, poller: Any = None) -> None:
        self._cfg = cfg
        self._ui = ui
        self._reporter = reporter
        self._logger = logger
        self._poller = poller
        self._processed_sigs = set()
        # 初始化下一次扫描未读的时间（当前时间 + 随机延迟），避免启动即扫描
        self._next_unread_scan_time = time.time() + random.uniform(2.0, 5.0)

    def process_cycle(self) -> None:
        try:
            self._logger.debug("进入处理循环 (process_cycle)")
            # self._ui.guard_popups()
            # self._logger.debug("Skipped guard_popups")
            
            now = time.time()
            
            # 1. 优先检查当前所在的会话窗口 (高频，像人一样盯着当前聊天看)
            #    这部分开销小，且能及时响应正在进行的对话
            try:
                current_title = self._ui.get_current_chat_title(self._ui.get_main_window())
                current_contact = self._ui._normalize_contact_name(current_title or "")
                if current_contact:
                    # self._logger.info(f"Scanning active window: {current_contact}")
                    self._fetch_and_report(current_contact)
            except Exception as e:
                self._logger.warning("扫描当前窗口异常: %s", e)

            # 2. 检查未读会话列表 (低频，像人一样偶尔瞟一眼左侧列表)
            #    不要每次循环都遍历，太频繁容易被风控且不像人
            if now >= self._next_unread_scan_time:
                self._logger.info("准备扫描未读会话...")
                unread = self._ui.find_unread_sessions()
                self._logger.info("发现 %d 个未读会话", len(unread))
                
                for item in unread:
                    contact = self._ui.click_session_item(item)
                    if not contact:
                        contact = self._ui.get_current_chat_title(self._ui.get_main_window())
                    contact = self._ui._normalize_contact_name(contact or "") or "unknown"
                    
                    # 点击后，模拟阅读时间，稍作停顿
                    time.sleep(random.uniform(1.0, 2.5))
                    
                    self._fetch_and_report(contact)
                
                # 重新计算下一次扫描时间
                interval = random.uniform(self._cfg.unread_scan_interval_min_seconds, self._cfg.unread_scan_interval_max_seconds)
                self._next_unread_scan_time = now + interval
                self._logger.info(f"下一次扫描安排在 {interval:.2f} 秒后")
            else:
                self._logger.debug(f"跳过未读扫描 (距离下次扫描还有 {self._next_unread_scan_time - now:.2f}s)")

        except Exception as e:  # noqa: BLE001
            self._logger.warning("监听循环异常: %s", e)

    def _fetch_and_report(self, contact: str) -> None:
        try:
            contact = self._ui._normalize_contact_name(contact)
            if not contact:
                return
            messages = self._ui.extract_latest_messages(contact)
            for msg in messages:
                # 优先使用 UI 元素的 RuntimeID + 内容 作为签名，防止 ID 复用或内容重复
                if msg.get('ui_id'):
                    sig = _sha1_text(f"{msg['ui_id']}|{msg['content']}")
                else:
                    sig = _sha1_text(f"{msg['contact']}|{msg['content']}|{msg.get('is_self', False)}")
                
                if sig in self._processed_sigs:
                    continue
                
                self._processed_sigs.add(sig)
                if len(self._processed_sigs) > 2000:
                    self._processed_sigs.clear()
                
                # 执行推送
                if self._poller is not None:
                    # 关键日志：如果这里打印了，AI 界面就一定能收到
                    self._logger.info(f"==> 正在推送到 AI 助手界面: {msg['content'][:15]}")
                    self._poller.enqueue(msg)
                
                self._reporter.submit(msg)
        except Exception as e:
            self._logger.warning(f"上报异常: {e}")

    def _signature(self, msg: Dict[str, Any]) -> str:
        if msg.get("type") == "image":
            meta = msg.get("meta", {})
            bbox = meta.get("bbox", {})
            raw = f"image|{bbox}|{msg.get('timestamp','')}"
            return _sha1_text(raw)
        raw = f"text|{msg.get('content','')}"
        return _sha1_text(raw)


class Poller:
    def __init__(self, ui: WeChatUI, logger: logging.Logger) -> None:
        self._ui = ui
        self._logger = logger
        self._queue: queue.Queue = queue.Queue()

    def enqueue(self, payload: Dict[str, Any]) -> None:
        self._queue.put(payload)

    def poll(self, timeout: float = 1.0) -> List[Dict[str, Any]]:
        messages = []
        try:
            # 非阻塞获取所有积压消息，或者阻塞等待第一条
            try:
                # 尝试获取第一条（带超时）
                first = self._queue.get(block=True, timeout=timeout)
                messages.append(first)
                # 获取剩余所有（不等待）
                while True:
                    messages.append(self._queue.get_nowait())
            except queue.Empty:
                pass
            return messages
        except Exception as e:
            self._logger.warning("poll 失败: %s", e)
            return []



class CommandHandler(BaseHTTPRequestHandler):
    server_version = "WeChatBridge/1.0"

    def _json_response(self, status: int, payload: Dict[str, Any]) -> None:
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self) -> None:  # noqa: N802
        if self.path.rstrip("/") == "/health":
            self._json_response(HTTPStatus.OK, {"ok": True})
            return

        # 其他接口涉及 COM 调用，需初始化 COM
        com_init = False
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
                com_init = True
            except Exception:
                pass

        try:
            if self.path.rstrip("/") == "/poll":
                poller: Poller = getattr(self.server, "poller")  # type: ignore[attr-defined]
                messages = poller.poll()
                self._json_response(HTTPStatus.OK, {"ok": True, "messages": messages})
                return
            self._json_response(HTTPStatus.NOT_FOUND, {"ok": False, "error": "not_found"})
        finally:
            if com_init and pythoncom is not None:
                pythoncom.CoUninitialize()

    def do_POST(self) -> None:  # noqa: N802
        if self.path.rstrip("/") != "/command":
            self._json_response(HTTPStatus.NOT_FOUND, {"ok": False, "error": "not_found"})
            return
        
        # 接口涉及 COM 调用，需初始化 COM
        com_init = False
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
                com_init = True
            except Exception:
                pass

        try:
            try:
                length = int(self.headers.get("Content-Length") or "0")
            except Exception:  # noqa: BLE001
                length = 0
            if length <= 0:
                self._json_response(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "empty_body"})
                return
            body = self.rfile.read(length)
            try:
                data = json.loads(body.decode("utf-8"))
            except Exception:  # noqa: BLE001
                self._json_response(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "invalid_json"})
                return
            target = str(data.get("target") or "").strip()
            content = str(data.get("content") or "")
            if not target:
                self._json_response(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "target_required"})
                return
            ui: WeChatUI = getattr(self.server, "ui")  # type: ignore[attr-defined]
            try:
                ok = ui.set_text_and_send(target, content)
                self._json_response(HTTPStatus.OK, {"ok": True, "success": bool(ok)})
            except Exception as e:  # noqa: BLE001
                logger: logging.Logger = getattr(self.server, "logger")  # type: ignore[attr-defined]
                logger.warning("发送指令执行失败: %s", e)
                self._json_response(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": "send_failed"})
        finally:
            if com_init and pythoncom is not None:
                pythoncom.CoUninitialize()

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        # 过滤掉框架自带的 Access Log
        pass


class CommandServer:
    def __init__(self, host: str, port: int, ui: WeChatUI, poller: Poller, logger: logging.Logger) -> None:
        self._host = host
        self._port = port
        self._ui = ui
        self._poller = poller
        self._logger = logger
        self._httpd: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None

    def start(self) -> None:
        httpd = ThreadingHTTPServer((self._host, self._port), CommandHandler)
        setattr(httpd, "ui", self._ui)
        setattr(httpd, "poller", self._poller)
        setattr(httpd, "logger", self._logger)
        self._httpd = httpd

        def _serve() -> None:
            self._logger.info("指令服务已启动: http://%s:%s", self._host, self._port)
            httpd.serve_forever(poll_interval=0.5)

        self._thread = threading.Thread(target=_serve, name="CommandServer", daemon=True)
        self._thread.start()

    def stop(self) -> None:
        if self._httpd is not None:
            self._httpd.shutdown()
            self._httpd.server_close()
            self._httpd = None


def dry_run(cfg: BridgeConfig, logger: logging.Logger) -> int:
    _ensure_deps()
    _ensure_com_initialized(logger)
    _warmup_uia(logger)
    ui = WeChatUI(cfg, logger)
    win = ui.get_main_window()
    if win is None:
        logger.error("未找到微信主窗口: ClassName=%s Name=%s", cfg.window_class_name, cfg.window_name)
        return 2
    try:
        title = getattr(win, "Name", "") or ""
        rect = getattr(win, "BoundingRectangle", None)
        logger.info("找到微信主窗口: title=%s rect=%s", title, rect)
    except Exception:  # noqa: BLE001
        logger.info("找到微信主窗口")

    try:
        unread = ui.find_unread_sessions()
        logger.info("未读会话数量(本轮): %s", len(unread))
    except Exception as e:  # noqa: BLE001
        logger.warning("未读扫描失败: %s", e)
    return 0


def inspect_run(cfg: BridgeConfig, logger: logging.Logger) -> int:
    _ensure_deps()
    _ensure_com_initialized(logger)
    _warmup_uia(logger)
    ui = WeChatUI(cfg, logger)
    win = ui.get_main_window()
    if win is None:
        logger.error("未找到微信主窗口: ClassName=%s Name=%s", cfg.window_class_name, cfg.window_name)
        return 2
    logger.info("开始深度遍历控件树")
    _inspect_window_tree(win, logger)
    return 0


def self_test(cfg: BridgeConfig, logger: logging.Logger) -> int:
    if auto is None:
        logger.info("Self-test 跳过 UIA：uiautomation 未加载")
        return 0
    else:
        _ensure_com_initialized(logger)
        _warmup_uia(logger)
        ui = WeChatUI(cfg, logger)
        poller = Poller(ui, logger)

    server = CommandServer(cfg.server_host, cfg.server_port, ui, poller, logger)
    server.start()
    logger.info("Self-test 运行中(3s)... 可请求 /health /poll /command")
    time.sleep(3)
    server.stop()
    logger.info("Self-test 完成")
    return 0


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(prog="wechat_bridge", add_help=True)
    parser.add_argument("--config", default=os.path.join(os.path.dirname(__file__), "config.yaml"))
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--inspect", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    logger = setup_logging(cfg)

    if args.dry_run:
        return dry_run(cfg, logger)
    if args.inspect:
        return inspect_run(cfg, logger)
    if args.self_test:
        return self_test(cfg, logger)

    _ensure_deps()
    _ensure_com_initialized(logger)
    _warmup_uia(logger)
    if auto is not None and hasattr(auto, "SetGlobalSearchTimeout"):
        try:
            auto.SetGlobalSearchTimeout(1.0)
        except Exception:  # noqa: BLE001
            pass

    ui = WeChatUI(cfg, logger)
    
    # 初始化主动上报组件
    java_client = JavaClient(cfg, logger)
    reporter = Reporter(java_client, logger)
    reporter.start()
    
    poller = Poller(ui, logger)
    listener = Listener(cfg, ui, reporter, logger, poller)
    
    server = CommandServer(cfg.server_host, cfg.server_port, ui, poller, logger)
    server.start()
    logger.info("wechat_bridge 已启动 (监听模式)")
    try:
        while True:
            listener.process_cycle()
            _sleep_with_jitter(cfg.scan_interval_seconds, cfg.scan_jitter_seconds)
    except KeyboardInterrupt:
        logger.info("收到退出信号，正在停止...")
    finally:
        reporter.stop()
        server.stop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
