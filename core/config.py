from __future__ import annotations
import dataclasses
import os
import sys
import re
from typing import Any, Dict, List, Tuple
from . import utils

@dataclasses.dataclass(frozen=True)
class BridgeConfig:
    java_receive_url: str = "http://localhost:8080/api/wechat/receive"
    java_timeout_seconds: float = 5.0
    java_retry_max: int = 3
    java_retry_backoff_base_seconds: float = 0.6

    # Changed default from mmui::MainWindow to WeChatMainWndForPC for version 3.9.12
    window_class_name: str = "WeChatMainWndForPC"
    window_name: str = "微信"

    scan_interval_seconds: float = 0.6
    scan_jitter_seconds: float = 0.3
    unread_max_per_round: int = 5
    message_scan_limit: int = 10

    send_delay_min_seconds: float = 0.5
    send_delay_max_seconds: float = 2.0
    click_move_min_seconds: float = 0.18
    click_move_max_seconds: float = 0.55

    unread_scan_interval_min_seconds: float = 1.5
    unread_scan_interval_max_seconds: float = 4.0

    server_host: str = "127.0.0.1"
    server_port: int = 51234

    log_file: str = "wechat_bridge.log"
    log_level: str = "INFO"
    log_max_bytes: int = 10 * 1024 * 1024
    log_backup_count: int = 3


def _parse_simple_yaml(text: str) -> Dict[str, Any]:
    root: Dict[str, Any] = {}
    stack: List[Tuple[int, Dict[str, Any]]] = [(0, root)]

    for raw in text.splitlines():
        line = utils.strip_yaml_comment(raw).rstrip("\r\n")
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
            current[key] = utils.parse_yaml_scalar(rest)

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
    except Exception:
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
        except Exception:
            pass

    if not os.path.isabs(cfg.log_file):
        # Determine target log path
        log_path = os.path.join(config_dir, cfg.log_file)
        
        # Check if we should use AppData instead
        use_appdata = False
        # If packaged (frozen), we are likely in Program Files which is read-only
        if getattr(sys, 'frozen', False):
            use_appdata = True
        # Or if the config directory is explicitly not writable
        elif not os.access(config_dir, os.W_OK):
            use_appdata = True
            
        if use_appdata:
            appdata = os.getenv('APPDATA')
            if appdata:
                # Use a specific subdirectory for the application logs
                log_dir = os.path.join(appdata, "ShijieAIAssistant", "logs")
                try:
                    os.makedirs(log_dir, exist_ok=True)
                    log_path = os.path.join(log_dir, os.path.basename(cfg.log_file))
                except Exception:
                    # Fallback to temp dir if AppData creation fails
                    import tempfile
                    log_path = os.path.join(tempfile.gettempdir(), os.path.basename(cfg.log_file))

        cfg = dataclasses.replace(cfg, log_file=log_path)
    return cfg
