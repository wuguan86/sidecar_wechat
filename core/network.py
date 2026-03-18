import json
import logging
import os
import queue
import sys
import threading
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from typing import Any, Dict, List, Optional, TYPE_CHECKING

from . import utils

if TYPE_CHECKING:
    from .ui import WeChatUI

try:
    import pythoncom
except ImportError:
    pythoncom = None


class Poller:
    def __init__(self, ui: Any, logger: logging.Logger) -> None:
        # ui dependency is kept for compatibility/future use, though not strictly used in current logic
        self._ui = ui
        self._logger = logger
        self._queue: queue.Queue = queue.Queue()

    def enqueue(self, payload: Dict[str, Any]) -> None:
        self._queue.put(payload)

    def poll(self, timeout: float = 0.2) -> List[Dict[str, Any]]:
        messages = []
        try:
            try:
                first = self._queue.get(block=True, timeout=timeout)
                messages.append(first)
                while True:
                    messages.append(self._queue.get_nowait())
            except queue.Empty:
                pass
            return messages
        except Exception as e:
            self._logger.warning("poll 失败: %s", e)
            return []


# Determine state file path relative to executable or script root
_BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if getattr(sys, 'frozen', False):
    _BASE_DIR = os.path.dirname(sys.executable)

_STATE_FILE = os.path.join(_BASE_DIR, "marketing_state.json")

def _load_state_from_file() -> dict:
    if os.path.exists(_STATE_FILE):
        try:
            with open(_STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def _save_state_to_file(data: dict) -> None:
    try:
        def default(o):
            if isinstance(o, set):
                return list(o)
            return str(o)
        with open(_STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=default)
    except Exception:
        pass


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
            state_provider = getattr(self.server, "state_provider", None)
            state = {}
            if callable(state_provider):
                try:
                    state = state_provider() or {}
                except Exception:
                    state = {}
            self._json_response(HTTPStatus.OK, {"ok": True, **state})
            return

        com_init = False
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
                com_init = True
            except Exception:
                pass

        try:
            if self.path.rstrip("/") == "/poll":
                poller: Poller = getattr(self.server, "poller")
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
            except Exception:
                length = 0
            if length <= 0:
                self._json_response(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "empty_body"})
                return
            body = self.rfile.read(length)
            try:
                data = json.loads(body.decode("utf-8"))
            except Exception:
                self._json_response(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "invalid_json"})
                return
            
            action = str(data.get("action") or "").strip().lower()
            target = str(data.get("target") or "").strip()
            content = str(data.get("content") or "")
            
            ui: "WeChatUI" = getattr(self.server, "ui")
            logger: logging.Logger = getattr(self.server, "logger")

            if action == "marketing_like":
                config = data.get("config") or {}
                state = getattr(self.server, "marketing_like_state", None)
                if not isinstance(state, dict):
                    state = {}
                    setattr(self.server, "marketing_like_state", state)
                try:
                    result = ui.execute_marketing_like(config, state)
                    self._json_response(HTTPStatus.OK, result if isinstance(result, dict) else {"ok": True, "success": True})
                    
                    # Save state
                    try:
                        like_state = getattr(self.server, "marketing_like_state", {})
                        comment_state = getattr(self.server, "marketing_comment_state", {})
                        _save_state_to_file({"like": like_state, "comment": comment_state})
                    except Exception as e:
                        logging.getLogger(__name__).warning("Failed to save state: %s", e)
                except Exception as e:
                    logger.warning("点赞指令执行失败: %s", e)
                    self._json_response(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": "marketing_like_failed"})
                return

            if action == "marketing_comment":
                config = data.get("config") or {}
                state = getattr(self.server, "marketing_comment_state", None)
                if not isinstance(state, dict):
                    state = {}
                    setattr(self.server, "marketing_comment_state", state)
                try:
                    result = ui.execute_marketing_comment(config, state)
                    self._json_response(HTTPStatus.OK, result if isinstance(result, dict) else {"ok": True, "success": True})

                    # Save state
                    try:
                        like_state = getattr(self.server, "marketing_like_state", {})
                        comment_state = getattr(self.server, "marketing_comment_state", {})
                        _save_state_to_file({"like": like_state, "comment": comment_state})
                    except Exception as e:
                        logging.getLogger(__name__).warning("Failed to save state: %s", e)
                except Exception as e:
                    logger.warning("评论指令执行失败: %s", e)
                    self._json_response(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": "marketing_comment_failed"})
                return

            if not target:
                self._json_response(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "target_required"})
                return
            try:
                ok = ui.set_text_and_send(target, content)
                self._json_response(HTTPStatus.OK, {"ok": True, "success": bool(ok)})
            except Exception as e:
                logger.warning("发送指令执行失败: %s", e)
                self._json_response(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": "send_failed"})
        finally:
            if com_init and pythoncom is not None:
                pythoncom.CoUninitialize()

    def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
        pass


class CommandServer:
    def __init__(
        self,
        host: str,
        port: int,
        ui: Any,
        poller: Poller,
        logger: logging.Logger,
        state_provider: Optional[Any] = None
    ) -> None:
        self._host = host
        self._port = port
        self._ui = ui
        self._poller = poller
        self._logger = logger
        self._state_provider = state_provider
        self._httpd: Optional[ThreadingHTTPServer] = None
        self._thread: Optional[threading.Thread] = None

    def start(self) -> None:
        httpd = ThreadingHTTPServer((self._host, self._port), CommandHandler)
        setattr(httpd, "ui", self._ui)
        setattr(httpd, "poller", self._poller)
        setattr(httpd, "logger", self._logger)
        setattr(httpd, "state_provider", self._state_provider)
        
        full_state = _load_state_from_file()
        setattr(httpd, "marketing_like_state", full_state.get("like", {}))
        setattr(httpd, "marketing_comment_state", full_state.get("comment", {}))
        
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
