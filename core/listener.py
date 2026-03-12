import logging
import random
import time
from typing import Any, Set, TYPE_CHECKING

from .config import BridgeConfig
from . import utils

if TYPE_CHECKING:
    from .ui import WeChatUI
    from .network import Reporter, Poller


class Listener:
    def __init__(self, cfg: BridgeConfig, ui: "WeChatUI", reporter: "Reporter", logger: logging.Logger, poller: "Poller" = None) -> None:
        self._cfg = cfg
        self._ui = ui
        self._reporter = reporter
        self._logger = logger
        self._poller = poller
        self._processed_sigs: Set[str] = set()
        self._next_unread_scan_time = time.time() + random.uniform(2.0, 5.0)

    def process_cycle(self) -> None:
        try:
            # self._logger.debug("进入处理循环 (process_cycle)")
            
            now = time.time()
            
            # 1. Check current chat window (High frequency)
            try:
                current_title = self._ui.get_current_chat_title(self._ui.get_main_window())
                current_contact = self._ui._normalize_contact_name(current_title or "")
                if current_contact:
                    self._fetch_and_report(current_contact)
            except Exception as e:
                self._logger.warning("扫描当前窗口异常: %s", e)

            # 2. Check unread session list (Low frequency)
            if now >= self._next_unread_scan_time:
                self._logger.info("准备扫描未读会话...")
                unread = self._ui.find_unread_sessions()
                self._logger.info("发现 %d 个未读会话", len(unread))
                
                if unread:
                    item = unread[0]
                    self._logger.info(f"本轮处理第一个未读会话 (共 {len(unread)} 个)")
                    
                    try:
                        contact = self._ui.click_session_item(item)
                        if not contact:
                            contact = self._ui.get_current_chat_title(self._ui.get_main_window())
                        contact = self._ui._normalize_contact_name(contact or "") or "unknown"
                        
                        time.sleep(random.uniform(1.0, 2.5))
                        
                        self._fetch_and_report(contact)
                    except Exception as e:
                        self._logger.error(f"处理未读会话时出错: {e}")

                interval = random.uniform(self._cfg.unread_scan_interval_min_seconds, self._cfg.unread_scan_interval_max_seconds)
                self._next_unread_scan_time = now + interval
                self._logger.info(f"下一次扫描安排在 {interval:.2f} 秒后")
            else:
                pass 

        except Exception as e:
            self._logger.warning("监听循环异常: %s", e)

    def _fetch_and_report(self, contact: str) -> None:
        try:
            contact = self._ui._normalize_contact_name(contact)
            if not contact:
                self._logger.warning("Fetched contact name is empty, skipping report.")
                return
            messages = self._ui.extract_latest_messages(contact)
            for msg in messages:
                if msg.get('ui_id'):
                    sig = utils.sha1_text(f"{msg['ui_id']}|{msg['content']}")
                else:
                    sig = utils.sha1_text(f"{msg['contact']}|{msg['content']}|{msg.get('is_self', False)}")
                
                if sig in self._processed_sigs:
                    continue
                
                self._processed_sigs.add(sig)
                if len(self._processed_sigs) > 2000:
                    self._processed_sigs.clear()
                
                if self._poller is not None:
                    self._logger.info(f"==> 正在推送到 AI 助手界面: {msg['content'][:15]}")
                    self._poller.enqueue(msg)
                
                self._reporter.submit(msg)
        except Exception as e:
            self._logger.warning(f"上报异常: {e}")
