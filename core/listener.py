import logging
import random
import time
from typing import Any, Dict, List, TYPE_CHECKING

from .config import BridgeConfig
from . import utils

if TYPE_CHECKING:
    from .ui import WeChatUI
    from .network import Poller


class Listener:
    def __init__(self, cfg: BridgeConfig, ui: "WeChatUI", logger: logging.Logger, poller: "Poller" = None) -> None:
        self._cfg = cfg
        self._ui = ui
        self._logger = logger
        self._poller = poller
        self._last_fingerprints_by_contact: Dict[str, List[str]] = {}
        self._next_unread_scan_time = time.time()
        self._managed_mode = "full"

    def set_managed_mode(self, mode: str) -> None:
        normalized = (mode or "").strip().lower()
        if normalized not in ("full", "semi"):
            self._logger.warning("收到无效托管模式，忽略: %s", mode)
            return
        if normalized == self._managed_mode:
            return
        self._managed_mode = normalized
        self._logger.info("托管模式已切换为: %s", self._managed_mode)

    def prime_initial_state(self) -> None:
        primed_contacts = set()
        try:
            window = self._ui.get_main_window()
            if window is None:
                self._logger.warning("启动预热失败：未找到微信主窗口")
                return

            original_contact = self._ui._normalize_contact_name(self._ui.get_current_chat_title(window) or "")
            if self._snapshot_contact(original_contact):
                primed_contacts.add(original_contact)

            if self._managed_mode == "full":
                unread_items = self._ui.find_unread_sessions()
                self._logger.info("启动预热发现 %d 个未读会话，开始建立历史基线", len(unread_items))
                for item in unread_items:
                    try:
                        contact = self._ui.click_session_item(item)
                        contact = self._ui._normalize_contact_name(contact or "")
                        if not contact or contact in primed_contacts:
                            continue
                        if self._snapshot_contact(contact):
                            primed_contacts.add(contact)
                    except Exception as e:
                        self._logger.warning("启动预热未读会话失败: %s", e)

                if original_contact:
                    try:
                        self._ui.ensure_chat_target(original_contact)
                    except Exception as e:
                        self._logger.warning("启动预热恢复原会话失败: %s", e)

            self._logger.info("启动预热完成，已建立 %d 个会话基线", len(primed_contacts))
        except Exception as e:
            self._logger.warning("启动预热异常: %s", e)

    def process_cycle(self) -> None:
        try:
            # self._logger.debug("进入处理循环 (process_cycle)")
            
            now = time.time()
            
            # 1. Check current chat window (High frequency)
            try:
                main_win = self._ui.get_main_window()
                current_title = self._ui.get_current_chat_title(main_win)
                current_contact = self._ui._normalize_contact_name(current_title or "")
                
                if current_contact in ("服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号", "文件传输助手"):
                    self._logger.info(f"检测到处于特殊会话 [{current_contact}]，尝试退出...")
                    if self._managed_mode == "full":
                        if current_contact in ("服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号"):
                            self._ui._check_and_exit_subscription_folder(main_win)

                        session_list = self._ui._locate_session_list(main_win)
                        if session_list:
                            children = session_list.GetChildren() or []
                            for item in children[:5]:
                                name = self._ui._normalize_contact_name(getattr(item, "Name", "") or "")
                                if name and name not in ("服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号", "文件传输助手"):
                                    self._logger.info(f"点击常规会话 [{name}] 以退出特殊会话")
                                    try:
                                        if not self._ui._click_control(item):
                                            item.Click(simulateMove=True)
                                    except Exception:
                                        pass
                                    break
                    else:
                        self._logger.info("当前为半托管模式，跳过特殊会话自动切换")
                elif current_contact:
                    self._fetch_and_report(current_contact)
            except Exception as e:
                self._logger.warning("扫描当前窗口异常: %s", e)

            # 2. Check unread session list (Low frequency)
            if now >= self._next_unread_scan_time:
                if self._managed_mode != "full":
                    interval = random.uniform(self._cfg.unread_scan_interval_min_seconds, self._cfg.unread_scan_interval_max_seconds)
                    self._next_unread_scan_time = now + interval
                    self._logger.info("当前为半托管模式，跳过未读会话自动点击")
                    return
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
                self._logger.warning("联系人为空，跳过上报")
                return
            messages = self._ui.extract_latest_messages(contact)
            if not messages:
                return

            current_fingerprints = [self._message_fingerprint(msg) for msg in messages]
            previous_fingerprints = self._last_fingerprints_by_contact.get(contact, [])
            overlap = self._calculate_overlap(previous_fingerprints, current_fingerprints)
            new_messages = messages[overlap:]

            self._last_fingerprints_by_contact[contact] = current_fingerprints
            if len(self._last_fingerprints_by_contact) > 300:
                self._last_fingerprints_by_contact.clear()

            for msg in new_messages:
                if self._poller is not None:
                    self._logger.info(f"==> 正在推送到 AI 助手界面: {msg['content'][:15]}")
                    self._poller.enqueue(msg)
        except Exception as e:
            self._logger.warning(f"上报异常: {e}")

    def _message_fingerprint(self, msg: dict) -> str:
        raw_ui_id = msg.get("ui_id")
        if isinstance(raw_ui_id, (list, tuple)):
            ui_id = "-".join(str(part) for part in raw_ui_id)
        else:
            ui_id = str(raw_ui_id or "").strip()
        content = str(msg.get("content") or "").strip()
        is_self = bool(msg.get("is_self", False))
        msg_type = str(msg.get("type") or "").strip()
        if ui_id:
            base = f"ui|{ui_id}|{msg_type}|{is_self}|{content}"
        else:
            base = f"text|{msg_type}|{is_self}|{content}"
        return utils.sha1_text(base)

    def _calculate_overlap(self, previous: List[str], current: List[str]) -> int:
        if not previous or not current:
            return 0
        max_overlap = min(len(previous), len(current))
        for size in range(max_overlap, 0, -1):
            if previous[-size:] == current[:size]:
                return size
        return 0

    def _snapshot_contact(self, contact: str) -> bool:
        normalized_contact = self._ui._normalize_contact_name(contact)
        if not normalized_contact:
            return False
        if normalized_contact in ("服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号", "文件传输助手"):
            return False

        messages = self._ui.extract_latest_messages(normalized_contact)
        if not messages:
            self._logger.info("启动预热跳过空会话: %s", normalized_contact)
            return False

        current_fingerprints = [self._message_fingerprint(msg) for msg in messages]
        self._last_fingerprints_by_contact[normalized_contact] = current_fingerprints
        if len(self._last_fingerprints_by_contact) > 300:
            self._last_fingerprints_by_contact.clear()
        self._logger.info("启动预热建立会话基线: %s, 消息数=%d", normalized_contact, len(messages))
        return True
