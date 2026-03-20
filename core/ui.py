from __future__ import annotations
import ctypes
import logging
import random
import re
import threading
import time
import subprocess
import io
import requests
from typing import Any, Dict, List, Optional, Tuple, Iterable

from .config import BridgeConfig
from . import utils

try:
    import pyperclip
except ImportError:
    pyperclip = None

try:
    import uiautomation as auto
except Exception:
    auto = None

try:
    import pythoncom
except Exception:
    pythoncom = None

try:
    from PIL import ImageGrab, ImageStat
except Exception:
    ImageGrab = None
    ImageStat = None

# Ensure DPI Awareness
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1) 
except:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except:
        pass


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
    except Exception:
        return ""


def _safe_attr(ctrl: Any, attr: str) -> str:
    try:
        val = getattr(ctrl, attr, "")
        if val is None:
            return ""
        return str(val)
    except Exception:
        return ""


def _rect_text(ctrl: Any) -> str:
    try:
        rect = getattr(ctrl, "BoundingRectangle", None)
        if rect is None:
            return ""
        return str(rect)
    except Exception:
        return ""


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
        except Exception:
            children = []
        for child in children:
            yield child
            q.append((child, depth + 1))


def inspect_window_tree(window: Any, logger: logging.Logger) -> None:
    keywords = ["未读", "消息", "输入", "发送", "会话"]
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
        except Exception:
            children = []
        for child in reversed(children):
            stack.append((child, depth + 1))


class WeChatUI:
    def __init__(self, cfg: BridgeConfig, logger: logging.Logger) -> None:
        self._cfg = cfg
        self._logger = logger
        self._uia_lock = threading.RLock()
        self._cached_main = None
        self._tree_logged_handle = None
        self._last_unread_debug_time = 0.0
        
        # ======== 新增：核心 UI 控件缓存 ========
        self._cached_session_list = None
        self._cached_message_list = None
        self._cached_input_box = None
        self._cached_chat_title_ctrl = None

    def _normalize_contact_name(self, name: str) -> str:
        if not name:
            return ""
        name = name.strip()
        if not name:
            return ""
        name = re.sub(r"\[\d+条\]", "", name)
        
        parts = [part.strip() for part in name.splitlines() if part.strip()]
        if parts:
            name = parts[0]
        name = re.sub(r"\s+", " ", name).strip()
        name = re.sub(r"(?:\d+\s*条新消息|未读)$", "", name).strip()
        return name

    def _brief_control(self, ctrl: Any) -> str:
        return f"{_control_type_name(ctrl)}|{_safe_attr(ctrl, 'ClassName')}|{_safe_attr(ctrl, 'Name')}|{_rect_text(ctrl)}"

    def _click_control(self, ctrl: Any) -> bool:
        if ctrl is None:
            return False
        
        # Ensure visible
        try:
            if not ctrl.IsOffscreen:
                 ctrl.SetFocus()
        except:
            pass

        try:
            # Try to get center point explicitly
            rect = getattr(ctrl, "BoundingRectangle", None)
            if rect:
                x = (rect.left + rect.right) // 2
                y = (rect.top + rect.bottom) // 2
                auto.Click(x, y)
                return True
            else:
                ctrl.Click(simulateMove=True)
                return True
        except Exception:
            pass
            
        try:
            ctrl.Click()
            return True
        except Exception:
            return False

    def _find_named_clickable(self, root: Any, names: List[str], search_depth: int = 18) -> Optional[Any]:
        if auto is None or root is None:
            return None
        normalized = [str(name).strip() for name in names if str(name).strip()]
        if not normalized:
            return None
        candidates: List[Tuple[int, Any]] = []
        for ctrl in _iter_descendants(root, max_depth=search_depth):
            text = (getattr(ctrl, "Name", "") or "").strip()
            if not text:
                continue
            matched = False
            for name in normalized:
                if text == name or name in text:
                    matched = True
                    break
            if not matched:
                continue
            rect = getattr(ctrl, "BoundingRectangle", None)
            bbox = utils.rect_to_bbox(rect)
            if bbox is None:
                continue
            candidates.append((bbox[1], ctrl))
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[0])
        return candidates[0][1]

    def _get_moments_window(self, main_window: Any) -> Optional[Any]:
        if auto is None:
            return None

        self._logger.info("尝试获取朋友圈窗口...")

        # Check if Moments window already exists
        # Use Win32 API directly to avoid UIA hanging on global search
        try:
            hwnd = ctypes.windll.user32.FindWindowW("SnsWnd", "朋友圈")
            if hwnd and hwnd != 0:
                self._logger.info(f"发现已存在的朋友圈窗口, HWND: {hwnd}")
                return auto.ControlFromHandle(hwnd)
        except Exception as e:
            self._logger.warning(f"检查朋友圈窗口存在性时出错: {e}")

        # Try to open it
        self._logger.info("尝试从主窗口进入朋友圈...")
        
        # Ensure main window is accessible
        try:
            if getattr(main_window, "IsIconic", False):
                self._logger.info("主窗口最小化中，尝试还原...")
                main_window.ShowWindow(auto.SW.Restore)
            main_window.SetFocus()
            self._logger.debug("主窗口已聚焦")
        except Exception as e:
            self._logger.warning(f"主窗口聚焦失败: {e}")

        # Strategy A: Look for "导航" ToolBar first (Based on control tree)
        self._logger.debug("正在查找[导航]工具栏(ToolBarControl)...")
        nav_toolbar = None
        
        # Search for ToolBarControl with Name='导航'
        try:
            for ctrl in _iter_descendants(main_window, max_depth=8):
                if _control_type_name(ctrl) == "ToolBarControl" and "导航" in _safe_attr(ctrl, "Name"):
                    nav_toolbar = ctrl
                    self._logger.debug(f"找到[导航]工具栏: {self._brief_control(ctrl)}")
                    break
        except Exception as e:
            self._logger.warning(f"遍历查找导航栏时出错: {e}")
        
        found_btn = False
        if nav_toolbar:
            self._logger.debug("在[导航]工具栏中查找[朋友圈]按钮...")
            try:
                # Direct children or slight depth
                moments_btn = self._find_named_clickable(nav_toolbar, ["朋友圈"], search_depth=4)
                if moments_btn:
                    self._logger.info(f"找到[朋友圈]按钮: {self._brief_control(moments_btn)}，尝试点击...")
                    if self._click_control(moments_btn):
                        self._logger.debug("点击操作已执行")
                        found_btn = True
                    else:
                        self._logger.warning("点击操作失败")
                else:
                    self._logger.warning("[导航]工具栏中未找到[朋友圈]按钮")
            except Exception as e:
                self._logger.warning(f"在导航栏中操作时出错: {e}")
        else:
            self._logger.warning("未找到[导航]工具栏，尝试全局搜索...")

        if not found_btn:
            # Strategy B: Original Logic (Global search or 'Discovery' tab)
            self._logger.debug("进入备用搜索策略...")
            # 1. Look for direct "朋友圈" button (sidebar)
            moments_btn = self._find_named_clickable(main_window, ["朋友圈"], search_depth=12)
            if moments_btn:
                self._logger.info("点击侧边栏[朋友圈]按钮 (全局搜索)")
                self._click_control(moments_btn)
            else:
                # 2. Look for "发现" -> "朋友圈"
                self._logger.info("未找到侧边栏[朋友圈]，尝试寻找[发现]...")
                discover = self._find_named_clickable(main_window, ["发现"], search_depth=12)
                if discover:
                    self._logger.info("点击[发现]按钮")
                    self._click_control(discover)
                    time.sleep(0.8) # Wait for animation
                    moments_btn = self._find_named_clickable(main_window, ["朋友圈"], search_depth=12)
                    if moments_btn:
                        self._logger.info("点击[发现]面板中的[朋友圈]按钮")
                        self._click_control(moments_btn)
                    else:
                        self._logger.warning("[发现]面板中未找到[朋友圈]按钮")
                else:
                    self._logger.warning("未找到[发现]按钮")

        # Wait for window to appear
        self._logger.debug("等待朋友圈窗口出现...")
        sns_wnd = None
        for i in range(12):  # 6 seconds max
            # Re-check via FindWindowW to avoid hang
            hwnd = ctypes.windll.user32.FindWindowW("SnsWnd", "朋友圈")
            if hwnd and hwnd != 0:
                self._logger.info(f"朋友圈窗口成功打开, HWND: {hwnd}")
                sns_wnd = auto.ControlFromHandle(hwnd)
                break
            time.sleep(0.5)

        if not sns_wnd:
            self._logger.warning("尝试打开操作后，仍未检测到朋友圈窗口")
            return None
            
        return sns_wnd

    def _extract_item_text(self, item: Any) -> str:
        parts: List[str] = []
        base = (getattr(item, "Name", "") or "").strip()
        if base:
            parts.append(base)
        for ctrl in _iter_descendants(item, max_depth=6):
            name = (getattr(ctrl, "Name", "") or "").strip()
            if not name:
                continue
            if len(name) > 100:
                continue
            parts.append(name)
        if not parts:
            return ""
        merged = "\n".join(parts)
        merged = re.sub(r"\n{2,}", "\n", merged).strip()
        return merged

    def _extract_item_author(self, text: str) -> str:
        if not text:
            return ""
        for line in text.splitlines():
            value = line.strip()
            if not value:
                continue
            if value in {"朋友圈", "发现", "微信"}:
                continue
            if re.fullmatch(r"(今天|昨天|星期.*|\d{1,2}:\d{2}|[上下]午.*)", value):
                continue
            return value
        return ""

    def _collect_moments_items(self, window: Any) -> List[Any]:
        items: List[Tuple[int, Any]] = []
        self._logger.debug("开始收集朋友圈条目(ListItem)...")
        
        for ctrl in _iter_descendants(window, max_depth=26):
            ct = _control_type_name(ctrl)
            cls_name = _safe_attr(ctrl, "ClassName")
            # Loose match for list items
            if ct != "ListItemControl" and "ListItem" not in cls_name and "listitem" not in cls_name.lower():
                continue
            rect = getattr(ctrl, "BoundingRectangle", None)
            bbox = utils.rect_to_bbox(rect)
            if bbox is None:
                continue
            left, top, right, bottom = bbox
            width = right - left
            height = bottom - top
            
            # self._logger.debug(f"Candidate Item: {cls_name} size={width}x{height} top={top}")
            
            if width < 300 or height < 40: # Relaxed width check
                continue
            if top < 80:
                continue
            items.append((top, ctrl))
            
        if not items:
            self._logger.warning("未找到任何朋友圈条目 (items is empty)")
            return []
            
        items.sort(key=lambda x: x[0])
        unique: List[Any] = []
        last_top = -99999
        for top, ctrl in items:
            if abs(top - last_top) < 8:
                continue
            unique.append(ctrl)
            last_top = top
            if len(unique) >= 30:
                break
        
        self._logger.info(f"收集到 {len(unique)} 个有效朋友圈条目")
        return unique

    def _find_interaction_button(self, item: Any) -> Optional[Any]:
        """Find the comment/interaction button within a moments item"""
        # Strategy 1: Look for "评论" named button
        btn = self._find_named_clickable(item, ["评论"], search_depth=8)
        if btn:
            return btn
            
        # Strategy 2: Look for button with specific characteristics (often bottom-right)
        # The interaction button is usually small and on the right side
        try:
            item_rect = getattr(item, "BoundingRectangle", None)
            item_bbox = utils.rect_to_bbox(item_rect)
            if not item_bbox:
                return None
                
            item_right = item_bbox[2]
            item_bottom = item_bbox[3]
            
            candidates = []
            for child in _iter_descendants(item, max_depth=8):
                ct = _control_type_name(child)
                if ct not in ["ButtonControl", "PaneControl", "ImageControl"]:
                    continue
                
                rect = getattr(child, "BoundingRectangle", None)
                bbox = utils.rect_to_bbox(rect)
                if not bbox:
                    continue
                    
                # Check if it's in the bottom-right area
                # (This is a heuristic and might need adjustment)
                l, t, r, b = bbox
                w = r - l
                h = b - t
                
                # Button size is usually small icon
                if 20 <= w <= 50 and 15 <= h <= 40:
                    candidates.append((child, r, b))
            
            # Sort by proximity to bottom-right
            # Maximize right and bottom
            if candidates:
                candidates.sort(key=lambda x: (x[1] + x[2]), reverse=True)
                # self._logger.debug(f"Found interaction button candidate via heuristic: {self._brief_control(candidates[0][0])}")
                return candidates[0][0]
                
        except Exception as e:
            self._logger.warning(f"Error finding interaction button: {e}")
            
        return None

    def execute_marketing_like(self, config: Dict[str, Any], state: Dict[str, Any]) -> Dict[str, Any]:
        if auto is None:
            return {"ok": False, "error": "uia_not_ready"}
        
        self._logger.info("开始执行朋友圈点赞任务...")
        
        like_start = int(config.get("likeIntervalStart") or 60)
        like_end = int(config.get("likeIntervalEnd") or like_start)
        if like_end < like_start:
            like_end = like_start
        per_friend_limit = max(1, int(config.get("maxDailyLikesPerFriend") or 3))
        total_limit = max(1, int(config.get("maxDailyTotalLikes") or 100))
        keyword_filter = [str(item).strip() for item in (config.get("keywordFilter") or []) if str(item).strip()]

        today = utils.today_key()
        if state.get("date") != today:
            state.clear()
            state.update({
                "date": today,
                "totalLikes": 0,
                "perFriendLikes": {},
                "actedSignatures": set()
            })

        total_likes = int(state.get("totalLikes") or 0)
        per_friend = state.get("perFriendLikes") or {}
        if not isinstance(per_friend, dict):
            per_friend = {}
            state["perFriendLikes"] = per_friend
        acted_signatures = state.get("actedSignatures")
        if isinstance(acted_signatures, list):
            acted_signatures = set(acted_signatures)
            state["actedSignatures"] = acted_signatures
        elif not isinstance(acted_signatures, set):
            acted_signatures = set()
            state["actedSignatures"] = acted_signatures

        if total_likes >= total_limit:
            self._logger.info("今日点赞总数已达上限，跳过任务")
            return {
                "ok": True,
                "skipped": True,
                "reason": "daily_total_limit_reached",
                "dailyTotalLikes": total_likes,
                "dailyTotalLimit": total_limit
            }

        with self._uia_lock:
            window = self.get_main_window()
            if window is None:
                return {"ok": False, "error": "wechat_window_not_found"}

            moments_window = self._get_moments_window(window)
            if not moments_window:
                return {"ok": False, "error": "moments_window_not_found"}

            items = self._collect_moments_items(moments_window)
            if not items:
                return {"ok": False, "error": "moments_item_not_found"}

            for index, item in enumerate(items):
                item_text = self._extract_item_text(item)
                # self._logger.info(f"处理第 {index+1} 条: {item_text[:30]}...")
                
                if not item_text:
                    self._logger.debug(f"第 {index+1} 条无文本，跳过")
                    continue
                signature = utils.sha1_text(item_text)
                if signature in acted_signatures:
                    self._logger.debug(f"第 {index+1} 条已处理过(Signature)，跳过")
                    continue

                if keyword_filter and any(word in item_text for word in keyword_filter):
                    self._logger.debug(f"第 {index+1} 条命中关键词过滤，跳过")
                    continue

                author = self._extract_item_author(item_text) or "unknown"
                friend_count = int(per_friend.get(author) or 0)
                if friend_count >= per_friend_limit:
                    self._logger.debug(f"用户 {author} 点赞数已达上限，跳过")
                    continue

                # Check if already liked based on text (if visible)
                if "取消" in item_text or "已赞" in item_text:
                    self._logger.debug(f"第 {index+1} 条检测到已赞状态，跳过")
                    acted_signatures.add(signature)
                    continue

                # Find the comment/interaction button to open the popup
                # Usually named "评论" or just a button at the bottom right
                self._logger.debug(f"正在寻找第 {index+1} 条的互动按钮...")
                menu_btn = self._find_interaction_button(item)
                if menu_btn is None:
                    self._logger.debug(f"第 {index+1} 条未找到互动按钮，跳过")
                    # Fallback: try to find the button by position or other property if needed
                    # For now, skip if not found
                    continue

                self._logger.debug(f"点击互动按钮: {self._brief_control(menu_btn)}")
                if not self._click_control(menu_btn):
                    self._logger.warning("点击互动按钮失败")
                    continue
                
                # Wait for the popup (SnsLikeToastWnd)
                # Use a loop or Exists with timeout
                self._logger.debug("等待点赞弹窗(SnsLikeToastWnd)...")
                popup = None
                for _ in range(5): # Increased wait time
                    popup_pane = auto.PaneControl(ClassName="SnsLikeToastWnd")
                    if popup_pane.Exists(0, 0):
                        popup = popup_pane
                        self._logger.debug("找到弹窗 (PaneControl)")
                        break
                    popup_wnd = auto.WindowControl(ClassName="SnsLikeToastWnd")
                    if popup_wnd.Exists(0, 0):
                        popup = popup_wnd
                        self._logger.debug("找到弹窗 (WindowControl)")
                        break
                    # Also try finding by Name="点赞" container? No, usually class name is reliable.
                    time.sleep(0.3)
                
                if popup is None or not popup.Exists(1, 0.2):
                    self._logger.warning("未找到点赞弹窗")
                    # Try searching as child of main window or moments window just in case
                    # But usually it is top level.
                    continue

                # Find "赞" button in popup
                self._logger.debug("在弹窗中查找[赞]按钮...")
                like_btn = self._find_named_clickable(popup, ["赞"], search_depth=5)
                
                if like_btn is None:
                    self._logger.debug("未找到[赞]按钮 (可能已赞或界面不匹配)")
                    # Close popup if possible or just continue
                    # Clicking elsewhere closes it usually
                    continue

                self._logger.info(f"找到[赞]按钮: {self._brief_control(like_btn)}，执行点击...")
                if not self._click_control(like_btn):
                    self._logger.warning("点击[赞]按钮失败")
                    continue

                total_likes += 1
                state["totalLikes"] = total_likes
                per_friend[author] = friend_count + 1
                acted_signatures.add(signature)
                self._logger.info(f"点赞成功! Author: {author}, Total: {total_likes}")
                return {
                    "ok": True,
                    "success": True,
                    "liked": True,
                    "author": author,
                    "dailyTotalLikes": total_likes,
                    "dailyTotalLimit": total_limit,
                    "friendLikes": int(per_friend.get(author) or 0),
                    "friendLimit": per_friend_limit,
                    "likeIntervalStart": like_start,
                    "likeIntervalEnd": like_end
                }

            self._logger.info("遍历完所有可见条目，未进行点赞 (可能是无合适条目或都已处理)")
            return {
                "ok": True,
                "success": True,
                "liked": False,
                "reason": "no_eligible_item",
                "dailyTotalLikes": total_likes,
                "dailyTotalLimit": total_limit
            }

    def execute_marketing_comment(self, config: Dict[str, Any], state: Dict[str, Any]) -> Dict[str, Any]:
        if auto is None:
            return {"ok": False, "error": "uia_not_ready"}
        
        self._logger.info("开始执行朋友圈评论任务...")
        
        backend_url = config.get("backendUrl")
        token = config.get("token")
        tenant_id = config.get("tenantId")
        if not backend_url or not token:
            return {"ok": False, "error": "missing_backend_config"}

        comment_start = int(config.get("commentIntervalStart") or 120)
        comment_end = int(config.get("commentIntervalEnd") or comment_start)
        if comment_end < comment_start:
            comment_end = comment_start
        per_friend_limit = max(1, int(config.get("maxDailyCommentsPerFriend") or 3))
        total_limit = max(1, int(config.get("maxDailyTotalComments") or 50))
        keyword_filter = [str(item).strip() for item in (config.get("keywordFilter") or []) if str(item).strip()]

        today = utils.today_key()
        if state.get("date") != today:
            state.clear()
            state.update({
                "date": today,
                "totalComments": 0,
                "perFriendComments": {},
                "actedSignatures": set()
            })

        total_comments = int(state.get("totalComments") or 0)
        per_friend = state.get("perFriendComments") or {}
        if not isinstance(per_friend, dict):
            per_friend = {}
            state["perFriendComments"] = per_friend
        acted_signatures = state.get("actedSignatures")
        if isinstance(acted_signatures, list):
            acted_signatures = set(acted_signatures)
            state["actedSignatures"] = acted_signatures
        elif not isinstance(acted_signatures, set):
            acted_signatures = set()
            state["actedSignatures"] = acted_signatures

        if total_comments >= total_limit:
            self._logger.info("今日评论总数已达上限，跳过任务")
            return {
                "ok": True,
                "skipped": True,
                "reason": "daily_total_limit_reached",
                "dailyTotalComments": total_comments,
                "dailyTotalLimit": total_limit
            }

        with self._uia_lock:
            window = self.get_main_window()
            if window is None:
                return {"ok": False, "error": "wechat_window_not_found"}

            moments_window = self._get_moments_window(window)
            if not moments_window:
                return {"ok": False, "error": "moments_window_not_found"}

            items = self._collect_moments_items(moments_window)
            if not items:
                return {"ok": False, "error": "moments_item_not_found"}

            for index, item in enumerate(items):
                item_text = self._extract_item_text(item)
                
                if not item_text:
                    self._logger.debug(f"第 {index+1} 条无文本，跳过")
                    continue
                signature = utils.sha1_text(item_text)
                if signature in acted_signatures:
                    self._logger.debug(f"第 {index+1} 条已处理过(Signature)，跳过")
                    continue

                if keyword_filter and any(word in item_text for word in keyword_filter):
                    self._logger.debug(f"第 {index+1} 条命中关键词过滤，跳过")
                    continue

                author = self._extract_item_author(item_text) or "unknown"
                friend_count = int(per_friend.get(author) or 0)
                if friend_count >= per_friend_limit:
                    self._logger.debug(f"用户 {author} 评论数已达上限，跳过")
                    continue

                # Find the interaction button
                self._logger.debug(f"正在寻找第 {index+1} 条的互动按钮...")
                menu_btn = self._find_interaction_button(item)
                if menu_btn is None:
                    self._logger.debug(f"第 {index+1} 条未找到互动按钮，跳过")
                    continue

                self._logger.debug(f"点击互动按钮: {self._brief_control(menu_btn)}")
                if not self._click_control(menu_btn):
                    self._logger.warning("点击互动按钮失败")
                    continue
                
                # Wait for popup
                self._logger.debug("等待评论弹窗(SnsLikeToastWnd)...")
                popup = None
                for _ in range(5):
                    popup_pane = auto.PaneControl(ClassName="SnsLikeToastWnd")
                    if popup_pane.Exists(0, 0):
                        popup = popup_pane
                        break
                    popup_wnd = auto.WindowControl(ClassName="SnsLikeToastWnd")
                    if popup_wnd.Exists(0, 0):
                        popup = popup_wnd
                        break
                    time.sleep(0.3)
                
                if popup is None or not popup.Exists(1, 0.2):
                    self._logger.warning("未找到评论弹窗")
                    continue

                # Find "评论" button in popup
                self._logger.debug("在弹窗中查找[评论]按钮...")
                comment_btn = self._find_named_clickable(popup, ["评论"], search_depth=5)
                
                if comment_btn is None:
                    self._logger.debug("未找到[评论]按钮")
                    continue

                self._logger.info(f"找到[评论]按钮，执行点击...")
                if not self._click_control(comment_btn):
                    self._logger.warning("点击[评论]按钮失败")
                    continue

                # Wait for input box
                # Usually "EditControl" or "PaneControl" with "评论" name or similar
                # Or just wait a bit and type
                time.sleep(1.0) # Wait for input box to appear

                # Call API to generate comment
                try:
                    api_url = f"{backend_url.rstrip('/')}/api/user/marketing/comment/generate"
                    self._logger.info(f"请求后端生成评论... URL: {api_url}")
                    
                    # Ensure token has Bearer prefix if needed
                    auth_header = token
                    if token and not token.lower().startswith("bearer "):
                        auth_header = f"Bearer {token}"
                    
                    headers = {
                        "Authorization": auth_header, 
                        "Content-Type": "application/json"
                    }
                    if tenant_id:
                        headers["X-Tenant-Id"] = str(tenant_id)

                    resp = requests.post(
                        api_url,
                        json={"postContent": item_text, "userNickname": author},
                        headers=headers,
                        timeout=60
                    )
                    
                    if resp.status_code != 200:
                        self._logger.error(f"生成评论失败: {resp.status_code} {resp.text}")
                        # Cancel comment input by pressing ESC
                        auto.SendKeys('{Esc}')
                        continue
                    
                    res_json = resp.json()

                    if res_json.get("code") != 0:
                        self._logger.error(f"生成评论API错误: {res_json}")
                        auto.SendKeys('{Esc}')
                        continue
                        
                    comment_content = res_json.get("data")
                    if not comment_content:
                        self._logger.warning("生成的评论内容为空")
                        auto.SendKeys('{Esc}')
                        continue
                        
                    self._logger.info(f"生成的评论: {comment_content}")
                    
                    # Find input box and type
                    # Often EditControl in Moments window
                    edit_box = moments_window.EditControl(searchDepth=10)
                    if not edit_box.Exists(0, 0):
                        # Try searching globally or deeper
                        edit_box = auto.EditControl(searchFromControl=moments_window, searchDepth=15)
                    
                    if edit_box.Exists(0, 0):
                        edit_box.Click(simulateMove=False)
                        edit_box.SendKeys(comment_content)
                        time.sleep(0.5)
                        
                        # Find Send button? Or just Enter?
                        # Usually there is a "发送" button nearby
                        # Or Enter key if configured? WeChat usually requires Ctrl+Enter or just Enter
                        # Safe bet: Find "发送" button
                        send_btn = self._find_named_clickable(moments_window, ["发送", "Send"])
                        if send_btn and send_btn.Exists(0, 0):
                            send_btn.Click()
                        else:
                            # Try Enter
                            auto.SendKeys('{Enter}')
                    else:
                        self._logger.warning("未找到输入框，无法评论")
                        continue

                except Exception as e:
                    self._logger.error(f"评论生成或发送过程异常: {e}")
                    continue

                total_comments += 1
                state["totalComments"] = total_comments
                per_friend[author] = friend_count + 1
                acted_signatures.add(signature)
                self._logger.info(f"评论成功! Author: {author}, Total: {total_comments}")
                return {
                    "ok": True,
                    "success": True,
                    "commented": True,
                    "author": author,
                    "dailyTotalComments": total_comments,
                    "dailyTotalLimit": total_limit,
                    "friendComments": int(per_friend.get(author) or 0),
                    "friendLimit": per_friend_limit,
                    "commentIntervalStart": comment_start,
                    "commentIntervalEnd": comment_end
                }

            self._logger.info("遍历完所有可见条目，未进行评论")
            return {
                "ok": True,
                "success": True,
                "commented": False,
                "reason": "no_eligible_item",
                "dailyTotalComments": total_comments,
                "dailyTotalLimit": total_limit
            }


    def _get_wechat_pids(self) -> List[int]:
        pids = []
        for name in ["WeChat.exe", "Weixin.exe"]:
            try:
                cmd = f'tasklist /FI "IMAGENAME eq {name}" /FO CSV /NH'
                output = subprocess.check_output(cmd, shell=True).decode("gbk", errors="ignore")
                for line in output.splitlines():
                    if not line.strip():
                        continue
                    parts = line.split(',')
                    if len(parts) >= 2:
                        pid_str = parts[1].strip('"')
                        if pid_str.isdigit():
                            pids.append(int(pid_str))
            except Exception as e:
                self._logger.warning(f"获取进程PID失败 {name}: {e}")
        return list(set(pids))

    def _find_main_window(self) -> Any:
        if auto is None:
            raise RuntimeError("uiautomation 未加载")
        
        self._logger.info("Debug: _find_main_window start (Adaptation for 3.9.12)")
        
        if hasattr(auto, "SetGlobalSearchTimeout"):
            try:
                auto.SetGlobalSearchTimeout(0.5)
            except Exception:
                pass

        if hasattr(auto, "SetTransactionTimeout"):
            try:
                auto.SetTransactionTimeout(500)
            except Exception:
                pass
        
        # Strategy 0: Direct Win32 FindWindow
        # Using self._cfg.window_class_name which is now WeChatMainWndForPC by default
        try:
            hwnd = ctypes.windll.user32.FindWindowW(self._cfg.window_class_name, self._cfg.window_name)
            if hwnd and hwnd != 0:
                self._logger.info(f"FindWindowW found EXACT match handle: {hwnd}")
                if ctypes.windll.user32.IsHungAppWindow(hwnd):
                    self._logger.warning(f"Handle {hwnd} is HUNG. Skipping.")
                else:
                    window = auto.ControlFromHandle(hwnd)
                    if window.Exists(0, 0):
                        return window
        except Exception as e:
            self._logger.warning(f"Win32 FindWindow failed: {e}")

        # Strategy 1: Global UIA Search
        try:
            window = auto.WindowControl(
                searchDepth=1,
                ClassName=self._cfg.window_class_name,
                Name=self._cfg.window_name,
            )
            if window.Exists(0, 0):
                self._logger.info("Found window via Global Search")
                return window
        except Exception:
            pass
        
        # Strategy 2: Fallback to PID search
        pids = self._get_wechat_pids()
        for pid in pids:
            try:
                window = auto.WindowControl(
                    searchDepth=1,
                    ProcessId=pid,
                    ClassName=self._cfg.window_class_name,
                    Name=self._cfg.window_name,
                )
                if window.Exists(0, 0):
                    return window
                
                # Loose match (just Name)
                window_loose = auto.WindowControl(
                    searchDepth=1,
                    ProcessId=pid,
                    Name=self._cfg.window_name,
                )
                if window_loose.Exists(0, 0):
                    return window_loose
            except Exception:
                pass
                
        self._logger.warning("未找到微信主窗口")
        return None

    def get_main_window(self) -> Any:
        with self._uia_lock:
            if self._cached_main is not None:
                try:
                    if self._cached_main.Exists(0, 0):
                        return self._cached_main
                except Exception:
                    self._cached_main = None
            win = self._find_main_window()
            self._cached_main = win
            if win is not None:
                self._log_window_tree(win)
            return win

    def _log_window_tree(self, window: Any) -> None:
        handle = getattr(window, "NativeWindowHandle", None)
        if handle is None:
            handle = id(window)
        
        if self._tree_logged_handle == handle:
            return
        self._tree_logged_handle = handle
        try:
            children = window.GetChildren() or []
            self._logger.info("Window has %d children", len(children))
        except Exception:
            pass

    def log_ready_snapshot(self, window: Any) -> None:
        if window is None:
            return
        with self._uia_lock:
            name = getattr(window, "Name", "") or ""
            class_name = getattr(window, "ClassName", "") or ""
            pid = getattr(window, "ProcessId", None)
            handle = getattr(window, "NativeWindowHandle", None)
            rect_text = _rect_text(window)
            self._logger.info("主窗口就绪快照: Name=%s Class=%s PID=%s Handle=%s Rect=%s", name, class_name, pid, handle, rect_text)
            
            session_list = self._locate_session_list(window)
            if session_list is not None:
                s_name = _safe_attr(session_list, "Name")
                s_class = _safe_attr(session_list, "ClassName")
                self._logger.info("会话列表控件: Class=%s Name=%s", s_class, s_name)
            else:
                self._logger.info("未找到会话列表控件")

            message_list = self._locate_message_list(window)
            if message_list is not None:
                m_name = _safe_attr(message_list, "Name")
                m_class = _safe_attr(message_list, "ClassName")
                self._logger.info("消息列表控件: Class=%s Name=%s", m_class, m_name)
            else:
                self._logger.info("未找到消息列表控件")

    def _locate_session_list(self, window: Any) -> Optional[Any]:
        if auto is None:
            return None
            
        # 1. 优先使用缓存 (0毫秒级响应)
        if self._cached_session_list and self._cached_session_list.Exists(0, 0):
            return self._cached_session_list
            
        # 2. Strategy 1: Standard Name="会话" (For 3.9.12) - 原生底层 C++ 搜索，极快
        try:
            target = auto.ListControl(searchFromControl=window, searchDepth=12, Name="会话")
            if target and target.Exists(0, 0):
                self._logger.debug("_locate_session_list Found Name='会话'")
                self._cached_session_list = target
                return target
        except Exception:
            pass

        # 3. Strategy 2: English Name="Session"
        try:
            target = auto.ListControl(searchFromControl=window, searchDepth=12, Name="Session")
            if target and target.Exists(0, 0):
                self._logger.debug("_locate_session_list Found Name='Session'")
                self._cached_session_list = target
                return target
        except Exception:
            pass
            
        self._logger.warning("未找到微信会话列表(Name='会话')，请确认微信版本。")
        return None

    def _locate_message_list(self, window: Any) -> Optional[Any]:
        if auto is None:
            return None
            
        # 1. 优先使用缓存，并验证其是否真的可用
        if self._cached_message_list:
            try:
                if self._cached_message_list.Exists(0, 0):
                    # 简单测试一下是否能获取到边界，如果抛出异常说明底层句柄已失效
                    _ = self._cached_message_list.BoundingRectangle
                    return self._cached_message_list
            except Exception:
                pass
            self._cached_message_list = None # 失效则清理
            
        # 2. 快速底层搜索 Name="消息"
        try:
            target = auto.ListControl(searchFromControl=window, searchDepth=15, Name="消息")
            if target and target.Exists(0, 0):
                self._cached_message_list = target
                return target
        except Exception:
            pass

        # 3. 英文系统兼容
        try:
            target = auto.ListControl(searchFromControl=window, searchDepth=15, Name="Message")
            if target and target.Exists(0, 0):
                self._cached_message_list = target
                return target
        except Exception:
            pass

        return None

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
        item_name = getattr(item, "Name", "") or ""
        
        if item_name in ("服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号", "文件传输助手"):
            return False
            
        if re.search(r"(\[\d+条\]|\d+条新消息|未读)", item_name):
            return True

        # 彻底废弃极其卡顿的 _iter_descendants
        # 微信的小红点通常在这个会话控件的最外层子节点中
        try:
            children = item.GetChildren() or []
            for ctrl in children:
                # 检查第一层和第二层子节点即可
                sub_children = ctrl.GetChildren() or []
                for sub_ctrl in sub_children:
                    name = (getattr(sub_ctrl, "Name", "") or "").strip()
                    if name and (re.fullmatch(r"\d+", name) or name == "99+"):
                        return True
                    if any(k in name for k in ("未读", "条新消息", "new message", "badge", "reddot")):
                        return True
        except Exception:
            pass
                
        return False

    def _check_and_exit_subscription_folder(self, window: Any) -> bool:
        if auto is None or window is None:
            return False
            
        targets = ["服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号"]
        for name in targets:
            try:
                btn = auto.ButtonControl(searchFromControl=window, searchDepth=12, Name=name)
                if btn.Exists(0, 0):
                    rect = getattr(btn, "BoundingRectangle", None)
                    bbox = utils.rect_to_bbox(rect) if rect else None
                    if bbox:
                        l, t, r, b = bbox
                        if l < 350 and t < 200:
                            btn.Click(simulateMove=True)
                            time.sleep(1.0)
                            return True
            except Exception:
                pass
        return False

    def click_session_item(self, item: Any) -> Optional[str]:
        if auto is None:
            return None
        with self._uia_lock:
            name = self._normalize_contact_name((getattr(item, "Name", "") or ""))
            
            if name in ("服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号", "文件传输助手"):
                self._logger.info(f"跳过聚合类/特殊会话: {name}")
                return None
                
            try:
                # Use _click_control for robust clicking
                if self._click_control(item):
                    time.sleep(random.uniform(self._cfg.click_move_min_seconds, self._cfg.click_move_max_seconds))
                else:
                    # Fallback
                    item.Click(simulateMove=True)
                
                time.sleep(random.uniform(0.2, 0.5))
            except Exception:
                pass
            
            window = self.get_main_window()
            if self._check_and_exit_subscription_folder(window):
                return None

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

            target_name = self._normalize_contact_name(target)
            if target_name in ("服务号", "订阅号", "Subscription Accounts", "订阅号消息", "公众号", "文件传输助手"):
                self._logger.info(f"拒绝向聚合类/特殊会话发送消息: {target_name}")
                return False
                
            current = self.get_current_chat_title(window)
            current_name = self._normalize_contact_name(current or "")
            
            if current_name and target_name and target_name in current_name:
                return True

            session_list = self._locate_session_list(window)
            if session_list is None:
                return False

            best = None
            children = session_list.GetChildren() or []
            for item in children:
                raw_name = getattr(item, "Name", "") or ""
                name_lines = raw_name.split('\n')
                name = name_lines[0].strip() if name_lines else ""
                name = self._normalize_contact_name(name)
                
                if not name:
                    continue
                
                if target_name == name or (target_name and target_name in name):
                    best = item
                    break
            
            if best is None:
                return False
                
            self.click_session_item(best)
            time.sleep(random.uniform(0.5, 1.0))
            
            current = self.get_current_chat_title(window)
            current_name = self._normalize_contact_name(current or "")
            return bool(current_name and target_name and target_name in current_name)

    def get_current_chat_title(self, window: Any) -> Optional[str]:
        if auto is None or window is None:
            return None

        # 1. 缓存拦截
        if self._cached_chat_title_ctrl:
            try:
                if self._cached_chat_title_ctrl.Exists(0, 0):
                    name = (getattr(self._cached_chat_title_ctrl, "Name", "") or "").strip()
                    if name and name not in ("微信", "文件传输助手", "聊天信息"):
                        return name
            except Exception:
                pass
            self._cached_chat_title_ctrl = None

        try:
            msg_list = self._locate_message_list(window)
            msg_bbox = utils.rect_to_bbox(getattr(msg_list, "BoundingRectangle", None)) if msg_list else None

            info_btn = auto.ButtonControl(searchFromControl=window, searchDepth=20, Name="聊天信息")
            if info_btn and info_btn.Exists(0, 0):
                info_bbox = utils.rect_to_bbox(getattr(info_btn, "BoundingRectangle", None))
                if info_bbox:
                    info_left, info_top, _, info_bottom = info_bbox
                    anchor_top = info_top - 18
                    anchor_bottom = info_bottom + 18
                    anchor_left_limit = (msg_bbox[0] - 40) if msg_bbox else 500
                    anchor_right_limit = info_left + 6

                    anchor_candidates = []
                    for ctrl in _iter_descendants(window, max_depth=18):
                        try:
                            if _control_type_name(ctrl) != "TextControl":
                                continue
                            name = (getattr(ctrl, "Name", "") or "").strip()
                            if not name:
                                continue
                            if name in ("微信", "聊天信息", "文件传输助手"):
                                continue
                            if re.search(r"20\d{2}年\d{1,2}月\d{1,2}日", name):
                                continue
                            if re.match(r"^([上下]午)?\s*\d{1,2}:\d{2}(?::\d{2})?$", name):
                                continue
                            if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", name):
                                continue
                            if name.isdigit():
                                continue
                            cbox = utils.rect_to_bbox(getattr(ctrl, "BoundingRectangle", None))
                            if not cbox:
                                continue
                            cl, ct, cr, cb = cbox
                            if ct < 0 or cb < 0:
                                continue
                            if cl < anchor_left_limit or cr > anchor_right_limit:
                                continue
                            if cb < anchor_top or ct > anchor_bottom:
                                continue
                            anchor_candidates.append((abs(ct - info_top), abs(cr - info_left), ct, cl, name, ctrl))
                        except Exception:
                            pass
                    if anchor_candidates:
                        anchor_candidates.sort(key=lambda x: (x[0], x[1], x[2], x[3]))
                        best_name = anchor_candidates[0][4]
                        best_ctrl = anchor_candidates[0][5]
                        self._cached_chat_title_ctrl = best_ctrl
                        self._logger.info(f"成功通过聊天信息锚点锁定当前聊天窗口标题: {best_name}")
                        return best_name

            if msg_bbox:
                msg_left, msg_top, _, _ = msg_bbox
                header_top = msg_top - 110
                header_bottom = msg_top + 22
                candidates = []
                for ctrl in _iter_descendants(window, max_depth=18):
                    try:
                        if _control_type_name(ctrl) != "TextControl":
                            continue
                        name = (getattr(ctrl, "Name", "") or "").strip()
                        if not name:
                            continue
                        if name in ("微信", "聊天信息", "文件传输助手"):
                            continue
                        if re.search(r"20\d{2}年\d{1,2}月\d{1,2}日", name):
                            continue
                        if re.match(r"^([上下]午)?\s*\d{1,2}:\d{2}(?::\d{2})?$", name):
                            continue
                        if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", name):
                            continue
                        if name.isdigit():
                            continue
                        cbox = utils.rect_to_bbox(getattr(ctrl, "BoundingRectangle", None))
                        if not cbox:
                            continue
                        cl, ct, cr, cb = cbox
                        if ct < 0 or cb < 0:
                            continue
                        if cl < msg_left - 40:
                            continue
                        if ct < header_top or cb > header_bottom:
                            continue
                        candidates.append((abs(ct - (msg_top - 44)), cl, name, ctrl))
                    except Exception:
                        pass
                if candidates:
                    candidates.sort(key=lambda x: (x[0], x[1]))
                    best_name = candidates[0][2]
                    best_ctrl = candidates[0][3]
                    self._cached_chat_title_ctrl = best_ctrl
                    self._logger.info(f"成功通过标题区域锁定当前聊天窗口标题: {best_name}")
                    return best_name
        except Exception as e:
            self._logger.warning(f"获取当前聊天标题异常: {e}")

        return None

    def _get_content_root(self, window: Any) -> Any:
        root = window
        try:
            for child in window.GetChildren() or []:
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
        
        # 1. 优先使用缓存
        if self._cached_input_box and self._cached_input_box.Exists(0, 0):
            return self._cached_input_box
            
        # 2. 策略：根据当前聊天标题查找 (EditControl Name通常等于聊天标题)
        current_title = self.get_current_chat_title(window)
        if current_title:
            try:
                # 限制搜索范围在右侧面板，或者全局搜索但检查位置
                # 增加深度以确保能找到
                edit = auto.EditControl(searchFromControl=window, searchDepth=20, Name=current_title)
                if edit and edit.Exists(0, 0):
                    # 校验位置：必须在右侧 (Left > 300) 且有一定宽度
                    rect = getattr(edit, "BoundingRectangle", None)
                    bbox = utils.rect_to_bbox(rect)
                    if bbox:
                        w = bbox[2] - bbox[0]
                        if bbox[0] > 300 and w > 300:
                            self._cached_input_box = edit
                            return edit
            except Exception:
                pass

        # 3. 策略：从右侧主面板向下查找 (最可靠)
        msg_list = self._locate_message_list(window)
        if msg_list:
            try:
                # msg_list -> chat_body -> right_main
                # 根据控件树结构，需要往上找几层
                p1 = msg_list.GetParentControl()
                if p1:
                    p2 = p1.GetParentControl() # right_main or container
                    if p2:
                        # 在右侧大容器里找 EditControl
                        # 通常输入框在底部，且尺寸较大
                        edits = p2.GetChildren() # 这里不能直接GetChildren，需要DeepSearch
                        # 使用 auto.EditControl 搜索
                        # 深度设为 10 应该足够
                        edit = auto.EditControl(searchFromControl=p2, searchDepth=12)
                        
                        # 可能找到搜索框(Name='搜索')，需要排除
                        # 搜索框通常在顶部，输入框在底部
                        # 如果找到多个，需要遍历筛选
                        
                        candidates = []
                        # 手动遍历 p2 下的 EditControl
                        for ctrl in _iter_descendants(p2, max_depth=12):
                            if _control_type_name(ctrl) == "EditControl":
                                rect = getattr(ctrl, "BoundingRectangle", None)
                                bbox = utils.rect_to_bbox(rect)
                                if bbox:
                                    w = bbox[2] - bbox[0]
                                    h = bbox[3] - bbox[1]
                                    t = bbox[1]
                                    # 排除搜索框 (通常高度较小 < 30 或 宽度较小 < 200，且位置靠上)
                                    if w > 300 and h > 40:
                                        candidates.append((t, ctrl))
                        
                        if candidates:
                            # 取 Top 最大的 (最下面的)
                            candidates.sort(key=lambda x: x[0], reverse=True)
                            best = candidates[0][1]
                            self._cached_input_box = best
                            return best
            except Exception:
                pass

        # 4. 策略：启发式查找 (全局遍历，位置+尺寸)
        try:
            candidates = []
            for ctrl in _iter_descendants(window, max_depth=20):
                if _control_type_name(ctrl) == "EditControl":
                    rect = getattr(ctrl, "BoundingRectangle", None)
                    bbox = utils.rect_to_bbox(rect)
                    if bbox:
                        l, t, r, b = bbox
                        w = r - l
                        h = b - t
                        # 输入框特征：右侧，宽大
                        if l > 350 and w > 300 and h > 40:
                            candidates.append((t, ctrl))
            
            if candidates:
                candidates.sort(key=lambda x: x[0], reverse=True)
                best = candidates[0][1]
                self._cached_input_box = best
                return best
        except Exception:
            pass
            
        # 5. 旧策略：查找 Name="输入" (兜底)
        try:
            edit = auto.EditControl(searchFromControl=window, searchDepth=15, Name="输入")
            if edit and edit.Exists(0, 0):
                self._cached_input_box = edit
                return edit
        except Exception:
            pass

        return None

    def find_send_button(self, window: Any) -> Optional[Any]:
        if auto is None:
            return None
            
        # 0. 策略：基于输入框的相对位置查找 (最稳)
        if self._cached_input_box:
            try:
                # 发送按钮通常是输入框的兄弟或叔叔节点，且位置在输入框下方
                p1 = self._cached_input_box.GetParentControl()
                if p1:
                    # 尝试在父节点找
                    btn = auto.ButtonControl(searchFromControl=p1, searchDepth=5, Name="发送(S)")
                    if btn.Exists(0, 0): return btn
                    
                    # 尝试在爷爷节点找
                    p2 = p1.GetParentControl()
                    if p2:
                        btn = auto.ButtonControl(searchFromControl=p2, searchDepth=8, Name="发送(S)")
                        if btn.Exists(0, 0): return btn
            except Exception:
                pass

        roots = []
        content_root = self._get_content_root(window)
        roots.append(content_root)
        if content_root != window:
            roots.append(window)

        try:
            # Strategy 1: Name="发送(S)" or "发送"
            for root in roots:
                for name in ("发送(S)", "发送"):
                    btn = auto.ButtonControl(searchFromControl=root, searchDepth=25, Name=name)
                    if btn and btn.Exists(0, 0):
                        return btn

            # Strategy 2: Fallback (Lowest button with '发送' in name)
            for root in roots:
                candidates = []
                for ctrl in _iter_descendants(root, max_depth=25):
                    if _control_type_name(ctrl) != "ButtonControl":
                        continue
                    text = (getattr(ctrl, "Name", "") or "")
                    if "发送" not in text:
                        continue
                    rect = getattr(ctrl, "BoundingRectangle", None)
                    b = utils.rect_to_bbox(rect)
                    if b is None:
                        continue
                    candidates.append((b[3], ctrl))
                if candidates:
                    candidates.sort(key=lambda x: x[0])
                    return candidates[-1][1]
        except Exception:
            pass
        return None

    def _analyze_item_alignment(self, item: Any, list_bbox: Tuple[int, int, int, int]) -> str:
        if item is None or list_bbox is None:
            return "other"
            
        try:
            # 列表中心线
            list_center_x = (list_bbox[0] + list_bbox[2]) / 2.0
            
            # 遍历 ListItem 的子孙节点，寻找头像 Button
            for ctrl in _iter_descendants(item, max_depth=4):
                if _control_type_name(ctrl) == "ButtonControl":
                    rect = getattr(ctrl, "BoundingRectangle", None)
                    if rect:
                        width = rect.right - rect.left
                        height = rect.bottom - rect.top
                        # 头像通常是正方形按钮，尺寸在 30 到 50 之间
                        if 30 <= width <= 50 and 30 <= height <= 50:
                            btn_center_x = (rect.left + rect.right) / 2.0
                            if btn_center_x > list_center_x:
                                return "self"
                            else:
                                return "other"
                                
            # 兜底：如果没找到头像，根据整个 item 的文本重心判断
            rect = getattr(item, "BoundingRectangle", None)
            if rect:
                if ((rect.left + rect.right) / 2.0) > list_center_x:
                    return "self"
        except Exception:
            pass
            
        return "other"

    def extract_latest_messages(self, contact_hint: str) -> List[Dict[str, Any]]:
        if auto is None: return []
        with self._uia_lock:
            window = self.get_main_window()
            if window is None: return []
            
            normalized_contact = self._normalize_contact_name(contact_hint)
            if not normalized_contact:
                current_title = self.get_current_chat_title(window)
                if not current_title:
                    return [] # 拿不到标题直接退出
                normalized_contact = self._normalize_contact_name(current_title)
                
            msg_list = self._locate_message_list(window)
            if msg_list is None: return []
            
            list_bbox = utils.rect_to_bbox(getattr(msg_list, "BoundingRectangle", None))
            if not list_bbox: return []

            items = msg_list.GetChildren() or []
            if not items: return []

            collected_messages = []
            scan_limit = max(5, int(self._cfg.message_scan_limit))
            scan_items = items[-scan_limit:]
            self._logger.info("开始提取最新消息: 会话=%s, 总节点=%d, 扫描节点=%d", normalized_contact, len(items), len(scan_items))

            for item in scan_items:
                msg = self._extract_message_from_item(normalized_contact or contact_hint, item)
                if not msg: continue 

                direction = self._analyze_item_alignment(item, list_bbox)
                msg['is_self'] = (direction == "self")
                collected_messages.append(msg)
            
            if collected_messages:
                self._logger.info("提取消息完成: 会话=%s, 有效消息=%d", normalized_contact, len(collected_messages))
                last_msg = collected_messages[-1]
                if not last_msg.get('is_self', False):
                    last_msg['trigger_reply'] = True
            else:
                sample_names = [((getattr(node, "Name", "") or "").strip())[:40] for node in scan_items]
                self._logger.warning("提取消息为空: 会话=%s, 最近节点样本=%s", normalized_contact, sample_names)
            
            return collected_messages

    def _extract_message_from_item(self, contact_hint: str, item: Any) -> Optional[Dict[str, Any]]:
        try:
            ui_id = item.GetRuntimeId()
        except:
            ui_id = None

        raw_name = (getattr(item, "Name", "") or "").strip()
        if not raw_name:
            return None
        
        # 1. 排除时间戳
        if self._is_time_separator_text(raw_name):
            return None
        
        # 2. 排除系统提示
        if self._is_new_message_divider(raw_name):
            self._logger.info("消息过滤-新消息分隔提示: %s", raw_name[:80])
            return None
        if raw_name in ("查看更多消息", "如果你要查看更多消息"):
            self._logger.info("消息过滤-历史消息入口: %s", raw_name)
            return None
        if ("你已添加了" in raw_name and "现在可以开始聊天了" in raw_name) or "以上是打招呼的内容" in raw_name:
            self._logger.info("消息过滤-打招呼系统提示: %s", raw_name[:80])
            return None
        
        # 排除撤回消息 (系统通知，非用户发言)
        if "撤回了一条消息" in raw_name:
            # 进一步确认是否为系统消息（系统消息没有头像按钮）
            is_system = True
            if auto:
                try:
                    # 正常消息ListItem下会有ButtonControl(头像)
                    # 结构通常为: ListItem -> Pane -> Button
                    btn = auto.ButtonControl(searchFromControl=item, searchDepth=4)
                    if btn.Exists(0, 0):
                        is_system = False
                except Exception:
                    pass
            
            if is_system:
                self._logger.info("消息过滤-撤回系统提示: %s", raw_name[:80])
                return None

        text_content = raw_name
        
        # 3. 极速提取富文本卡片 (视频号/小程序/链接等)
        # 不再使用 Python 遍历，改用 UIA 底层 C++ 搜索寻找附带文字
        if raw_name.startswith("[") and raw_name.endswith("]"):
            try:
                # 限制深度为 8，瞬间找出卡片里带的文本（比如：AI-魔方视界）
                text_ctrl = auto.TextControl(searchFromControl=item, searchDepth=8)
                if text_ctrl and text_ctrl.Exists(0, 0):
                    found_text = (text_ctrl.Name or "").strip()
                    if found_text and found_text != raw_name:
                        text_content = f"{raw_name} {found_text}"
            except Exception:
                pass

        if text_content:
             return {"contact": contact_hint, "type": "text", "content": text_content, "timestamp": utils.now_iso(), "ui_id": ui_id}
        
        return None

    def _is_new_message_divider(self, text: str) -> bool:
        normalized = re.sub(r"\s+", "", text or "")
        if not normalized:
            return False
        normalized = normalized.strip("：:。.!！?？-—_~～·•|｜")
        return normalized in ("以下是新消息", "以下为新消息", "以下是最新消息", "以下为最新消息")

    def _is_time_separator_text(self, text: str) -> bool:
        normalized = re.sub(r"\s+", "", text or "")
        if not normalized:
            return False
        if re.fullmatch(r"\d{1,2}:\d{2}(:\d{2})?", normalized):
            return True
        if re.fullmatch(r"(今天|昨天)(\d{1,2}:\d{2}(:\d{2})?)?", normalized):
            return True
        if re.fullmatch(r"星期[一二三四五六日天]([上下]午)?\d{1,2}:\d{2}(:\d{2})?", normalized):
            return True
        if re.fullmatch(r"20\d{2}年\d{1,2}月\d{1,2}日([上下]午)?\d{1,2}:\d{2}(:\d{2})?", normalized):
            return True
        return False

    def set_text_and_send(self, target: str, text: str) -> bool:
        if auto is None:
            return False
        
        time.sleep(random.uniform(0.2, 0.6))

        with self._uia_lock:
            window = self.get_main_window()
            if window is None:
                return False
            
            try:
                window.SetActive()
            except:
                pass

            if not self.ensure_chat_target(target):
                self._logger.warning(f"无法切换到目标会话: {target}")
                return False

            edit = self.find_input_box(window)
            if edit is None:
                self._logger.warning("未找到输入框，无法发送消息")
                return False

            ok = self._set_edit_value(edit, text)
            if not ok:
                self._logger.warning("无法在输入框中粘贴文本")
                return False

            time.sleep(random.uniform(0.3, 0.8))

            send_btn = self.find_send_button(window)
            if send_btn is not None:
                try:
                    send_btn.Click(simulateMove=True)
                    return True
                except Exception:
                    pass

            # Fallback 1: Alt+S (Common shortcut for Send)
            try:
                auto.SendKeys("{Alt}s")
                return True
            except:
                pass

            # Fallback 2: Enter
            try:
                auto.SendKeys("{Enter}")
                return True
            except Exception:
                return False

    def _set_edit_value(self, edit: Any, text: str) -> bool:
        if pyperclip is None:
            self._logger.error("缺少 pyperclip 依赖")
            return False

        try:
            # Use _click_control for more robust clicking (center point)
            self._click_control(edit)
            time.sleep(0.2)
            edit.SendKeys("{Ctrl}a", waitTime=0.1) 
            edit.SendKeys("{Delete}", waitTime=0.1)
            pyperclip.copy(text)
            edit.SendKeys("{Ctrl}v", waitTime=0.2)
            time.sleep(0.1)
            return True
        except Exception as e:
            self._logger.error(f"输入文本异常: {e}")
            return False
