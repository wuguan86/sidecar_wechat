import argparse
import os
import sys
import time
import logging

from .config import load_config
from .logger import setup_logging
from .ui import WeChatUI, inspect_window_tree
from .network import JavaClient, Reporter, Poller, CommandServer
from .listener import Listener
from . import utils

try:
    import uiautomation as auto
except ImportError:
    auto = None

try:
    import pythoncom
except ImportError:
    pythoncom = None


def _ensure_deps() -> None:
    missing = []
    if auto is None:
        missing.append("uiautomation")
    try:
        from PIL import ImageGrab
    except ImportError:
        missing.append("Pillow(ImageGrab)")
    if missing:
        raise RuntimeError("缺少依赖: " + ", ".join(missing))


def _ensure_com_initialized(logger: logging.Logger) -> None:
    if pythoncom is None:
        logger.warning("pythoncom 未安装，无法显式初始化 COM")
        return
    try:
        pythoncom.CoInitialize()
    except Exception as e:
        logger.warning("COM 初始化失败: %s", e)


def _warmup_uia(logger: logging.Logger) -> None:
    if auto is None:
        return
    
    try:
        SPI_SETSCREENREADER = 0x0047
        SPIF_SENDCHANGE = 0x0002
        SPIF_UPDATEINIFILE = 0x0001
        import ctypes
        ctypes.windll.user32.SystemParametersInfoW(SPI_SETSCREENREADER, 1, None, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE)
        logger.info("已尝试设置 SPI_SETSCREENREADER 以激活无障碍支持")
    except Exception as e:
        logger.warning("设置 SPI_SETSCREENREADER 失败: %s", e)

    try:
        auto.GetRootControl()
    except Exception as e:
        logger.warning("UIA 核心组件预加载失败: %s", e)


def dry_run(cfg, logger) -> int:
    _ensure_deps()
    _ensure_com_initialized(logger)
    _warmup_uia(logger)
    ui_inst = WeChatUI(cfg, logger)
    win = ui_inst.get_main_window()
    if win is None:
        logger.error("未找到微信主窗口: ClassName=%s Name=%s", cfg.window_class_name, cfg.window_name)
        return 2
    try:
        title = getattr(win, "Name", "") or ""
        rect = getattr(win, "BoundingRectangle", None)
        logger.info("找到微信主窗口: title=%s rect=%s", title, rect)
    except Exception:
        logger.info("找到微信主窗口")

    try:
        unread = ui_inst.find_unread_sessions()
        logger.info("未读会话数量(本轮): %s", len(unread))
    except Exception as e:
        logger.warning("未读扫描失败: %s", e)
    return 0


def inspect_run(cfg, logger) -> int:
    _ensure_deps()
    _ensure_com_initialized(logger)
    _warmup_uia(logger)
    ui_inst = WeChatUI(cfg, logger)
    win = ui_inst.get_main_window()
    if win is None:
        logger.error("未找到微信主窗口: ClassName=%s Name=%s", cfg.window_class_name, cfg.window_name)
        return 2
    logger.info("开始深度遍历控件树")
    inspect_window_tree(win, logger)
    return 0


def self_test(cfg, logger) -> int:
    if auto is None:
        logger.info("Self-test 跳过 UIA：uiautomation 未加载")
        return 0
    else:
        _ensure_com_initialized(logger)
        _warmup_uia(logger)
        ui_inst = WeChatUI(cfg, logger)
        
        win = ui_inst.get_main_window()
        if win:
            class_name = getattr(win, "ClassName", "")
            if class_name != cfg.window_class_name:
                logger.warning(f"Self-test 检测到 '黑盒' 微信 (ClassName={class_name})，正式运行可能会有问题")
                
        poller = Poller(ui_inst, logger)

    server = CommandServer(cfg.server_host, cfg.server_port, ui_inst, poller, logger)
    server.start()
    logger.info("Self-test 运行中(3s)... 可请求 /health /poll /command")
    time.sleep(3)
    server.stop()
    logger.info("Self-test 完成")
    return 0


def main() -> int:
    default_config = "config.yaml"
    if getattr(sys, "frozen", False):
        default_config = os.path.join(os.path.dirname(sys.executable), "config.yaml")
    else:
        # Assuming run from parent directory or similar, but let's be safe
        # If this file is in sidecar_wechat/core/main.py, config is in sidecar_wechat/config.yaml
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        default_config = os.path.join(base_dir, "config.yaml")

    parser = argparse.ArgumentParser(prog="wechat_bridge", add_help=True)
    parser.add_argument("--config", default=default_config)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--inspect", action="store_true")
    args = parser.parse_args()

    cfg = load_config(args.config)
    logger_inst = setup_logging(cfg)

    if args.dry_run:
        return dry_run(cfg, logger_inst)
    if args.inspect:
        return inspect_run(cfg, logger_inst)
    if args.self_test:
        return self_test(cfg, logger_inst)

    _ensure_deps()
    _ensure_com_initialized(logger_inst)
    _warmup_uia(logger_inst)
    if auto is not None and hasattr(auto, "SetGlobalSearchTimeout"):
        try:
            auto.SetGlobalSearchTimeout(1.0)
        except Exception:
            pass

    ui_inst = WeChatUI(cfg, logger_inst)

    def wait_for_wechat_window() -> None:
        last_tip_time = 0.0
        last_state_time = 0.0
        last_flag_time = 0.0
        while True:
            if time.time() - last_flag_time > 20.0:
                _warmup_uia(logger_inst)
                last_flag_time = time.time()
            win = ui_inst.get_main_window()
            if win is not None:
                class_name = getattr(win, "ClassName", "")
                if class_name == cfg.window_class_name:
                    try:
                        children = win.GetChildren() or []
                    except Exception:
                        children = []
                    if len(children) > 0:
                        logger_inst.info("微信主窗口已就绪: ClassName=%s", class_name)
                        try:
                            ui_inst.log_ready_snapshot(win)
                        except Exception:
                            pass
                        return
                    if time.time() - last_state_time > 4.0:
                        logger_inst.info("检测到微信窗口，但控件树未就绪，等待中...")
                        last_state_time = time.time()
                else:
                    if time.time() - last_state_time > 4.0:
                        logger_inst.warning("检测到微信窗口类名异常: %s (预期 %s)，等待你手动重启微信", class_name, cfg.window_class_name)
                        last_state_time = time.time()
                    ui_inst._cached_main = None
                    ui_inst._tree_logged_handle = None
            else:
                if time.time() - last_tip_time > 6.0:
                    logger_inst.info("未检测到微信窗口，请手动启动并登录微信")
                    last_tip_time = time.time()
            time.sleep(1.5)
    
    try:
        logger_inst.info("检查微信环境状态...")
        win = ui_inst.get_main_window()
        if win:
            class_name = getattr(win, "ClassName", "")
            if class_name != cfg.window_class_name:
                logger_inst.warning(f"检测到黑盒/异常微信 (ClassName={class_name} != {cfg.window_class_name})，请手动关闭微信后再启动。")
                ui_inst._cached_main = None
                ui_inst._tree_logged_handle = None
            else:
                logger_inst.info(f"检测到正常的微信窗口 (ClassName={class_name})")
    except Exception as e:
        logger_inst.warning(f"环境检查异常: {e}")

    wait_for_wechat_window()
    
    java_client = JavaClient(cfg, logger_inst)
    reporter = Reporter(java_client, logger_inst)
    reporter.start()
    
    poller = Poller(ui_inst, logger_inst)
    listener_inst = Listener(cfg, ui_inst, reporter, logger_inst, poller)
    
    server = CommandServer(cfg.server_host, cfg.server_port, ui_inst, poller, logger_inst)
    server.start()
    logger_inst.info("wechat_bridge 已启动 (监听模式) - 适配版本 3.9.12")
    try:
        while True:
            listener_inst.process_cycle()
            utils.sleep_with_jitter(cfg.scan_interval_seconds, cfg.scan_jitter_seconds)
    except KeyboardInterrupt:
        logger_inst.info("收到退出信号，正在停止...")
    finally:
        reporter.stop()
        server.stop()
    return 0

if __name__ == "__main__":
    sys.exit(main())
