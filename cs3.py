import time
import subprocess
import uiautomation as auto
import ctypes
import os
import shutil
import sys

# ==========================================
# 1. 环境配置与伪装工具 (静默无感)
# ==========================================

def set_screen_reader_flag(status: bool):
    """静默修改系统底层标志位"""
    ctypes.windll.user32.SystemParametersInfoW(0x0047, status, None, 1 | 2)

def create_fake_narrator():
    """创建一个伪装的无障碍进程文件 (nvda.exe)"""
    temp_dir = os.environ.get('TEMP')
    fake_nvda = os.path.join(temp_dir, "nvda.exe")
    if not os.path.exists(fake_nvda):
        shutil.copy(sys.executable, fake_nvda) 
    return fake_nvda

def start_fake_process(path):
    """在后台静默启动伪装进程"""
    print("[1] 正在启动静默伪装服务 (nvda.exe)...")
    return subprocess.Popen([path, "-c", "import time; time.sleep(60)"], 
                            creationflags=subprocess.CREATE_NO_WINDOW)

def kill_wechat():
    print("[0] 正在强杀残留微信进程以刷新环境...")
    subprocess.run(['taskkill', '/F', '/IM', 'WeChat.exe'], capture_output=True)
    time.sleep(1.5)

# ==========================================
# 2. 深度树探测函数
# ==========================================

def dump_detailed_tree(control, depth=0, max_depth=6):
    """
    递归打印 UI 树详细结构
    """
    indent = "  " * depth
    try:
        # 提取关键属性
        name = control.Name if control.Name else "<空>"
        class_name = control.ClassName if control.ClassName else "<空>"
        ctrl_type = control.ControlTypeName
        
        # 打印当前节点
        print(f"{indent}【{ctrl_type}】 Name: {name} | Class: {class_name}")
        
        # 如果未达到最大深度，继续向下探测
        if depth < max_depth:
            for child in control.GetChildren():
                dump_detailed_tree(child, depth + 1, max_depth)
    except Exception:
        pass

# ==========================================
# 3. 核心执行流程
# ==========================================

def main():
    print("="*50)
    print(">>> 视界 AI - 终极静默爆破 & 深度探测器 <<<")
    print("="*50)
    
    # 清场并设置环境
    kill_wechat()
    fake_exe = create_fake_narrator()
    set_screen_reader_flag(True)
    p_fake = start_fake_process(fake_exe)
    
    try:
        print("\n💡 [操作指南]: 请现在启动微信并登录，进入聊天界面。")
        
        # 等待窗口变异
        target_win = None
        wechat_win = auto.WindowControl(searchDepth=1, Name='微信')
        
        print("⏳ 正在监控窗口状态...")
        for i in range(40):
            if wechat_win.Exists(0, 0):
                if wechat_win.ClassName == 'mmui::MainWindow':
                    if len(wechat_win.GetChildren()) > 0:
                        target_win = wechat_win
                        break
                else:
                    print(f"  > 当前状态: {wechat_win.ClassName} (等待变异...)")
            time.sleep(2)
            
        if target_win:
            print("\n" + "*"*50)
            print("🎉 爆破成功！mmui::MainWindow 已解锁，开始生成深度结构图...")
            print("*"*50 + "\n")
            
            # --- 核心：深度打印 UI 树 ---
            # 建议 max_depth 设为 6-8，Qt 微信的树非常深
            dump_detailed_tree(target_win, depth=0, max_depth=8)
            
            print("\n" + "="*50)
            print("✅ 探测完毕！上方打印的信息即为微信内部所有控件的‘地图’。")
            print(">> 请根据 Name 和 Class 编写你的自动回复逻辑。")
            print("="*50)
            
            # 任务完成，关闭假进程
            p_fake.terminate()
            
            print("\n>> 脚本保持运行中 (维持 mmui 状态)，按 Ctrl+C 退出。")
            while True: time.sleep(1)
        else:
            print("\n❌ 爆破超时。请确保微信已完全退出并重新登录。")

    except KeyboardInterrupt:
        print("\n[!] 用户中止。")
    finally:
        set_screen_reader_flag(False)
        try: p_fake.terminate()
        except: pass
        print("[*] 环境已还原，安全退出。")

if __name__ == '__main__':
    # 设置全局搜索速度
    auto.SetGlobalSearchTimeout(5)
    main()