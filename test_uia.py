import uiautomation as auto
import time

def fast_diagnose():
    print("--- 正在进入极速定点诊断模式 (V2) ---")
    print("【操作指引】")
    print("1. 脚本倒计时开始后，请立即切换到微信。")
    print("2. 在微信聊天窗口内点击一下，使其获得焦点。")
    print("--------------------------------")
    
    for i in range(3, 0, -1):
        print(f"请在 {i} 秒内点击微信窗口...")
        time.sleep(1)

    try:
        # 1. 先获取当前前台窗口的句柄 (HWND)
        hwnd = auto.GetForegroundWindow()
        print(f"捕获到前台窗口句柄: {hwnd}")

        # 2. 将句柄转换为 uiautomation 的 Control 对象
        active_win = auto.ControlFromHandle(hwnd)
        
        if not active_win:
            print("错误：无法根据句柄创建控件对象。")
            return

        print(f"\n成功识别前台窗口：")
        print(f"  - 名称 (Name): {active_win.Name}")
        print(f"  - 类名 (ClassName): {active_win.ClassName}")
        
        # 3. 核心判断
        if active_win.ClassName == "WeChatMainWndForPC":
            print("\n!!! 成功：已确认该窗口为微信主窗口 !!!")
            print("正在尝试读取第一层子控件，看看响应速度...")
            
            # 使用列表推导式快速获取子控件，避免深层遍历
            children = active_win.GetChildren()
            print(f"微信顶层控件数量: {len(children)}")
            for i, child in enumerate(children[:10]): # 只打印前10个防止刷屏
                print(f"  [{i}] 类型: {child.ControlTypeName}, 名称: {child.Name}")
            
            print("\n诊断完成：微信响应正常，可以进行下一步开发。")
        else:
            print("\n警告：当前窗口类名不匹配。")
            print("微信主窗口的类名通常应为: WeChatMainWndForPC")
            print(f"当前窗口的类名是: {active_win.ClassName}")

    except Exception as e:
        print(f"\n发生错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    fast_diagnose()