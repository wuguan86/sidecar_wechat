import uiautomation as auto
import sys
import time

def print_control(control, depth):
    indent = "  " * depth
    try:
        # 获取基本属性
        name = control.Name
        class_name = control.ClassName
        automation_id = control.AutomationId
        control_type = control.ControlTypeName
        rect = control.BoundingRectangle
        
        # 构造输出字符串
        info = f"{indent}[{control_type}] Name='{name}' Class='{class_name}' AutomationId='{automation_id}' Rect={rect}"
        print(info)
        
        # 递归遍历子控件
        children = control.GetChildren()
        for child in children:
            print_control(child, depth + 1)
            
    except Exception as e:
        print(f"{indent}Error reading control: {e}")

def main():
    print("正在寻找微信窗口 (WeChatMainWndForPC)...")
    sys.stdout.flush()
    
    # 尝试查找微信主窗口
    # 注意：微信 3.9.12+ 的类名通常是 WeChatMainWndForPC
    wechat_window = auto.WindowControl(ClassName="WeChatMainWndForPC", searchDepth=1)
    
    if not wechat_window.Exists(maxSearchSeconds=2):
        print("未找到 WeChatMainWndForPC，尝试通过名称 '微信' 查找...")
        wechat_window = auto.WindowControl(Name="微信", searchDepth=1)
        
    if not wechat_window.Exists(maxSearchSeconds=2):
        print("错误：未找到微信主窗口！请确保微信已登录并显示在桌面上。")
        return

    print(f"找到窗口: {wechat_window.Name} (ClassName={wechat_window.ClassName})")
    print("Handle:", wechat_window.NativeWindowHandle)
    print("-" * 50)
    print("开始遍历控件树 (可能需要几秒钟)...")
    print("-" * 50)
    
    # 开始遍历
    print_control(wechat_window, 0)
    
    print("-" * 50)
    print("遍历完成。")
    
    # 保持窗口打开，方便查看输出 (如果是在终端运行，可以直接按回车退出)
    try:
        input("按回车键退出...")
    except:
        pass

if __name__ == "__main__":
    try:
        # 检查管理员权限 (推荐，但不是必须，取决于微信是否以管理员运行)
        if not auto.IsUserAnAdmin():
            print("提示: 当前脚本未以管理员权限运行，如果发现无法获取某些控件，请尝试以管理员身份运行。")
            # time.sleep(1) # Remove sleep to avoid delay
            
        main()
    except Exception as e:
        print(f"致命错误: {e}")
        import traceback
        traceback.print_exc()
    except KeyboardInterrupt:
        print("用户中断")
