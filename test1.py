import uiautomation as auto

def inspect_deep_wechat():
    print("正在寻找微信窗口...")
    wechat_win = auto.WindowControl(ClassName="mmui::MainWindow", searchDepth=1)
    
    if not wechat_win.Exists(1):
        print("未找到微信窗口！")
        return

    print(f"锁定窗口: {wechat_win.Name}")
    print("-" * 50)

    # 策略：直接找可能包含内容的容器
    # 微信的主界面通常由一个 Splitter 分割（左边联系人，右边聊天框）
    # 我们尝试寻找这个核心容器，然后只遍历它的内部
    target_container = wechat_win.Control(ClassName="mmui::XSplitterView")
    
    if not target_container.Exists(1):
        # 如果找不到 Splitter，就找 StackedWidget
        target_container = wechat_win.Control(ClassName="QStackedWidget")
    
    if target_container.Exists(1):
        print(f"锁定核心容器: {target_container.ClassName}，正在深挖内部结构...")
        
        def walk(control, level):
            # 增加深度限制到 12 层，Qt 结构通常很深
            if level > 12: return 
            
            children = control.GetChildren()
            for child in children:
                indent = "  " * level
                
                # 过滤掉一些没用的纯容器，只显示有名字的或者关键控件
                # EditControl = 输入框, ListControl = 消息列表, TextControl = 文字
                is_important = child.ControlType in [
                    auto.ControlType.ListControl, 
                    auto.ControlType.EditControl, 
                    auto.ControlType.TextControl,
                    auto.ControlType.ButtonControl
                ] or child.Name != ""
                
                # 为了调试，暂时全部打印，但标记出重要类型
                prefix = ">>> " if is_important else ""
                
                print(f"{indent}{prefix}[{child.ControlType}] Name='{child.Name}' Class='{child.ClassName}'")
                
                # 递归
                walk(child, level + 1)

        walk(target_container, 0)
    else:
        print("奇怪，没找到核心容器，尝试全局搜索输入框...")
        # 最后的保底手段：直接查找输入框，看看它在哪里
        edit = wechat_win.EditControl(searchDepth=15)
        if edit.Exists(2):
            print(f"找到输入框！Name={edit.Name}, 位于: {edit.GetParentControl().ClassName}")
        else:
            print("完全无法透视内部结构。可能需要管理员权限运行 VSCode/CMD。")

if __name__ == "__main__":
    inspect_deep_wechat()