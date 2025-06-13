import tkinter as tk
import tool_left
import pyautogui
import os


def get_window_size():
    size = pyautogui.size()
    return size


def tool_window():
    # 获取当前路径
    # now_path = os.getcwd()
    # now_path = resource_path()
    tool = tk.Tk()
    tool.title('use your tool')
    # tool.iconbitmap('../tool_picture/dog_icon.ico')
    # tool.iconphoto(False, tk.PhotoImage(file='tool_picture/dog_head.png'))

    sw = tool.winfo_screenwidth()
    sh = tool.winfo_screenheight()

    w = get_window_size()[0] // 2
    h = get_window_size()[1] // 3 * 2
    x = 200
    y = (sh - h) // 2
    tool.geometry(f'{w}x{h}+{x}+{y}')
    tool.resizable(False, False)

    # 左边功能选择 需要在这里配置切换功能页
    tool_left.tool_left_window(tool)

    tool.mainloop()


# pyinstaller -F tool_client.py -p D:\ATTCommonTool\tool_manage\simple_tiantu_forward.py -p D:\ATTCommonTool\tool_manage\tool_default.py -i D:\ATTCommonTool\tool_picture\dog_icon.ico --noconsole  隐藏控制台窗口

# pyinstaller -F tool_client.py -p D:\ATTCommonTool\tool_manage\simple_tiantu_forward.py -p D:\ATTCommonTool\tool_manage\tool_default.py -i D:\ATTCommonTool\tool_picture\dog_icon.ico

# pyinstaller D:\ATTCommonTool\tool_client.spec
if __name__ == '__main__':
    tool_window()
