import ctypes
from ctypes import wintypes
import time
import os
import tkinter as tk
from tkinter import scrolledtext, messagebox

# ================= 局域网配置区 =================
LAN_PWD_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\password.txt" 
LAN_LOG_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\log.txt"
# ===============================================

# --- Windows API 预设 ---
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
kernel32.GlobalLock.restype = wintypes.LPVOID
kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
user32.OpenClipboard.argtypes = [wintypes.HWND]
user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]

VK_RETURN, VK_CONTROL, VK_V, VK_A = 0x0D, 0x11, 0x56, 0x41
KEYEVENTF_KEYUP = 0x0002

def press_key(vk):
    user32.keybd_event(vk, 0, 0, 0)
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)

def hotkey_ctrl(vk):
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(vk, 0, 0, 0)
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)

def set_clipboard(text):
    encoded = text.encode('utf-16le') + b'\x00\x00'
    h_mem = kernel32.GlobalAlloc(0x0002, len(encoded))
    p_mem = kernel32.GlobalLock(h_mem)
    if p_mem:
        ctypes.memmove(p_mem, encoded, len(encoded))
        kernel32.GlobalUnlock(h_mem)
        if user32.OpenClipboard(0):
            user32.EmptyClipboard()
            user32.SetClipboardData(13, h_mem)
            user32.CloseClipboard()

# --- 录入逻辑 ---
def start_auto_input(text_area, root, mode_var, s1, s2, s3):
    raw_data = text_area.get("1.0", tk.END).strip()
    if not raw_data:
        messagebox.showwarning("提示", "条码不能为空")
        return
    codes = [line.strip() for line in raw_data.split('\n') if line.strip()]
    mode = mode_var.get()
    root.withdraw() 
    time.sleep(5)
    for code in codes:
        set_clipboard(code)
        hotkey_ctrl(VK_A)
        time.sleep(0.05)
        hotkey_ctrl(VK_V)
        time.sleep(s1.get()) 
        press_key(VK_RETURN)
        if mode == "double":
            time.sleep(s2.get()) 
            press_key(VK_RETURN)
        time.sleep(s3.get()) 
    messagebox.showinfo("完成", "录入结束")
    root.deiconify()

# --- 主界面 ---
def open_main_window():
    root = tk.Tk()
    root.title("82023703的进站摸鱼工具")
    root.geometry("420x760") 
    
    # 1. 条码输入区
    tk.Label(root, text="条码列表 (一行一个):", font=("微软雅黑", 9, "bold")).pack(pady=5)
    text_area = scrolledtext.ScrolledText(root, width=48, height=12)
    text_area.pack(padx=15, pady=5)
    
    # --- 新增：操作按钮栏 (清空 & 粘贴) ---
    btn_tool_frame = tk.Frame(root)
    btn_tool_frame.pack(fill=tk.X, padx=15, pady=2)
    
    def clear_text():
        text_area.delete("1.0", tk.END)

    def paste_text():
        try:
            clipboard_data = root.clipboard_get()
            text_area.insert(tk.END, clipboard_data)
        except:
            messagebox.showwarning("提示", "剪贴板为空或无法读取")

    tk.Button(btn_tool_frame, text="🗑️ 清空列表", command=clear_text, font=("微软雅黑", 8), bg="#f5f5f5", width=12).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_tool_frame, text="📋 一键粘贴", command=paste_text, font=("微软雅黑", 8), bg="#f5f5f5", width=12).pack(side=tk.LEFT)

    # 2. 模式与速度面板
    config_frame = tk.Frame(root)
    config_frame.pack(fill=tk.X, padx=20)
    
    mode_frame = tk.LabelFrame(config_frame, text="回车模式", padx=10, pady=5)
    mode_frame.pack(fill=tk.X, pady=5)
    mode_var = tk.StringVar(value="single")
    tk.Radiobutton(mode_frame, text="单次回车", variable=mode_var, value="single").pack(side=tk.LEFT, padx=20)
    tk.Radiobutton(mode_frame, text="两次回车(FDT)", variable=mode_var, value="double").pack(side=tk.LEFT, padx=20)
    
    speed_frame = tk.LabelFrame(root, text="速度调节(秒)", padx=10, pady=5)
    speed_frame.pack(fill=tk.X, padx=20, pady=5)
    def create_slider(parent, label, default):
        tk.Label(parent, text=label, font=("微软雅黑", 8)).pack(anchor="w")
        s = tk.Scale(parent, from_=0.0, to=3.0, resolution=0.1, orient=tk.HORIZONTAL)
        s.set(default)
        s.pack(fill=tk.X)
        return s
    s1 = create_slider(speed_frame, "粘贴后等待提交", 0.2)
    s2 = create_slider(speed_frame, "双回车确认间隔", 0.6)
    s3 = create_slider(speed_frame, "下一条循环间隔", 0.8)

    # 3. 更新日志区
    log_frame = tk.LabelFrame(root, text="📢 内网通知", padx=10, pady=5, fg="#555")
    log_frame.pack(fill=tk.X, padx=20, pady=5)
    log_display = scrolledtext.ScrolledText(log_frame, width=48, height=4, font=("微软雅黑", 8), bg="#F0F0F0")
    log_display.pack(fill=tk.X)
    
    try:
        if os.path.exists(LAN_LOG_PATH):
            with open(LAN_LOG_PATH, "r", encoding="utf-8-sig") as f:
                log_display.insert(tk.END, f.read())
        else:
            log_display.insert(tk.END, "未发现内网日志...")
    except:
        log_display.insert(tk.END, "加载失败")
    log_display.config(state=tk.DISABLED)
    
    # 4. 开始按钮
    tk.Button(root, text="🔥 开始自动化录入 (5s准备)", bg="#2E7D32", fg="white", 
              font=("微软雅黑", 11, "bold"), pady=8,
              command=lambda: start_auto_input(text_area, root, mode_var, s1, s2, s3)).pack(pady=15, fill=tk.X, padx=20)
    
    root.mainloop()

# --- 局域网验证界面 ---
def login_screen():
    login_root = tk.Tk()
    login_root.title("安全验证")
    login_root.geometry("300x180")
    sw, sh = login_root.winfo_screenwidth(), login_root.winfo_screenheight()
    login_root.geometry(f"+{int((sw-300)/2)}+{int((sh-180)/2)}")

    tk.Label(login_root, text="请输入局域网授权码", font=("微软雅黑", 10)).pack(pady=15)
    pwd_entry = tk.Entry(login_root, show="*", justify='center', font=("Arial", 12))
    pwd_entry.pack(pady=5, padx=20)
    pwd_entry.focus_set()

    def check_password(event=None):
        try:
            if not os.path.exists(LAN_PWD_PATH):
                messagebox.showerror("连接错误", "无法连接10.1.93.32服务器")
                return
            with open(LAN_PWD_PATH, "r", encoding="utf-8-sig") as f:
                lan_pwd = f.read().strip()
            if pwd_entry.get() == lan_pwd:
                login_root.destroy()
                open_main_window()
            else:
                messagebox.showerror("错误", "授权码错误")
        except:
            messagebox.showerror("失败", "文件权限受阻")

    tk.Button(login_root, text="验证进入", command=check_password, bg="#1565C0", fg="white", width=15).pack(pady=15)
    login_root.bind('<Return>', check_password)
    login_root.mainloop()

if __name__ == "__main__":
    login_screen()

