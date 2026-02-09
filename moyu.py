import ctypes
from ctypes import wintypes
import time
import os
import tkinter as tk
from tkinter import messagebox, ttk, scrolledtext
import sys
import subprocess
import hashlib
import re

# ================= 局域网配置区 =================
LAN_PWD_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\password.txt" 
LAN_LOG_PATH = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\log.txt"
LAN_UPDATE_SRC = r"\\10.1.93.32\DT_HU_RDteam_F\视频\Z\密码\update\摸鱼进站工具.exe"
# ===============================================

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
VK_RETURN, VK_CONTROL, VK_V, VK_A = 0x0D, 0x11, 0x56, 0x41
KEYEVENTF_KEYUP = 0x0002

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

def hotkey_ctrl(vk):
    user32.keybd_event(VK_CONTROL, 0, 0, 0)
    user32.keybd_event(vk, 0, 0, 0)
    user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
    user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)

def open_main_window():
    root = tk.Tk()
    root.title("邹秋的摸鱼进站工具 v2.8")
    root.geometry("400x820") # 稍微拉高高度给日志留空间
    root.attributes("-topmost", True)

    # 1. 列表区
    tree_frame = tk.Frame(root)
    tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    tree = ttk.Treeview(tree_frame, columns=("check", "sn"), show="headings", height=10)
    tree.heading("check", text="选")
    tree.heading("sn", text="序列号 SN")
    tree.column("check", width=30, anchor="center")
    tree.column("sn", width=330)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    sb = tk.Scrollbar(tree_frame, command=tree.yview)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    tree.config(yscrollcommand=sb.set)

    def get_all_sns(): return [tree.set(k, "sn") for k in tree.get_children()]

    def refresh_and_sort():
        sns = list(set(get_all_sns()))
        def nat_sort(t): return [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', t)]
        sns.sort(key=nat_sort)
        for i in tree.get_children(): tree.delete(i)
        for s in sns: tree.insert("", tk.END, values=("☐", s))

    def paste_sn():
        try:
            data = root.clipboard_get()
            for line in data.split('\n'):
                if line.strip(): tree.insert("", tk.END, values=("☐", line.strip()))
            refresh_and_sort()
        except: pass

    def toggle_check(event):
        item = tree.identify_row(event.y)
        if item:
            cur = tree.set(item, "check")
            tree.set(item, "check", "☑" if cur == "☐" else "☐")
    tree.bind("<Button-1>", toggle_check)

    # 2. 按钮区
    btn_f = tk.Frame(root)
    btn_f.pack(fill=tk.X, padx=5)
    tk.Button(btn_f, text="📋 粘贴排序", command=paste_sn, bg="#E1F5FE", font=("微软雅黑", 9)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
    tk.Button(btn_f, text="❌ 删除勾选", command=lambda: [tree.delete(i) for i in tree.get_children() if tree.set(i, "check") == "☑"], bg="#FFEBEE", font=("微软雅黑", 9)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)
    tk.Button(btn_f, text="🗑️ 清空", command=lambda: tree.delete(*tree.get_children()), font=("微软雅黑", 9)).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=2)

    # 3. 调速滑块区
    speed_frame = tk.LabelFrame(root, text="🚀 速度控制(秒)", font=("微软雅黑", 8))
    speed_frame.pack(fill=tk.X, padx=10, pady=5)
    def create_s(label, default):
        f = tk.Frame(speed_frame)
        f.pack(fill=tk.X)
        tk.Label(f, text=label, width=12, font=("微软雅黑", 8)).pack(side=tk.LEFT)
        s = tk.Scale(f, from_=0.0, to=2.0, resolution=0.1, orient=tk.HORIZONTAL)
        s.set(default); s.pack(side=tk.RIGHT, fill=tk.X, expand=True); return s
    s1 = create_s("粘贴后提交:", 0.2)
    s2 = create_s("双回车间隔:", 0.5)
    s3 = create_s("下一条循环:", 0.7)

    mode_var = tk.StringVar(value="single")
    mode_f = tk.Frame(root)
    mode_f.pack(pady=2)
    tk.Radiobutton(mode_f, text="单回车", variable=mode_var, value="single").pack(side=tk.LEFT, padx=15)
    tk.Radiobutton(mode_f, text="双回车(FDT)", variable=mode_var, value="double").pack(side=tk.LEFT)

    # --- 新增：日志显示区 (log.txt) ---
    log_frame = tk.LabelFrame(root, text="📢 内网通知 (log.txt)", padx=5, pady=5, font=("微软雅黑", 8), fg="#555")
    log_frame.pack(fill=tk.X, padx=10, pady=5)
    log_display = scrolledtext.ScrolledText(log_frame, height=4, font=("微软雅黑", 8), bg="#F5F5F5")
    log_display.pack(fill=tk.X)
    try:
        if os.path.exists(LAN_LOG_PATH):
            with open(LAN_LOG_PATH, "r", encoding="utf-8-sig") as f:
                log_display.insert(tk.END, f.read())
        else: log_display.insert(tk.END, "未发现日志文件")
    except: log_display.insert(tk.END, "读取日志失败")
    log_display.config(state=tk.DISABLED)

    # 4. 执行逻辑
    def start_work():
        sns = get_all_sns()
        if not sns: return
        root.withdraw()
        fw = tk.Toplevel()
        fw.overrideredirect(True); fw.attributes("-topmost", True)
        fw.geometry(f"130x30+{root.winfo_screenwidth()-140}+20")
        msg = tk.Label(fw, text="准备...", bg="black", fg="white", font=("微软雅黑", 9))
        msg.pack(fill=tk.BOTH, expand=True)

        time.sleep(5)
        for i, sn in enumerate(sns):
            msg.config(text=f"进度: {i+1}/{len(sns)}"); fw.update()
            set_clipboard(str(sn))
            hotkey_ctrl(VK_A); time.sleep(0.05); hotkey_ctrl(VK_V)
            time.sleep(s1.get())
            user32.keybd_event(VK_RETURN, 0, 0, 0); user32.keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
            if mode_var.get() == "double":
                time.sleep(s2.get()); user32.keybd_event(VK_RETURN, 0, 0, 0); user32.keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0)
            time.sleep(s3.get())
        fw.destroy(); messagebox.showinfo("完成", "录入结束"); root.deiconify()

    tk.Button(root, text="🔥 开始录入 (5s准备)", bg="#2E7D32", fg="white", font=("微软雅黑", 10, "bold"), pady=10, command=start_work).pack(fill=tk.X, padx=10, pady=10)
    root.mainloop()

# 更新检测与登录
def get_file_md5(f):
    if not os.path.exists(f): return None
    h = hashlib.md5()
    with open(f, "rb") as _f:
        for c in iter(lambda: _f.read(4096), b""): h.update(c)
    return h.hexdigest()

def check_for_updates():
    try:
        src = LAN_UPDATE_SRC
        if not os.path.exists(src): return
        cur = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
        if get_file_md5(src) != get_file_md5(cur):
            if messagebox.askyesno("更新", "检测到新版本，是否升级？"):
                with open("updater.bat", "w") as f:
                    f.write(f'@echo off\ntimeout /t 1\ncopy /y "{src}" "{cur}"\nstart "" "{cur}"\ndel %0')
                subprocess.Popen("updater.bat", shell=True); sys.exit()
    except: pass

def login_screen():
    lr = tk.Tk(); lr.title("验证"); lr.geometry("240x120")
    lr.eval('tk::PlaceWindow . center')
    tk.Label(lr, text="授权码:").pack()
    pw = tk.Entry(lr, show="*"); pw.pack(); pw.focus_set()
    def go():
        try:
            with open(LAN_PWD_PATH, "r", encoding="utf-8-sig") as f:
                if pw.get() == f.read().strip():
                    lr.withdraw(); check_for_updates(); lr.destroy(); open_main_window()
                else: messagebox.showerror("!", "错")
        except: messagebox.showerror("!", "连不上内网")
    tk.Button(lr, text="进入", command=go).pack(pady=5)
    lr.bind('<Return>', lambda e: go()); lr.mainloop()

if __name__ == "__main__":
    login_screen()
