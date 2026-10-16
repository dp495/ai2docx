import tkinter as tk
from tkinter import font
import os
import platform
import subprocess
import datetime
import atexit
import re
import time

# 记录文件路径及其初始修改时间，用于判断是否被修改
temp_files_info = {}

def print_log(msg, err=True):
    now = datetime.datetime.now().strftime("%X")
    try:
        log_entry.config(state='normal')
        log_entry.delete(0, tk.END)
        if err:
            log = now + ' 错误：' + msg
        else:
            log = now + ' ' + msg
        log_entry.insert(0, log)
        log_entry.config(state='readonly')
    except tk.TclError:
        pass  # 防止窗口已关闭时调用报错

def delete_temp_files():
    for file, initial_mtime in temp_files_info.items():
        try:
            if os.path.exists(file):
                current_mtime = os.path.getmtime(file)
                # 如果当前修改时间比初始时间大 1 秒以上，认为文件被修改过
                if current_mtime > initial_mtime + 1.0:
                    print(f"文件 {os.path.basename(file)} 已被修改，保留不删")
                    continue
                os.remove(file)
        except PermissionError:
            print(f"文件 {os.path.basename(file)} 正在被使用，无法删除")
        except Exception as e:
            print(f'删除文件失败-{e}')

# 注册退出时的清理函数
atexit.register(delete_temp_files)

def convert_to_word():
    md = text_area.get("1.0", tk.END)
    text_area.delete("1.0", tk.END)

    if re.fullmatch(r'^\s*$', md):
        print_log('输入内容为空')
        return
        
    md = re.sub(r'\\\( *', '$', md)
    md = re.sub(r' *\\\)', '$', md)
    md = md.replace('\\[', '$$')
    md = md.replace('\\]', '$$')

    filename = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', '_', re.sub(r'\s+', ' ', md[0:11].rstrip()))
    if not filename:
        filename = "untitled"
    
    # 1. 创建分类输出目录
    md_dir = "output_md"
    docx_dir = "output_docx"
    os.makedirs(md_dir, exist_ok=True)
    os.makedirs(docx_dir, exist_ok=True)
    
    md_path = os.path.join(md_dir, f'{filename}.md')
    docx_path = os.path.join(docx_dir, f'{filename}.docx')
    
    try:
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write(md)    
    except Exception as e:
        print_log(f'写文件失败-{e}')
        return 
    
    # 记录 md 文件的初始修改时间
    time.sleep(0.1)  # 确保时间戳精度
    temp_files_info[md_path] = os.path.getmtime(md_path)
    
    try:
        # 调用 pandoc 转换
        cmd = f'pandoc -f commonmark_x -t docx "{md_path}" -o "{docx_path}" --reference-doc=ref.docx'
        pcp = subprocess.run(cmd, bufsize=0, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception as e:
        print_log(f'调用pandoc失败-{e}')
        return
        
    if pcp.returncode != 0:
        err_msg = pcp.stdout.decode('utf-8', errors='ignore') if pcp.stdout else "（无输出）"
        print_log(f'转换失败-{re.sub(r"\s", " ", err_msg)}')
        return
    
    # 记录 docx 文件的初始修改时间
    time.sleep(0.1)
    temp_files_info[docx_path] = os.path.getmtime(docx_path)
    
    # 3. 根据开关决定是否自动打开
    if auto_open_var.get():
        try:
            os.startfile(docx_path)
            print_log(f'转换成功并打开：{os.path.basename(docx_path)}', 0)
        except Exception as e:
            print_log(f'文件打开失败-{e}')
            return
    else:
        print_log(f'转换成功：{os.path.basename(docx_path)}', 0)


# ================= UI 布局部分 =================

# DPI适配
os_name = platform.system()
scale_factor = 1.0
if os_name == "Windows":
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
        scale_factor *= 0.018
    except Exception as e:
        print(f"适配DPI失败-{e}")
        scale_factor = 1.0

pad0 = int(3 * scale_factor)
pad1 = int(5 * scale_factor)

# 创建主窗口
root = tk.Tk()
root.tk.call('tk', 'scaling', scale_factor)
root.title("语言模型输出转为 Word")
root.geometry(f"{int(320 * scale_factor)}x{int(240 * scale_factor)}")

# 字体设置
SmallFixed = font.nametofont('TkFixedFont').copy()
SmallFixed['size'] = int(font.nametofont('TkFixedFont').cget('size'))

SmallFont = font.nametofont('TkDefaultFont').copy()
SmallFont['size'] = int(font.nametofont('TkDefaultFont').cget('size') * 0.85)

# 使用 grid 布局
root.grid_columnconfigure(0, weight=1)
root.grid_rowconfigure(2, weight=1)

# 标题
instruction_label = tk.Label(root, text="请粘贴复制按钮复制的文字：")
instruction_label.grid(row=0, column=0, sticky='w', padx=pad1, pady=pad1)

# 提示
hint_text = "注：转换会清空输入框，退出时自动清除未被修改的输出文件。"
hint_label = tk.Label(root, text=hint_text, font=SmallFont, wraplength=int(320 * scale_factor) - 2 * pad1, justify="left")
hint_label.grid(row=1, column=0, sticky='w', padx=pad1)

def update_warplength(event):
    instruction_label.config(wraplength=root.winfo_width() - 2 * pad1)
    hint_label.config(wraplength=root.winfo_width() - 2 * pad1)
root.bind("<Configure>", update_warplength)

# 输入框和滚动条
input_frame = tk.Frame(root)
input_frame.grid(row=2, column=0, sticky='nsew')
input_frame.grid_columnconfigure(0, weight=1)
input_frame.grid_rowconfigure(0, weight=1)

text_area = tk.Text(input_frame, height=0, width=0, borderwidth=0, font='TkFixedFont')
text_area.grid(row=0, column=0, sticky='nsew')

scrollbar = tk.Scrollbar(input_frame, command=text_area.yview)
scrollbar.grid(row=0, column=1, sticky='ns')
text_area['yscrollcommand'] = scrollbar.set

# 日志
log_entry = tk.Entry(root, font=SmallFixed, state='readonly')
log_entry.grid(row=3, column=0, sticky='ew', padx=pad0, pady=pad0)

# 控制区域（复选框和按钮）
control_frame = tk.Frame(root)
control_frame.grid(row=4, column=0, sticky='ew', padx=pad1, pady=pad0)

# 自动打开开关
auto_open_var = tk.BooleanVar(value=True)
auto_open_check = tk.Checkbutton(control_frame, text="自动打开.docx文件", variable=auto_open_var, font=SmallFont)
auto_open_check.pack(side='left')

# 转换按钮
read_button = tk.Button(control_frame, text="转换", command=convert_to_word)
read_button.pack(side='right')

# 运行主循环
root.mainloop()
