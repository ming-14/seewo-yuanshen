###################################################
# @file manager.py
# @brief 启动图片设置
# @author Rikka Github/ming-14
###################################################

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import os
import shutil
import json

class ImageSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("启动图片设置工具")
        self.root.geometry("900x700")
        
        # 初始化变量 - 调整尺寸参数
        self.image_library = []
        self.current_image_path = "./img.pngx"
        self.thumbnail_size = (180, 180)
        self.current_preview_size = (280, 210)
        
        # 创建菜单栏
        self.create_menu()
        
        # 创建UI框架
        self.create_widgets()
        
        # 加载图片库
        self.load_image_library()
        
        # 显示当前启动图片
        self.show_current_image()
        
        # 显示图片库
        self.display_image_library()

    def create_menu(self):
        # 创建菜单栏
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 创建"文件"菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 添加"刷新"菜单项
        file_menu.add_command(label="刷新", command=self.refresh_app, accelerator="F5")
        
        # 添加"关于"菜单项
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)
        
        # 添加键盘快捷键
        self.root.bind("<F5>", lambda event: self.refresh_app())

    def refresh_app(self):
        """刷新应用程序，重新加载图片库"""
        try:
            # 重新加载图片库
            self.load_image_library()
            
            # 更新图片库显示
            self.display_image_library()
            
            # 更新当前图片显示
            self.show_current_image()
        except Exception as e:
            messagebox.showerror("刷新错误", f"刷新过程中出错: {str(e)}")

    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "希沃启动图片设置工具",
            "本程序应与 seewo-yuanshen 配套使用\n\n"
            "开发者: Rikka github/ming-14\n"
            "GitHub: https://github.com/ming-14\n"
        )

    def create_widgets(self):
        # 主框架（垂直布局）
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 当前图片预览区（上方）- 减小高度
        current_frame = ttk.LabelFrame(main_frame, text="当前启动图片")
        current_frame.pack(fill=tk.X, padx=5, pady=(0, 10))  # 减少垂直间距
        
        # 当前图片显示区域
        current_img_frame = ttk.Frame(current_frame)
        current_img_frame.pack(fill=tk.X, pady=5)
        
        self.current_label = ttk.Label(current_img_frame)
        self.current_label.pack(padx=10, pady=5)  # 减少内边距
        
        self.current_info = ttk.Label(current_frame, text="未检测到启动图片")
        self.current_info.pack(pady=(0, 5))  # 减少底部间距
        
        # 图片库区域（下方）- 增大比例
        library_frame = ttk.LabelFrame(main_frame, text="图片库")
        library_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(library_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 画布用于容纳图片库
        self.canvas = tk.Canvas(library_frame, yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=self.canvas.yview)
        
        # 图片容器框架
        self.image_container = ttk.Frame(self.canvas)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.image_container, anchor="nw")
        
        # 绑定事件
        self.image_container.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        
        # 添加鼠标滚轮支持
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        self.image_container.bind("<MouseWheel>", self.on_mousewheel)

    def on_mousewheel(self, event):
        """处理鼠标滚轮事件，实现滚动功能"""
        if event.delta:
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        else:
            # 兼容Linux系统
            if event.num == 4:
                self.canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(1, "units")

    def on_frame_configure(self, event):
        """更新滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def on_canvas_configure(self, event):
        """调整画布宽度"""
        self.canvas.itemconfig(self.canvas_frame, width=event.width)

    def load_image_library(self):
        json_path = "./img/files.json"
        try:
            with open(json_path, "r", encoding="utf-8-sig") as f:
                self.image_library = json.load(f)
        except Exception as e:
            messagebox.showerror("错误", f"加载图片库失败: {str(e)}")

    def show_current_image(self):
        if os.path.exists(self.current_image_path):
            try:
                img = Image.open(self.current_image_path)
                # 保持宽高比调整大小
                img.thumbnail(self.current_preview_size, Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                self.current_label.config(image=photo)
                self.current_label.image = photo
                self.current_info.config(text="当前启动图片: " + os.path.basename(self.current_image_path))
            except Exception as e:
                self.current_label.config(image='')
                self.current_info.config(text=f"无法加载图片: {str(e)}")
        else:
            self.current_label.config(image='')
            self.current_info.config(text="未检测到启动图片")

    def display_image_library(self):
        # 清除现有内容
        for widget in self.image_container.winfo_children():
            widget.destroy()
        
        # 显示每张图片 - 每行显示3张（缩略图增大后每行显示数量不变）
        for idx, img_info in enumerate(self.image_library):
            frame = ttk.Frame(self.image_container)
            frame.grid(row=idx // 3, column=idx % 3, padx=15, pady=15, sticky="nsew")
            
            # 配置网格列权重
            self.image_container.grid_columnconfigure(idx % 3, weight=1)
            
            # 显示缩略图 - 使用更大的尺寸
            try:
                img_path = os.path.join("./img", img_info["name"])
                img = Image.open(img_path)
                img.thumbnail(self.thumbnail_size, Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                img_label = ttk.Label(frame, image=photo)
                img_label.image = photo
                img_label.pack(padx=8, pady=8)  # 增加内边距使图片更突出
            except Exception as e:
                img_label = ttk.Label(frame, text=f"无法加载图片\n{str(e)}", foreground="red")
                img_label.pack(padx=5, pady=5)
            
            # 显示图片信息
            info_text = f"{img_info['brief']}"
            info_label = ttk.Label(frame, text=info_text, wraplength=180)  # 增加换行宽度
            info_label.pack(padx=5, pady=5)
            
            # 设置按钮
            set_button = ttk.Button(
                frame, 
                text="设为启动图片",
                command=lambda path=img_path, name=img_info["name"]: self.set_as_start_image(path, name)
            )
            set_button.pack(padx=5, pady=5, fill=tk.X)

    def set_as_start_image(self, source_path, image_name):
        try:
            # 复制文件
            shutil.copy2(source_path, self.current_image_path)
            
            # 刷新显示
            self.show_current_image()
            messagebox.showinfo("成功", f"已设置 {image_name} 为启动图片")
        except Exception as e:
            messagebox.showerror("错误", f"设置启动图片失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ImageSelectorApp(root)
    root.mainloop()