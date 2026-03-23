#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF转Word工具 - 修复版
修复内容：
1. 添加编码声明
2. 使用多线程避免界面冻结
3. 使用上下文管理器确保资源释放
4. 添加文件存在性验证
5. 添加文件覆盖提示
6. 自动创建输出目录
7. 验证文件格式
8. 友好的错误信息
9. 窗口居中显示
10. 转换过程中禁用按钮防止重复点击
"""
import os
import threading
from pdf2docx import Converter
from tkinter import Tk, filedialog, messagebox, Label, Button, Entry, StringVar, BooleanVar

class PDFtoWordConverter:
    def __init__(self, master):
        self.master = master
        master.title("PDF转Word工具")
        
        # 设置窗口大小和居中显示
        window_width = 500
        window_height = 300
        screen_width = master.winfo_screenwidth()
        screen_height = master.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        master.geometry(f"{window_width}x{window_height}+{x}+{y}")
        master.resizable(False, False)
        
        # 转换状态标志
        self.is_converting = BooleanVar(value=False)
        
        # PDF文件路径
        self.label_pdf = Label(master, text="选择PDF文件:")
        self.label_pdf.pack(pady=5)
        
        self.pdf_path = StringVar()
        self.entry_pdf = Entry(master, textvariable=self.pdf_path, width=50)
        self.entry_pdf.pack(pady=5)
        
        self.btn_browse_pdf = Button(master, text="浏览", command=self.browse_pdf)
        self.btn_browse_pdf.pack(pady=5)
        
        # Word保存路径
        self.label_word = Label(master, text="保存Word文件到:")
        self.label_word.pack(pady=5)
        
        self.word_path = StringVar()
        self.entry_word = Entry(master, textvariable=self.word_path, width=50)
        self.entry_word.pack(pady=5)
        
        self.btn_browse_word = Button(master, text="浏览", command=self.browse_word)
        self.btn_browse_word.pack(pady=5)
        
        # 转换按钮
        self.btn_convert = Button(master, text="转换", command=self.start_convert, 
                                  bg="#4CAF50", fg="white", width=15, height=2)
        self.btn_convert.pack(pady=20)
        
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            # 验证文件扩展名
            if not file_path.lower().endswith('.pdf'):
                messagebox.showwarning("警告", "请选择PDF格式的文件！")
                return
            self.pdf_path.set(file_path)
            # 自动设置Word文件路径
            word_path = os.path.splitext(file_path)[0] + ".docx"
            self.word_path.set(word_path)
    
    def browse_word(self):
        file_path = filedialog.asksaveasfilename(
            title="保存Word文件",
            defaultextension=".docx",
            filetypes=[("Word文件", "*.docx"), ("所有文件", "*.*")]
        )
        if file_path:
            self.word_path.set(file_path)
    
    def start_convert(self):
        """启动转换（在新线程中执行以避免界面冻结）"""
        if self.is_converting.get():
            messagebox.showinfo("提示", "转换正在进行中，请稍候...")
            return
        
        pdf_file = self.pdf_path.get().strip()
        word_file = self.word_path.get().strip()
        
        # 输入验证
        if not pdf_file:
            messagebox.showerror("错误", "请选择PDF文件！")
            return
        
        if not word_file:
            messagebox.showerror("错误", "请指定Word文件保存路径！")
            return
        
        # 验证PDF文件是否存在
        if not os.path.exists(pdf_file):
            messagebox.showerror("错误", f"PDF文件不存在：\n{pdf_file}")
            return
        
        # 验证是否为PDF文件
        if not pdf_file.lower().endswith('.pdf'):
            messagebox.showerror("错误", "请选择有效的PDF文件！")
            return
        
        # 检查文件覆盖
        if os.path.exists(word_file):
            result = messagebox.askyesno("确认", f"文件已存在：\n{word_file}\n\n是否覆盖？")
            if not result:
                return
        
        # 确保输出目录存在
        output_dir = os.path.dirname(word_file)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir, exist_ok=True)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录：\n{str(e)}")
                return
        
        # 禁用按钮，防止重复点击
        self.set_ui_state(False)
        
        # 在新线程中执行转换
        convert_thread = threading.Thread(
            target=self.convert_worker,
            args=(pdf_file, word_file)
        )
        convert_thread.daemon = True
        convert_thread.start()
    
    def convert_worker(self, pdf_file, word_file):
        """转换工作线程"""
        self.is_converting.set(True)
        
        try:
            # 使用上下文管理器确保资源正确释放
            with Converter(pdf_file) as cv:
                # 转换整个PDF文件
                cv.convert(word_file, start=0, end=None)
            
            # 在主线程中显示成功消息
            self.master.after(0, lambda: self.show_success(word_file))
            
        except Exception as e:
            # 在主线程中显示错误消息
            error_msg = self.get_friendly_error(e)
            self.master.after(0, lambda: self.show_error(error_msg))
        
        finally:
            self.is_converting.set(False)
            # 在主线程中恢复UI状态
            self.master.after(0, lambda: self.set_ui_state(True))
    
    def set_ui_state(self, enabled):
        """设置UI组件状态"""
        state = "normal" if enabled else "disabled"
        self.btn_convert.config(state=state, text="转换" if enabled else "转换中...")
        self.btn_browse_pdf.config(state=state)
        self.btn_browse_word.config(state=state)
        self.entry_pdf.config(state=state)
        self.entry_word.config(state=state)
        if enabled:
            self.btn_convert.config(bg="#4CAF50")
        else:
            self.btn_convert.config(bg="#9E9E9E")
    
    def show_success(self, word_file):
        """显示成功消息"""
        messagebox.showinfo("成功", f"转换完成！\n\n文件已保存到：\n{word_file}")
    
    def show_error(self, error_msg):
        """显示错误消息"""
        messagebox.showerror("错误", f"转换失败：\n\n{error_msg}")
    
    def get_friendly_error(self, exception):
        """获取友好的错误信息"""
        error_str = str(exception).lower()
        
        # 常见错误的友好提示
        if "password" in error_str or "encrypted" in error_str:
            return "PDF文件已加密，无法转换，请先解密。"
        elif "permission" in error_str or "access" in error_str:
            return "无法访问文件，请检查文件权限或是否被其他程序占用。"
        elif "corrupt" in error_str or "damaged" in error_str or "invalid" in error_str:
            return "PDF文件已损坏或格式无效。"
        elif "disk" in error_str or "space" in error_str:
            return "磁盘空间不足，请清理磁盘后重试。"
        else:
            # 返回原始错误信息，但截断过长的内容
            msg = str(exception)
            if len(msg) > 200:
                msg = msg[:200] + "..."
            return f"发生未知错误：{msg}"

if __name__ == "__main__":
    root = Tk()
    app = PDFtoWordConverter(root)
    root.mainloop()
