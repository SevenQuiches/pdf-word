import os
import threading
from pdf2docx import Converter
from tkinter import Tk, filedialog, messagebox, Label, Button, Entry, StringVar, Toplevel, ttk

class PDFtoWordConverter:
    def __init__(self, master):
        self.master = master
        master.title("PDF转Word工具")
        
        # 设置窗口大小和位置
        master.geometry("500x300")
        master.resizable(False, False)
        
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
        self.btn_convert = Button(master, text="转换", command=self.convert, bg="#4CAF50", fg="white")
        self.btn_convert.pack(pady=20)
        
        # 进度窗口（初始隐藏）
        self.progress_window = None
        self.progress_bar = None
        
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            # 检查文件是否存在
            if not os.path.exists(file_path):
                messagebox.showerror("错误", "选择的文件不存在!")
                return
            # 检查是否是PDF文件
            if not file_path.lower().endswith('.pdf'):
                messagebox.showwarning("警告", "选择的文件不是PDF格式，请确认!")
            
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
    
    def show_progress(self):
        """显示进度窗口"""
        self.progress_window = Toplevel(self.master)
        self.progress_window.title("转换中...")
        self.progress_window.geometry("300x100")
        self.progress_window.resizable(False, False)
        self.progress_window.transient(self.master)
        self.progress_window.grab_set()
        
        Label(self.progress_window, text="正在转换，请稍候...").pack(pady=10)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_window, 
            mode='indeterminate',
            length=250
        )
        self.progress_bar.pack(pady=10)
        self.progress_bar.start(10)
        
        # 禁用主窗口按钮
        self.btn_convert.config(state='disabled')
        self.btn_browse_pdf.config(state='disabled')
        self.btn_browse_word.config(state='disabled')
    
    def hide_progress(self):
        """隐藏进度窗口"""
        if self.progress_window:
            self.progress_window.destroy()
            self.progress_window = None
        
        # 启用主窗口按钮
        self.btn_convert.config(state='normal')
        self.btn_browse_pdf.config(state='normal')
        self.btn_browse_word.config(state='normal')
    
    def convert_in_thread(self, pdf_file, word_file):
        """在后台线程中执行转换"""
        try:
            # 创建输出目录（如果不存在）
            output_dir = os.path.dirname(word_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # 创建转换器对象
            cv = Converter(pdf_file)
            
            # 转换整个PDF文件
            cv.convert(word_file, start=0, end=None)
            
            # 关闭转换器
            cv.close()
            
            # 在主线程中显示成功消息
            self.master.after(0, lambda: self.show_success(word_file))
        except Exception as e:
            # 在主线程中显示错误消息
            self.master.after(0, lambda: self.show_error(str(e)))
    
    def show_success(self, word_file):
        """显示成功消息"""
        self.hide_progress()
        messagebox.showinfo("成功", f"转换完成!\n文件已保存到:\n{word_file}")
    
    def show_error(self, error_msg):
        """显示错误消息"""
        self.hide_progress()
        messagebox.showerror("错误", f"转换失败:\n{error_msg}")
    
    def convert(self):
        pdf_file = self.pdf_path.get().strip()
        word_file = self.word_path.get().strip()
        
        if not pdf_file:
            messagebox.showerror("错误", "请选择PDF文件!")
            return
        
        if not word_file:
            messagebox.showerror("错误", "请指定Word文件保存路径!")
            return
        
        # 检查PDF文件是否存在
        if not os.path.exists(pdf_file):
            messagebox.showerror("错误", "PDF文件不存在!\n请重新选择文件。")
            return
        
        # 检查PDF文件扩展名
        if not pdf_file.lower().endswith('.pdf'):
            if not messagebox.askyesno("确认", "选择的文件可能不是PDF格式，是否继续转换?"):
                return
        
        # 检查输出路径是否可写
        output_dir = os.path.dirname(word_file)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录:\n{str(e)}")
                return
        
        # 如果输出文件已存在，询问是否覆盖
        if os.path.exists(word_file):
            if not messagebox.askyesno("确认", "输出文件已存在，是否覆盖?"):
                return
        
        # 显示进度窗口
        self.show_progress()
        
        # 在后台线程中执行转换
        thread = threading.Thread(
            target=self.convert_in_thread,
            args=(pdf_file, word_file),
            daemon=True
        )
        thread.start()

if __name__ == "__main__":
    root = Tk()
    app = PDFtoWordConverter(root)
    root.mainloop()
