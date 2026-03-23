import os
import threading
from pdf2docx import Converter
from tkinter import Tk, filedialog, messagebox, Label, Button, Entry, StringVar

class PDFtoWordConverter:
    def __init__(self, master):
        self.master = master
        master.title("PDF转Word工具")
        
        master.geometry("500x350")
        master.resizable(False, False)
        
        self.is_converting = False
        
        self.label_pdf = Label(master, text="选择PDF文件:")
        self.label_pdf.pack(pady=5)
        
        self.pdf_path = StringVar()
        self.entry_pdf = Entry(master, textvariable=self.pdf_path, width=50)
        self.entry_pdf.pack(pady=5)
        
        self.btn_browse_pdf = Button(master, text="浏览", command=self.browse_pdf)
        self.btn_browse_pdf.pack(pady=5)
        
        self.label_word = Label(master, text="保存Word文件到:")
        self.label_word.pack(pady=5)
        
        self.word_path = StringVar()
        self.entry_word = Entry(master, textvariable=self.word_path, width=50)
        self.entry_word.pack(pady=5)
        
        self.btn_browse_word = Button(master, text="浏览", command=self.browse_word)
        self.btn_browse_word.pack(pady=5)
        
        self.btn_convert = Button(master, text="转换", command=self.start_convert, bg="#4CAF50", fg="white", width=15)
        self.btn_convert.pack(pady=20)
        
        self.status_label = Label(master, text="", fg="blue")
        self.status_label.pack(pady=5)
        
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
            self.pdf_path.set(file_path)
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
    
    def validate_input(self):
        pdf_file = self.pdf_path.get().strip()
        word_file = self.word_path.get().strip()
        
        if not pdf_file:
            messagebox.showerror("错误", "请选择PDF文件!")
            return None, None
        
        if not word_file:
            messagebox.showerror("错误", "请指定Word文件保存路径!")
            return None, None
        
        if not os.path.exists(pdf_file):
            messagebox.showerror("错误", "所选PDF文件不存在!")
            return None, None
        
        if not os.path.isfile(pdf_file):
            messagebox.showerror("错误", "所选路径不是文件!")
            return None, None
        
        if not pdf_file.lower().endswith('.pdf'):
            messagebox.showerror("错误", "请选择PDF格式的文件!")
            return None, None
        
        output_dir = os.path.dirname(word_file)
        if output_dir and not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except OSError as e:
                messagebox.showerror("错误", f"无法创建输出目录:\n{str(e)}")
                return None, None
        
        return pdf_file, word_file
    
    def start_convert(self):
        if self.is_converting:
            messagebox.showwarning("提示", "正在转换中，请稍候...")
            return
        
        pdf_file, word_file = self.validate_input()
        if pdf_file is None:
            return
        
        self.is_converting = True
        self.btn_convert.config(state="disabled", text="转换中...")
        self.status_label.config(text="正在转换，请稍候...")
        
        thread = threading.Thread(
            target=self.convert_thread,
            args=(pdf_file, word_file),
            daemon=True
        )
        thread.start()
    
    def convert_thread(self, pdf_file, word_file):
        cv = None
        try:
            cv = Converter(pdf_file)
            cv.convert(word_file, start=0, end=None)
            
            self.master.after(0, lambda: messagebox.showinfo("成功", f"转换完成!\n文件已保存到:\n{word_file}"))
            self.master.after(0, lambda: self.status_label.config(text="转换完成!", fg="green"))
            
        except Exception as e:
            error_msg = str(e)
            if "password" in error_msg.lower() or "encrypted" in error_msg.lower():
                error_msg = "PDF文件已加密，无法转换"
            elif "permission" in error_msg.lower() or "access" in error_msg.lower():
                error_msg = "文件访问权限不足"
            elif "corrupt" in error_msg.lower() or "invalid" in error_msg.lower():
                error_msg = "PDF文件已损坏或格式无效"
            
            self.master.after(0, lambda: messagebox.showerror("错误", f"转换失败:\n{error_msg}"))
            self.master.after(0, lambda: self.status_label.config(text="转换失败", fg="red"))
            
        finally:
            if cv is not None:
                try:
                    cv.close()
                except Exception:
                    pass
            
            self.master.after(0, self.reset_convert_button)
    
    def reset_convert_button(self):
        self.is_converting = False
        self.btn_convert.config(state="normal", text="转换")

if __name__ == "__main__":
    root = Tk()
    app = PDFtoWordConverter(root)
    root.mainloop()
