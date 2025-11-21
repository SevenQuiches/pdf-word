import os
from pdf2docx import Converter
from tkinter import Tk, filedialog, messagebox, Label, Button, Entry, StringVar

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
        
    def browse_pdf(self):
        file_path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if file_path:
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
    
    def convert(self):
        pdf_file = self.pdf_path.get()
        word_file = self.word_path.get()
        
        if not pdf_file:
            messagebox.showerror("错误", "请选择PDF文件!")
            return
        
        if not word_file:
            messagebox.showerror("错误", "请指定Word文件保存路径!")
            return
        
        try:
            # 创建转换器对象
            cv = Converter(pdf_file)
            
            # 转换整个PDF文件
            cv.convert(word_file, start=0, end=None)
            
            # 关闭转换器
            cv.close()
            
            messagebox.showinfo("成功", f"转换完成!\n文件已保存到:\n{word_file}")
        except Exception as e:
            messagebox.showerror("错误", f"转换失败:\n{str(e)}")

if __name__ == "__main__":
    root = Tk()
    app = PDFtoWordConverter(root)
    root.mainloop()