import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import PyPDF2
import fitz  # PyMuPDF
import re

class CertificateClassifier:
    def __init__(self, root):
        self.root = root
        self.root.title("奖状分类器")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 设置主题
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # 主框架
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        self.title_label = ttk.Label(self.main_frame, text="奖状分类器", font=("微软雅黑", 24, "bold"))
        self.title_label.pack(pady=20)
        
        # 选择文件夹部分
        self.folder_frame = ttk.LabelFrame(self.main_frame, text="选择文件夹", padding="10")
        self.folder_frame.pack(fill=tk.X, pady=10)
        
        self.source_label = ttk.Label(self.folder_frame, text="奖状源文件夹:")
        self.source_label.grid(row=0, column=0, sticky=tk.W, pady=5)
        
        self.source_entry = ttk.Entry(self.folder_frame, width=50)
        self.source_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        self.source_button = ttk.Button(self.folder_frame, text="浏览", command=self.select_source_folder)
        self.source_button.grid(row=0, column=2, padx=10, pady=5)
        
        self.dest_label = ttk.Label(self.folder_frame, text="分类目标文件夹:")
        self.dest_label.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        self.dest_entry = ttk.Entry(self.folder_frame, width=50)
        self.dest_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        self.dest_button = ttk.Button(self.folder_frame, text="浏览", command=self.select_dest_folder)
        self.dest_button.grid(row=1, column=2, padx=10, pady=5)
        
        # 分类方式部分
        self.classify_frame = ttk.LabelFrame(self.main_frame, text="分类方式", padding="10")
        self.classify_frame.pack(fill=tk.X, pady=10)
        
        self.classify_var = tk.StringVar()
        self.classify_var.set("name")  # 默认按姓名分类
        
        classify_options = [
            ("按姓名分类", "name"),
            ("按奖项类型分类", "type"),
            ("按年份分类", "year"),
            ("按颁发机构分类", "organization")
        ]
        
        for text, value in classify_options:
            ttk.Radiobutton(self.classify_frame, text=text, variable=self.classify_var, value=value).pack(side=tk.LEFT, padx=15)
        
        # 操作按钮
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(pady=20)
        
        self.start_button = ttk.Button(self.button_frame, text="开始分类", command=self.start_classification, style="Accent.TButton")
        self.start_button.pack(side=tk.LEFT, padx=10)
        
        self.exit_button = ttk.Button(self.button_frame, text="退出", command=root.quit)
        self.exit_button.pack(side=tk.LEFT, padx=10)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 自定义样式
        self.style.configure("Accent.TButton", foreground="white", background="#0078d7")
    
    def select_source_folder(self):
        folder = filedialog.askdirectory(title="选择奖状源文件夹")
        if folder:
            self.source_entry.delete(0, tk.END)
            self.source_entry.insert(0, folder)
    
    def select_dest_folder(self):
        folder = filedialog.askdirectory(title="选择分类目标文件夹")
        if folder:
            self.dest_entry.delete(0, tk.END)
            self.dest_entry.insert(0, folder)
    
    def extract_text_from_pdf(self, pdf_path):
        """从PDF中提取文本"""
        try:
            # 使用PyMuPDF提取文本
            doc = fitz.open(pdf_path)
            text = ""
            for page in doc:
                text += page.get_text()
            doc.close()
            return text
        except Exception as e:
            messagebox.showerror("错误", f"提取PDF文本失败: {str(e)}")
            return ""
    
    def extract_info(self, text):
        """从文本中提取信息"""
        info = {
            "name": "未知",
            "type": "未知",
            "year": "未知",
            "organization": "未知"
        }
        
        # 尝试提取姓名（简单示例，实际可能需要更复杂的规则）
        name_match = re.search(r"姓名[:：]?\s*([\u4e00-\u9fa5]+)", text)
        if not name_match:
            # 尝试其他模式
            name_match = re.search(r"([\u4e00-\u9fa5]{2,4})", text)
        if name_match:
            info["name"] = name_match.group(1)
        
        # 尝试提取年份
        year_match = re.search(r"(20[0-2]\d)", text)
        if year_match:
            info["year"] = year_match.group(1)
        
        # 尝试提取奖项类型
        award_types = ["一等奖", "二等奖", "三等奖", "优秀奖", "特等奖", "金奖", "银奖", "铜奖"]
        for award_type in award_types:
            if award_type in text:
                info["type"] = award_type
                break
        
        # 尝试提取颁发机构
        org_match = re.search(r"颁发机构[:：]?\s*([\u4e00-\u9fa5]+)", text)
        if not org_match:
            # 尝试其他模式
            org_match = re.search(r"(.*?[大学|学院|学校|公司|协会|委员会])", text)
        if org_match:
            info["organization"] = org_match.group(1)
        
        return info
    
    def start_classification(self):
        source_folder = self.source_entry.get()
        dest_folder = self.dest_entry.get()
        classify_by = self.classify_var.get()
        
        if not source_folder or not dest_folder:
            messagebox.showerror("错误", "请选择源文件夹和目标文件夹")
            return
        
        if not os.path.exists(source_folder):
            messagebox.showerror("错误", "源文件夹不存在")
            return
        
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
        
        self.status_var.set("正在分类...")
        self.root.update()
        
        try:
            pdf_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.pdf')]
            total_files = len(pdf_files)
            
            if total_files == 0:
                messagebox.showinfo("提示", "源文件夹中没有PDF文件")
                self.status_var.set("就绪")
                return
            
            processed_files = 0
            for pdf_file in pdf_files:
                pdf_path = os.path.join(source_folder, pdf_file)
                text = self.extract_text_from_pdf(pdf_path)
                info = self.extract_info(text)
                
                # 根据分类方式创建文件夹
                if classify_by == "name":
                    folder_name = info["name"]
                elif classify_by == "type":
                    folder_name = info["type"]
                elif classify_by == "year":
                    folder_name = info["year"]
                else:  # organization
                    folder_name = info["organization"]
                
                # 清理文件夹名称中的非法字符
                folder_name = re.sub(r'[<>:"/\\|?*]', '', folder_name)
                
                # 创建目标文件夹
                target_folder = os.path.join(dest_folder, folder_name)
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)
                
                # 复制文件到目标文件夹
                target_path = os.path.join(target_folder, pdf_file)
                shutil.copy2(pdf_path, target_path)
                
                processed_files += 1
                self.status_var.set(f"正在分类... ({processed_files}/{total_files})")
                self.root.update()
            
            messagebox.showinfo("完成", f"分类完成！共处理 {processed_files} 个PDF文件")
            self.status_var.set("就绪")
        except Exception as e:
            messagebox.showerror("错误", f"分类过程中出错: {str(e)}")
            self.status_var.set("就绪")

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateClassifier(root)
    root.mainloop()