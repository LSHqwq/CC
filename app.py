import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz
import re
import json
from PIL import Image, ImageTk
import pytesseract
import cv2

class ModernUI:
    def __init__(self, parent):
        self.parent = parent
        self.setup_styles()
        self.create_widgets()
    
    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.style.configure("Title.TLabel", font=("Microsoft YaHei", 28, "bold"), foreground="#2C3E50")
        self.style.configure("Subtitle.TLabel", font=("Microsoft YaHei", 12), foreground="#7F8C8D")
        
        self.style.configure("Card.TLabelframe", background="#ECF0F1", relief="flat")
        self.style.configure("Card.TLabelframe.Label", font=("Microsoft YaHei", 11, "bold"), foreground="#2C3E50", background="#ECF0F1")
        
        self.style.configure("Modern.TButton", font=("Microsoft YaHei", 10), padding=(20, 10))
        self.style.configure("Primary.TButton", font=("Microsoft YaHei", 11, "bold"), padding=(30, 15))
        
        self.style.configure("Info.TLabel", font=("Microsoft YaHei", 10), foreground="#34495E")
        self.style.configure("Hint.TLabel", font=("Microsoft YaHei", 9), foreground="#95A5A6")
        self.style.configure("Format.TLabel", font=("Microsoft YaHei", 9), foreground="#3498DB")
        
        self.style.configure("Status.TLabel", font=("Microsoft YaHei", 9), foreground="#7F8C8D", anchor=tk.W)
        
    def create_widgets(self):
        pass

class CertificateClassifier:
    def __init__(self, root):
        self.root = root
        self.root.title("图像识别分类器")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        self.root.configure(bg="#FFFFFF")
        
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.setup_styles()
        self.create_ui()
        
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        
        self.progress_var = tk.DoubleVar()
        
    def setup_styles(self):
        self.style.configure("Title.TLabel", font=("Microsoft YaHei", 32, "bold"), foreground="#2C3E50")
        self.style.configure("Subtitle.TLabel", font=("Microsoft YaHei", 13), foreground="#7F8C8D")
        
        self.style.configure("Header.TLabel", font=("Microsoft YaHei", 14, "bold"), foreground="#2C3E50")
        self.style.configure("Info.TLabel", font=("Microsoft YaHei", 11), foreground="#34495E")
        self.style.configure("Hint.TLabel", font=("Microsoft YaHei", 9), foreground="#95A5A6", wraplength=400)
        self.style.configure("Format.TLabel", font=("Microsoft YaHei", 9), foreground="#3498DB", wraplength=400)
        
        self.style.configure("Modern.TButton", font=("Microsoft YaHei", 10, "bold"), padding=(20, 12))
        self.style.configure("Primary.TButton", font=("Microsoft YaHei", 12, "bold"), padding=(35, 18))
        self.style.configure("Exit.TButton", font=("Microsoft YaHei", 10), padding=(20, 12))
        
        self.style.configure("Card.TLabelframe", background="#F8F9FA", relief="flat", borderwidth=0)
        self.style.configure("Card.TLabelframe.Label", font=("Microsoft YaHei", 12, "bold"), foreground="#2C3E50", background="#F8F9FA")
        
        self.style.configure("Status.TLabel", font=("Microsoft YaHei", 9), foreground="#7F8C8D", anchor=tk.W)
        
        self.style.configure("Progress.Horizontal.TProgressbar", thickness=8)
        
    def create_ui(self):
        self.main_frame = tk.Frame(self.root, bg="#FFFFFF")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.header_frame = tk.Frame(self.main_frame, bg="#FFFFFF")
        self.header_frame.pack(fill=tk.X, pady=(30, 20))
        
        self.title_label = ttk.Label(self.header_frame, text="图像识别分类器", style="Title.TLabel", background="#FFFFFF")
        self.title_label.pack(pady=(0, 5))
        
        self.subtitle_label = ttk.Label(self.header_frame, text="智能奖状分类工具 - 支持PDF、图片等多种文件格式", style="Subtitle.TLabel", background="#FFFFFF")
        self.subtitle_label.pack(pady=(0, 10))
        
        self.format_label = ttk.Label(self.header_frame, text="支持格式: PDF | JPG | PNG | BMP | TIFF | JPEG", style="Format.TLabel", background="#FFFFFF")
        self.format_label.pack()
        
        content_frame = tk.Frame(self.main_frame, bg="#FFFFFF")
        content_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=20)
        
        folder_card = ttk.LabelFrame(content_frame, text="📁 文件夹选择", style="Card.TLabelframe", padding=25)
        folder_card.pack(fill=tk.X, pady=(0, 20))
        
        source_row = tk.Frame(folder_card, bg="#F8F9FA")
        source_row.pack(fill=tk.X, pady=10)
        ttk.Label(source_row, text="奖状源文件夹:", style="Info.TLabel", background="#F8F9FA").pack(side=tk.LEFT)
        self.source_entry = ttk.Entry(source_row, width=55, font=("Microsoft YaHei", 11))
        self.source_entry.pack(side=tk.LEFT, padx=15)
        self.source_button = ttk.Button(source_row, text="浏览", command=self.select_source_folder, style="Modern.TButton")
        self.source_button.pack(side=tk.LEFT)
        
        dest_row = tk.Frame(folder_card, bg="#F8F9FA")
        dest_row.pack(fill=tk.X, pady=10)
        ttk.Label(dest_row, text="分类目标文件夹:", style="Info.TLabel", background="#F8F9FA").pack(side=tk.LEFT)
        self.dest_entry = ttk.Entry(dest_row, width=55, font=("Microsoft YaHei", 11))
        self.dest_entry.pack(side=tk.LEFT, padx=15)
        self.dest_button = ttk.Button(dest_row, text="浏览", command=self.select_dest_folder, style="Modern.TButton")
        self.dest_button.pack(side=tk.LEFT)
        
        classify_card = ttk.LabelFrame(content_frame, text="🔍 分类条件", style="Card.TLabelframe", padding=25)
        classify_card.pack(fill=tk.X, pady=(0, 20))
        
        condition_row = tk.Frame(classify_card, bg="#F8F9FA")
        condition_row.pack(fill=tk.X, pady=10)
        ttk.Label(condition_row, text="输入条件:", style="Info.TLabel", background="#F8F9FA").pack(side=tk.LEFT)
        self.classify_entry = ttk.Entry(condition_row, width=40, font=("Microsoft YaHei", 13))
        self.classify_entry.pack(side=tk.LEFT, padx=15, ipady=5)
        self.classify_entry.insert(0, "按姓名分类")
        
        hint_row = tk.Frame(classify_card, bg="#F8F9FA")
        hint_row.pack(fill=tk.X, pady=5)
        self.hint_label = ttk.Label(hint_row, text="💡 支持模糊输入，如'姓名'、'名字'、'名'、'按姓'等都能识别 | 可混合使用多个条件", style="Hint.TLabel", background="#F8F9FA")
        self.hint_label.pack(anchor=tk.W)
        
        examples_row = tk.Frame(classify_card, bg="#F8F9FA")
        examples_row.pack(fill=tk.X, pady=5)
        self.examples_label = ttk.Label(examples_row, text="示例: '谁获奖' | '几年' | '按名' | '姓+年' | '名和级' | '人或时间' | '省级' | '市级'", style="Hint.TLabel", background="#F8F9FA")
        self.examples_label.pack(anchor=tk.W)
        
        self.progress_frame = tk.Frame(content_frame, bg="#FFFFFF")
        self.progress_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100, style="Progress.Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X, pady=5)
        
        self.status_label = ttk.Label(self.progress_frame, textvariable=self.status_var, style="Status.TLabel", background="#FFFFFF")
        self.status_label.pack(anchor=tk.W)
        
        button_frame = tk.Frame(self.main_frame, bg="#FFFFFF")
        button_frame.pack(pady=30)
        
        self.start_button = ttk.Button(button_frame, text="🎯 开始分类", command=self.start_classification, style="Primary.TButton")
        self.start_button.pack(side=tk.LEFT, padx=20)
        
        self.exit_button = ttk.Button(button_frame, text="退出", command=self.root.quit, style="Exit.TButton")
        self.exit_button.pack(side=tk.LEFT, padx=20)
        
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
    
    def parse_classify_condition(self, user_input):
        user_input = user_input.lower().strip()
        
        synonyms_map = {
            '姓名': ['姓名', '名字', '名', '人名', '用户名', '获奖人', '参赛者', '选手', 'who', 'person', 'name', 'nm'],
            '年份': ['年份', '年', '年度', '时间', '什么时候', '哪年', '何时', 'year', 'yr', 'date', 'time', 'y'],
            '奖项类型': ['奖项', '奖', '奖种', '奖类型', '奖项类型', '几等奖', '奖项名', 'type', 'kind', 'award', 'prize'],
            '颁发机构': ['机构', '单位', '发证', '颁奖', '主办', '颁发', '组织', 'organization', 'org', 'issuer'],
            '奖项级别': ['省级', '市级', '县级', '镇级', '区级', '国家级', '省', '市', '县', '镇', '区', '国家', '级别', 'level', 'grade', 'rank'],
        }
        
        separators = ['和', '与', '加', '跟', '及', '/', '、', ',', '+', '&', '或', '或者']
        words = [user_input]
        for sep in separators:
            new_words = []
            for word in words:
                new_words.extend(word.split(sep))
            words = new_words
        
        matched_conditions = []
        for word in words:
            word = word.strip()
            if not word:
                continue
            
            for condition, keywords in synonyms_map.items():
                if word in keywords:
                    if condition not in matched_conditions:
                        matched_conditions.append(condition)
                    break
            else:
                for condition, keywords in synonyms_map.items():
                    matched = False
                    for keyword in keywords:
                        if keyword in word or word in keyword:
                            if condition not in matched_conditions:
                                matched_conditions.append(condition)
                            matched = True
                            break
                    if matched:
                        break
        
        if not matched_conditions:
            intent_patterns = {
                '姓名': ['谁', '何人', '哪个人', '哪个人的'],
                '年份': ['什么时候', '何时', '什么时间', '什么年'],
                '奖项类型': ['什么奖', '哪个奖', '奖是什么'],
                '颁发机构': ['谁发的', '哪个单位', '哪里的'],
                '奖项级别': ['什么级别', '多大的', '什么层次'],
            }
            
            for condition, patterns in intent_patterns.items():
                for pattern in patterns:
                    if pattern in user_input:
                        matched_conditions.append(condition)
                        break
        
        if not matched_conditions:
            if '名' in user_input:
                matched_conditions.append('姓名')
            if '年' in user_input:
                matched_conditions.append('年份')
            if '奖' in user_input:
                matched_conditions.append('奖项类型')
            if '级' in user_input:
                matched_conditions.append('奖项级别')
        
        if not matched_conditions:
            matched_conditions = ['姓名']
        
        return matched_conditions
    
    def extract_text_from_pdf(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            text = ""
            for page in doc:
                text += page.get_text()
            doc.close()
            return text
        except Exception as e:
            return ""
    
    def extract_text_from_image(self, image_path):
        try:
            img = cv2.imread(image_path)
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, thresh = cv2.threshold(gray, 127, 255, cv2.THRESH_BINARY)
            temp_image = "temp_image.png"
            cv2.imwrite(temp_image, thresh)
            text = pytesseract.image_to_string(Image.open(temp_image), lang='chi_sim')
            if os.path.exists(temp_image):
                os.remove(temp_image)
            return text
        except Exception as e:
            return ""
    
    def extract_info(self, text):
        info = {
            "姓名": "未知",
            "奖项类型": "未知",
            "年份": "未知",
            "颁发机构": "未知",
            "奖项级别": "未知"
        }
        
        name_match = re.search(r"姓名[:：]?\s*([\u4e00-\u9fa5]{2,4})", text)
        if name_match:
            info["姓名"] = name_match.group(1)
        else:
            name_patterns = [
                r"获奖人[:：]?\s*([\u4e00-\u9fa5]{2,4})",
                r"参赛者[:：]?\s*([\u4e00-\u9fa5]{2,4})",
                r"选手[:：]?\s*([\u4e00-\u9fa5]{2,4})",
            ]
            for pattern in name_patterns:
                match = re.search(pattern, text)
                if match:
                    info["姓名"] = match.group(1)
                    break
        
        year_match = re.search(r"(20[0-2]\d)", text)
        if year_match:
            info["年份"] = year_match.group(1)
        
        award_types = ["特等奖", "一等奖", "二等奖", "三等奖", "优秀奖", "金奖", "银奖", "铜奖", "冠军", "亚军", "季军"]
        for award_type in award_types:
            if award_type in text:
                info["奖项类型"] = award_type
                break
        
        org_patterns = [
            r"颁发单位[:：]?\s*([\u4e00-\u9fa5]+(?:大学|学院|学校|公司|协会|委员会|组委会|教育局))",
            r"发证单位[:：]?\s*([\u4e00-\u9fa5]+(?:大学|学院|学校|公司|协会|委员会|组委会|教育局))",
            r"主办单位[:：]?\s*([\u4e00-\u9fa5]+(?:大学|学院|学校|公司|协会|委员会|组委会|教育局))",
            r"([\u4e00-\u9fa5]+(?:大学|学院|学校|公司|协会|委员会))",
        ]
        for pattern in org_patterns:
            match = re.search(pattern, text)
            if match:
                info["颁发机构"] = match.group(1)
                break
        
        level_patterns = [
            r"([\u4e00-\u9fa5]+)省(级)?(大赛|竞赛|评选|比赛|奖)",
            r"([\u4e00-\u9fa5]+)市(级)?(大赛|竞赛|评选|比赛|奖)",
            r"([\u4e00-\u9fa5]+)县(级)?(大赛|竞赛|评选|比赛|奖)",
            r"([\u4e00-\u9fa5]+)镇(级)?(大赛|竞赛|评选|比赛|奖)",
            r"([\u4e00-\u9fa5]+)区(级)?(大赛|竞赛|评选|比赛|奖)",
            r"(国家级|省级|市级|县级|镇级|区级)(大赛|竞赛|评选|比赛|奖)",
        ]
        
        for pattern in level_patterns:
            match = re.search(pattern, text)
            if match:
                for group in match.groups():
                    if group and group.strip() in ['国家级', '省级', '市级', '县级', '镇级', '区级']:
                        info["奖项级别"] = group.strip()
                        break
                if info["奖项级别"] != "未知":
                    break
        
        if info["奖项级别"] == "未知":
            if "省" in text and ("大赛" in text or "竞赛" in text or "评选" in text):
                info["奖项级别"] = "省级"
            elif "市" in text and ("大赛" in text or "竞赛" in text or "评选" in text):
                info["奖项级别"] = "市级"
            elif "县" in text and ("大赛" in text or "竞赛" in text or "评选" in text):
                info["奖项级别"] = "县级"
            elif "镇" in text and ("大赛" in text or "竞赛" in text or "评选" in text):
                info["奖项级别"] = "镇级"
            elif "区" in text and ("大赛" in text or "竞赛" in text or "评选" in text):
                info["奖项级别"] = "区级"
        
        return info
    
    def generate_folder_name(self, info, conditions):
        if len(conditions) == 1:
            return info.get(conditions[0], "未知")
        
        folder_parts = []
        for condition in conditions:
            value = info.get(condition, "未知")
            if value != "未知":
                folder_parts.append(value)
        
        if not folder_parts:
            return "未知"
        
        folder_name = "_".join(folder_parts)
        return folder_name
    
    def start_classification(self):
        source_folder = self.source_entry.get()
        dest_folder = self.dest_entry.get()
        user_condition = self.classify_entry.get()
        
        if not source_folder or not dest_folder:
            messagebox.showerror("错误", "请选择源文件夹和目标文件夹")
            return
        
        if not user_condition.strip():
            messagebox.showerror("错误", "请输入分类条件")
            return
        
        if not os.path.exists(source_folder):
            messagebox.showerror("错误", "源文件夹不存在")
            return
        
        if not os.path.exists(dest_folder):
            os.makedirs(dest_folder)
        
        conditions = self.parse_classify_condition(user_condition)
        
        self.status_var.set("正在分类...")
        self.progress_var.set(0)
        self.root.update()
        
        try:
            supported_extensions = ['.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']
            pdf_files = [f for f in os.listdir(source_folder) if any(f.lower().endswith(ext) for ext in supported_extensions)]
            total_files = len(pdf_files)
            
            if total_files == 0:
                messagebox.showinfo("提示", "源文件夹中没有支持的文件")
                self.status_var.set("就绪")
                self.progress_var.set(0)
                return
            
            processed_files = 0
            for pdf_file in pdf_files:
                pdf_path = os.path.join(source_folder, pdf_file)
                
                ext = os.path.splitext(pdf_file)[1].lower()
                if ext == '.pdf':
                    text = self.extract_text_from_pdf(pdf_path)
                else:
                    text = self.extract_text_from_image(pdf_path)
                
                info = self.extract_info(text)
                
                folder_name = self.generate_folder_name(info, conditions)
                folder_name = re.sub(r'[<>:"/\\|?*]', '', folder_name)
                
                if not folder_name or folder_name.strip() == "":
                    folder_name = "未分类"
                
                target_folder = os.path.join(dest_folder, folder_name)
                if not os.path.exists(target_folder):
                    os.makedirs(target_folder)
                
                target_path = os.path.join(target_folder, pdf_file)
                shutil.copy2(pdf_path, target_path)
                
                processed_files += 1
                progress = (processed_files / total_files) * 100
                self.progress_var.set(progress)
                self.status_var.set(f"正在分类... {processed_files}/{total_files} ({int(progress)}%)")
                self.root.update()
            
            self.progress_var.set(100)
            self.status_var.set("分类完成！")
            messagebox.showinfo("完成", f"分类完成！\n\n共处理 {processed_files} 个文件\n分类条件: {', '.join(conditions)}")
            self.status_var.set("就绪")
            self.progress_var.set(0)
        except Exception as e:
            messagebox.showerror("错误", f"分类过程中出错: {str(e)}")
            self.status_var.set("就绪")
            self.progress_var.set(0)

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateClassifier(root)
    root.mainloop()