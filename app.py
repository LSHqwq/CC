import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz
import re
import json

class CertificateClassifier:
    def __init__(self, root):
        self.root = root
        self.root.title("奖状分类器")
        self.root.geometry("800x650")
        self.root.resizable(True, True)

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.title_label = ttk.Label(self.main_frame, text="奖状分类器", font=("微软雅黑", 24, "bold"))
        self.title_label.pack(pady=20)

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

        self.classify_frame = ttk.LabelFrame(self.main_frame, text="输入分类条件", padding="10")
        self.classify_frame.pack(fill=tk.X, pady=10)

        self.classify_label = ttk.Label(self.classify_frame, text="分类条件:", font=("微软雅黑", 11))
        self.classify_label.grid(row=0, column=0, sticky=tk.W, pady=10)

        self.classify_entry = ttk.Entry(self.classify_frame, width=60, font=("微软雅黑", 11))
        self.classify_entry.grid(row=0, column=1, sticky=tk.W, pady=10, padx=10)
        self.classify_entry.insert(0, "按姓名分类")

        self.hint_label = ttk.Label(self.classify_frame, text="提示: 支持模糊输入，如'姓名'、'名字'、'名'、'按姓'等都能识别为姓名分类", 
                                    foreground="gray", font=("微软雅黑", 9))
        self.hint_label.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.examples_label = ttk.Label(self.classify_frame, text="示例: '谁获奖'、'几年'、'按名'、'加奖'、'姓+年'、'名和级'、'人或时间'", 
                                        foreground="gray", font=("微软雅黑", 9))
        self.examples_label.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(pady=20)

        self.start_button = ttk.Button(self.button_frame, text="开始分类", command=self.start_classification, style="Accent.TButton")
        self.start_button.pack(side=tk.LEFT, padx=10)

        self.exit_button = ttk.Button(self.button_frame, text="退出", command=root.quit)
        self.exit_button.pack(side=tk.LEFT, padx=10)

        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

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

    def parse_classify_condition(self, user_input):
        """智能解析用户输入的分类条件，支持模糊匹配"""
        user_input = user_input.lower().strip()

        # 同义词和变体词库
        synonyms_map = {
            '姓名': ['姓名', '名字', '名', '人名', '用户名', '获奖人', '参赛者', '选手', 'who', 'person', 'name', 'nm'],
            '年份': ['年份', '年', '年度', '时间', '什么时候', '哪年', '何时', 'year', 'yr', 'date', 'time', 'y'],
            '奖项类型': ['奖项', '奖', '奖种', '奖类型', '奖项类型', '几等奖', '奖项名', 'type', 'kind', 'award', 'prize'],
            '颁发机构': ['机构', '单位', '发证', '颁奖', '主办', '颁发', '组织', 'organization', 'org', 'issuer', 'issuer'],
            '奖项级别': ['省级', '市级', '县级', '镇级', '区级', '国家级', '省', '市', '县', '镇', '区', '国家', '级别', 'level', 'grade', 'rank'],
        }

        # 连接词和分隔符
        separators = ['和', '与', '加', '跟', '及', '/', '、', ',', '+', '&', '或', '或者']

        # 将输入按分隔符拆分
        words = [user_input]
        for sep in separators:
            new_words = []
            for word in words:
                new_words.extend(word.split(sep))
            words = new_words

        # 对每个词进行模糊匹配
        matched_conditions = []
        for word in words:
            word = word.strip()
            if not word:
                continue

            # 先精确匹配
            for condition, keywords in synonyms_map.items():
                if word in keywords:
                    if condition not in matched_conditions:
                        matched_conditions.append(condition)
                    break
            else:
                # 模糊匹配 - 检查词是否包含关键词或关键词包含该词
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

        # 如果仍然没有匹配，尝试理解意图
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

        # 再次尝试单字符匹配（提高容错性）
        if not matched_conditions:
            if '名' in user_input:
                matched_conditions.append('姓名')
            if '年' in user_input:
                matched_conditions.append('年份')
            if '奖' in user_input:
                matched_conditions.append('奖项类型')
            if '级' in user_input:
                matched_conditions.append('奖项级别')

        # 默认按姓名分类
        if not matched_conditions:
            matched_conditions = ['姓名']

        return matched_conditions

    def extract_text_from_pdf(self, pdf_path):
        """从PDF中提取文本"""
        try:
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
        """根据分类条件生成文件夹名称"""
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
        print(f"解析到的分类条件: {conditions}")

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
                self.status_var.set(f"正在分类... ({processed_files}/{total_files})")
                self.root.update()

            messagebox.showinfo("完成", f"分类完成！共处理 {processed_files} 个PDF文件\n分类条件: {', '.join(conditions)}")
            self.status_var.set("就绪")
        except Exception as e:
            messagebox.showerror("错误", f"分类过程中出错: {str(e)}")
            self.status_var.set("就绪")

if __name__ == "__main__":
    root = tk.Tk()
    app = CertificateClassifier(root)
    root.mainloop()