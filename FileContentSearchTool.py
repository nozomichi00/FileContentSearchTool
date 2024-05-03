import os
import re
from PyPDF2 import PdfReader
from pptx import Presentation
import customtkinter as ctk
from tkinter import filedialog, messagebox, Listbox, END
import xml.etree.ElementTree as ET

class FolderSearchApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("File Content Search Tool")
        self.geometry("1024x800")

        self.file_path_mapping = {}
        self.folder_list = []
        self.load_folders()

        self.create_widgets()

    def create_widgets(self):
        self.folder_list_label = ctk.CTkLabel(self, text="路徑設定", font=("Microsoft JhengHei", 14))
        self.folder_list_label.pack(pady=5)

        folder_frame = ctk.CTkFrame(self)
        folder_frame.pack(fill=ctk.BOTH, expand=True, padx=int(self.winfo_width() * 0.025), pady=5)

        self.folder_listbox = Listbox(folder_frame, height=1)
        self.folder_listbox.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True)

        self.folder_scrollbar = ctk.CTkScrollbar(folder_frame, orientation='vertical', command=self.folder_listbox.yview)
        self.folder_scrollbar.pack(side=ctk.RIGHT, fill=ctk.Y)

        self.folder_listbox.config(yscrollcommand=self.folder_scrollbar.set)
        self.update_folder_listbox()

        self.button_frame = ctk.CTkFrame(self, fg_color='transparent')
        self.button_frame.pack(fill=ctk.X, pady=5)

        self.add_folder_button = ctk.CTkButton(self.button_frame, text="Add Folder", command=self.add_folder)
        self.add_folder_button.pack(side=ctk.LEFT, padx=10, expand=True, anchor='center')

        self.delete_folder_button = ctk.CTkButton(self.button_frame, text="Delete Folder", command=self.delete_folder)
        self.delete_folder_button.pack(side=ctk.LEFT, padx=10, expand=True, anchor='center')

        self.search_results_label = ctk.CTkLabel(self, text="搜尋功能 (結果雙點即可自動開啟)", font=("Microsoft JhengHei", 14))
        self.search_results_label.pack(pady=(20, 5))

        self.search_frame = ctk.CTkFrame(self, fg_color='transparent')
        self.search_frame.pack(fill=ctk.X, pady=5)

        self.search_entry = ctk.CTkEntry(self.search_frame, placeholder_text="Enter keyword")
        self.search_entry.pack(side=ctk.LEFT, expand=True, fill=ctk.X, padx=10)

        self.search_button = ctk.CTkButton(self.search_frame, text="Search", command=self.search)
        self.search_button.pack(side=ctk.LEFT, padx=10)

        result_frame = ctk.CTkFrame(self)
        result_frame.pack(fill=ctk.BOTH, expand=True, padx=int(self.winfo_width() * 0.025), pady=5)

        self.search_results = Listbox(result_frame, height=23)
        self.search_results.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True)

        self.search_scrollbar = ctk.CTkScrollbar(result_frame, orientation='vertical', command=self.search_results.yview)
        self.search_scrollbar.pack(side=ctk.RIGHT, fill=ctk.Y)

        self.search_results.config(yscrollcommand=self.search_scrollbar.set)

        self.search_results.bind("<Double-Button-1>", self.open_file)

        self.progress_bar = ctk.CTkProgressBar(self, height=20, progress_color="#90EE90")
        self.progress_bar.pack(fill=ctk.X, padx=self.winfo_width() * 0.025, pady=(5, 5))
        self.progress_bar.set(0)

    def add_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            if folder_path not in self.folder_list:
                self.folder_list.append(folder_path)
                self.update_folder_listbox()
                self.save_folders()

    def delete_folder(self):
        selected_indices = self.folder_listbox.curselection()
        if selected_indices:
            selected_folder = self.folder_listbox.get(selected_indices[0])
            if selected_folder in self.folder_list:
                self.folder_list.remove(selected_folder)
            self.folder_listbox.delete(selected_indices[0])
            self.save_folders()

    def update_folder_listbox(self):
        self.folder_listbox.delete(0, END)
        for folder in self.folder_list:
            self.folder_listbox.insert(0, folder)

    def search(self):
        keyword = self.search_entry.get()
        if not keyword:
            messagebox.showwarning("Warning", "請輸入關鍵字進行搜索。")
            return

        # 將關鍵字轉換為小寫，以進行區分大小寫的搜索
        keyword_lower = keyword.lower()

        # 清空搜索結果列表框
        self.search_results.delete(0, END)

        # 用於跟蹤已找到的文件
        found_files = set()

        # 計算所有文件總數，用於進度條
        total_files = 0
        for folder in self.folder_list:
            for root, dirs, files in os.walk(folder):
                total_files += len(files)
        
        processed_files = 0

        # 遍歷文件夾列表
        for folder in self.folder_list:
            for root, dirs, files in os.walk(folder):
                for file in files:
                    # 進行文件處理計數
                    processed_files += 1
                    progress = processed_files / total_files
                    self.progress_bar.set(progress)
                    self.update()

                    if file.endswith(".pdf") or file.endswith(".pptx"):
                        file_path = os.path.join(root, file)
                        if file_path not in found_files:
                            try:
                                if file.endswith(".pdf"):
                                    # 處理PDF文件
                                    reader = PdfReader(file_path)
                                    for page_num, page in enumerate(reader.pages):
                                        # 提取原始文本，不轉換為小寫
                                        original_text = page.extract_text()

                                        # 將原始文本轉換為小寫以進行搜索匹配
                                        text = original_text.lower()

                                        # 檢查關鍵字是否存在於文本中
                                        if keyword_lower in text:
                                            lines = original_text.split("\n")
                                            # 遍歷每一行，找到包含關鍵字的行
                                            for line in lines:
                                                if keyword in line:
                                                    file_name = os.path.basename(file_path)
                                                    # 顯示原始字串
                                                    result_text = f"檔案名稱: {file_name}、原始字串「{line}」"
                                                    self.search_results.insert(END, result_text)

                                                    # 保存結果與文件路徑的映射
                                                    self.file_path_mapping[result_text] = file_path
                                                    found_files.add(file_path)
                                                    break  # 找到匹配行後跳出

                                        # 如果找到匹配文件，則結束當前頁面的搜索
                                        if file_path in found_files:
                                            break

                                elif file.endswith(".pptx"):
                                    # 處理PPT文件
                                    presentation = Presentation(file_path)
                                    for slide_num, slide in enumerate(presentation.slides):
                                        # 提取原始文本，不轉換為小寫
                                        original_text = []

                                        for shape in slide.shapes:
                                            if shape.has_text_frame:
                                                original_text.append(shape.text_frame.text)
                                        slide_text = " ".join(original_text)

                                        # 將原始文本轉換為小寫以進行搜索匹配
                                        text_lower = slide_text.lower()

                                        # 檢查關鍵字是否存在於文本中
                                        if keyword_lower in text_lower:
                                            lines = slide_text.split("\n")
                                            # 遍歷每一行，找到包含關鍵字的行
                                            for line in lines:
                                                if keyword in line:
                                                    file_name = os.path.basename(file_path)
                                                    # 顯示原始字串
                                                    result_text = f"檔案名稱: {file_name}、原始字串「{line}」"
                                                    self.search_results.insert(END, result_text)

                                                    # 保存結果與文件路徑的映射
                                                    self.file_path_mapping[result_text] = file_path
                                                    found_files.add(file_path)
                                                    break  # 找到匹配行後跳出
                                        
                                        # 如果找到匹配文件，則結束當前頁面的搜索
                                        if file_path in found_files:
                                            break
                                            
                            except Exception as e:
                                messagebox.showinfo(f"Error", f"處理文件時出現錯誤：{file_path}\n錯誤原因：{e}")

        if not self.search_results.size():
            messagebox.showinfo("Information", "搜索完畢，但結果是空的。")

        self.progress_bar.set(1)

    def open_file(self, event):
        selected_index = self.search_results.curselection()
        if selected_index:
            selected_text = self.search_results.get(selected_index)
            selected_file_path = self.file_path_mapping.get(selected_text)
            if selected_file_path:
                selected_file_path = re.sub(r"/", r"\\", selected_file_path)
                os.startfile(selected_file_path)

    def save_folders(self):
        root = ET.Element("folders")
        for folder in self.folder_list:
            ET.SubElement(root, "folder").text = folder

        tree = ET.ElementTree(root)
        tree.write("folders.xml")

    def load_folders(self):
        try:
            tree = ET.parse("folders.xml")
            root = tree.getroot()
            self.folder_list = [folder.text for folder in root.findall("folder")]
        except FileNotFoundError:
            self.folder_list = []

if __name__ == "__main__":
    app = FolderSearchApp()
    app.mainloop()
