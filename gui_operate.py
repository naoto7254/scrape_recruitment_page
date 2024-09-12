import tkinter as tk
from tkinter import filedialog
import excel_processor as excel_p

class GUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("スクレイパーアプリ")
        self.label = tk.Label(self.root, text="Excelファイルを選択してください")
        self.label.pack(pady=20)

        self.upload_button = tk.Button(self.root, text="Excelファイルをアップロード", command=self.upload_file)
        self.upload_button.pack(pady=10)
    
    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            processor = excel_p.ExcelProcessor(file_path)
            elements_data, company_name, urls = processor.process_data()
            
            save_path = filedialog.asksaveasfilename(
                initialfile=f"{company_name}_スクレイピング結果.xlsx",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if save_path:
                processor.save_excel(save_path)
                self.label.config(text="ファイルが保存されました: " + save_path)

    def run(self):
        self.root.mainloop()
