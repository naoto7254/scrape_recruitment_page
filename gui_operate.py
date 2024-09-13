import tkinter as tk
from tkinter import filedialog, messagebox, PhotoImage
from tkinter import ttk
import excel_processor as excel_p

class GUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("抽出アプリ")
        self.root.geometry("400x150")
        self.root.configure(bg="#f0f0f0")
        
        self.title_label = tk.Label(self.root, text="抽出ツール", font=("Arial", 16, "bold"), bg="#f0f0f0")
        self.title_label.pack(pady=10)
        
        logo = PhotoImage(file='images/logo_c.png')
        
        self.label = tk.Label(
            self.root,
            text="Excelファイルを選択してください",
            font=("Arial", 12), bg="#f0f0f0",
            # image = logo,
        )
        
        self.label.pack(pady=10)

        self.upload_button = ttk.Button(self.root, text="Excelファイルをアップロード", command=self.upload_file)
        self.upload_button.pack(pady=10, padx=20)
        
        self.status_label = tk.Label(self.root, text="", font=("Arial", 10), fg="green", bg="#f0f0f0")
        self.status_label.pack(pady=10)
    
        # self.root.iconbitmap('images/logo_c.png')
    
    def upload_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            processor = excel_p.ExcelProcessor(file_path)
            elements_data, company_name, urls = processor.input_data()
            header_array, urls_array, search_elements_array, scrape_result = processor.processing_data_for_scraping()
            wb = processor.processing_data_for_save_excel(scrape_result, header_array, urls_array)
            
            
            save_path = filedialog.asksaveasfilename(
                initialfile=f"{company_name}_スクレイピング結果.xlsx",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if save_path:
                processor.save_excel(save_path, wb)
                self.label.config(text="ファイルが保存されました: " + save_path)

    def run(self):
        self.root.mainloop()
