# -*- coding: utf-8 -*-
"""
ỨNG DỤNG XỬ LÝ MÃ CERT VÀ TẠO BÁO CÁO
Tích hợp toàn bộ quy trình trong 1 app
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
import os
import sys
import threading
from datetime import datetime

class AwardsProcessingApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🎓 HỆ THỐNG XỬ LÝ MÃ CERT ASMO")
        self.root.geometry("900x700")
        self.root.resizable(False, False)
        
        # Biến lưu trữ đường dẫn file
        self.file_full = tk.StringVar()
        self.file_trao_giai = tk.StringVar()
        self.output_dir = tk.StringVar(value=os.getcwd())
        
        # Biến trạng thái
        self.is_processing = False
        
        # Tạo giao diện
        self.create_widgets()
        
    def create_widgets(self):
        """Tạo các widget cho giao diện"""
        
        # ========== HEADER ==========
        header_frame = tk.Frame(self.root, bg="#2c3e50", height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="🎓 HỆ THỐNG XỬ LÝ MÃ CERT ASMO",
            font=("Arial", 18, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title_label.pack(pady=20)
        
        # ========== MAIN CONTENT ==========
        main_frame = tk.Frame(self.root, bg="#ecf0f1", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # --- File Input Section ---
        input_frame = tk.LabelFrame(
            main_frame,
            text="📂 CHỌN FILE ĐẦU VÀO",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            fg="#2c3e50",
            padx=15,
            pady=10
        )
        input_frame.pack(fill=tk.X, pady=(0, 15))
        
        # File 1: Awards_Template_Full.xlsx
        tk.Label(
            input_frame,
            text="File đầy đủ (Awards_Template_Full.xlsx):",
            font=("Arial", 10),
            bg="#ecf0f1"
        ).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        file1_entry = tk.Entry(
            input_frame,
            textvariable=self.file_full,
            width=50,
            font=("Arial", 9)
        )
        file1_entry.grid(row=0, column=1, padx=10, pady=5)
        
        tk.Button(
            input_frame,
            text="Chọn file",
            command=lambda: self.browse_file(self.file_full),
            bg="#3498db",
            fg="white",
            font=("Arial", 9, "bold"),
            cursor="hand2"
        ).grid(row=0, column=2, pady=5)
        
        # File 2: Awards_TRAO GIAI.xlsx
        tk.Label(
            input_frame,
            text="File trao giải (Awards_TRAO GIAI.xlsx):",
            font=("Arial", 10),
            bg="#ecf0f1"
        ).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        file2_entry = tk.Entry(
            input_frame,
            textvariable=self.file_trao_giai,
            width=50,
            font=("Arial", 9)
        )
        file2_entry.grid(row=1, column=1, padx=10, pady=5)
        
        tk.Button(
            input_frame,
            text="Chọn file",
            command=lambda: self.browse_file(self.file_trao_giai),
            bg="#3498db",
            fg="white",
            font=("Arial", 9, "bold"),
            cursor="hand2"
        ).grid(row=1, column=2, pady=5)
        
        # Output Directory
        tk.Label(
            input_frame,
            text="Thư mục lưu kết quả:",
            font=("Arial", 10),
            bg="#ecf0f1"
        ).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        output_entry = tk.Entry(
            input_frame,
            textvariable=self.output_dir,
            width=50,
            font=("Arial", 9)
        )
        output_entry.grid(row=2, column=1, padx=10, pady=5)
        
        tk.Button(
            input_frame,
            text="Chọn thư mục",
            command=self.browse_directory,
            bg="#3498db",
            fg="white",
            font=("Arial", 9, "bold"),
            cursor="hand2"
        ).grid(row=2, column=2, pady=5)
        
        # --- Progress Section ---
        progress_frame = tk.LabelFrame(
            main_frame,
            text="📊 TIẾN TRÌNH XỬ LÝ",
            font=("Arial", 12, "bold"),
            bg="#ecf0f1",
            fg="#2c3e50",
            padx=15,
            pady=10
        )
        progress_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Progress bar
        self.progress = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=800
        )
        self.progress.pack(fill=tk.X, pady=10)
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            height=15,
            width=95,
            font=("Consolas", 9),
            bg="#2c3e50",
            fg="#2ecc71",
            insertbackground="white"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # --- Action Buttons ---
        button_frame = tk.Frame(main_frame, bg="#ecf0f1")
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.process_btn = tk.Button(
            button_frame,
            text="▶ BẮT ĐẦU XỬ LÝ",
            command=self.start_processing,
            bg="#27ae60",
            fg="white",
            font=("Arial", 12, "bold"),
            height=2,
            width=20,
            cursor="hand2"
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="🗑 XÓA LOG",
            command=self.clear_log,
            bg="#e67e22",
            fg="white",
            font=("Arial", 12, "bold"),
            height=2,
            width=15,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="📁 MỞ THƯ MỤC",
            command=self.open_output_folder,
            bg="#3498db",
            fg="white",
            font=("Arial", 12, "bold"),
            height=2,
            width=15,
            cursor="hand2"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="❌ THOÁT",
            command=self.root.quit,
            bg="#c0392b",
            fg="white",
            font=("Arial", 12, "bold"),
            height=2,
            width=15,
            cursor="hand2"
        ).pack(side=tk.RIGHT, padx=5)
        
    def browse_file(self, var):
        """Chọn file"""
        filename = filedialog.askopenfilename(
            title="Chọn file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            var.set(filename)
            self.log(f"✅ Đã chọn file: {os.path.basename(filename)}")
    
    def browse_directory(self):
        """Chọn thư mục"""
        directory = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if directory:
            self.output_dir.set(directory)
            self.log(f"✅ Đã chọn thư mục: {directory}")
    
    def log(self, message):
        """Ghi log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def clear_log(self):
        """Xóa log"""
        self.log_text.delete(1.0, tk.END)
    
    def open_output_folder(self):
        """Mở thư mục output"""
        output_dir = self.output_dir.get()
        if os.path.exists(output_dir):
            os.startfile(output_dir)
        else:
            messagebox.showerror("Lỗi", "Thư mục không tồn tại!")
    
    def update_progress(self, value):
        """Cập nhật progress bar"""
        self.progress['value'] = value
        self.root.update_idletasks()
    
    def start_processing(self):
        """Bắt đầu xử lý"""
        if self.is_processing:
            messagebox.showwarning("Cảnh báo", "Đang xử lý, vui lòng đợi!")
            return
        
        # Kiểm tra file input
        if not self.file_full.get() or not self.file_trao_giai.get():
            messagebox.showerror("Lỗi", "Vui lòng chọn đầy đủ 2 file đầu vào!")
            return
        
        if not os.path.exists(self.file_full.get()):
            messagebox.showerror("Lỗi", "File Awards_Template_Full.xlsx không tồn tại!")
            return
        
        if not os.path.exists(self.file_trao_giai.get()):
            messagebox.showerror("Lỗi", "File Awards_TRAO GIAI.xlsx không tồn tại!")
            return
        
        # Chạy xử lý trong thread riêng
        self.is_processing = True
        self.process_btn.config(state=tk.DISABLED, text="⏳ ĐANG XỬ LÝ...")
        
        thread = threading.Thread(target=self.process_all)
        thread.daemon = True
        thread.start()
    
    def process_all(self):
        """Xử lý toàn bộ quy trình"""
        try:
            self.clear_log()
            self.log("="*80)
            self.log("🎓 BẮT ĐẦU QUY TRÌNH XỬ LÝ MÃ CERT")
            self.log("="*80)
            
            output_dir = self.output_dir.get()
            
            # Định nghĩa đường dẫn file output
            file_step1 = os.path.join(output_dir, "Awards_Comparison_Result.xlsx")
            file_step2 = os.path.join(output_dir, "Awards_Comparison_WITH_RANK.xlsx")
            file_step3 = os.path.join(output_dir, "Awards_Comparison_WITH_CERT.xlsx")
            file_step4 = os.path.join(output_dir, "Awards_Comparison_WITH_REPORT.xlsx")
            
            # BƯỚC 1: So sánh và tách file
            self.update_progress(10)
            self.log("\n📌 BƯỚC 1/4: Tạo file so sánh...")
            self.step1_compare_files(file_step1)
            self.update_progress(25)
            
            # BƯỚC 2: Thêm rank và sắp xếp
            self.update_progress(30)
            self.log("\n📌 BƯỚC 2/4: Thêm RANK NHẬN GIẢI và sắp xếp...")
            self.step2_add_rank(file_step1, file_step2)
            self.update_progress(50)
            
            # BƯỚC 3: Tạo mã CERT
            self.update_progress(55)
            self.log("\n📌 BƯỚC 3/4: Tạo mã CERT...")
            self.step3_generate_cert(file_step2, file_step3)
            self.update_progress(75)
            
            # BƯỚC 4: Tạo báo cáo thống kê
            self.update_progress(80)
            self.log("\n📌 BƯỚC 4/4: Tạo báo cáo thống kê...")
            self.step4_create_report(file_step3, file_step4)
            self.update_progress(100)
            
            # Hoàn thành
            self.log("\n" + "="*80)
            self.log("🎉 HOÀN THÀNH TOÀN BỘ QUY TRÌNH!")
            self.log("="*80)
            self.log(f"\n✅ Các file đã tạo:")
            self.log(f"   1. {os.path.basename(file_step1)}")
            self.log(f"   2. {os.path.basename(file_step2)}")
            self.log(f"   3. {os.path.basename(file_step3)}")
            self.log(f"   4. {os.path.basename(file_step4)}")
            
            messagebox.showinfo("Thành công", "Đã hoàn thành toàn bộ quy trình!\n\nCác file đã được lưu vào thư mục output.")
            
        except Exception as e:
            self.log(f"\n❌ LỖI: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra:\n{str(e)}")
        
        finally:
            self.is_processing = False
            self.process_btn.config(state=tk.NORMAL, text="▶ BẮT ĐẦU XỬ LÝ")
            self.update_progress(0)
    
    # ========== CÁC HÀM XỬ LÝ ==========
    
    def step1_compare_files(self, output_file):
        """Bước 1: So sánh và tách file"""
        df_trao_giai = pd.read_excel(self.file_trao_giai.get())
        sbd_trao_giai = set(df_trao_giai['SBD'].dropna().astype(str))
        self.log(f"   ✓ Đọc file TRAO GIẢI: {len(df_trao_giai)} dòng")
        
        df_full = pd.read_excel(self.file_full.get())
        self.log(f"   ✓ Đọc file FULL: {len(df_full)} dòng")
        
        df_full['SBD_str'] = df_full['SBD'].astype(str)
        df_sheet1 = df_full[df_full['SBD_str'].isin(sbd_trao_giai)].drop('SBD_str', axis=1)
        df_sheet2 = df_full[~df_full['SBD_str'].isin(sbd_trao_giai)].drop('SBD_str', axis=1)
        
        self.log(f"   ✓ Sheet TRAO GIẢI: {len(df_sheet1)} học sinh")
        self.log(f"   ✓ Sheet KO ĐK: {len(df_sheet2)} học sinh")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_sheet1.to_excel(writer, sheet_name='TRAO GIẢI', index=False)
            df_sheet2.to_excel(writer, sheet_name='KO ĐK', index=False)
        
        self.log(f"   ✅ Đã lưu: {os.path.basename(output_file)}")
    
    def step2_add_rank(self, input_file, output_file):
        """Bước 2: Thêm rank và sắp xếp"""
        df_trao_giai = pd.read_excel(input_file, sheet_name='TRAO GIẢI')
        df_ko_dk = pd.read_excel(input_file, sheet_name='KO ĐK')
        
        # Tính RANK NHẬN GIẢI
        rank_cols = ['RANK T', 'RANK S', 'RANK E']
        df_trao_giai['RANK NHẬN GIẢI'] = df_trao_giai[rank_cols].min(axis=1, skipna=True)
        
        all_nan_mask = df_trao_giai[rank_cols].isna().all(axis=1)
        df_trao_giai.loc[all_nan_mask, 'RANK NHẬN GIẢI'] = np.nan
        
        # Sắp xếp
        df_trao_giai = df_trao_giai.sort_values(
            ['RANK NHẬN GIẢI', 'KHỐI', 'TRƯỜNG'],
            na_position='last'
        ).reset_index(drop=True)
        
        self.log(f"   ✓ Đã thêm RANK NHẬN GIẢI và sắp xếp")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_trao_giai.to_excel(writer, sheet_name='TRAO GIẢI', index=False)
            df_ko_dk.to_excel(writer, sheet_name='KO ĐK', index=False)
        
        self.log(f"   ✅ Đã lưu: {os.path.basename(output_file)}")
    
    def step3_generate_cert(self, input_file, output_file):
        """Bước 3: Tạo mã CERT"""
        df_trao_giai = pd.read_excel(input_file, sheet_name='TRAO GIẢI')
        df_ko_dk = pd.read_excel(input_file, sheet_name='KO ĐK')
        
        # Sắp xếp sheet KO ĐK theo MÃ TRƯỜNG (cột AC), sau đó đến KHỐI (cột E)
        if 'MÃ TRƯỜNG' in df_ko_dk.columns and 'KHỐI' in df_ko_dk.columns:
            df_ko_dk['MÃ TRƯỜNG'] = df_ko_dk['MÃ TRƯỜNG'].astype(str)
            df_ko_dk = df_ko_dk.sort_values(['MÃ TRƯỜNG', 'KHỐI'], na_position='last').reset_index(drop=True)
        
        # Xử lý sheet TRAO GIẢI
        df_trao_giai = self.process_sheet_cert(df_trao_giai, 1)
        bags1 = df_trao_giai['STT TÚI'].max() if 'STT TÚI' in df_trao_giai.columns else 0
        self.log(f"   ✓ Sheet TRAO GIẢI: {len(df_trao_giai)} HS, {int(bags1)} túi")
        
        # Xử lý sheet KO ĐK
        start_bag = int(bags1) + 1
        df_ko_dk = self.process_sheet_cert(df_ko_dk, start_bag)
        bags2 = df_ko_dk['STT TÚI'].max() - bags1 if 'STT TÚI' in df_ko_dk.columns else 0
        self.log(f"   ✓ Sheet KO ĐK: {len(df_ko_dk)} HS, {int(bags2)} túi")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_trao_giai.to_excel(writer, sheet_name='TRAO GIẢI', index=False)
            df_ko_dk.to_excel(writer, sheet_name='KO ĐK', index=False)
        
        self.log(f"   ✅ Đã lưu: {os.path.basename(output_file)}")
    
    def process_sheet_cert(self, df, start_bag_num):
        """Xử lý tạo mã CERT cho 1 sheet"""
        df['MÃ CERT ĐẦY ĐỦ'] = df.apply(self.generate_cert_code_full, axis=1)
        df['MÃ CERT'] = df.apply(self.generate_cert_code_short, axis=1)
        df['Rank nhận giải'] = df.apply(self.get_highest_rank, axis=1)
        df['SL GCN'] = df.apply(self.count_certificates, axis=1)
        
        # Phân túi
        bag_series = self.assign_bags(df, start_bag_num)
        df['STT TÚI'] = bag_series
        
        # Cập nhật mã CERT với STT túi
        df['MÃ CERT ĐẦY ĐỦ'] = df.apply(
            lambda row: f"{row['MÃ CERT ĐẦY ĐỦ']}*{int(row['STT TÚI'])}" if row['STT TÚI'] > 0 else row['MÃ CERT ĐẦY ĐỦ'],
            axis=1
        )
        df['MÃ CERT'] = df.apply(
            lambda row: f"{row['MÃ CERT']}*{int(row['STT TÚI'])}" if row['STT TÚI'] > 0 else row['MÃ CERT'],
            axis=1
        )
        
        return df
    
    def step4_create_report(self, input_file, output_file):
        """Bước 4: Tạo báo cáo thống kê"""
        df_trao_giai = pd.read_excel(input_file, sheet_name='TRAO GIẢI')
        df_ko_dk = pd.read_excel(input_file, sheet_name='KO ĐK')
        
        # Báo cáo TRAO GIẢI theo khối
        report1 = self.create_report_by_khoi(df_trao_giai)
        self.log(f"   ✓ Báo cáo TRAO GIẢI: {len(report1)-1} khối")
        
        # Báo cáo KO ĐK theo mã trường
        report2 = self.create_report_by_truong(df_ko_dk)
        self.log(f"   ✓ Báo cáo KO ĐK: {len(report2)-1} trường")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            report1.to_excel(writer, sheet_name='BÁO CÁO TRAO GIẢI', index=False)
            report2.to_excel(writer, sheet_name='BÁO CÁO KO ĐK', index=False)
        
        self.log(f"   ✅ Đã lưu: {os.path.basename(output_file)}")
    
    # ========== CÁC HÀM HỖ TRỢ ==========
    
    @staticmethod
    def map_result_to_code(result, subject):
        if pd.isna(result) or result == '':
            return f'NULL-{subject}'
        result = str(result).strip().upper()
        if 'VÀNG' in result or 'VANG' in result:
            return f'V-{subject}'
        elif 'BẠC' in result or 'BAC' in result:
            return f'B-{subject}'
        elif 'ĐỒNG' in result or 'DONG' in result:
            return f'D-{subject}'
        elif 'KHUYẾN KHÍCH' in result or 'KHUYEN KHICH' in result or 'KK' in result:
            return f'KK-{subject}'
        elif 'CHỨNG NHẬN' in result or 'CHUNG NHAN' in result or 'CN' in result:
            return f'CN-{subject}'
        return f'NULL-{subject}'
    
    def generate_cert_code_full(self, row):
        """Tạo mã CERT đầy đủ theo thứ tự: MATH → ENGLISH → SCIENCE"""
        khoi = str(int(row['KHỐI'])) if not pd.isna(row['KHỐI']) else 'X'
        math = self.map_result_to_code(row['KQ VQG TOÁN'], 'MATH')
        english = self.map_result_to_code(row['KQ VQG TIẾNG ANH'], 'ENGLISH')
        science = self.map_result_to_code(row['KQ VQG KHOA HỌC'], 'SCIENCE')
        return f"{khoi}*{math}*{english}*{science}"
    
    def generate_cert_code_short(self, row):
        """Tạo mã CERT rút gọn theo thứ tự: M → E → S"""
        khoi = str(int(row['KHỐI'])) if not pd.isna(row['KHỐI']) else 'X'
        math = self.map_result_to_code(row['KQ VQG TOÁN'], 'M')
        english = self.map_result_to_code(row['KQ VQG TIẾNG ANH'], 'E')
        science = self.map_result_to_code(row['KQ VQG KHOA HỌC'], 'S')
        
        parts = [khoi]
        # Thứ tự: MATH → ENGLISH → SCIENCE
        if not math.startswith('NULL'):
            parts.append(math)
        if not english.startswith('NULL'):
            parts.append(english)
        if not science.startswith('NULL'):
            parts.append(science)
        
        return '*'.join(parts)
    
    def get_highest_rank(self, row):
        math = self.map_result_to_code(row['KQ VQG TOÁN'], 'M')
        science = self.map_result_to_code(row['KQ VQG KHOA HỌC'], 'S')
        english = self.map_result_to_code(row['KQ VQG TIẾNG ANH'], 'E')
        
        rank_priority = {'V': 3, 'B': 2, 'D': 1}
        awards = []
        
        for code, priority in [(math, 3), (english, 2), (science, 1)]:
            if code.startswith(('V-', 'B-', 'D-')):
                rank_type = code.split('-')[0]
                awards.append((rank_priority.get(rank_type, 0), priority, code))
        
        if not awards:
            return ''
        
        awards.sort(key=lambda x: (x[0], x[1]), reverse=True)
        return awards[0][2]
    
    @staticmethod
    def count_certificates(row):
        count = 0
        for col in ['KQ VQG TOÁN', 'KQ VQG KHOA HỌC', 'KQ VQG TIẾNG ANH']:
            if not pd.isna(row[col]) and row[col] != '':
                count += 1
        return count
    
    @staticmethod
    def assign_bags(df, start_bag_number=1, max_gcn=30):
        current_bag = start_bag_number
        current_gcn = 0
        bag_assignments = {}
        
        for idx, row in df.iterrows():
            student_gcn = row['SL GCN']
            
            if student_gcn == 0:
                bag_assignments[idx] = 0
                continue
            
            if current_gcn + student_gcn <= max_gcn:
                current_gcn += student_gcn
                bag_assignments[idx] = current_bag
            else:
                current_bag += 1
                current_gcn = student_gcn
                bag_assignments[idx] = current_bag
        
        return pd.Series(bag_assignments, name='STT TÚI')
    
    def create_report_by_khoi(self, df):
        """Tạo báo cáo theo khối"""
        khoi_list = sorted(df['KHỐI'].dropna().unique())
        report = []
        
        for khoi in khoi_list:
            df_khoi = df[df['KHỐI'] == khoi]
            
            # Đếm số học sinh theo giải cao nhất
            hs_vang = self.count_students_by_highest_award(df_khoi, 'VÀNG')
            hs_bac = self.count_students_by_highest_award(df_khoi, 'BẠC')
            hs_dong = self.count_students_by_highest_award(df_khoi, 'ĐỒNG')
            
            # Đếm số GCN
            vang_gcn = self.count_gcn_for_award(df_khoi, 'VÀNG|VANG')
            bac_gcn = self.count_gcn_for_award(df_khoi, 'BẠC|BAC')
            dong_gcn = self.count_gcn_for_award(df_khoi, 'ĐỒNG|DONG')
            kk_gcn = self.count_gcn_for_award(df_khoi, 'KHUYẾN KHÍCH|KHUYEN KHICH')
            cn_gcn = self.count_gcn_for_award(df_khoi, 'CHỨNG NHẬN|CHUNG NHAN')
            
            report.append({
                'Khối': int(khoi),
                'Tổng HS': len(df_khoi),
                'Số HS VÀNG': hs_vang,
                'Số HS BẠC': hs_bac,
                'Số HS ĐỒNG': hs_dong,
                'GCN VÀNG': vang_gcn,
                'GCN BẠC': bac_gcn,
                'GCN ĐỒNG': dong_gcn,
                'GCN KHUYẾN KHÍCH': kk_gcn,
                'GCN CHỨNG NHẬN': cn_gcn,
                'TỔNG GCN': int(df_khoi['SL GCN'].sum())
            })
        
        # Tổng cộng
        report.append({
            'Khối': 'TỔNG CỘNG',
            'Tổng HS': sum([r['Tổng HS'] for r in report]),
            'Số HS VÀNG': sum([r['Số HS VÀNG'] for r in report]),
            'Số HS BẠC': sum([r['Số HS BẠC'] for r in report]),
            'Số HS ĐỒNG': sum([r['Số HS ĐỒNG'] for r in report]),
            'GCN VÀNG': sum([r['GCN VÀNG'] for r in report]),
            'GCN BẠC': sum([r['GCN BẠC'] for r in report]),
            'GCN ĐỒNG': sum([r['GCN ĐỒNG'] for r in report]),
            'GCN KHUYẾN KHÍCH': sum([r['GCN KHUYẾN KHÍCH'] for r in report]),
            'GCN CHỨNG NHẬN': sum([r['GCN CHỨNG NHẬN'] for r in report]),
            'TỔNG GCN': sum([r['TỔNG GCN'] for r in report])
        })
        
        return pd.DataFrame(report)
    
    def create_report_by_truong(self, df):
        """Tạo báo cáo theo trường"""
        # Sắp xếp theo Mã trường trước, sau đó đến Khối
        df_sorted = df.sort_values(['MÃ TRƯỜNG', 'KHỐI'], na_position='last')
        
        # Lấy danh sách kết hợp (Mã trường, Khối)
        group_keys = df_sorted.groupby(['MÃ TRƯỜNG', 'KHỐI'], dropna=False).size().index.tolist()
        
        report = []
        
        for ma_truong, khoi in group_keys:
            df_truong = df[(df['MÃ TRƯỜNG'] == ma_truong) & (df['KHỐI'] == khoi)]
            
            vang = self.count_gcn_for_award(df_truong, 'VÀNG|VANG')
            bac = self.count_gcn_for_award(df_truong, 'BẠC|BAC')
            dong = self.count_gcn_for_award(df_truong, 'ĐỒNG|DONG')
            
            ten_truong = df_truong['TRƯỜNG'].iloc[0] if 'TRƯỜNG' in df_truong.columns and len(df_truong) > 0 else ''
            khoi_display = int(khoi) if not pd.isna(khoi) else ''
            
            report.append({
                'MÃ TRƯỜNG': str(ma_truong),
                'TÊN TRƯỜNG': ten_truong,
                'Khối': khoi_display,
                'Tổng HS': len(df_truong),
                'GCN VÀNG': vang,
                'GCN BẠC': bac,
                'GCN ĐỒNG': dong,
                'TỔNG GCN': int(df_truong['SL GCN'].sum())
            })
        
        # Tổng cộng
        report.append({
            'MÃ TRƯỜNG': 'TỔNG CỘNG',
            'TÊN TRƯỜNG': '',
            'Khối': '',
            'Tổng HS': sum([r['Tổng HS'] for r in report]),
            'GCN VÀNG': sum([r['GCN VÀNG'] for r in report]),
            'GCN BẠC': sum([r['GCN BẠC'] for r in report]),
            'GCN ĐỒNG': sum([r['GCN ĐỒNG'] for r in report]),
            'TỔNG GCN': sum([r['TỔNG GCN'] for r in report])
        })
        
        return pd.DataFrame(report)
    
    @staticmethod
    def count_gcn_for_award(df, award_type):
        """Đếm số GCN cho loại giải"""
        count = 0
        for col in ['KQ VQG TOÁN', 'KQ VQG KHOA HỌC', 'KQ VQG TIẾNG ANH']:
            if col in df.columns:
                count += df[col].astype(str).str.upper().str.contains(award_type, na=False).sum()
        return count
    
    @staticmethod
    def count_students_by_highest_award(df, award_level):
        """
        Đếm số học sinh theo giải cao nhất
        - VÀNG: có ít nhất 1 giải VÀNG
        - BẠC: giải cao nhất là BẠC (không có VÀNG)
        - ĐỒNG: chỉ có huy chương ĐỒNG (không có VÀNG hoặc BẠC)
        """
        count = 0
        for idx, row in df.iterrows():
            has_vang = False
            has_bac = False
            has_dong = False
            
            for col in ['KQ VQG TOÁN', 'KQ VQG KHOA HỌC', 'KQ VQG TIẾNG ANH']:
                if col in df.columns:
                    val = str(row[col]).upper()
                    if 'VÀNG' in val or 'VANG' in val:
                        has_vang = True
                    elif 'BẠC' in val or 'BAC' in val:
                        has_bac = True
                    elif 'ĐỒNG' in val or 'DONG' in val:
                        has_dong = True
            
            if award_level == 'VÀNG' and has_vang:
                count += 1
            elif award_level == 'BẠC' and not has_vang and has_bac:
                count += 1
            elif award_level == 'ĐỒNG' and not has_vang and not has_bac and has_dong:
                count += 1
        
        return count


def main():
    root = tk.Tk()
    app = AwardsProcessingApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
