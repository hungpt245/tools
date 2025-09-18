# gop_diem_gui.py
import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

# ---------- Cấu hình / helper ----------
def extract_info_from_filename(filename):
    """Lấy mã môn học và mã nhóm từ tên file dạng: (251-CPS201-07)"""
    match = re.search(r"\((\d+)-([A-Z0-9]+)-(\d+)\)", filename.upper())
    if match:
        return match.group(2), match.group(3)
    # fallback: tìm token kiểu CPS201 trong tên file
    match2 = re.search(r"([A-Z]{2,}\d{2,})", filename.upper())
    if match2:
        return match2.group(1), ''
    return '', ''

def format_mssv_value(v):
    """Chuẩn hóa MSSV: nếu là số nguyên như 12345.0 -> '12345'"""
    if pd.isna(v):
        return ''
    try:
        f = float(v)
        if f.is_integer():
            return str(int(f))
        else:
            return str(v).strip()
    except Exception:
        return str(v).strip()

# ---------- Xử lý file ----------
def merge_files(paths, log_func):
    all_parts = []
    for p in paths:
        basename = os.path.basename(p)
        log_func(f">>> Xử lý file: {basename}")
        ma_mon, ma_nhom = extract_info_from_filename(basename)
        try:
            # header=7 => dùng dòng 8 làm tiêu đề (user nói dòng 8)
            df = pd.read_excel(p, header=7, engine='openpyxl')
        except Exception as e:
            log_func(f"  Lỗi đọc file: {e}")
            continue

        # Tìm cột MSSV (dựa vào tiêu đề)
        mssv_col = None
        for c in df.columns:
            cs = str(c).upper()
            if ("MÃ" in cs and "SINH" in cs) or "MSSV" in cs or "MÃ SV" in cs or "MÃSINH" in cs:
                mssv_col = c
                break

        if mssv_col is None:
            log_func("  ❌ Không tìm thấy cột MSSV theo tên. Bỏ file này.")
            continue

        # Lấy điểm từ cột L (index 11) nếu có, ngược lại fallback tìm theo tên chứa 'TBC' hoặc 'ĐIỂM'
        if df.shape[1] >= 12:
            tbc_series = df.iloc[:, 11]   # cột L
            tbc_col_header = df.columns[11]
            log_func(f"  Lấy điểm từ cột L (header: '{tbc_col_header}').")
        else:
            # fallback: tìm bằng tên cột
            tbc_col_header = None
            for c in df.columns:
                cs = str(c).upper()
                if "TBC" in cs or "TBC ĐTP" in cs or "TỔNG" in cs or "ĐIỂM TRUNG" in cs or "ĐIỂM" in cs:
                    tbc_col_header = c
                    break
            if tbc_col_header is None:
                log_func("  ❌ Không tìm thấy cột điểm (cột L và không có cột theo tên). Bỏ file này.")
                continue
            tbc_series = df[tbc_col_header]
            log_func(f"  Lấy điểm từ cột theo tên: '{tbc_col_header}'.")

        # Tạo DataFrame tạm
        temp = pd.DataFrame()
        temp['MSSV'] = df[mssv_col].apply(format_mssv_value)
        # convert điểm numeric (nếu dạng text hay có dấu) -> numeric, lỗi -> NaN
        temp['Điểm trung bình cộng'] = pd.to_numeric(tbc_series, errors='coerce')
        temp['Mã môn học'] = ma_mon
        temp['Mã nhóm'] = ma_nhom

        # Lọc: MSSV không rỗng, điểm không rỗng
        before = len(temp)
        temp = temp[temp['MSSV'].notna()]
        temp = temp[temp['MSSV'].str.replace(r'\s+', '', regex=True) != '']
        temp = temp[temp['Điểm trung bình cộng'].notna()]
        after = len(temp)

        log_func(f"  Dòng trước lọc: {before}, sau lọc hợp lệ: {after}")
        if after > 0:
            all_parts.append(temp)

    if not all_parts:
        return None
    merged = pd.concat(all_parts, ignore_index=True)
    # Sắp cột cho dễ nhìn
    merged = merged[['MSSV', 'Mã môn học', 'Mã nhóm', 'Điểm trung bình cộng']]
    return merged

# ---------- Giao diện Tkinter ----------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Gộp dữ liệu điểm - Lấy cột L cho Điểm trung bình cộng")
        root.geometry("750x480")

        frm_top = tk.Frame(root)
        frm_top.pack(fill='x', padx=10, pady=8)

        btn_add = tk.Button(frm_top, text="Chọn file Excel", width=18, command=self.select_files)
        btn_add.pack(side='left', padx=5)

        btn_clear = tk.Button(frm_top, text="Xóa danh sách", width=12, command=self.clear_list)
        btn_clear.pack(side='left', padx=5)

        btn_merge = tk.Button(frm_top, text="Gộp dữ liệu và lưu", width=20, command=self.process_files)
        btn_merge.pack(side='right', padx=5)

        # Listbox các file
        self.lb = tk.Listbox(root, width=110, height=8)
        self.lb.pack(padx=10, pady=(0,8))

        # Log area
        lbl_log = tk.Label(root, text="Nhật ký xử lý:")
        lbl_log.pack(anchor='w', padx=10)
        self.log = scrolledtext.ScrolledText(root, width=100, height=14, state='disabled')
        self.log.pack(padx=10, pady=(0,10))

    def log_write(self, text):
        self.log.configure(state='normal')
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.configure(state='disabled')

    def select_files(self):
        paths = filedialog.askopenfilenames(
            title="Chọn các file Excel (có header ở dòng 8)",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if paths:
            for p in paths:
                self.lb.insert(tk.END, p)

    def clear_list(self):
        self.lb.delete(0, tk.END)
        self.log.configure(state='normal')
        self.log.delete('1.0', tk.END)
        self.log.configure(state='disabled')

    def process_files(self):
        paths = list(self.lb.get(0, tk.END))
        if not paths:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn ít nhất 1 file Excel.")
            return
        self.log_write("Bắt đầu gộp...")

        merged = merge_files(paths, self.log_write)
        if merged is None or merged.empty:
            messagebox.showerror("Kết quả", "Không tìm thấy dữ liệu hợp lệ để gộp.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            title="Lưu file gộp"
        )
        if save_path:
            try:
                merged.to_excel(save_path, index=False)
                messagebox.showinfo("Hoàn tất", f"Đã gộp dữ liệu và lưu tại:\n{save_path}")
                self.log_write(f"Hoàn tất. File lưu tại: {save_path}")
            except Exception as e:
                messagebox.showerror("Lỗi lưu", f"Lỗi khi lưu file: {e}")
                self.log_write(f"Lỗi lưu file: {e}")

# ---------- Chạy ứng dụng ----------
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
