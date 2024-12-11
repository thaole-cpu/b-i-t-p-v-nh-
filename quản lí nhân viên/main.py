import tkinter as tk
from tkinter import messagebox, ttk
import csv
from datetime import datetime
import pandas as pd

# File lưu trữ dữ liệu
CSV_FILE = "employees.csv"

# Hàm lưu thông tin nhân viên vào file CSV
def save_employee():
    employee_data = {
        "Mã số": entry_id.get(),
        "Tên": entry_name.get(),
        "Đơn vị": entry_unit.get(),
        "Chức danh": entry_role.get(),
        "Ngày sinh": entry_birth.get(),
        "Giới tính": gender_var.get(),
        "Số CMND": entry_id_number.get(),
        "Nơi cấp": entry_place_of_issue.get(),
        "Ngày cấp": entry_issue_date.get(),
        "Loại": "Khách hàng" if customer_var.get() else ("Nhà cung cấp" if supplier_var.get() else "")
    }

    with open(CSV_FILE, mode="a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=employee_data.keys())
        if file.tell() == 0:  # Ghi tiêu đề nếu file trống
            writer.writeheader()
        writer.writerow(employee_data)

    messagebox.showinfo("Thông báo", "Lưu thông tin thành công!")
    clear_entries()

# Hàm xóa dữ liệu trên giao diện
def clear_entries():
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_unit.delete(0, tk.END)
    entry_role.delete(0, tk.END)
    entry_birth.delete(0, tk.END)
    entry_id_number.delete(0, tk.END)
    entry_place_of_issue.delete(0, tk.END)
    entry_issue_date.delete(0, tk.END)
    gender_var.set("Nam")
    customer_var.set(0)
    supplier_var.set(0)

# Hàm hiển thị danh sách sinh nhật hôm nay
def show_today_birthdays():
    today = datetime.now().strftime("%d/%m/%Y")
    try:
        with open(CSV_FILE, mode="r", encoding="utf-8") as file:
            reader = csv.DictReader(file)
            birthdays = [row for row in reader if row["Ngày sinh"] == today]

        if birthdays:
            result = "\n".join([f"Mã số: {emp['Mã số']}, Tên: {emp['Tên']}" for emp in birthdays])
            messagebox.showinfo("Sinh nhật hôm nay", result)
        else:
            messagebox.showinfo("Thông báo", "Không có nhân viên nào sinh nhật hôm nay.")
    except FileNotFoundError:
        messagebox.showwarning("Lỗi", "Chưa có dữ liệu nhân viên.")

# Hàm xuất danh sách nhân viên ra file Excel
def export_to_excel():
    try:
        df = pd.read_csv(CSV_FILE)
        df["Tuổi"] = df["Ngày sinh"].apply(lambda x: datetime.now().year - datetime.strptime(x, "%d/%m/%Y").year)
        df.sort_values(by="Tuổi", ascending=False, inplace=True)
        df.drop(columns=["Tuổi"], inplace=True)
        excel_file = "employees.xlsx"
        df.to_excel(excel_file, index=False, encoding="utf-8")
        messagebox.showinfo("Thông báo", f"Xuất danh sách thành công vào file {excel_file}!")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")

# Tạo giao diện chính
root = tk.Tk()
root.title("Quản lý thông tin nhân viên")

# Các biến
entry_id = ttk.Entry(root)
entry_name = ttk.Entry(root)
entry_unit = ttk.Entry(root)
entry_role = ttk.Entry(root)
entry_birth = ttk.Entry(root)
entry_id_number = ttk.Entry(root)
entry_place_of_issue = ttk.Entry(root)
entry_issue_date = ttk.Entry(root)
gender_var = tk.StringVar(value="Nam")
customer_var = tk.IntVar()
supplier_var = tk.IntVar()

# Bố cục giao diện
frame_info = tk.LabelFrame(root, text="Thông tin nhân viên", padx=10, pady=10)
frame_info.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

frame_options = tk.LabelFrame(root, text="Loại", padx=10, pady=10)
frame_options.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

# Khung thông tin nhân viên - phía bên trái
left_frame = tk.Frame(frame_info)
left_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

tk.Label(left_frame, text="Mã số:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
entry_id.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

tk.Label(left_frame, text="Tên:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
entry_name.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

tk.Label(left_frame, text="Đơn vị:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
entry_unit.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

tk.Label(left_frame, text="Chức danh:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
entry_role.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

# Khung thông tin nhân viên - phía bên phải
right_frame = tk.Frame(frame_info)
right_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

tk.Label(right_frame, text="Ngày sinh:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
entry_birth.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

tk.Label(right_frame, text="Giới tính:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
tk.Radiobutton(right_frame, text="Nam", variable=gender_var, value="Nam").grid(row=1, column=1, sticky="w", padx=5, pady=5)
tk.Radiobutton(right_frame, text="Nữ", variable=gender_var, value="Nữ").grid(row=1, column=2, sticky="w", padx=5, pady=5)

tk.Label(right_frame, text="Số CMND:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
entry_id_number.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

tk.Label(right_frame, text="Nơi cấp:").grid(row=3, column=0, sticky="w", padx=5, pady=5)
entry_place_of_issue.grid(row=3, column=1, padx=5, pady=5, sticky="ew")

tk.Label(right_frame, text="Ngày cấp:").grid(row=4, column=0, sticky="w", padx=5, pady=5)
entry_issue_date.grid(row=4, column=1, padx=5, pady=5, sticky="ew")

# Khung loại nhân viên
tk.Checkbutton(frame_options, text="Là khách hàng", variable=customer_var).grid(row=0, column=0, sticky="w", padx=5, pady=5)
tk.Checkbutton(frame_options, text="Là nhà cung cấp", variable=supplier_var).grid(row=0, column=1, sticky="w", padx=5, pady=5)

# Các nút chức năng
tk.Button(root, text="Lưu thông tin", command=save_employee).grid(row=2, column=0, padx=5, pady=10, sticky="ew")
tk.Button(root, text="Sinh nhật hôm nay", command=show_today_birthdays).grid(row=3, column=0, padx=5, pady=10, sticky="ew")
tk.Button(root, text="Xuất danh sách", command=export_to_excel).grid(row=4, column=0, padx=5, pady=10, sticky="ew")

root.mainloop()
