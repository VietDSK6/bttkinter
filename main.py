import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import csv
import datetime
import pandas as pd
from tkcalendar import DateEntry  # Import DateEntry for date picker


class EmployeeManager:
    def __init__(self, root):
        self.root = root
        self.root.title("Quản lý thông tin nhân viên")
        self.root.geometry("1200x550")

        # Initialize data storage
        self.data_file = "employees.csv"
        self.fields = ["Mã", "Tên", "Đơn vị", "Chức danh", "Ngày sinh", "Giới tính", "Số CMND", "Nơi cấp", "Ngày cấp",
                       "Khách hàng", "Nhà cung cấp"]

        # Create UI
        self.create_ui()

    def create_ui(self):
        # Main frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill="both", expand=True)

        # Labels and Entry widgets
        self.entries = {}

        tk.Label(main_frame, text="Mã *", anchor="w", width=15).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.entries["Mã"] = tk.Entry(main_frame, width=30)
        self.entries["Mã"].grid(row=0, column=1, padx=5, pady=5)

        tk.Label(main_frame, text="Tên *", anchor="w", width=15).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.entries["Tên"] = tk.Entry(main_frame, width=30)
        self.entries["Tên"].grid(row=0, column=3, padx=5, pady=5)

        tk.Label(main_frame, text="Đơn vị *", anchor="w", width=15).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.entries["Đơn vị"] = ttk.Combobox(main_frame, values=["Phân xưởng que hàn", "Văn phòng", "Kho vận"],
                                              width=28)
        self.entries["Đơn vị"].grid(row=1, column=1, padx=5, pady=5)

        tk.Label(main_frame, text="Chức danh", anchor="w", width=15).grid(row=1, column=2, padx=5, pady=5, sticky="w")
        self.entries["Chức danh"] = ttk.Combobox(main_frame, values=["Nhân viên", "Quản lý", "Giám đốc"], width=28)
        self.entries["Chức danh"].grid(row=1, column=3, padx=5, pady=5)

        tk.Label(main_frame, text="Ngày sinh", anchor="w", width=15).grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.entries["Ngày sinh"] = DateEntry(main_frame, width=28, date_pattern='dd/mm/yyyy')  # Date picker
        self.entries["Ngày sinh"].grid(row=2, column=1, padx=5, pady=5)

        tk.Label(main_frame, text="Giới tính", anchor="w", width=15).grid(row=2, column=2, padx=5, pady=5, sticky="w")
        self.gender_var = tk.StringVar(value="Nam")
        gender_frame = tk.Frame(main_frame)
        gender_frame.grid(row=2, column=3, sticky="w")
        tk.Radiobutton(gender_frame, text="Nam", variable=self.gender_var, value="Nam").pack(side="left")
        tk.Radiobutton(gender_frame, text="Nữ", variable=self.gender_var, value="Nữ").pack(side="left")

        tk.Label(main_frame, text="Số CMND", anchor="w", width=15).grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.entries["Số CMND"] = tk.Entry(main_frame, width=30)
        self.entries["Số CMND"].grid(row=3, column=1, padx=5, pady=5)

        tk.Label(main_frame, text="Nơi cấp", anchor="w", width=15).grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.entries["Nơi cấp"] = tk.Entry(main_frame, width=30)
        self.entries["Nơi cấp"].grid(row=3, column=3, padx=5, pady=5)

        tk.Label(main_frame, text="Ngày cấp", anchor="w", width=15).grid(row=4, column=0, padx=5, pady=5, sticky="w")
        self.entries["Ngày cấp"] = tk.Entry(main_frame, width=30)
        self.entries["Ngày cấp"].grid(row=4, column=1, padx=5, pady=5)

        # Checkboxes for customer and supplier
        self.is_customer = tk.IntVar()
        self.is_supplier = tk.IntVar()
        tk.Checkbutton(main_frame, text="Là khách hàng", variable=self.is_customer).grid(row=5, column=0, padx=5,
                                                                                         pady=5, sticky="w")
        tk.Checkbutton(main_frame, text="Là nhà cung cấp", variable=self.is_supplier).grid(row=5, column=1, padx=5,
                                                                                           pady=5, sticky="w")

        # Buttons
        button_frame = tk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=4, pady=15)

        tk.Button(button_frame, text="Lưu thông tin", command=self.save_data, width=20).pack(side="left", padx=10)
        tk.Button(button_frame, text="Sinh nhật hôm nay", command=self.check_today_birthday, width=20).pack(side="left",
                                                                                                            padx=10)
        tk.Button(button_frame, text="Xuất danh sách", command=self.export_data, width=20).pack(side="left", padx=10)

    def save_data(self):
        # Gather data from the form
        data = {field: self.entries[field].get() for field in self.entries}
        data["Giới tính"] = self.gender_var.get()
        data["Khách hàng"] = "Có" if self.is_customer.get() else "Không"
        data["Nhà cung cấp"] = "Có" if self.is_supplier.get() else "Không"

        # Validate required fields
        if not data["Mã"] or not data["Tên"] or not data["Đơn vị"]:
            messagebox.showerror("Lỗi", "Vui lòng điền đầy đủ các trường bắt buộc (*)")
            return

        # Save to CSV file
        try:
            file_exists = False
            try:
                with open(self.data_file, "r", encoding="utf-8") as file:
                    file_exists = True
            except FileNotFoundError:
                pass

            with open(self.data_file, "a", newline="", encoding="utf-8") as file:
                writer = csv.DictWriter(file, fieldnames=self.fields)
                if not file_exists:
                    writer.writeheader()
                writer.writerow(data)
            messagebox.showinfo("Thành công", "Lưu thông tin thành công!")

        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể lưu thông tin: {e}")

    def check_today_birthday(self):
        today = datetime.datetime.today().strftime("%d/%m/%Y")
        try:
            with open(self.data_file, "r", encoding="utf-8") as file:
                reader = csv.DictReader(file)
                results = [row for row in reader if row["Ngày sinh"] == today]

            if results:
                result_text = "\n".join([f"{row['Tên']} - {row['Đơn vị']}" for row in results])
                messagebox.showinfo("Sinh nhật hôm nay", result_text)
            else:
                messagebox.showinfo("Sinh nhật hôm nay", "Không có nhân viên nào sinh nhật hôm nay.")

        except FileNotFoundError:
            messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên.")

    def export_data(self):
        try:
            data = pd.read_csv(self.data_file)
            data['Tuổi'] = data['Ngày sinh'].apply(lambda x: self.calculate_age(x))
            sorted_data = data.sort_values(by='Tuổi', ascending=False)

            export_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[["Excel files", "*.xlsx"]])
            if export_file:
                sorted_data.to_excel(export_file, index=False)
                messagebox.showinfo("Thành công", f"Danh sách nhân viên đã được xuất ra {export_file}")

        except FileNotFoundError:
            messagebox.showerror("Lỗi", "Chưa có dữ liệu nhân viên.")

    def calculate_age(self, birthdate):
        try:
            birthdate = datetime.datetime.strptime(birthdate, "%d/%m/%Y")
            today = datetime.datetime.today()
            age = today.year - birthdate.year - ((today.month, today.day) < (birthdate.month, birthdate.day))
            return age
        except Exception:
            return 0


if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeManager(root)
    root.mainloop()
