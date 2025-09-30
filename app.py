from colorama import Fore, Back, init
import os
from openpyxl import Workbook, load_workbook
import pandas as pd
from pathlib import Path
import time
# Web dependencies (optional, only used when RUN_WEB=1)
try:
    from flask import Flask, render_template, request
except Exception:
    Flask = None
    render_template = None

def ID_kh():
    wb=load_workbook(filename="ThongTinKhachHang.xlsx")
    sheet=wb.active
    
    #Rut ra danh sach ma khach hang
    ma_kh_list = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is not None:
            ma_kh_list.append(row[0])
            
    #tim ma khach hang lon nhat
    if not ma_kh_list:
        ID_kh = "DLT" + "0"*4 + "1"
        return ID_kh
    else:
        max_id = 0
        for ma_kh in ma_kh_list:
            try:
                num_part = int(ma_kh[3:])  # ma khach hang co dang DLT00001
                if num_part > max_id:
                    max_id = num_part
            except ValueError:
                continue  # Bỏ qua các mã không hợp lệ
        new_id_num = max_id + 1
        new_id = "DLU" + str(new_id_num).zfill(5)  # id luon co 5 chu so
        return new_id

def is_data_none():
    wb=load_workbook(filename="ThongTinKhachHang.xlsx")
    sheet=wb.active
    
    for row in sheet.iter_rows(min_row=2, max_row=2, min_col=1, values_only=True):
        if all(cell is None for cell in row):
            return True
        else:
            return False
        
def is_recycle_bin():
    # Dùng Path join để tránh lỗi escape chuỗi trên Windows
    file_path = Path("Recycle Bin") / "ThongTinKhachHang.xlsx"
    if file_path.exists() and file_path.is_file():
        return True

def check_file():
    if is_recycle_bin() == True: #False
        return False
    elif os.path.exists("ThongTinKhachHang.xlsx") == False:
        return False
    elif os.path.exists("ThongTinKhachHang.xlsx") and is_data_none() == True:
        return False
    else:
        return True
        

def show_customer_information():
    # Load excel file
    wb=load_workbook(filename="ThongTinKhachHang.xlsx")
    sheet=wb.active

    #create header from first row
    header = [cell.value for cell in sheet[1]]

    #create list of dictionaries from excel file
    data_rows = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data_rows.append(dict(zip(header, row)))
    #create dataframe from excel file
    df = pd.DataFrame(data_rows)
    return print(df)

def create_file():
    wb = Workbook()
    ws = wb.active
    ws['A1'] = 'Mã KH'
    ws['B1'] = 'Họ Tên'
    ws['C1'] = 'Số ĐT'
    ws['D1'] = 'Email'
    ws['E1'] = 'Địa Chỉ'
    
    wb.save("ThongTinKhachHang.xlsx")
    
def add_customer():
    new_customer = []
    #load excel file
    wb=load_workbook(filename="ThongTinKhachHang.xlsx")
    sheet=wb.active
    
    #get next row
    next_row = sheet.max_row + 1
    print(Fore.CYAN + Back.BLACK + "\n ----------THÊM KHÁCH HÀNG----------")
    ma_kh = ID_kh()
    ho_ten = input("Nhập họ tên khách hàng: ")
    so_dt = input("Nhập số điện thoại khách hàng: ")
    dia_chi = input("Nhập địa chỉ khách hàng: ")
    email = input("Nhập email khách hàng: ")
    new_customer.append(ma_kh)
    new_customer.append(ho_ten)
    new_customer.append(so_dt)
    new_customer.append(email)
    new_customer.append(dia_chi)
    Ten_truong = ["Mã KH", "Họ Tên", "Số ĐT", "Email", "Địa Chỉ"]
    
    while True:
        for i in range(len(new_customer)):
            if new_customer[i] == "":
                print(Fore.RED + Back.BLACK + f"Trường '{Ten_truong[i]}' không được để trống. Vui lòng nhập lại.")
                if Ten_truong[i] == "Họ Tên":
                    ho_ten = input("Nhập họ tên khách hàng: ")
                    new_customer[i] = ho_ten
                elif Ten_truong[i] == "Số ĐT":
                    so_dt = input("Nhập số điện thoại khách hàng: ")
                    new_customer[i] = so_dt
                elif Ten_truong[i] == "Email":
                    email = input("Nhập email khách hàng: ")
                    new_customer[i] = email
                elif Ten_truong[i] == "Địa Chỉ":
                    dia_chi = input("Nhập địa chỉ khách hàng: ")
                    new_customer[i] = dia_chi
        if all(value != "" for value in new_customer):
            break
    for col, value in enumerate(new_customer, start=1):
        sheet.cell(row=next_row, column=col, value=str(value))
            
    wb.save("ThongTinKhachHang.xlsx")
    
    return print(Fore.GREEN + Back.BLACK + "Thêm khách hàng thành công.")
    
#########################
# Flask web integration #
#########################

def load_first_customer_for_web() -> dict | None:
    try:
        excel_path = Path("ThongTinKhachHang.xlsx")
        if not excel_path.exists():
            return None
        wb = load_workbook(filename=str(excel_path))
        sheet = wb.active
        header = [cell.value for cell in sheet[1]]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row and any(cell is not None for cell in row):
                return dict(zip(header, row))
        return None
    except Exception:
        return None

def load_all_customers_for_web() -> list:
    """Load tất cả khách hàng từ Excel và phân nhóm"""
    try:
        excel_path = Path("ThongTinKhachHang.xlsx")
        if not excel_path.exists():
            return []
        
        wb = load_workbook(filename=str(excel_path))
        sheet = wb.active
        header = [cell.value for cell in sheet[1]]
        
        customers = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row and any(cell is not None for cell in row):
                raw_data = dict(zip(header, row))
                
                # Chuẩn hóa dữ liệu
                customer = {
                    "code": raw_data.get("Mã KH") or raw_data.get("Ma KH") or raw_data.get("Code") or "",
                    "name": raw_data.get("Họ Tên") or raw_data.get("Ho Ten") or raw_data.get("Ten") or "",
                    "phone": raw_data.get("Số ĐT") or raw_data.get("So DT") or raw_data.get("SĐT") or "",
                    "email": raw_data.get("Email") or "",
                    "address": raw_data.get("Địa Chỉ") or raw_data.get("Dia Chi") or "",
                    "total_amount": raw_data.get("Tổng tiền mua") or raw_data.get("Tong Tien Mua") or "0đ",
                    "last_purchase": raw_data.get("Ngày cuối mua") or raw_data.get("Ngay Cuoi Mua") or "",
                }
                
                # Phân nhóm dựa trên tổng tiền mua
                total_amount_str = str(customer["total_amount"]).replace("đ", "").replace(",", "")
                try:
                    total_amount = float(total_amount_str) if total_amount_str else 0
                except ValueError:
                    total_amount = 0
                
                if total_amount >= 10000000:  # >= 10 triệu
                    customer["group"] = "vip"
                elif total_amount >= 5000000:  # >= 5 triệu
                    customer["group"] = "loyal"
                else:
                    customer["group"] = "potential"
                
                # Phân trạng thái Active/Inactive dựa trên lần cuối mua hàng
                last_purchase = customer["last_purchase"]
                if last_purchase and last_purchase != "Chưa có":
                    # Giả sử nếu có mua hàng trong 6 tháng gần đây thì Active
                    customer["status"] = "active"
                else:
                    customer["status"] = "inactive"
                
                customers.append(customer)
        
        # Sắp xếp theo alphabet dựa trên chữ cái cuối của tên
        customers.sort(key=lambda x: x["name"][-1].lower() if x["name"] else "z")
        
        return customers
    except Exception:
        return []

def load_customer_by_code_for_web(code: str) -> dict | None:
    """Tải 1 khách hàng theo mã từ Excel (phục vụ trang web)."""
    try:
        if not code:
            return None
        excel_path = Path("ThongTinKhachHang.xlsx")
        if not excel_path.exists():
            return None

        wb = load_workbook(filename=str(excel_path))
        sheet = wb.active
        header = [cell.value for cell in sheet[1]]

        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or all(cell is None for cell in row):
                continue
            raw = dict(zip(header, row))
            row_code = (
                raw.get("Mã KH")
                or raw.get("Ma KH")
                or raw.get("Code")
                or ""
            )
            if str(row_code) == str(code):
                customer = {
                    "name": raw.get("Họ Tên") or raw.get("Ho Ten") or raw.get("Ten") or "",
                    "dob": raw.get("Ngày sinh") or raw.get("Ngay Sinh") or "",
                    "phone": raw.get("Số ĐT") or raw.get("So DT") or raw.get("SĐT") or raw.get("Phone") or "",
                    "email": raw.get("Email") or "",
                    "address": raw.get("Địa Chỉ") or raw.get("Dia Chi") or raw.get("Address") or "",
                    "code": row_code,
                    "total": raw.get("Tổng tiền mua") or raw.get("Tong Tien Mua") or "0đ",
                    "last_purchase": raw.get("Ngày cuối mua") or raw.get("Ngay Cuoi Mua") or "",
                }

                # Xác định group/status giống logic danh sách
                total_amount_str = str(customer["total"]).replace("đ", "").replace(",", "")
                try:
                    total_amount = float(total_amount_str) if total_amount_str else 0
                except ValueError:
                    total_amount = 0

                if total_amount >= 10000000:
                    customer["group"] = "vip"
                elif total_amount >= 5000000:
                    customer["group"] = "loyal"
                else:
                    customer["group"] = "potential"

                last_purchase = customer["last_purchase"]
                customer["status"] = "active" if last_purchase and last_purchase != "Chưa có" else "inactive"

                return customer
        return None
    except Exception:
        return None

def get_customer_stats(customers: list) -> dict:
    """Tính toán thống kê khách hàng"""
    active_customers = len([c for c in customers if c["status"] == "active"])
    
    # Tính tổng tiền của khách hàng Active
    active_amount = 0
    for customer in customers:
        if customer["status"] == "active":
            total_amount_str = str(customer["total_amount"]).replace("đ", "").replace(",", "")
            try:
                active_amount += float(total_amount_str) if total_amount_str else 0
            except ValueError:
                pass
    
    # Giả sử số đơn hàng hoạt động = số khách hàng active * 2 (trung bình)
    active_orders = active_customers * 2
    
    stats = {
        "total_customers": len(customers),
        "active_customers": active_customers,
        "active_orders": active_orders,
        "active_amount": f"{active_amount:,.0f}đ" if active_amount > 0 else "0đ",
    }
    return stats

# Create Flask app only if Flask is available
app = None
if Flask:
    base_dir = Path(__file__).resolve().parent
    template_dir = base_dir / "templates"
    static_dir = base_dir / "static"
    app = Flask(
        __name__,
        template_folder=str(template_dir),
        static_folder=str(static_dir),
    )

if app:
    @app.route("/customer-dashboard")
    def customer_dashboard():
        code = request.args.get("code") if request else None
        customer = load_customer_by_code_for_web(code) if code else None
        if not customer:
            raw = load_first_customer_for_web()
            if raw:
                customer = {
                    "name": raw.get("Họ Tên") or raw.get("Ho Ten") or raw.get("Ten") or "",
                    "dob": raw.get("Ngày sinh") or raw.get("Ngay Sinh") or "",
                    "phone": raw.get("Số ĐT") or raw.get("So DT") or raw.get("SĐT") or raw.get("Phone") or "",
                    "email": raw.get("Email") or "",
                    "address": raw.get("Địa Chỉ") or raw.get("Dia Chi") or raw.get("Address") or "",
                    "code": raw.get("Mã KH") or raw.get("Ma KH") or raw.get("Code") or "",
                    "total": raw.get("Tổng tiền mua") or raw.get("Tong Tien Mua") or "0đ",
                    "last_purchase": raw.get("Ngày cuối mua") or raw.get("Ngay Cuoi Mua") or "",
                }

                # Bổ sung group/status để đồng bộ với UI
                total_amount_str = str(customer["total"]).replace("đ", "").replace(",", "")
                try:
                    total_amount = float(total_amount_str) if total_amount_str else 0
                except ValueError:
                    total_amount = 0
                if total_amount >= 10000000:
                    customer["group"] = "vip"
                elif total_amount >= 5000000:
                    customer["group"] = "loyal"
                else:
                    customer["group"] = "potential"
                last_purchase = customer["last_purchase"]
                customer["status"] = "active" if last_purchase and last_purchase != "Chưa có" else "inactive"

        sample_orders = [
            {
                "code": "DLU00001",
                "customer": (customer["name"] if customer and customer.get("name") else "Nguyễn Phước Lộc"),
                "export_status": "Đã xuất",
                "value": "1,500,000đ",
                "date": "28/09/2025",
                "status": "done",
                "status_label": "Hoàn thành",
            },
            {
                "code": "DLU00002",
                "customer": (customer["name"] if customer and customer.get("name") else "Nguyễn Phước Lộc"),
                "export_status": "Chưa xuất",
                "value": "2,000,000đ",
                "date": "29/09/2025",
                "status": "processing",
                "status_label": "Đang xử lý",
            },
        ]

        return render_template(
            "customer_dashboard_jinja.html",
            active="customers",
            user_name="Admin",
            customer=customer,
            orders=sample_orders,
        )

    @app.route("/customers-list")
    def customers_list():
        customers = load_all_customers_for_web()
        stats = get_customer_stats(customers)
        
        return render_template(
            "customer_list.html",
            active="customers",
            customers=customers,
            stats=stats,
        )

##########################
# HAM CHINH CHUONG TRINH #
##########################
def main():
    
    init(autoreset=False)
    print(Fore.YELLOW + Back.BLACK + "\n ----------QUẢN LÝ KHÁCH HÀNG----------")
    print("1. Thêm khách hàng")
    print("2. Hiển thị khách hàng")
    print("3. Tìm kiếm khách hàng")
    print("4. Cập nhật khách hàng")
    print("5. Xóa khách hàng")
    print("0. Thoát")
    print(Fore.CYAN + Back.BLACK)
    choice = input("Nhập lựa chọn của bạn: ")
    
    if choice == '1':
        if not os.path.exists("ThongTinKhachHang.xlsx"):
            print(Fore.YELLOW + Back.BLACK + "Chưa có file dữ liệu khách hàng. Đang tạo file...")
            time.sleep(1)
            create_file()
            print(Fore.GREEN + Back.BLACK + "Tạo file thành công.")
        elif os.path.exists("ThongTinKhachHang.xlsx") and is_data_none() == True:
            print(Fore.YELLOW + Back.BLACK + "File dữ liệu khách hàng hiện tại đang trống. Vui lòng thêm khách hàng.")
            add_customer()
        else:
            add_customer()
    elif choice == '2':
        if check_file() == False:
            print(Fore.RED + Back.BLACK + "Chưa có dữ liệu khách hàng. Vui lòng thêm khách hàng trước!!!")
        else:
            show_customer_information()
    elif choice == '3':
        if not os.path.exists("ThongTinKhachHang.xlsx"):
            print(Fore.YELLOW + Back.BLACK + "Chưa có file dữ liệu khách hàng. Đang tạo file...")
            time.sleep(1)
            create_file()
            print(Fore.GREEN + Back.BLACK + "Tạo file thành công.")
        elif os.path.exists("ThongTinKhachHang.xlsx") and is_data_none() == True:
            print(Fore.YELLOW + Back.BLACK + "File dữ liệu khách hàng hiện tại đang trống. Vui lòng thêm khách hàng.")
    elif choice == '4':
        pass
    elif choice == '5':
        pass
    elif choice == '0':
        print(Fore.WHITE + Back.BLACK + "Thoát chương trình.")
        exit()
    else: print(Fore.RED + Back.BLACK + "Lựa chọn không hợp lệ. Vui lòng thử lại!!!")
    
if __name__ == "__main__":
    # Set RUN_WEB=1 to start Flask server; otherwise run CLI loop as before
    if os.getenv("RUN_WEB") == "1" and app is not None:
        app.run(debug=True)
    else:
        while True:
            main()