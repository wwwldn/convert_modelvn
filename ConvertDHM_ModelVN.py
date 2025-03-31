import os
import sys
import pandas as pd

def show_info():
    print("="*40)
    print("Liên hệ hỗ trợ: 0858.66.88.44")
    print("="*40)

# Hiển thị thông tin khi khởi động
show_info()

# Xác định thư mục làm việc hiện tại
if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(sys.executable)
else:
    current_dir = os.path.dirname(__file__)

file_import = os.path.join(current_dir, 'Import_Excel_DonHangMua_GREE.xlsx')
file_data = os.path.join(current_dir, 'Data.xlsx')

if not os.path.exists(file_import):
    raise FileNotFoundError(f"❌ Không tìm thấy file '{file_import}'")
if not os.path.exists(file_data):
    raise FileNotFoundError(f"❌ Không tìm thấy file '{file_data}'")

# Đọc dữ liệu vào DataFrame (chỉ định engine để tránh lỗi)
df_import = pd.read_excel(file_import, header=None, engine='openpyxl')
df_data = pd.read_excel(file_data, engine='openpyxl')

# Kiểm tra cột trong file Data
required_columns = ['ModelVN', 'InventoryID']
missing_columns = [col for col in required_columns if col not in df_data.columns]
if missing_columns:
    raise KeyError(f"❌ Không tìm thấy cột {missing_columns} trong file Data.")

# Chuẩn hóa dữ liệu trong file Data
df_data['ModelVN'] = df_data['ModelVN'].astype(str).str.strip()
df_data['InventoryID'] = df_data['InventoryID'].astype(str).str.strip()

# Tạo từ điển ánh xạ (loại bỏ giá trị nan)
mapping_dict = df_data.set_index('ModelVN')['InventoryID'].to_dict()
mapping_dict = {k: (v if pd.notna(v) and v != 'nan' else '') for k, v in mapping_dict.items()}

# 🔥 Xử lý từ dòng thứ 10 và cột Y (cột 30 → chỉ số là 29)
start_row = 9
column_Y = 29

if column_Y >= df_import.shape[1]:
    raise IndexError(f"❌ Không tìm thấy cột thứ {column_Y + 1} trong file Import (chỉ có {df_import.shape[1]} cột).")

# 👉 Chuyển số cột thành tên cột Excel
def get_excel_column_name(col):
    name = ''
    while col >= 0:
        name = chr((col % 26) + 65) + name
        col = (col // 26) - 1
    return name

# column_Y = 29 → 'AD'
column_excel = get_excel_column_name(column_Y)

# 📝 Tính toán số lượng dòng cần xử lý
total_rows = len(df_import) - start_row

# ✅ THÔNG BÁO TRẠNG THÁI TRƯỚC KHI XỬ LÝ
print("\n📌 Tổng số hàng từ dòng 10:", total_rows)
print(f"📌 Đang xử lý từ cột: '{column_excel}'")

# Đếm số lượng thành công và thất bại
count_updated = 0
count_failed = 0
log_updated = []
log_failed = []

# 🚀 Thay thế giá trị theo ánh xạ
for i in range(start_row, len(df_import)):
    value = str(df_import.iloc[i, column_Y]).strip()

    # Bỏ qua nếu giá trị trống
    if value == '' or value.lower() == 'nan':
        continue

    if value in mapping_dict:
        new_value = mapping_dict[value]

        # Nếu giá trị ánh xạ là rỗng → Ghi là 'FAIL'
        if new_value == '':
            df_import.iloc[i, column_Y] = 'FAIL'
            count_failed += 1
            log_failed.append(f"❌ Dòng {i + 1}: '{value}' → 'FAIL'")
        else:
            df_import.iloc[i, column_Y] = new_value
            count_updated += 1
            log_updated.append(f"✅ Dòng {i + 1}: '{value}' → '{new_value}'")
    else:
        # Nếu không tìm thấy → Ghi là 'FAIL'
        df_import.iloc[i, column_Y] = 'FAIL'
        count_failed += 1
        log_failed.append(f"❌ Dòng {i + 1}: '{value}' → 'FAIL' (Không tìm thấy trong file Data)")

# ✅ Xóa file cũ nếu đã tồn tại để tránh lỗi ghi đè
output_file = os.path.join(current_dir, 'Import_Converted.xlsx')
if os.path.exists(output_file):
    try:
        os.remove(output_file)
    except PermissionError:
        input(f"❌ File '{output_file}' đang được mở. Hãy đóng file và nhấn Enter để thử lại...")
        sys.exit(1)


# 💾 Lưu file mới (đóng đúng cách)
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    df_import.to_excel(writer, index=False, header=False)

# ✅ Kết quả
print(f"\n✅ Đã thay thế {count_updated}/{total_rows} mã thành công.")
print(f"❌ Có {count_failed}/{total_rows} mã không tìm thấy hoặc không có giá trị ánh xạ.")

if log_updated:
    print("\n--- Chi tiết thay thế thành công ---")
    for log in log_updated:
        print(log)

if log_failed:
    print("\n--- Các mã không tìm thấy hoặc giá trị trống ---")
    for log in log_failed:
        print(log)

print(f"\n✅ File đã được xử lý và lưu thành '{output_file}'")
input("\nNhấn Enter để thoát...")
