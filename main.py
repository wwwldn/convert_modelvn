import os
import sys
import pandas as pd

# Xác định thư mục làm việc hiện tại (khi chạy dưới dạng .exe)
if getattr(sys, 'frozen', False):
    current_dir = os.path.dirname(sys.executable)  # Khi chạy dưới dạng .exe
else:
    current_dir = os.path.dirname(__file__)  # Khi chạy dưới dạng .py

file_import = os.path.join(current_dir, 'Import_Excel_DonHangBan.xlsx')
file_data = os.path.join(current_dir, 'Data.xlsx')

# Kiểm tra xem file có tồn tại hay không
if not os.path.exists(file_import):
    raise FileNotFoundError(f"❌ Không tìm thấy file '{file_import}'")
if not os.path.exists(file_data):
    raise FileNotFoundError(f"❌ Không tìm thấy file '{file_data}'")

# Đọc dữ liệu vào DataFrame
df_import = pd.read_excel(file_import, header=None)  # Đọc không lấy tiêu đề để giữ nguyên template
df_data = pd.read_excel(file_data)

# Kiểm tra cột tồn tại trong file Data
required_columns = ['ModelVN', 'InventoryID']
missing_columns = [col for col in required_columns if col not in df_data.columns]
if missing_columns:
    raise KeyError(f"❌ Không tìm thấy cột {missing_columns} trong file Data.")

# Chuẩn hóa dữ liệu trong file Data (loại bỏ khoảng trắng và ký tự thừa)
df_data['ModelVN'] = df_data['ModelVN'].astype(str).str.strip()
df_data['InventoryID'] = df_data['InventoryID'].astype(str).str.strip()

# Tạo từ điển ánh xạ từ 'ModelVN' → 'InventoryID'
mapping_dict = df_data.set_index('ModelVN')['InventoryID'].to_dict()

# Xử lý từ dòng thứ 10 và cột Y (cột 24 - chỉ số là 23)
start_row = 9
column_Y = 24

# Thay đổi giá trị từ dòng thứ 10 theo ánh xạ
count_updated = 0
count_not_found = 0
log_updated = []
log_not_found = []

for i in range(start_row, len(df_import)):
    value = str(df_import.iloc[i, column_Y]).strip()
    if value in mapping_dict:
        new_value = mapping_dict[value]
        log_updated.append(f"✅ Dòng {i + 1}: '{value}' → '{new_value}'")
        df_import.iloc[i, column_Y] = new_value
        count_updated += 1
    else:
        log_not_found.append(f"⚠️ Dòng {i + 1}: '{value}' không tìm thấy trong file Data.")
        count_not_found += 1

# Xuất ra file mới
output_file = os.path.join(current_dir, 'Import_Converted.xlsx')
df_import.to_excel(output_file, index=False, header=False)

# ✅ Kết quả
print(f"\n✅ Đã thay thế {count_updated} mã thành công.")
if log_updated:
    print("\n--- Chi tiết thay thế thành công ---")
    for log in log_updated:
        print(log)

if count_not_found > 0:
    print(f"\n⚠️ Có {count_not_found} mã không tìm thấy trong từ điển:")
    for log in log_not_found:
        print(log)

print(f"\n✅ File đã được xử lý và lưu thành '{output_file}'")
input("\nNhấn Enter để thoát...")
