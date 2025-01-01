import json
import os

import pandas as pd
from openpyxl import load_workbook


def json_to_excel(json_file, excel_file):
    try:
        # Đọc dữ liệu từ tệp JSON
        with open(json_file, 'r', encoding='utf-8') as file:
            data = json.load(file)

        # Kiểm tra dữ liệu JSON có ở dạng danh sách hay không
        if isinstance(data, list):
            # Chuyển đổi dữ liệu thành DataFrame
            df = pd.DataFrame(data)
        else:
            # Nếu không phải danh sách, chuyển đổi từ dictionary
            df = pd.DataFrame([data])

        # Xuất dữ liệu ra tệp Excel
        df.to_excel(excel_file, index=False, engine='openpyxl')

        # Tự động điều chỉnh độ rộng cột
        wb = load_workbook(excel_file)
        ws = wb.active
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter  # Lấy ký tự cột
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            adjusted_width = max_length + 2
            ws.column_dimensions[column_letter].width = adjusted_width
        wb.save(excel_file)

        print(f"Dữ liệu đã được chuyển từ {json_file} sang {excel_file} thành công.")
    except Exception as e:
        print(f"Đã xảy ra lỗi: {e}")


def convert_all_json_to_excel(input_folder, output_folder):
    try:
        # Tạo thư mục đầu ra nếu chưa tồn tại
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Lặp qua tất cả các tệp trong thư mục đầu vào
        for file_name in os.listdir(input_folder):
            if file_name.endswith('.json'):
                json_file = os.path.join(input_folder, file_name)
                excel_file = os.path.join(output_folder, f"{os.path.splitext(file_name)[0]}.xlsx")

                # Chuyển đổi tệp JSON sang Excel
                json_to_excel(json_file, excel_file)
    except Exception as e:
        print(f"Đã xảy ra lỗi khi chuyển đổi thư mục: {e}")


# Ví dụ sử dụng
input_folder = 'data'  # Thư mục chứa các tệp JSON
output_folder = 'Output'  # Thư mục lưu các tệp Excel
convert_all_json_to_excel(input_folder, output_folder)
