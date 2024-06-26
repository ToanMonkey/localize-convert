import os
import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog
import math
# from tkinter import tk

output_folder_path = None  # Đảm bảo biến output_folder_path được định nghĩa trong phạm vi toàn cục
input_file_path = None  # Đảm bảo biến input_file_path được định nghĩa trong phạm vi toàn cục
checkboxes = []  # Đảm bảo biến checkboxes được định nghĩa trong phạm vi toàn cục
root = tk.Tk()
result_label = tk.Label(root, text="")
result_label.pack()
df = pd.DataFrame()

def excel_to_json(input_file, output_folder):
    # df = pd.read_excel(input_file)
    # print(df)

    languages = df.columns[1:]  # Lấy các cột từ cột thứ 2 trở đi là các ngôn ngữ

    # Tạo thư mục đích nếu nó không tồn tại
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for lang in languages:
        lang_data = {}
        for index, row in df.iterrows():
            key = row['key']
            lang_data[key] = row[lang]  # Sử dụng cột ngôn ngữ tương ứng

        # Xuất dữ liệu ra file JSON
        output_file = os.path.join(output_folder, f"localize-{lang}.json")
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(lang_data, f, ensure_ascii=False, indent=4)

def select_input_file():
    # global input_file_path
    input_file_path = filedialog.askopenfilename()
    print(input_file_path)
    display_data(input_file_path)

def select_output_folder():
    global output_folder_path
    output_folder_path = filedialog.askdirectory()
    # output_label.config(text=output_folder_path)

def convert():
    # global input_file_path
    # global output_folder_path
    if input_file_path != "" and output_folder_path != "":
        # df = pd.read_excel(input_file_path)

        excel_to_json(input_file_path, output_folder_path)
        result_label.config(text="Chuyển đổi hoàn thành!")
    else:
        result_label.config(text="Vui lòng chọn tệp đầu vào và thư mục đầu ra trước khi chuyển đổi.")

def display_data(input_file):
    # global checkboxes
    global df
    df = pd.read_excel(input_file).dropna()

    # Tạo cửa sổ giao diện
    data_frame = tk.Toplevel()
    data_frame.title("Dữ liệu từ tệp Excel")

    # Tạo canvas để chứa bảng dữ liệu
    canvas = tk.Canvas(data_frame, width=300, height=300,scrollregion=(0,0,500,500))
    canvas.pack(fill=tk.BOTH, expand=True)

            # Tạo thanh cuộn cho frame dọc
    y_scrollbar = tk.Scrollbar(data_frame, orient=tk.VERTICAL)
    y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # Tạo thanh cuộn cho frame ngang
    x_scrollbar = tk.Scrollbar(data_frame, orient=tk.HORIZONTAL)
    x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

    # Tạo frame để chứa bảng dữ liệu bên trong canvas
    table_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=table_frame)



    # Kết nối thanh cuộn với canvas
    y_scrollbar.config(command=canvas.yview)
    x_scrollbar.config(command=canvas.xview)

    # Kết nối canvas với thanh cuộn
    canvas.config(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set, width=300,height=300)

    # Khởi tạo danh sách checkboxes
    # checkboxes = []
    
    for index, row in df.iterrows():

        # print(index,row)
        check_var = tk.BooleanVar()
        check_var.set(False)
        data = ''
        for value in row:
            print(value)
            data += value + ' \t '

        checkbox = tk.Checkbutton(table_frame, variable=check_var, text=data)
        checkbox.grid(row=index, column=0)
        checkboxes.append(check_var)  # Thêm biến kiểm soát vào danh sách checkboxes

        # break

# Tạo cửa sổ giao diện
def select_all():
    # global checkboxes
    for checkbox in checkboxes:
        checkbox.set(True)

def deselect_all():
    # global checkboxes
    for checkbox in checkboxes:
        checkbox.set(False)

def main():
    # Tạo cửa sổ giao diện
    
    var = tk.IntVar()
    root.title("Excel to JSON Converter")

    # # Frame chứa các nút bấm
    button_frame = tk.Frame(root)
    button_frame.pack(side=tk.TOP, padx=10, pady=10)

    # Tạo checkbox để chọn tất cả
    select_all_checkbox = tk.Checkbutton(root, text="Chọn Tất Cả", command=select_all, variable=var)
    select_all_checkbox.pack(side=tk.LEFT)

    # Tạo checkbox để bỏ chọn tất cả
    deselect_all_checkbox = tk.Checkbutton(button_frame, text="Bỏ Chọn Tất Cả", command=deselect_all)
    deselect_all_checkbox.pack(side=tk.LEFT, padx=10)

    # Label và Button cho tệp đầu vào
    # input_label = tk.Label(button_frame, text="Chưa chọn tệp đầu vào")
    # input_label.pack(side=tk.LEFT)
    input_button = tk.Button(button_frame, text="Chọn Tệp Đầu Vào", command=select_input_file)
    input_button.pack(side=tk.LEFT, padx=10)

    # # Label và Button cho thư mục đầu ra
    # output_label = tk.Label(button_frame, text="Chưa chọn thư mục đầu ra")
    # output_label.pack(side=tk.LEFT)
    output_button = tk.Button(button_frame, text="Chọn Thư Mục Đầu Ra", command=select_output_folder)
    output_button.pack(side=tk.LEFT, padx=10)

    # # Button để chuyển đổi
    convert_button = tk.Button(button_frame, text="Chuyển Đổi", command=convert)
    convert_button.pack(side=tk.LEFT, padx=10)

    # # Label hiển thị kết quả
    

    # Khởi chạy giao diện
    root.mainloop()

if __name__ == '__main__':
    main()
