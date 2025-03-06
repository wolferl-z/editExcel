import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import platform
from decimal import Decimal, ROUND_UP


# 处理 Excel 文件的函数
def process_excel(file_path, word_to_add, series_name, postage_fee):
    try:
        # 读取 Excel 文件
        df = read_excel(file_path)
        # 修改列名
        df = modify_columns(df, word_to_add)
        # 更新数据内容
        df = update_data(df)
        # 处理前三列数据
        df_first_three_columns = process_first_three_columns(df, series_name, word_to_add)
        # 处理邮费
        handle_postage_fee(df_first_three_columns, postage_fee)

        # 生成文件名
        output_file_name = generate_output_filename(series_name, word_to_add)
        # 保存并打开文件
        save_and_open_file(df_first_three_columns, output_file_name)
        messagebox.showinfo("成功", "处理已完成！生成的文件已保存并打开。")

    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")


# 读取 Excel 文件并返回 DataFrame
def read_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    df.columns = df.iloc[1]  # 将第二行作为列名
    df = df.drop(index=1)  # 删除第二行
    return df


# 修改列名的函数，添加指定的词语
def modify_columns(df, word_to_add):
    new_columns = []
    for i, col in enumerate(df.columns):
        # 从第三列开始修改列名
        if i >= 2 and word_to_add:
            new_columns.append(col + word_to_add)
        else:
            new_columns.append(col)
    df.columns = new_columns
    return df


# 更新数据内容的函数，给数据添加乘法运算
def update_data(df):
    for row_idx in range(3, len(df)):
        for col_idx in range(2, len(df.columns)):
            if df.iloc[row_idx, col_idx] > 0:
                df.iloc[row_idx, col_idx] = str(df.columns[col_idx]) + "*" + str(df.iloc[row_idx, col_idx])
    # 将“制品”列加入并更新
    df['制品'] = df.iloc[3:, 2:].apply(lambda row: ','.join(row.dropna().astype(str)), axis=1)
    df.insert(2, '制品', df.pop('制品'))  # 将制品列插入到第二列位置
    return df


# 处理前三列数据并更新列名
def process_first_three_columns(df, series_name, word_to_add):
    df_first_three_columns = df.iloc[3:, :3]  # 提取前三列数据
    df_first_three_columns.columns = ['金额', 'cn', '制品']  # 修改列名
    header_row = pd.DataFrame([[series_name + (word_to_add if word_to_add else ""), "", ""]], columns=df_first_three_columns.columns)
    df_first_three_columns = pd.concat([header_row, df_first_three_columns], ignore_index=True)  # 插入系列名为第一行
    return df_first_three_columns


# 处理邮费
def handle_postage_fee(df_first_three_columns, postage_fee):
    # 检查邮费是否大于 0
    if postage_fee > 0:
        # 计算有效的单元格数量（第二列从第三行开始的非空单元格）
        valid_cells_count = df_first_three_columns.iloc[1:, 1].count()  # 第二列从第二行开始的非空单元格数量

        if valid_cells_count > 0:
            # 计算每个有效行应分配的平均邮费
            average_postage = Decimal(float(postage_fee)) / Decimal(float(valid_cells_count))
            average_postage = average_postage.quantize(Decimal('0.01'), rounding=ROUND_UP)

            # 将邮费信息写入表格的第一行
            df_first_three_columns.iloc[0, 1] = f"邮{postage_fee}，{average_postage}/人"

            # 遍历每一行数据，将邮费加到金额列中
            for row_idx in range(1, len(df_first_three_columns)):
                original_value = df_first_three_columns.iloc[row_idx, 0]
                if isinstance(original_value, (int, float)):
                    # 计算最终金额 = 原金额 + 平均邮费
                    final_amount = Decimal(original_value) + average_postage
                    # 将计算后的金额写回到表格中，保留两位小数，向上取整
                    df_first_three_columns.iloc[row_idx, 0] = final_amount.quantize(Decimal('0.01'), rounding=ROUND_UP)


# 生成输出文件名
def generate_output_filename(series_name, word_to_add):
    return f"{series_name}{word_to_add if word_to_add else ''}.xlsx"


# 保存文件并打开
def save_and_open_file(df_first_three_columns, output_file_name):
    df_first_three_columns.to_excel(output_file_name, index=False, header=True)
    print(f"包含前三列的数据已保存为 {output_file_name}")
    open_generated_file(output_file_name)


# 打开生成的文件
def open_generated_file(file_path):
    if platform.system() == "Windows":
        os.startfile(file_path)  # Windows 中打开文件
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", file_path])
    else:  # Linux
        subprocess.run(["xdg-open", file_path])


# 选择单个文件的函数
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file_entry.delete(0, tk.END)  # 清空现有路径
    file_entry.insert(0, file_path)  # 显示选择的文件路径


# 运行处理单个表格的函数
def run_processing():
    file_path = file_entry.get()
    word_to_add = word_entry.get()
    series_name = series_entry.get()

    # 尝试获取邮费，如果为空则设置为0
    postage_fee_input = postage_entry.get().strip()
    try:
        postage_fee = float(postage_fee_input) if postage_fee_input else 0
    except ValueError:
        messagebox.showwarning("输入错误", "邮费必须是一个数字！")
        return

    if not file_path or not series_name:
        messagebox.showwarning("输入错误", "请确保输入文件路径和系列名！")
        return

    process_excel(file_path, word_to_add, series_name, postage_fee)


# 合并多个表格的函数
def merge_multiple_files(file_paths, series_name, postage_fee):
    try:
        merged_df = pd.DataFrame()  # 创建一个空的 DataFrame 用于存储合并后的数据

        for file_path in file_paths:
            # 对每个文件进行单表处理
            word_to_add = ""  # 如果需要添加制品类型，可以在这里设置
            df_processed = process_excel_and_return_df(file_path, word_to_add, series_name)
            df_processed = df_processed.drop(index=0)  # 删除第一行
            # 将处理后的数据添加到 merged_df
            merged_df = pd.concat([merged_df, df_processed], ignore_index=True)

        # 检查是否有重复的 cn，如果有，则合并金额和制品
        merged_df_grouped = merged_df.groupby('cn').agg(
            {'金额': 'sum', '制品': lambda x: ','.join(x)}
        ).reset_index()

        # 交换金额和 cn 列的位置
        merged_df_grouped = merged_df_grouped[['金额', 'cn', '制品']]  # 交换金额和 cn 的位置

        # 插入系列名为第一行
        merged_df_grouped.loc[-1] = [series_name, '', '']  # 插入系列名
        merged_df_grouped.index = merged_df_grouped.index + 1  # 调整索引
        merged_df_grouped = merged_df_grouped.sort_index()  # 排序

        # 处理邮费
        if postage_fee > 0:
            handle_postage_fee(merged_df_grouped, postage_fee)

        # 生成输出文件名
        output_file_name = f"{series_name}.xlsx"

        # 保存合并后的数据
        merged_df_grouped.to_excel(output_file_name, index=False, header=True)
        open_generated_file(output_file_name)
        messagebox.showinfo("成功", "多表合并已完成！生成的文件已保存并打开。")

    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")

# 辅助函数：处理单个文件并返回 DataFrame
def process_excel_and_return_df(file_path, word_to_add, series_name):
    try:
        # 读取 Excel 文件
        df = read_excel(file_path)
        # 修改列名
        df = modify_columns(df, word_to_add)
        # 更新数据内容
        df = update_data(df)
        # 处理前三列数据
        df_first_three_columns = process_first_three_columns(df, series_name, word_to_add)

        return df_first_three_columns

    except Exception as e:
        messagebox.showerror("错误", f"处理文件 {file_path} 时发生错误: {e}")
        return pd.DataFrame()  # 返回空的 DataFrame

# 打开生成的文件
def open_generated_file(file_path):
    if platform.system() == "Windows":
        os.startfile(file_path)  # Windows 中打开文件
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", file_path])
    else:  # Linux
        subprocess.run(["xdg-open", file_path])


# 浏览多个文件的函数
def browse_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file_list.delete(0, tk.END)  # 清空现有文件列表
    for file_path in file_paths:
        file_list.insert(tk.END, file_path)  # 将每个选择的文件路径添加到 Listbox 中


# 运行合并的函数
def run_merge():
    file_paths = file_list.get(0, tk.END)  # 获取所有项
    series_name = series_entry_multi.get().strip()  # 获取系列名
    postage_fee_input = postage_entry_multi.get().strip()  # 获取邮费输入并去除空白

    # 如果邮费框为空，则将邮费设置为 0
    try:
        postage_fee = float(postage_fee_input) if postage_fee_input else 0
    except ValueError:
        messagebox.showwarning("输入错误", "邮费必须是一个数字！")
        return

    # 检查输入路径和系列名是否为空
    if not file_paths or file_paths == [''] or not series_name:
        messagebox.showwarning("输入错误", "请确保输入文件路径和系列名！")
        return

    # 调用多表合并函数并传递邮费参数
    merge_multiple_files(file_paths, series_name, postage_fee)


# 显示单表处理界面
def show_single_table_processing():
    single_table_frame.pack(fill="both", expand=True)
    multi_table_frame.pack_forget()


# 显示多表合并界面
def show_multi_table_processing():
    single_table_frame.pack_forget()
    multi_table_frame.pack(fill="both", expand=True)


def open_new_window():
    # 创建新窗口
    new_window = tk.Toplevel(root)
    new_window.title("使用说明")

    # 设置新窗口的大小
    new_window_width = 400
    new_window_height = 300

    # 获取屏幕宽度和高度
    screen_width = new_window.winfo_screenwidth()
    screen_height = new_window.winfo_screenheight()

    # 计算窗口位置，使其居中
    position_top = (screen_height - new_window_height) // 2
    position_right = (screen_width - new_window_width) // 2

    # 设置窗口大小和位置
    new_window.geometry(f'{new_window_width}x{new_window_height}+{position_right}+{position_top}')

    # 加载图片
    try:
        image = Image.open("logo.png")  # 替换为你的图片路径
        image = image.resize((150, 150), Image.ANTIALIAS)  # 调整图片大小
        photo = ImageTk.PhotoImage(image)

        # 创建 Label 用于显示图片
        image_label = tk.Label(new_window, image=photo)
        image_label.image = photo  # 保持图片引用，避免被垃圾回收
        image_label.pack(pady=10)
    except Exception as e:
        # 如果图片加载失败，显示错误信息
        error_label = tk.Label(new_window, text="图片加载失败", fg="red")
        error_label.pack(pady=10)

    # 添加作者信息
    author_label = tk.Label(new_window, text="作者：坐标系\n联系方式：2194970133@qq.com")
    author_label.pack(pady=10)

# 创建主窗口
root = tk.Tk()
root.title("打表工具")

# 获取屏幕宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 设置窗口的大小
window_width = 500
window_height = 350

# 计算窗口位置，使其居中
position_top = (screen_height - window_height) // 2
position_right = (screen_width - window_width) // 2

# 设置窗口大小和位置
root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

# 单表处理界面
single_table_frame = tk.Frame(root)

# 文件路径输入框及浏览按钮
file_label = tk.Label(single_table_frame, text="选择 Excel 文件:")
file_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

file_entry = tk.Entry(single_table_frame, width=40)
file_entry.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(single_table_frame, text="浏览", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# 系列名输入框
series_label = tk.Label(single_table_frame, text="输入系列名:")
series_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

series_entry = tk.Entry(single_table_frame, width=40)
series_entry.grid(row=1, column=1, padx=10, pady=10)

# 制品类型输入框（允许为空）
word_label = tk.Label(single_table_frame, text="输入制品类型:")
word_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

word_entry = tk.Entry(single_table_frame, width=40)
word_entry.grid(row=2, column=1, padx=10, pady=10)

# 邮费输入框（允许为空）
postage_label = tk.Label(single_table_frame, text="输入邮费:")
postage_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")

postage_entry = tk.Entry(single_table_frame, width=40)
postage_entry.grid(row=3, column=1, padx=10, pady=10)

# 运行按钮
run_button = tk.Button(single_table_frame, text="开始处理", command=run_processing)
run_button.grid(row=4, column=0, columnspan=3, pady=20)

# 多表合并界面
multi_table_frame = tk.Frame(root)

# 文件路径输入框及浏览按钮
file_label = tk.Label(multi_table_frame, text="选择多个 Excel 文件:")
file_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

file_list = tk.Listbox(multi_table_frame, height=4, width=40)
file_list.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(multi_table_frame, text="浏览", command=browse_files)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# 系列名输入框
series_label_multi = tk.Label(multi_table_frame, text="输入系列名:")
series_label_multi.grid(row=1, column=0, padx=10, pady=10, sticky="w")

series_entry_multi = tk.Entry(multi_table_frame, width=40)
series_entry_multi.grid(row=1, column=1, padx=10, pady=10)

# 邮费输入框
postage_label_multi = tk.Label(multi_table_frame, text="输入邮费:")
postage_label_multi.grid(row=2, column=0, padx=10, pady=10, sticky="w")

postage_entry_multi = tk.Entry(multi_table_frame, width=40)
postage_entry_multi.grid(row=2, column=1, padx=10, pady=10)

# 运行合并按钮
merge_button = tk.Button(multi_table_frame, text="开始合并", command=run_merge)
merge_button.grid(row=3, column=0, columnspan=3, pady=20)

# 切换按钮
switch_button_frame = tk.Frame(root)
single_button = tk.Button(switch_button_frame, text="单表处理", command=show_single_table_processing)
single_button.grid(row=0, column=0, padx=10, pady=10)
multi_button = tk.Button(switch_button_frame, text="多表合并", command=show_multi_table_processing)
multi_button.grid(row=0, column=1, padx=10, pady=10)
question_button = tk.Button(switch_button_frame, text="?", command=open_new_window)
question_button.grid(row=0, column=2, padx=10, pady=10)

switch_button_frame.pack(pady=10)
single_table_frame.pack(fill="both", expand=True)



# 启动窗口
root.mainloop()
