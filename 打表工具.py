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
        # 读取 Excel 文件，保留所有原始数据（header=None）
        df = pd.read_excel(file_path, header=None)

        # 将第二行作为列名
        df.columns = df.iloc[1]  # 将第二行数据作为列名
        df = df.drop(index=1)  # 删除第二行（原列名所在的行）

        # 获取列名，并找到从第三列开始的列名
        new_columns = []
        for i, col in enumerate(df.columns):
            if i >= 2 and word_to_add:  # 从第三列开始修改列名（0-based index）并且词语不为空
                new_columns.append(col + word_to_add)
            else:
                new_columns.append(col)

        # 更新 DataFrame 的列名
        df.columns = new_columns

        # 从第五行（索引为 3）和第三列（索引为 2）开始处理数据
        for row_idx in range(3, len(df)):
            for col_idx in range(2, len(df.columns)):
                # 如果该单元格的值大于 0，则在该单元格内容前面加上列名
                if df.iloc[row_idx, col_idx] > 0:
                    df.iloc[row_idx, col_idx] = str(df.columns[col_idx]) + str(df.iloc[row_idx, col_idx])

        # 合并同一行从第三列开始的单元格内容，分隔符为“，”并生成新的一列
        df['制品'] = df.iloc[3:, 2:].apply(lambda row: ','.join(row.dropna().astype(str)), axis=1)

        # 将“制品”插入到第二列后面
        df.insert(2, '制品', df.pop('制品'))

        # 将前几列提取出来生成新的文件
        df_first_three_columns = df.iloc[5:, :3]  # 提取前三列（从第五行开始）

        # 修改列名为指定的名称
        df_first_three_columns.columns = ['金额', 'cn', '制品']

        # 创建一个新的 DataFrame，作为新文件的第一行
        header_row = pd.DataFrame([[series_name + (word_to_add if word_to_add else ""), "", ""]],
                                  columns=df_first_three_columns.columns)

        # 将新的第一行插入到现有 DataFrame 中，作为列名行之上
        df_first_three_columns = pd.concat([header_row, df_first_three_columns], ignore_index=True)

        # 统计第二列第三行开始的单元格个数
        valid_cells_count = df_first_three_columns.iloc[1:, 1].count()  # 第二列从第三行开始的非空单元格数量

        # 计算人均邮费
        if postage_fee > 0 and valid_cells_count > 0:
            # 使用 Decimal 来处理邮费，并进行进位
            average_postage = Decimal(float(postage_fee)) / Decimal(float(valid_cells_count))
            # 将结果保留两位小数并进行进位
            average_postage = average_postage.quantize(Decimal('0.01'), rounding=ROUND_UP)

            # 将人均邮费插入到第二行第二列单元格
            df_first_three_columns.iloc[0, 1] = f"邮{average_postage}/人"  # 保留两位小数并进位处理

            # 对第一列从第二行开始的每个单元格加上人均邮费
            for row_idx in range(1, len(df_first_three_columns)):
                original_value = df_first_three_columns.iloc[row_idx, 0]
                # 如果原值是数字，则添加邮费
                if isinstance(original_value, (int, float)):
                    # 计算最终金额（原金额 + 人均邮费）
                    final_amount = Decimal(original_value) + average_postage
                    # 更新金额列，保留两位小数并进位
                    df_first_three_columns.iloc[row_idx, 0] = final_amount.quantize(Decimal('0.01'), rounding=ROUND_UP)

        # 动态生成文件名，结合输入的系列名和词语
        output_file_name = f"{series_name}{word_to_add if word_to_add else ''}.xlsx"

        # 保存包含前三列的数据
        df_first_three_columns.to_excel(output_file_name, index=False, header=True)
        print(f"包含前三列的数据已保存为 {output_file_name}")

        # 自动打开生成的文件
        open_generated_file(output_file_name)

        messagebox.showinfo("成功", "处理已完成！生成的文件已保存并打开。")

    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")


# 打开生成的文件的函数
def open_generated_file(file_path):
    if platform.system() == "Windows":
        os.startfile(file_path)  # Windows 中打开文件
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", file_path])
    else:  # Linux
        subprocess.run(["xdg-open", file_path])


# 选择文件的函数
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    file_entry.delete(0, tk.END)  # 清空现有路径
    file_entry.insert(0, file_path)  # 显示选择的文件路径


# 运行处理的函数
def run_processing():
    file_path = file_entry.get()
    word_to_add = word_entry.get()
    series_name = series_entry.get()
    try:
        postage_fee = float(postage_entry.get())  # 获取邮费输入并转换为浮动类型
    except ValueError:
        messagebox.showwarning("输入错误", "邮费必须是一个数字！")
        return

    # 验证输入
    if not file_path or not series_name:
        messagebox.showwarning("输入错误", "请确保输入文件路径和系列名！")
        return

    process_excel(file_path, word_to_add, series_name, postage_fee)


# 创建主窗口
root = tk.Tk()
root.title("打表工具")

# 获取屏幕宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 设置窗口的大小
window_width = 550
window_height = 250

# 计算窗口位置，使其居中
position_top = (screen_height - window_height) // 2
position_right = (screen_width - window_width) // 2

# 设置窗口大小和位置
root.geometry(f'{window_width}x{window_height}+{position_right}+{position_top}')

# 文件路径输入框及浏览按钮
file_label = tk.Label(root, text="选择 Excel 文件:")
file_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

file_entry = tk.Entry(root, width=40)
file_entry.grid(row=0, column=1, padx=10, pady=10)

browse_button = tk.Button(root, text="浏览", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# 系列名输入框
series_label = tk.Label(root, text="输入系列名:")
series_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")

series_entry = tk.Entry(root, width=40)
series_entry.grid(row=1, column=1, padx=10, pady=10)

# 词语输入框（允许为空）
word_label = tk.Label(root, text="输入制品类型（允许为空）:")
word_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")

word_entry = tk.Entry(root, width=40)
word_entry.grid(row=2, column=1, padx=10, pady=10)

# 邮费输入框
postage_label = tk.Label(root, text="输入邮费:")
postage_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")

postage_entry = tk.Entry(root, width=40)
postage_entry.grid(row=3, column=1, padx=10, pady=10)

# 运行按钮
run_button = tk.Button(root, text="开始处理", command=run_processing)
run_button.grid(row=4, column=0, columnspan=3, pady=20)

# 启动窗口
root.mainloop()
