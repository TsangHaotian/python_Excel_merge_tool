import os
import pandas as pd
from tkinter import Tk, filedialog, messagebox, Button

def merge_excel_files(folder_path):
    # 检查文件夹路径是否有效
    if not os.path.isdir(folder_path):
        messagebox.showerror("错误", "提供的文件夹路径无效")
        return

    # 获取文件夹内所有Excel文件
    excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx') or f.endswith('.xls')]
    if not excel_files:
        messagebox.showinfo("信息", "在文件夹中没有找到Excel文件")
        return

    # 读取并合并所有Excel文件
    merged_df = pd.DataFrame()
    for file in excel_files:
        file_path = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(file_path)
            merged_df = pd.concat([merged_df, df], ignore_index=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法读取文件 {file}: {e}")

    # 如果有数据则保存到新的Excel文件
    if not merged_df.empty:
        output_file = os.path.join(folder_path, '合并数据.xlsx')
        merged_df.to_excel(output_file, index=False)
        messagebox.showinfo("完成", f"所有数据已合并并保存到 {output_file}")
    else:
        messagebox.showinfo("信息", "没有数据可以保存")

# 创建Tkinter窗口
root = Tk()
root.title("Excel文件合并工具")
root.geometry("300x150")

# 添加选择文件夹按钮
def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        merge_excel_files(folder_path)

select_folder_button = Button(root, text="选择文件夹", command=select_folder)
select_folder_button.pack(pady=20)

# 运行Tkinter事件循环
root.mainloop()