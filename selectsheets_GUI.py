import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def select_sheets_from_excel():
    def open_excel_file():
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', ('*.xlsx', '*.xls'))])
        if not file_path:
            return None
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        excel_file.close()
        return sheet_names

    def display_sheets(sheet_names):
        selection_window = tk.Toplevel(root)
        selection_window.title("选择工作表")
        selection_window.geometry("400x600")  # 设置窗口大小

        checkboxes_vars = {}
        for sheet in sheet_names:
            var = tk.IntVar()
            chk = tk.Checkbutton(selection_window, text=sheet, variable=var)
            chk.pack(anchor='w')
            checkboxes_vars[sheet] = var

        def confirm_selection():
            selected_sheets = [sheet for sheet, var in checkboxes_vars.items() if var.get() == 1]
            if not selected_sheets:
                messagebox.showwarning("警告", "至少选择一个工作表！")
                return
            selection_window.destroy()  # 关闭选择窗口
            return selected_sheets

        confirm_button = tk.Button(selection_window, text="确认", command=lambda: selection_window.quit())
        confirm_button.pack()

        selection_window.mainloop()  # 这里启动了局部事件循环

        # 返回选择的工作表
        return [sheet for sheet, var in checkboxes_vars.items() if var.get() == 1]

    root = tk.Tk()
    root.withdraw()  # 隐藏根窗口

    sheet_names = open_excel_file()
    if sheet_names is not None:
        selected_sheets = display_sheets(sheet_names)
        root.destroy()  # 确保所有窗口都关闭
        return selected_sheets
    else:
        root.destroy()  # 用户没有选择文件
        return None

# 这个函数可以被其他程序调用，如下所示：
# selected_sheet_list = select_sheets_from_excel()
# print(selected_sheet_list)