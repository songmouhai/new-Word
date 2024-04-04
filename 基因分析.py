import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd
import openpyxl


def open_u_1():
    ui_1.destroy()
    def open_file():
        global A
        file_path = filedialog.askopenfilename()
        file_label.config(text=file_path)
        A =str(file_path)

    def lie_text():
        global a1
        global a2
        a1 = lie0.get()
        a2 = lie1.get()

    def on_button_click():
        result = output()
        if result:
            result_label.config(text="文件处理成功")
        else:
            result_label.config(text="文件处理失败")
          
    def output():
        try:
            global A
            global a1
            global a2
            b = "相关性"
            # 读取表格数据
            df = pd.read_excel(A)
            correlation = "相关性"
            # 获取基因a和基因b列的数据
            gene_a = df[a1]
            gene_b = df[a2]
            
            for index, row in df.iterrows():
                a1_value = row[a1]
                for i in range(len(df)):
                    if not pd.isna(a1_value):
                        matched_row = df[df[a2] == a1_value]
                        if not matched_row.empty:
                            df.at[index, b] = matched_row[a1].values[0]
                        else:
                            df.at[index, b] = float("55555155555")
            workbook = openpyxl.load_workbook(A)
            sheet = workbook.active
            search_value = "55555155555"
            replace_value = "无"
            for row in sheet.iter_rows():
                for cell in row:
                    # 如果单元格中的内容与要查找的内容相同，则替换为新内容
                    if cell.value == search_value:
                        cell.value = replace_value

            df.to_excel('modified_file.xlsx', index=False)
            return True
        except Exception as e:
            print(f"处理文件时出现错误: {str(e)}")
            return False
    u_1 = tk.Tk()
    u_1.title('基因相关分析')
    u_1.geometry("1000x500")
    file_label =  tk.Label(u_1)
    file_label.pack()
    open_button = tk.Button(u_1, text="选择表格", command=open_file)
    open_button.pack()
    text = tk.Label(u_1,text="基因ID列")
    text.pack()
    lie0 = tk.Entry(u_1)
    lie0.pack()
    text1 = tk.Label(u_1,text="目标基因列")
    text1.pack()
    lie1 = tk.Entry(u_1)
    lie1.pack()
    btn1 = tk.Button(u_1,text="输出", command=output and lie_text and on_button_click)
    btn1.pack()
    # 创建处理结果标签
    result_label = tk.Label(u_1, text="")
    result_label.pack()
    u_1.mainloop()


def open_u_2():
    ui_1.destroy()

    def open_file():
        global A
        file_path = filedialog.askopenfilename()
        file_label.config(text=file_path)
        A =str(file_path)
    def open_file1():
        global B
        file_path1 = filedialog.askopenfilename()
        file_label1.config(text=file_path1)
        B =str(file_path1)

    def on_focus_out_1():
        if lie0.get() == "":
            lie0.config(bg="red")
        else:
            lie0.config(bg="white")
        #第一个输入框判断，如果没有输入，框变红色
        if lie1.get() == "":
            lie1.config(bg="red")
        else:
            lie1.config(bg="white")
        #第二个输入框判断，如果没有输入，框变红色
        if lie2.get() == "":
            lie2.config(bg="red")
        else:
            lie2.config(bg="white")

    def save():
        global a1
        global b1
        global b2
        a1 = lie0.get()
        b1 = lie1.get()
        b2 = lie2.get()

    def on_button_click():
        result = output()
        if result:
            result_label.config(text="处理成功")
        else:
            result_label.config(text="处理失败")


          
    def output():
        try:
            global A
            global B
            global a1
            global b1
            global b2
            a2 = b2 + "(1)"
            a3 = b2 + "(2)"
            a_df = pd.read_excel(A)
            b_df = pd.read_excel(B)
            df_a = pd.read_excel(A)
            a_column_format = b_df[ b1 ].dtype
            a_df[[ a1 ]] = a_df[[ a1  ]].astype(a_column_format)
            a_df = pd.read_excel(A)
            b_df = pd.read_excel(B)
            df_a[a2] = df_a[a1] + str(1)
            df_a[a3] = df_a[a1] + str(2)
            
            for index, row in df_a.iterrows():
                a1_value = row[a1]
                if not pd.isna(a1_value):
                    matched_row = b_df[b_df[b1] == a1_value]
                    if not matched_row.empty:
                        df_a.at[index, a2] = matched_row[b2].values[0]
                    else:
                        df_a.at[index, a2] = float("55555155555")

            for index, row in df_a.iterrows():
                a1_value = row[a1]
                if not pd.isna(a1_value):
                    matched_row = b_df[b_df[b2] == a1_value]
                    if not matched_row.empty:
                        df_a.at[index, a3] = matched_row[b1].values[0]
                    else:
                        df_a.at[index, a3] = float("55555155555")

            df_a.to_excel('new.xlsx', index=False, engine='openpyxl')
            return True
        except Exception as e:
            print(f"处理文件时出现错误: {str(e)}")
            return False

    u_2 = tk.Tk()
    u_2.title('共线基因筛选')
    u_2.geometry("1000x500")
    file_label =  tk.Label(u_2)
    file_label.pack()
    open_button = tk.Button(u_2, text="填入表格", command=open_file)
    open_button.pack()
    file_label1 =  tk.Label(u_2)
    file_label1.pack()
    open_button = tk.Button(u_2, text="基因表格", command=open_file1)
    open_button.pack()
    text = tk.Label(u_2,text="填入表基因ID列")
    text.pack()
    lie0 = tk.Entry(u_2)
    lie0.pack()
    text1 = tk.Label(u_2,text="基因表a列")
    text1.pack()
    lie1 = tk.Entry(u_2)
    lie1.pack()
    text2 = tk.Label(u_2,text="基因表b列")
    text2.pack()
    lie2 = tk.Entry(u_2)
    lie2.pack()
    btn2 = tk.Button(u_2,text="保存", command=save and on_focus_out_1)
    btn2.pack()
    btn3 = tk.Button(u_2,text="输出", command=output and on_button_click)
    btn3.pack()
    # 创建处理结果标签
    result_label = tk.Label(u_2, text="")
    result_label.pack()

    u_2.mainloop()

ui_1 = tk.Tk()
ui_1.title('基因脚本')
ui_1.geometry("1000x500")
text_label = tk.Label(ui_1, text="说明！""\n基因相关分析:同一个表格两个指定列的查找，并在在新的列生成有无相关性""\n共线基因筛选:两个表格寻找相同列的基因，并生成两列显示这些基因""\n“55555155555”为无")
text_label.pack()
button1 = tk.Button(ui_1, text="基因相关分析", command=open_u_1 )
button1.pack()
button2 = tk.Button(ui_1, text="共线基因筛选", command=open_u_2)
button2.pack()


ui_1.mainloop()

