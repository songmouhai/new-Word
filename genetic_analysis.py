import tkinter
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import pandas as pd

def open_file(*args):
    def drop_down(event):
        global A1
        A1 = combo.get()
    global A
    file_path = filedialog.askopenfilename()
    file_label.insert("insert", file_path)
    A =str(file_path)
    sheet_names = pd.ExcelFile(A).sheet_names   
    combo = ttk.Combobox(u_2, values=sheet_names,textvariable=tkinter.StringVar(),state="readonly")
    combo.bind("<<ComboboxSelected>>", drop_down)
    combo.grid(row=1, column=2)
    
def open_file1():
    global B
    file_path1 = filedialog.askopenfilename()
    file_label1.insert("insert", file_path1)
    B =str(file_path1)

def select_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    file_label2.insert("insert", folder_path)    

def on_button_click():
    if lie0.get() == "":
        lie0.config(bg="red")
    else:
        lie0.config(bg="white")
    if lie1.get() == "":
        lie1.config(bg="red")
    else:
        lie1.config(bg="white")
    if lie2.get() == "":
        lie2.config(bg="red")
    else:
        lie2.config(bg="white")
    result = output()
    if result:
        result_label0.insert("end","运行成功")
        result_label0.config(bg="green")
    else:
        result_label0.insert("end","处理失败")
        result_label0.config(bg="red")

def output():
    try:
        B1 = lie3.get()
        a1 = lie0.get()
        b1 = lie1.get()
        b2 = lie2.get()
        a2 = b2 + "(1)"
        a3 = b2 + "(2)"
        a_df = pd.read_excel(A, sheet_name= A1 )
        b_df = pd.read_excel(B)
        df_a = pd.read_excel(A, sheet_name= A1 )
        a_column_format = b_df[ b1 ].dtype
        a_df[[ a1 ]] = a_df[[ a1  ]].astype(a_column_format)
        a_df = pd.read_excel(A, sheet_name= A1 )
        b_df = pd.read_excel(B, sheet_name= 0)
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

        file_path =folder_path + '/' + B1 + '.xlsx'
        df_a.to_excel(file_path, index=False, engine='openpyxl')
        return True
    except Exception as e:
        result_label1.config(text=f"处理文件时出现错误: {str(e)}",bg="red")
        print(f"处理文件时出现错误: {str(e)}")
        return False
    
global A
u_2 = tk.Tk()
u_2.title('表格信息筛选')
u_2.geometry("720x300")
text_label = tk.Label(u_2, text="说明！""\n有疑问联系QQ:3474284784")
text_label.grid(row=0,column=0,columnspan=3)
file_label =  tk.Text(u_2, width=60, height=2)
file_label.grid(row=1, column=1, padx=10)
open_button1 = tk.Button(u_2, text="填入表格", command=open_file)
open_button1.grid(row=1, column=0)
file_label1 =  tk.Text(u_2, width=60, height=2)
file_label1.grid(row=3, column=1, padx=10)
open_button2 = tk.Button(u_2, text="信息表格", command=open_file1)
open_button2.grid(row=3, column=0)
text = tk.Label(u_2,text="填入表格信息列")
text.grid(row=4, column=0)
lie0 = tk.Entry(u_2)
lie0.grid(row=4, column=1)
text1 = tk.Label(u_2,text="基因表a列")
text1.grid(row=5, column=0)
lie1 = tk.Entry(u_2)
lie1.grid(row=5, column=1)
text2 = tk.Label(u_2,text="基因表b列")
text2.grid(row=6, column=0)
lie2 = tk.Entry(u_2)
lie2.grid(row=6, column=1)
file_label2 =  tk.Text(u_2, width=60, height=2)
file_label2.grid(row=7, column=1, padx=10)
button = tk.Button(u_2, text="保存位置", command=select_folder)
button.grid(row=7,column=0)
text2 = tk.Label(u_2,text="新表格名称")
text2.grid(row=6, column=2)
lie3 = tk.Entry(u_2)
lie3.grid(row=7, column=2)
btn3 = tk.Button(u_2,text="输出", command=output and on_button_click)
btn3.grid(row=8, column=1)
text3 = tk.Label(u_2,text="运行结果")
text3.grid(row=8, column=2)
result_label0 = tk.Text(u_2,width=25, height=1)
result_label0.grid(row=9, column=2)
result_label1 = tk.Label(u_2, text="")
result_label1.grid(row=9, column=0,columnspan=2)
    
u_2.mainloop()

