import xlrd, xlwt, tkinter
from tkinter import ttk, filedialog, messagebox
path = ""
def choose_click():
    global path
    path = filedialog.askopenfilename()
    entry.delete(0, tkinter.END)
    entry.insert(0, path)
def confirm_click():
    global path
    path = entry.get()
    if(path == ""):
        return
    if(path.split(".")[-1] != "xls"):
        tkinter.messagebox.showwarning("提示", "不受支持的文件格式！请选择.xls文件")
        return
    if(path.find("\\") != -1):
        path.replace("\\", "/")
    convert()
def convert():
    global path
    try:
        data_in = xlrd.open_workbook(path)
    except Exception as errinfo:
        tkinter.messagebox.showwarning("提示", str(errinfo).split("\n  (Session")[0])
    data_out = xlwt.Workbook(encoding="utf-8")
    sheet = data_in.sheets()[2]
    output = data_out.add_sheet("sheet1")
    i = 0
    n = sheet.nrows
    while(42+i*96 < n):
        a = sheet.cell(i*96+42,3).value
        b = sheet.cell(i*96+43,3).value
        c = sheet.cell(i*96+44,3).value
        d = sheet.cell(i*96+45,3).value
        output.write(i,0,a)
        output.write(i,1,b)
        output.write(i,2,c)
        output.write(i,3,d)
        i += 1
    try:
        data_out.save(path.split(".")[0] + "_converted.xls")
    except Exception as errinfo:
        tkinter.messagebox.showwarning("提示", str(errinfo).split("\n  (Session")[0])
    tkinter.messagebox.showinfo("提示", "转换完成！")
window = tkinter.Tk()
window.resizable(False, False)
x = int((window.winfo_screenwidth() - window.winfo_reqwidth()) / 2)
y = int((window.winfo_screenheight() - window.winfo_reqheight()) / 2)
window.geometry(f"+{x}+{y}")
window.title("")
entry = ttk.Entry(window)
choose_button = ttk.Button(window, text="选择文件", command=choose_click, takefocus=False)
confirm_button = ttk.Button(window, text="转换", command=confirm_click, takefocus=False)
entry.grid(row=0, column=0, columnspan=2, ipadx=50, padx=(5,3), pady=(5,3))
choose_button.grid(row=0, column=2, padx=(0,3), pady=(3,0))
confirm_button.grid(row=1, column=1, pady=(1,3))
window.mainloop()