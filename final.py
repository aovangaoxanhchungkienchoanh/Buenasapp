import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.filedialog import asksaveasfilename
from tkinter.messagebox import showerror
import pandas as pd
from tkinter import filedialog, simpledialog
from tkinter import *
from datetime import datetime, timedelta
import pytz
from pandas import ExcelWriter
import clipboard as cb
from requests import get
import pymysql, os, sys


def get_path(filename):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, filename)
    else:
        return filename

# initalise the tkinter GUI
root = tk.Tk()
root.geometry("500x600") # set the root dimensions
root.title("App Buenas")
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.
root.iconbitmap(get_path('setting_folder_icon_219451.ico'))
pd.options.mode.chained_assignment = None  # default='warn'
#mutiple tab
# window = Tk()
tabsystem = ttk.Notebook(root)
tab1 = Frame(tabsystem)
tab2 = Frame(tabsystem)
tab3 = Frame(tabsystem)

tabsystem.add(tab1, text='QR Tab')
tabsystem.add(tab2, text='Kiểm kho')
tabsystem.add(tab3, text='Đơn web')
tabsystem.pack(expand=1, fill="both")

##TAB 1
# Frame for TreeView
frame1_tab1 = tk.LabelFrame(tab1, text="Excel Data")
frame1_tab1.place(height=250, width=500)

# Frame for open file dialog
file_frame_tab1 = tk.LabelFrame(tab1, text="Open File")
file_frame_tab1.place(height=250, width=400, rely=0.5, relx=0.09)


# Buttons
button1_tab1 = tk.Button(file_frame_tab1, text="File Kiot sáng", command=lambda: File_dialog_kiot())
button1_tab1.place(rely=0.3, relx=0.1)

button2_tab1 = tk.Button(file_frame_tab1, text="File QR sáng", command=lambda: File_dialog_QR())
button2_tab1.place(rely=0.3, relx=0.4)

button3_tab1 = tk.Button(file_frame_tab1, text="Xuất File sáng", command=lambda: Save_excel_data_QR_morning())
button3_tab1.place(rely=0.3, relx=0.65)

button4_tab1 = tk.Button(file_frame_tab1, text="File Kiot chiều", command=lambda: File_dialog_kiot())
button4_tab1.place(rely=0.5, relx=0.1)

button5_tab1 = tk.Button(file_frame_tab1, text="File QR chiều", command=lambda: File_dialog_QR())
button5_tab1.place(rely=0.5, relx=0.4)

button6_tab1 = tk.Button(file_frame_tab1, text="Xuất file chiều", command=lambda: Save_excel_data_QR_affternoon())
button6_tab1.place(rely=0.5, relx=0.65)

button7_tab1 = tk.Button(file_frame_tab1, text="Lấy hàng chiều", command=lambda: Save_excel_data_handled_affternoon())
button7_tab1.place(rely=0.7, relx=0.65)

button8_tab1 = tk.Button(file_frame_tab1, text="Lấy hàng sáng", command=lambda: Save_excel_data_handled())
button8_tab1.place(rely=0.7, relx=0.3)

button9_tab1 = tk.Button(file_frame_tab1, text="Lấy giờ", command=lambda: Get_max_time())
button9_tab1.place(rely=0.7, relx=0.1)

button10_tab1 = tk.Button(file_frame_tab1, text="XUẤT FILE CẢ NGÀY", bg="#353A5F", fg="white", font=" sans 10 bold", command=lambda: Concat_data_QR())
button10_tab1.place(rely=0.88, relx=0, width=395)


# The file/file path text
label_file1_tab1 = ttk.Label(file_frame_tab1, text="No File Selected Kiot")
label_file1_tab1.place(rely=0.05, relx=0)
label_file2_tab1 = ttk.Label(file_frame_tab1, text="No File Selected QR")
label_file2_tab1.place(rely=0.15, relx=0)


## Treeview Widgetx
tv1 = ttk.Treeview(frame1_tab1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly_tab1 = tk.Scrollbar(frame1_tab1, orient="vertical", command=tv1.yview) # command means update the yaxis view of the widget
treescrollx_tab1 = tk.Scrollbar(frame1_tab1, orient="horizontal", command=tv1.xview) # command means update the xaxis view of the widget
tv1.configure(xscrollcommand=treescrollx_tab1.set, yscrollcommand=treescrolly_tab1.set) # assign the scrollbars to the Treeview Widget
treescrollx_tab1.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly_tab1.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget


##TAB 2
# Frame for TreeView
# Frame for TreeView
frame1_tab2 = tk.LabelFrame(tab2, text="Excel Data")
frame1_tab2.place(height=250, width=500)

# Frame for open file dialog
file_frame_tab2 = tk.LabelFrame(tab2, text="Open File")
file_frame_tab2.place(height=200, width=400, rely=0.6, relx=0.09)


# Buttons
button1_tab2 = tk.Button(file_frame_tab2, text="1. Chọn file kiot", command=lambda: File_dialog_stock())
button1_tab2.place(rely=0.4, relx=0.1)

button2_tab2 = tk.Button(file_frame_tab2, text="2. Chọn file kiểm", command=lambda: File_dialog_QR_stock())
button2_tab2.place(rely=0.4, relx=0.4)

button3_tab2 = tk.Button(file_frame_tab2, text="3. Thống kê", command=lambda: Get_data_information_stock())
button3_tab2.place(rely=0.4, relx=0.69)

button4_tab2 = tk.Button(file_frame_tab2, text="4. Đối soát", command=lambda: File_compare_stock())
button4_tab2.place(rely=0.69, relx=0.1)
#
button5_tab2 = tk.Button(file_frame_tab2, text="5. File upload kho mới", command=lambda: Save_excel_data_stock())
button5_tab2.place(rely=0.69, relx=0.4)

# button6_tab2 = tk.Button(file_frame_tab2, text="6. Get total", command=lambda: Save_excel_data())
# button6_tab2.place(rely=0.69, relx=0.69)


# The file/file path text
label_file1_tab2 = ttk.Label(file_frame_tab2, text="No File Selected hanlded")
label_file1_tab2.place(rely=0.1, relx=0)
label_file2_tab2 = ttk.Label(file_frame_tab2, text="No File Selected QR")
label_file2_tab2.place(rely=0.2, relx=0)

## Treeview Widget
tv2 = ttk.Treeview(frame1_tab2)
tv2.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly_tab2 = tk.Scrollbar(frame1_tab2, orient="vertical", command=tv2.yview) # command means update the yaxis view of the widget
treescrollx_tab2 = tk.Scrollbar(frame1_tab2, orient="horizontal", command=tv2.xview) # command means update the xaxis view of the widget
tv2.configure(xscrollcommand=treescrollx_tab2.set, yscrollcommand=treescrolly_tab2.set) # assign the scrollbars to the Treeview Widget
treescrollx_tab2.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly_tab2.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget

#tab3
# Frame for TreeView
frame1_tab3 = tk.LabelFrame(tab3, text="Excel Data")
frame1_tab3.place(height=250, width=500)
#
# # Frame for open file dialog
file_frame_tab3 = tk.LabelFrame(tab3, text="Open File")
file_frame_tab3.place(height=200, width=400, rely=0.5, relx=0.09)

# Buttons
button1 = tk.Button(file_frame_tab3, text="Lấy IP public", command=lambda: Get_my_IP())
button1.place(rely=0.3, relx=0.1)

button2 = tk.Button(file_frame_tab3, text="Cấp quyền", command=lambda: Access_buenas())
button2.place(rely=0.3, relx=0.4)

button3 = tk.Button(file_frame_tab3, text="Kiểm tra kết nối", command=lambda: connection_web())
button3.place(rely=0.3, relx=0.69)

button4 = tk.Button(file_frame_tab3, text="Lưu File",bg="#353A5F", fg="white", font=" sans 10 bold", command=lambda: Save_file_web())
button4.place(rely=0.55, relx=0, width=395)



## Treeview Widget
tv3 = ttk.Treeview(frame1_tab3)
tv3.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly_tab3 = tk.Scrollbar(frame1_tab3, orient="vertical", command=tv3.yview) # command means update the yaxis view of the widget
treescrollx_tab3 = tk.Scrollbar(frame1_tab3, orient="horizontal", command=tv3.xview) # command means update the xaxis view of the widget
tv3.configure(xscrollcommand=treescrollx_tab3.set, yscrollcommand=treescrolly_tab3.set) # assign the scrollbars to the Treeview Widget
treescrollx_tab3.pack(side="bottom", fill="x") # make the scrollbar fill the x axis of the Treeview widget
treescrolly_tab3.pack(side="right", fill="y") # make the scrollbar fill the y axis of the Treeview widget


#Function tab 1
def File_dialog_kiot():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file1_tab1["text"] = filename
    return None

def File_dialog_QR():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file2_tab1["text"] = filename
    return None

def Processing_data_time_package():
    file_path1 = label_file1_tab1["text"]
    excel_filename1 = r"{}".format(file_path1)
    df1 = pd.read_excel(excel_filename1)
    format_time = get_data_timer()
    getNow = datetime.strftime(datetime.now(pytz.timezone('Asia/Ho_Chi_Minh')), '{0}'.format(format_time))
    df1 = df1.loc[df1["Trạng thái giao hàng"] == "Chờ xử lý"]
    df1 = df1[
        ['Mã hóa đơn', 'Mã vận đơn', 'Tên hàng', 'Mã hàng', 'Thời gian tạo', 'Số lượng', 'Thương hiệu',
         'Trạng thái giao hàng',
         'Ghi chú', 'Tên khách hàng', 'Địa chỉ (Khách hàng)']]
    df1['Thời gian tạo'] = pd.to_datetime(df1['Thời gian tạo']).apply(lambda x: x.strftime('%d/%m/%Y %H:%M:%S'))
    df1 = df1.loc[df1['Thời gian tạo'] > getNow]
    return df1


def Processing_excel_data_QR_morning():
    file_path1 = label_file1_tab1["text"]
    file_path2 = label_file2_tab1["text"]
    excel_filename1 = r"{}".format(file_path1)
    excel_filename2 = r"{}".format(file_path2)
    df1 = pd.read_excel(excel_filename1)
    df2 = pd.read_excel(excel_filename2)
    df2 = df2.astype({'Mã vận đơn': 'str'})
    df2 = df2.loc[df2["Mã vận đơn"] != "Demo - please subscribe to full version"]
    df2 = df2.drop_duplicates(subset=['Mã vận đơn'], keep='first')
    df2['Đã xử lý'] = "TRUE"
    df2 = df2[['Mã vận đơn', 'Đã xử lý', 'Thời gian quét']]
    df1 = df1.loc[df1["Trạng thái giao hàng"] == "Chờ xử lý"]
    df1 = df1[
        ['Mã hóa đơn', 'Mã vận đơn', 'Thời gian tạo', 'Tên hàng', 'Mã hàng', 'Số lượng','Thương hiệu', 'Trạng thái giao hàng',
         'Ghi chú',
         'Tên khách hàng', 'Điện thoại', 'Địa chỉ (Khách hàng)', 'Thành tiền']]
    inner = pd.merge(df1, df2, on="Mã vận đơn", how="left")
    return inner

def Save_excel_data_QR_morning():
    inner = Processing_excel_data_QR_morning()
    datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%d_%m_%Y_%Hh')
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                       ("All files", "*.*")), initialfile="Xuly" + datestring)
    finalTable = inner.to_excel("{}".format(savefile) + ".xlsx", sheet_name="Orders")

def Processing_excel_data_QR_affternoon():
    file_path1 = label_file1_tab1["text"]
    file_path2 = label_file2_tab1["text"]
    excel_filename1 = r"{}".format(file_path1)
    excel_filename2 = r"{}".format(file_path2)
    df1 = pd.read_excel(excel_filename1)
    df2 = pd.read_excel(excel_filename2)
    df2 = df2.loc[df2["Mã vận đơn"] != "Demo - please subscribe to full version"]
    df2 = df2.astype({'Mã vận đơn': 'str'})
    df2 = df2.drop_duplicates(subset=['Mã vận đơn'], keep='first')
    df2['Đã xử lý'] = "TRUE"
    df2 = df2[['Mã vận đơn', 'Đã xử lý', 'Thời gian quét']]
    format_time = get_data_timer()
    getNow = datetime.strftime(datetime.now(pytz.timezone('Asia/Ho_Chi_Minh')), '{0}'.format(format_time))
    df1 = df1.loc[df1["Trạng thái giao hàng"] == "Chờ xử lý"]
    df1 = df1[
        ['Mã hóa đơn', 'Mã vận đơn', 'Thời gian tạo', 'Tên hàng', 'Mã hàng', 'Số lượng', 'Thương hiệu', 'Trạng thái giao hàng',
         'Ghi chú', 'Tên khách hàng', 'Điện thoại', 'Địa chỉ (Khách hàng)', 'Thành tiền']]
    df1['Thời gian tạo'] = pd.to_datetime(df1['Thời gian tạo']).apply(lambda x: x.strftime('%d/%m/%Y %H:%M:%S'))
    df1 = df1.loc[df1['Thời gian tạo'] > getNow]
    inner = pd.merge(df1, df2, on="Mã vận đơn", how="left")
    return inner

def Save_excel_data_QR_affternoon():
    inner = Processing_excel_data_QR_affternoon()
    # sortedDF = df1.sort_values("Thương hiệu", ascending=False)
    datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%d_%m_%Y_%Hh')
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                       ("All files", "*.*")), initialfile="XulyChieu" + datestring)
    finalTable = inner.to_excel("{}".format(savefile) + ".xlsx", sheet_name="Orders")

def Save_excel_data_handled():
    file_path1 = label_file1_tab1["text"]
    excel_filename1 = r"{}".format(file_path1)
    df1 = pd.read_excel(excel_filename1)
    df1 = df1.loc[df1["Trạng thái giao hàng"] == "Chờ xử lý"]
    df1 = df1[
        ['Mã hóa đơn', 'Mã vận đơn', 'Tên hàng', 'Mã hàng', 'Thời gian tạo', 'Số lượng', 'Thương hiệu', 'Trạng thái giao hàng',
         'Ghi chú', 'Tên khách hàng', 'Địa chỉ (Khách hàng)']]
    df1['Thời gian tạo'] = pd.to_datetime(df1['Thời gian tạo']).apply(lambda x: x.strftime('%d-%m-%Y'))
    inner = df1.groupby(by=["Mã hàng", "Thương hiệu", "Tên hàng"]).sum().reset_index()
    sortedDF = inner.sort_values("Thương hiệu", ascending=False)
    datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%d_%m_%Y_%Hh')
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                       ("All files", "*.*")),initialfile="LayHang"+datestring)
    # writer = pd.ExcelWriter('dataframes.xlsx', engine='xlsxwriter')
    # finalTable = sortedDF.to_excel("{}".format(savefile) + ".xlsx", sheet_name="Orders")
    with pd.ExcelWriter("{0}".format(savefile) + ".xlsx") as writer:
        # df1.to_excel(writer, sheet_name="Đơn xử lý")
        inner.to_excel(writer, sheet_name="Lấy hàng")

def Save_excel_data_handled_affternoon():
    df = Processing_data_time_package()
    df['Thời gian tạo'] = pd.to_datetime(df['Thời gian tạo']).apply(lambda x: x.strftime('%d/%m/%Y %H:%M:%S'))
    inner = df.groupby(by=["Mã hàng", "Thương hiệu", "Tên hàng"]).sum().reset_index()
    datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%d_%m_%Y_%Hh')
    sortedDF = inner.sort_values("Thương hiệu", ascending=False)
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                       ("All files", "*.*")), initialfile="LayHangChieu" + datestring)
    # writer = pd.ExcelWriter('dataframes.xlsx', engine='xlsxwriter')
    # finalTable = sortedDF.to_excel("{}".format(savefile) + ".xlsx", sheet_name="Orders")
    with pd.ExcelWriter("{0}".format(savefile) + ".xlsx") as writer:
        # df1.to_excel(writer, sheet_name="Đơn xử lý")
        inner.to_excel(writer, sheet_name="Lấy hàng")

def Get_data_information():
    filename1 = filedialog.askopenfilename(initialdir="/",
                                           title="Chọn file xử lý ca sáng",
                                           filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file1_tab1["text"] = filename1
    df = pd.read_excel(filename1)
    #tổng các đơn hàng chưa được xử lý
    processed = df['Đã xử lý'].isna().sum()
    revenue = df['Thành tiền'].sum()
    bills = df['Mã hàng'].count()
    tv1.insert('', tk.END, text='Tổng đơn hàng cần xử lý: {0} đơn'.format(bills), iid=0, open=False)
    tv1.insert('', tk.END, text='Đơn hàng đã xử lý hôm nay: {0} / {1} đơn'.format(bills - processed, bills), iid=1, open=False)
    tv1.insert('', tk.END, text='Doanh thu hôm nay: {0}'.format(int(revenue)), iid=2, open=False)
    tv1.insert('', tk.END, text='Giá trị trung bình trên 1 đơn hàng {0}'.format(revenue/bills), iid=3, open=False)

    # adding children of first node
    # tv1.insert('', tk.END, text='Buenas: {0} đơn'.format(len(inner.loc[inner['Thương hiệu'] == "BUENAS"])),
    #            iid=4, open=False)
    # tv1.insert('', tk.END, text='Shondo: {0} đơn'.format(len(inner.loc[inner['Thương hiệu'] == "SHONDO"])),
    #            iid=5, open=False)
    # tv1.insert('', tk.END, text='Vento: {0} đơn'.format(len(inner.loc[inner['Thương hiệu'] == "VENTO"])),
    #            iid=6, open=False)
    # tv1.move(4, 3, 0)
    # tv1.move(5, 3, 1)
    # tv1.move(6, 3, 2)


def Concat_data_QR():
    filename1 = filedialog.askopenfilename(initialdir="/",
                                          title="Chọn file xử lý ca sáng",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file1_tab1["text"] = filename1

    filename2 = filedialog.askopenfilename(initialdir="/",
                                          title="Chọn file xử lý ca chiều",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file2_tab1["text"] = filename2
    df = pd.DataFrame()
    df = df.append(pd.read_excel("{0}".format(filename1)))
    df = df.append(pd.read_excel("{0}".format(filename2)))
    df = df.reset_index()
    # df = df.drop(['Unnamed: 0'], axis=1)
    df = df[
        ['Mã hóa đơn', 'Mã vận đơn', 'Thời gian tạo', 'Tên hàng', 'Mã hàng', 'Số lượng', 'Thương hiệu',
         'Trạng thái giao hàng',
         'Ghi chú',
         'Tên khách hàng', 'Điện thoại', 'Địa chỉ (Khách hàng)', 'Thành tiền', 'Đã xử lý','Thời gian quét']]
    datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%Y_%m_%d_%Hh')
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                       ("All files", "*.*")), initialfile="FileXuLy" + datestring)
    finalTable = df.to_excel("{}".format(savefile) + ".xlsx", sheet_name="Xử lý toàn ngày {0}".format(datestring))


def get_data_timer():
    a =simpledialog.askstring(title="Nhập giờ", prompt="Nhập điều kiện giờ \n(format nhập: hh:mm:ss [vd: 08:05:00])")
    return a

def Get_max_time():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file1_tab1["text"] = filename
    df1 = pd.read_excel("{}".format(filename))
    df1 = df1.loc[df1["Trạng thái giao hàng"] == "Chờ xử lý"]
    df1 = df1[
        ['Mã hóa đơn', 'Mã vận đơn', 'Thời gian tạo', 'Tên hàng', 'Mã hàng', 'Số lượng', 'Thương hiệu',
         'Trạng thái giao hàng',
         'Ghi chú', 'Tên khách hàng', 'Điện thoại', 'Địa chỉ (Khách hàng)', 'Thành tiền']]
    df1['Thời gian tạo'] = pd.to_datetime(df1['Thời gian tạo']).apply(lambda x: x.strftime('%d/%m/%Y %H:%M:%S'))
    max_get_time = df1['Thời gian tạo'].max()
    top = Toplevel(tab1)
    top.geometry("250x100")
    top.title("Show time")
    Label(top, text="{}".format(max_get_time)).place(relx=0.30, rely=0.1)
    # Label(top, text="alo alo alo alo").place(relx=0.35, rely=0.1)
    button = cb.copy(max_get_time)
    button = tk.Button(top, text="Copy time", command=lambda: messagebox.showinfo('Copy thời gian', 'Copy thành công'))
    button.place(rely=0.4, relx=0.37)
    # messagebox.showinfo('information', 'Copy thành công')

##Function tab 2

def File_dialog_stock():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file1_tab2["text"] = filename
    return None

def File_dialog_QR_stock():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file2_tab2["text"] = filename
    return None

def excel_data_processing_stock():
    file_path1 = label_file1_tab2["text"]
    file_path2 = label_file2_tab2["text"]
    excel_filename1 = r"{}".format(file_path1)
    excel_filename2 = r"{}".format(file_path2)
    df1 = pd.read_excel(excel_filename1)
    df2 = pd.read_excel(excel_filename2)

    # xử lý file tồn kho
    df2['Tồn kho thực'] = 1
    df2_count = df2.groupby(df2["Mã hàng"]).sum()

    # merge 2 file lại với nhau theo file kiot
    inner = pd.merge(df1, df2_count, on="Mã hàng", how="inner")

    #xử lý dữ liệu
    for i in range(0, len(inner['Tồn kho thực'])):
        if pd.isna(inner['Tồn kho thực'][i]) == True:
            inner['Tồn kho thực'][i] = 0
    inner = inner.astype({'Tồn kho thực': 'int'})
    return inner

def Get_data_information_stock():
    # clear_data()
    pd.options.mode.chained_assignment = None  # default='warn'
    data = excel_data_processing_stock()
    data['Chênh lệch'] = data['Tồn kho thực'] - data["Tồn kho"]
    inner1 = data[["Mã hàng", "Thương hiệu", "Tồn kho", "Tồn kho thực", "Chênh lệch"]]
    inner1 = inner1.sort_values("Thương hiệu", ascending=True)
    inner2 = inner1.copy()
    notequals = inner2.loc[inner2["Chênh lệch"] != 0].reset_index()
    equals = inner1.loc[inner2["Chênh lệch"] == 0].reset_index()

    tv2.insert('', tk.END, text='Tổng số hàng đã kiểm: {0}'.format(int(sum(inner1["Tồn kho thực"]))), iid=0, open=False)
    tv2.insert('', tk.END,
               text='Tổng số hàng chênh lệch trong kioviet: {0} so với thực tế'.format(int(sum(inner1["Chênh lệch"]))),
               iid=1, open=False)
    tv2.insert('', tk.END, text='Số hàng đúng với thực tế: {0}'.format(int(len(equals["Mã hàng"]))), iid=2, open=False)

def File_compare_stock():
    pd.options.mode.chained_assignment = None  # default='warn'
    data = excel_data_processing_stock()
    data['Chênh lệch'] = data['Tồn kho thực'] - data["Tồn kho"]
    inner1 = data[["Mã hàng", "Thương hiệu", "Tồn kho", "Tồn kho thực", "Chênh lệch"]]
    inner1 = inner1.sort_values("Thương hiệu", ascending=True)
    inner2 = inner1.copy()
    notequals = inner2.loc[inner2["Chênh lệch"] != 0].reset_index(drop=True)
    equals = inner1.loc[inner2["Chênh lệch"] == 0].reset_index(drop=True)
    datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%d_%m_%Y_%Hh')
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                       ("All files", "*.*")), initialfile="Doisoat" + datestring)
    with pd.ExcelWriter("{0}".format(savefile) + ".xlsx") as writer:
        inner1.to_excel(writer, sheet_name="Total")
        equals.to_excel(writer, sheet_name="Not Difference")
        notequals.to_excel(writer, sheet_name="Difference")


def Save_excel_data_stock():
    data = excel_data_processing_stock()
    data["Tồn kho"] = data["Tồn kho thực"]
    data = data.drop(columns='Tồn kho thực')
    datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%d_%m_%Y_%Hh')
    savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                       ("All files", "*.*")), initialfile="KiemKho" + datestring)
    finalTable = data.to_excel("{}".format(savefile) + ".xlsx", sheet_name="Hang hoa", index=False)

def clear_data():
    tv1.delete(*tv1.get_children())
    return None

##tab3
def Get_my_IP():
    ip = get('https://api.ipify.org').text
    top = Toplevel(root)
    top.geometry("250x100")
    top.title("Show my ip")
    Label(top, text="{}".format(ip)).place(relx=0.30, rely=0.1)
    # button = cb.copy(ip)
    button = tk.Button(top, text="Copy IP", command=lambda: messagebox.showinfo('Copy my ip public', 'Copy thành công') and cb.copy(ip))
    # button = cb.copy(ip)
    button.place(rely=0.4, relx=0.37)
    # print(f'My public IP address is: {ip}')

def Access_buenas():
    os.system("start \"\" https://buenas.vn:2083/cpsess9173809802/frontend/paper_lantern/sql/managehost.html")

def connection_web():
    clear_data_tab3()
    try:
        connection = pymysql.connect(host='45.252.251.69',
                                     user='ezqmxegv_admin2',
                                     password='Buenasthang10',
                                     db='ezqmxegv_buenas',
                                     charset='utf8mb4',
                                     cursorclass=pymysql.cursors.DictCursor)
        cursor = connection.cursor()
        tv3.insert('', tk.END, text='Kết nối THÀNH CÔNG', iid=2, open=False)
    except Exception as e:
        tv3.insert('', tk.END, text='Kết nối không thành công', iid=0, open=False)
        tv3.insert('', tk.END, text='Lỗi: {0}'.format(e), iid=1, open=False)
        # print(e)
def Get_information_web():
    try:
        connection = pymysql.connect(host='45.252.251.69',
                                                 user='ezqmxegv_admin2',
                                                 password='Buenasthang10',
                                                 db='ezqmxegv_buenas',
                                                 charset='utf8mb4',
                                                 cursorclass=pymysql.cursors.DictCursor)
        cursor = connection.cursor()
        getNow = datetime.strftime(datetime.now(pytz.timezone('Asia/Ho_Chi_Minh')), '%Y-%m-%d')
        getYesterday = datetime.strftime(datetime.now(pytz.timezone('Asia/Ho_Chi_Minh')) - timedelta(1), '%Y-%m-%d')
        query = "SELECT p.ID, post_date as order_id, p.post_status, " \
                "p.post_date," \
                "max( CASE WHEN pm.meta_key = '_billing_email' and p.ID = pm.post_id THEN pm.meta_value END ) as Email, " \
                "max( CASE WHEN pm.meta_key = '_billing_last_name' and p.ID = pm.post_id THEN pm.meta_value END ) as Name, " \
                "max( CASE WHEN pm.meta_key = '_billing_phone' and p.ID = pm.post_id THEN pm.meta_value END ) as Phone, " \
                "max( CASE WHEN pm.meta_key = '_billing_address_1' and p.ID = pm.post_id THEN pm.meta_value END ) as Address, " \
                "max( CASE WHEN pm.meta_key = '_billing_address_2' and p.ID = pm.post_id THEN pm.meta_value END ) as Comumune," \
                "max( CASE WHEN pm.meta_key = '_billing_city' and p.ID = pm.post_id THEN pm.meta_value END ) as District, " \
                "max( CASE WHEN pm.meta_key = '_billing_state' and p.ID = pm.post_id THEN pm.meta_value END ) as city, " \
                "max( CASE WHEN pm.meta_key = '_billing_address_index' and p.ID = pm.post_id THEN pm.meta_value END ) as Address_total, " \
                "max( CASE WHEN pm.meta_key = '_order_total' and p.ID = pm.post_id THEN pm.meta_value END ) as Total, " \
                "( select group_concat( order_item_name separator '|' ) from wp_woocommerce_order_items where order_id = p.ID ) as order_items " \
                "FROM " \
                "wp_posts p  " \
                "join wp_postmeta pm on p.ID = pm.post_id " \
                "join wp_woocommerce_order_items oi on p.ID = oi.order_id " \
                "WHERE " \
                "post_type = 'shop_order' and " \
                "post_date BETWEEN '{0}' AND '{1}' and " \
                "post_status IN ('wc-processing', 'wc-on-hold') " \
                "group by " \
                "p.ID ".format(getYesterday, getNow)
        cursor.execute(query)
        myresult = cursor.fetchall()
                    # myOrders = dict()
                    # for x in myresult:
                    # myOrders.append(x)
                    #   print(x.get('post_date'))

        df = pd.DataFrame(myresult)
        df = df[['ID', 'post_date', 'post_status', 'Email', 'Name', 'Phone', 'Address_total', 'Total', 'order_items']]
        df["Address_total"] = df["Address_total"].astype(str)
        df["Name"] = df["Name"].astype(str)
        df["Email"] = df["Email"].astype(str)
        df["Phone"] = df["Phone"].astype(str)

        for i in range(0, len(df['Address_total'])):
            if df['Name'][i] in df['Address_total'][i]:
                df['Address_total'][i] = df['Address_total'][i].replace(df['Name'][i], "")
            if df['Email'][i] in df['Address_total'][i]:
                df['Address_total'][i] = df['Address_total'][i].replace(df['Email'][i], "")
            if df['Phone'][i] in df['Address_total'][i]:
                df['Address_total'][i] = df['Address_total'][i].replace(df['Phone'][i], "")

        for i in range(0, len(df['Address_total'])):
            df['Address_total'][i] = df['Address_total'][i].strip()
            df = pd.DataFrame(df)
    except Exception as e:
        clear_data_tab3()
        tv3.insert('', tk.END, text='Lưu file không thành công do chưa được cấp quyền hoặc '
                                    'không có đơn hàng nào\ntrong ngày {0} và {1}'.format(getYesterday,getNow), iid=0, open=False)
    return df


def Save_file_web():
    df = Get_information_web()
    if df.empty is False:
        datestring = datetime.strftime(datetime.now(pytz.timezone("Asia/Ho_Chi_Minh")), '%d_%m_%Y_%Hh')
        savefile = filedialog.asksaveasfilename(filetypes=(("Excel files", "*.xlsx"),
                                                           ("All files", "*.*")), initialfile="DonWeb" + datestring)
        finalTable = df.to_excel("{}".format(savefile) + ".xlsx", sheet_name="Orders", index=False)
    else:
        messagebox.showinfo('Show errors', 'Chưa được cấp quyền hoặc hôm nay không có đơn hàng')
def clear_data_tab3():
    tv3.delete(*tv3.get_children())
    return None
root.mainloop()