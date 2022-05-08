from tkinter import *
from tkinter import filedialog, messagebox, ttk
import matplotlib.pyplot as plt
from PIL import ImageTk, Image
import pandas as pd

root = Tk()
root.title("Data Visualizer using Python")
root.iconbitmap('images/favicon.ico')
root.geometry("800x500")
root.pack_propagate(False)
root.resizable(0, 0)


my_img = Image.open('images/logo.png')
resize = my_img.resize((60, 70), Image.ANTIALIAS)
new_img = ImageTk.PhotoImage(resize)
my_label = Label(root, image=new_img)
my_label.place(relx=0.05, rely=0.01)

h_name = Label(root, text="Data Visualization using Python", fg="black", font=('roman', 24, 'bold')).place(relx=0.22, rely=0.01)
m_name = Label(root, text="", fg="blue", font=('times', 21, 'underline')).place(relx=0.40, rely=0.07)


frame1 = LabelFrame(root, text="Excel Data")
frame1.place(height=300, width=550, rely=0.13)


file_frame = LabelFrame(root, text="Open File")
file_frame.place(height=100, width=400, rely=0.58, relx=0.14)

button1 = Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button1.place(rely=0.44, relx=0.50)

button2 = Button(file_frame, text="Load File", command=lambda: Load_excel_data())
button2.place(rely=0.44, relx=0.30)


label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)

menu_frame = LabelFrame(root, text="Options")
menu_frame.place(height=150, width=400, rely=0.74, relx=0.14)
text = Label(menu_frame, text="Choose Column: ").place(rely=0.01, relx=0)


#Treeview
tv1 = ttk.Treeview(frame1)
tv1.place(relheight=1, relwidth=1) # set the height and width of the widget to 100% of its container (frame1).

treescrolly = Scrollbar(frame1, orient="vertical", command=tv1.yview)
treescrollx = Scrollbar(frame1, orient="horizontal", command=tv1.xview)
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill="x")
treescrolly.pack(side="right", fill="y")


def File_dialog():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("csv files", "*.csv"), ("xlsx files", "*.xlsx"), ("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        if excel_filename[-4:] == ".csv":

            df = pd.read_csv(excel_filename)
            clicked = StringVar()
            clicked.set("Choose Column")
            column = list(df.columns)
            ddmenu = OptionMenu(menu_frame, clicked, *column).place(relx=0.1, rely=0.25)
            clicked2 = StringVar()
            clicked2.set("Choose Column")
            column2 = list(df.columns)
            ddmenu2 = OptionMenu(menu_frame, clicked2, *column2).place(relx=0.5, rely=0.25)
            view3 = Button(menu_frame, text="View", width="40" ,command=lambda: viewdata(clicked,clicked2)).place(relx=0.1,rely=0.55)

            def viewdata(a,b):
                try:
                    df.plot(kind='bar', x=a.get(), y=b.get(), color='skyblue')
                    plt.show()
                except TypeError:
                    df.plot(kind='bar', x=b.get(), y=a.get(), color='skyblue')
                    plt.show()
                except KeyError:
                    messagebox.showerror("Error", "Invalid Choice")

        else:
            df = pd.read_excel(excel_filename)
            clicked = StringVar()
            clicked.set("Choose Column")
            column = list(df.columns)
            ddmenu = OptionMenu(menu_frame, clicked, *column).pack()
            clicked2 = StringVar()
            clicked2.set("Choose Column")
            column2 = list(df.columns)
            ddmenu2 = OptionMenu(menu_frame, clicked2, *column2).pack()
            view3 = Button(menu_frame, text="View", command=lambda: viewdata(clicked, clicked2)).pack()

            def viewdata(a, b):
                try:
                    df.plot(kind='bar', x=a.get(), y=b.get(), color='skyblue')
                    plt.show()
                except TypeError:
                    df.plot(kind='bar', x=b.get(), y=a.get(), color='skyblue')
                    plt.show()
                except KeyError:
                    messagebox.showerror("Error", "Invaid Choice")

    except ValueError:
        messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        messagebox.showerror("Information", f"No such file as {file_path}")
        return None

    clear_data()
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tv1.insert("", "end", values=row)
    return None


def clear_data():
    tv1.delete(*tv1.get_children())
    return None


root.mainloop()