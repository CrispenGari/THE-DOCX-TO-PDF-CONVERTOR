"""
This aplication is a tkinter app that allows windows users to convert
docx and docs file to pdf, because most people prefer to work with
pdf docs than word docs therefore this app can help them to convert
their docs to make them user friendly.
"""
class ConvertingError(Exception):
    pass
def notification(message):
    try:
        notification = ToastNotifier()
        notification.show_toast(title="THE DOCX TO PDF CONVERTOR", msg=message, duration=8, icon_path='main.ico')
    except Exception as e:
        pass
    return
try:
    import docx2pdf
except ImportError:
    from  pip._internal import main as install
    install(['install', 'docx2pdf'])
from tkinter import *
import datetime
from tkinter import messagebox, filedialog, scrolledtext
from win10toast import ToastNotifier
import tkinter.ttk as ttk
import os
root = Tk()
window_width = 610
window_height = 300
positionRight = int(root.winfo_screenwidth()/2 - window_width/2)
positionBottom = int(root.winfo_screenheight()/2 - window_height/2)
root.geometry('{}x{}+{}+{}'.format(window_width, window_height, positionRight, positionBottom))
root.resizable(False, False)
root.title("DOCX TO PDF CONVERTOR")
root.iconbitmap("main.ico")
def exitfunct():
    confirm = messagebox.askyesnocancel("CLOSING THE DOCX TO PDF CONVERTOR",
                                        "ARE YOU SURE YOU WANT TO CLOSE DOCX TO PDF CONVERTOR?")
    if(confirm):
        root.destroy()
        notification("Thank you for using THE DOCX TO PDF CONVERTOR.\n We have hope that you find this APP usefull. Give us feedback!!.-Crispen Gari")
        return
    else:
        root.focus()
        return
    return
def reset():
    messagebox.showwarning("RESETING THE DOCX TO PDF CONVERTOR", "RESETING THE DOCX TO PDF CONVERTOR WILL REMOVE ALL THE INFOMATION ON THE APP!")
    confirm = messagebox.askyesno("RESETING THE DOCX TO PDF CONVERTOR",
                                        "ARE YOU SURE YOU WANT TO RESERT DOCX TO PDF CONVERTOR?")
    if(confirm):
        docx_path.config(state=NORMAL)
        docx_path.delete(0, END)
        docx_path.config(state=DISABLED)
        output_path['state'] = NORMAL
        output_path.delete(0, END)
        output_path.insert(0, str(os.getcwd()))
        output_path['state'] = DISABLED
    else:
        root.focus()
        return
    return
def convert():
    file_name = docx_path.get()
    saving_path = output_path.get()
    output_file =  os.path.join(os.getcwd(), f'{datetime.datetime.now().date()}-T-H-{datetime.datetime.now().hour}-M-{datetime.datetime.now().minute}-S-{datetime.datetime.now().hour}.pdf')
    print(output_file)
    try:
        if(file_name):
            if(saving_path):
               from docx2pdf import convert
               convert(file_name, output_file)
               messagebox.showinfo("THE DOCX TO PDF CONVERTOR",
                                   "A document has been created and saved in {}.".format(saving_path))
            else:
                raise ConvertingError("Saving directory can not be empty!")
        else:
            raise ConvertingError("You must choose at least one file to convert!")

    except ConvertingError as e:
        messagebox.showerror("THE DOCX TO PDF CONVERTOR", e)
        return
    return
def change():
    filename = filedialog.askdirectory(initialdir="/", title="CHOOSE SAVING DIRECTORY")
    if(filename):
        output_path['state'] = NORMAL
        output_path.delete(0, END)
        output_path.insert(0, str(filename))
        output_path['state'] = DISABLED
    else:
        root.focus()
        return
    return
def choose():
    filename= filedialog.askopenfilename(
                                         title="CHOOSE A DOCX FILE", filetypes=(('DOCX FILES', '*.docx'),
                                                                                ('DOCX FILES', '*.doc')))
    docx_path.config(state=NORMAL)
    docx_path.insert(0, filename)
    docx_path.config(state=DISABLED)
    return
Label(root, text="WELCOME TO THE DOCX TO PDF CONVETOR",
      font=('Arial', '12', 'bold')).grid(row=0, column=0, columnspan=5, sticky=W, padx=10)
Label(root, text="DOCOMENT PATH",
      font=('Arial', 10)).grid(row=1, column=0, sticky=W, padx=10)
Label(root, text="OUTPUT PATH",
      font=('Arial', 10)).grid(row=2, column=0,  sticky=W, padx=10)
docx_path = ttk.Entry(root, width=50, font=("arial", 10), state=DISABLED)
docx_path.grid(row=1, column=1,sticky=W, padx=10)
output_path = ttk.Entry(root, width=50, font=("arial", 10), state=DISABLED)
output_path['state'] =NORMAL
output_path.insert(0,str(os.getcwd()))
output_path['state'] = DISABLED
output_path.grid(row=2, column=1,sticky=W, padx=10)

output_screen = scrolledtext.ScrolledText(root, width=60, height=10, state=DISABLED)
output_screen.grid(row=3, column=0, columnspan=2, rowspan=3)
choose_btn = Button(root, text="Choose", bg="green", fg="#fff",
                    relief=SOLID,width=10, font=('arial', 10, 'bold'), borderwidth=1, command=choose)
choose_btn.grid(row=1, column=2,  pady=10)
change_btn =Button(root, text="Change", bg="green", fg="#fff",
                   relief=SOLID,width=10, font=('arial', 10, 'bold'), borderwidth=1, command=change)
change_btn.grid(row=2, column=2, pady=10)

convert_btn =Button(root, text="Convert", bg="green", fg="#fff",
                   relief=SOLID,width=10, font=('arial', 10, 'bold'), borderwidth=1, command=convert)
convert_btn.grid(row=3, column=2, pady=10)
reset_btn =Button(root, text="Reset", bg="green", fg="#fff",
                   relief=SOLID,width=10, font=('arial', 10, 'bold'), borderwidth=1, command=reset)
reset_btn.grid(row=4, column=2, pady=10)
exit_btn =Button(root, text="Close", bg="green", fg="#fff",
                   relief=SOLID,width=10, font=('arial', 10, 'bold'), borderwidth=1, command=exitfunct)
exit_btn.grid(row=5, column=2, pady=10)
root.mainloop()
