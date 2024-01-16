from tkinter import *
from tkinter import filedialog
from pdf2ppt import pdf2ppt
from pdf2docx import Converter
from tkinter import messagebox
root = Tk()
root.geometry("600x250")
root.title("Converter")
root.resizable(0,0)
root.config(bg="#ff5722")
root.iconbitmap("dw.ico")
image=PhotoImage(file ="logo.png").subsample(2)
limage=Label(root,image=image)
limage.config(bg="#ff5722")
limage.place(x=40,y=15)
#---------------lables---------------------
a1=Label(root,text="PDF to DOCX Converter",font=("Calibri",26,"bold"), fg
         ="black", bg ="#ff5722")
a1.place(x=120, y=30)
#------------------------------------------------
l1=Label(root,text="Select PDF File:" ,font=("Calibri",12,"bold"), fg = "black", bg ="#ff5722")
l1.place(x=20, y=100)
en1=Entry(root, width=55 ,bd=2 ,bg="#757575" )
en1.place(x=143, y=103)
#------------------------------------------------
status=Label(root,text= "Ready...",font=("Calibri",12,"italic"), fg="black" , bg="white", anchor="w")
status.place(rely=1,anchor="sw",relwidth=1)
#------------------------------------------------
def browsepdf():
    file_path = filedialog.askopenfilename()
    en1.delete(0, 'end')
    en1.insert(0, file_path)
    status.config(text="Ready...")
def convert_to_docx():
    input_file = en1.get()
    output_file = en1.get().replace(".pdf", ".docx")
    if not input_file:
        messagebox.showerror("Error", "No file selected")
        return
    try:
        status.config(text="Converting...")
        status.config(fg="black")
        status.config(bg="white")
        cv = Converter(input_file)
        cv.convert(output_file, start=0, end=None, encoding='UTF-8')
        cv.close()
        status.config(text="Conversion Complete")
        status.config(fg="green")
        status.config(bg="white")
    except Exception as e:
        messagebox.showerror("Error", str(e))
     
#------------------------------------------------
btn1=Button(root, text ="browse",bg="#37474f",fg="white", font=("Calibri",12),command=browsepdf)
btn1.place(x=500, y=100)
#------------------------------------------------
btn1=Button(root, text ="Convert to docx",bg="#37474f",fg="white", font=("Calibri",12),command=convert_to_docx)
btn1.place(x=250, y=140)
#------------------------------------------------
root.mainloop()
