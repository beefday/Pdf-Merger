import os
import sys
import glob
import threading
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import tkinter as tk
import tkinter.messagebox

def compress_pdf():
    reader = PdfReader(path + "merged.pdf")
    writer = PdfWriter()

    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)

    with open(path + "mergercompressed.pdf", "wb") as f:
        writer.write(f)

class Drag_and_Drop_Listbox(tk.Listbox):
    def __init__(self, master, **kw):
        kw['selectmode'] = tk.MULTIPLE
        kw['activestyle'] = 'none'
        tk.Listbox.__init__(self, master, kw)
        self.bind('<Button-1>', self.getState, add='+')
        self.bind('<Button-1>', self.setCurrent, add='+')
        self.bind('<B1-Motion>', self.shiftSelection)
        self.curIndex = None
        self.curState = None
    def setCurrent(self, event):
        self.curIndex = self.nearest(event.y)
    def getState(self, event):
        i = self.nearest(event.y)
        self.curState = self.selection_includes(i)
    def shiftSelection(self, event):
        i = self.nearest(event.y)
        if self.curState == 1:
            self.selection_set(self.curIndex)
        else:
            self.selection_clear(self.curIndex)
        if i < self.curIndex:
            x = self.get(i)
            selected = self.selection_includes(i)
            self.delete(i)
            self.insert(i+1, x)
            if selected:
                self.selection_set(i+1)
            self.curIndex = i
        elif i > self.curIndex:
            x = self.get(i)
            selected = self.selection_includes(i)
            self.delete(i)
            self.insert(i-1, x)
            if selected:
                self.selection_set(i-1)
            self.curIndex = i

def pdf_merge():
    pdf_files_to_merge = []
    for i in range(listbox.size()):
        pdf_files_to_merge.append(listbox.get(i))
    
    merger = PdfMerger()

    for file in pdf_files_to_merge:
        merger.append(file)

    final = str(path + "merged.pdf")
    merger.write(final)
    merger.close()
    tkinter.messagebox.showinfo("Info", "PDFs merged successfully")

def loader():
    listbox.delete(0,listbox.size())
    
    if os.path.exists(path + "merged.pdf"):
        os.remove(path + "merged.pdf")

    pdf_files = []
    for file in glob.glob(path + "*.pdf"):
        pdf_files.append(file)
    
    pdf_files = sorted(pdf_files, key=lambda x: int(os.path.basename(x).split('.')[0]))
    
    for file in pdf_files:
        listbox.insert(tk.END, str(file))
    
def loader_thread():
    threading.Thread(target=loader).start()

def pdf_merge_thread():
    threading.Thread(target=pdf_merge).start()

if __name__ == '__main__':
    #path = str(os.path.realpath(os.path.dirname(__file__)) + "/")
    
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    path = os.path.join(application_path, "")
    #print(f"\"{path}\"")

    root = tk.Tk()
    root.geometry("600x400")
    listbox = Drag_and_Drop_Listbox(root)
    
    A = tk.Button(root, text ="Load", command = loader_thread)
    B = tk.Button(root, text ="Merge", command = pdf_merge_thread)
    #C = tk.Button(root, text ="Compress", command = compress_pdf)
    A.pack()
    B.pack()
    #C.pack()
    listbox.pack(fill=tk.BOTH, expand=True)
    root.mainloop()