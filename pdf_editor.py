from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox
import os
from miner import PDFMiner
from miner import PDFOvervieuwer
from tkinter import PhotoImage
import PyPDF2


class PDFViewer:
    def __init__(self, master):
        self.path = None
        self.fileisopen = None
        self.author = None
        self.name = None
        self.current_page = 0
        self.zoom_level = 0  # add this line
        self.numPages = None
        self.master = master
        self.master.title("PDF Viewer")
        self.master.geometry("580x730+540+180")
        self.master.iconbitmap(self.master, "pdf.ico")
        
         # the bindings
        self.master.bind('<Control-Shift-O>', self.open_file)
        self.master.bind('<Control-o>', self.open_file)
        self.master.bind('<Up>', self.previous_page)
        self.master.bind('<Down>', self.next_page)
        self.master.bind('<Control-equal>', self.zoom_in)
        self.master.bind('<Control-minus>', self.zoom_out)



        # the menu
        self.menu = Menu(self.master)
        self.master.config(menu = self.menu)
        self.filemenu = Menu(self.menu)
        self.menu.add_cascade(label = "File", menu = self.filemenu)
        self.filemenu.add_command(label="Open File", accelerator='Ctrl+O' , command=self.open_file)
        self.filemenu.add_command(label="Exit", command=self.master.destroy)
        self.filemenu.add_command(label = "Save", command = self.save_file)
        self.filemenu.add_command(label = "Save As")
        self.filemenu.add_command(label = "zoom in", accelerator = "Ctrl + +", command = self.zoom_in)
        self.filemenu.add_command(label = "zoom out", accelerator = "Ctrl + -", command = self.zoom_out)
        self.filemenu.add_command(label = "Close")
        self.optionsmenu = Menu(self.menu)
        self.menu.add_cascade(label = "Options", menu = self.optionsmenu)
        self.optionsmenu.add_command(label = "merge", command = self.merge_pdf)
        self.optionsmenu.add_command(label = "split", command = self.split_pdf)
        #self.optionsmenu.add_command(label = "rotate", command = self.rotate_pdf)
        self.optionsmenu.add_command(label = "overvieuw", command = self.overvieuw)        

        # the frames
        self.toolbar_frame = ttk.Frame(self.master, width = 580, height = 30)
        self.toolbar_frame.grid(row = 0, column = 0)
        self.toolbar_frame.grid_propagate(False)
        
        self.top_frame = ttk.Frame(self.master, width = 580, height = 650)
        self.top_frame.grid(row = 1, column = 0)
        self.top_frame.grid_propagate(False)
        self.bottom_frame = ttk.Frame(self.master, width = 580, height = 50)
        self.bottom_frame.grid(row = 2, column = 0)
        self.bottom_frame.grid_propagate(False)
        self.zoomin_icon = PhotoImage(file = "zoomin2.png", width = 20, height = 20)
        self.zoomout_icon = PhotoImage(file = "zoomout2.png", width = 20, height = 20)
        self.open_icon = PhotoImage(file = "folder.png", width = 20, height = 20)
        self.save_icon = PhotoImage(file = "save2.png", width = 20, height = 20)
        self.zoomin_button = ttk.Button(self.toolbar_frame, image = self.zoomin_icon, command = self.zoom_in)
        self.zoomin_button.grid(row = 0, column = 2)
        self.zoomout_button = ttk.Button(self.toolbar_frame, image = self.zoomout_icon, command = self.zoom_out)
        self.zoomout_button.grid(row = 0, column = 3)   
        self.open_button = ttk.Button(self.toolbar_frame, image = self.open_icon, command = self.open_file)
        self.open_button.grid(row = 0, column = 0)
        self.save_button = ttk.Button(self.toolbar_frame, image = self.save_icon, command = self.save_file)
        self.save_button.grid(row = 0, column = 1)
        self.scrolly = Scrollbar(self.top_frame, orient = VERTICAL)
        self.scrolly.grid(row = 0, column = 1, sticky = (N, S))
        self.scrollx = Scrollbar(self.top_frame, orient = HORIZONTAL)
        self.scrollx.grid(row = 1, column = 0, sticky = (W, E))
        self.output = Canvas(self.top_frame, bg='white', width = 560, height = 630)
        self.output.configure(yscrollcommand = self.scrolly.set, xscrollcommand = self.scrollx.set)
        self.output.grid(row = 0, column = 0)
        self.scrolly.config(command = self.output.yview)
        self.scrollx.config(command = self.output.xview)

        # the buttons
        self.uparrow_icon = PhotoImage(file='uparrow.png')
        self.downarrow_icon = PhotoImage(file='downarrow.png')
        self.uparrow = self.uparrow_icon.subsample(3, 3)
        self.downarrow = self.downarrow_icon.subsample(3, 3)
        self.upbutton = ttk.Button(self.bottom_frame, image=self.uparrow, command=self.previous_page)
        self.upbutton.grid(row=0, column=1, padx=(270, 5), pady=8)
        self.downbutton = ttk.Button(self.bottom_frame, image=self.downarrow, command=self.next_page)
        self.downbutton.grid(row=0, column=3, pady=8)
        self.page_label = ttk.Label(self.bottom_frame, text='page')
        self.page_label.grid(row=0, column=4, padx=5)
        self.merge_button = ttk.Button(self.bottom_frame, text = "Merge PDF", command = self.merge_pdf)
        self.merge_button.grid(row = 0, column = 5, padx = 5)
    
    def open_file(self, event=None, filepath=None):
        if filepath is None:
            filepath = fd.askopenfilename(title='Select a PDF file', filetypes=(('PDF', '*.pdf'), ))

        if filepath:
            self.path = filepath
            filename = os.path.basename(self.path)
            self.miner = PDFMiner(self.path)
            self.over = PDFOvervieuwer(self.path)
            data, numPages = self.miner.get_metadata()
            self.current_page = int(0)
            if numPages:
                self.name = data.get('title', filename[:-4])
                self.author = data.get('author', None)
                self.numPages = numPages
                self.fileisopen = True
                self.display_page()
                self.master.title(self.name)

    # a overvieuw that you can see 8 pages at the same time
    def overvieuw(self):
        self.overvieuw_window = Toplevel(self.master)
        self.overvieuw_window.title("Overvieuw")
        self.overvieuw_window.geometry("1200x750")
        self.overvieuw_window.resizable(True, True)
        self.overvieuw_window.iconbitmap(self.overvieuw_window, "pdf.ico")
        self.overvieuw_frame = ttk.Frame(self.overvieuw_window, width = 1200, height = 800)
        self.overvieuw_frame.grid(row = 0, column = 0)
        self.overvieuw_frame.grid_propagate(False)
        self.overvieuw_output = Canvas(self.overvieuw_frame, bg='white', width = 1150, height = 730)
        self.overvieuw_output.grid(row = 0, column = 0)
        try:
            totalpages = int(self.over.get_total_pages())
            self.img_file_overvieuw = []

            for i in range(totalpages):
                self.img_file_overvieuw.append(self.over.pages(i))

            rangey = [0, 0, 0, 0, 0, 300, 300, 300, 300, 300, 600, 600, 600, 600, 600]
            rangex = [0, 200, 400, 600, 800, 0, 200, 400, 600, 800, 0, 200, 400, 600, 800]
            rangeycheckbox = [250, 250, 250, 250, 250, 550, 550, 550, 550, 550, 850, 850, 850, 850, 850]
            rangexcheckbox = [50, 250, 450, 650, 850, 50, 250, 450, 650, 850, 50, 250, 450, 650, 850]

            for i in range(totalpages if totalpages < 15 else 15):
                self.overvieuw_output.create_image(rangex[i], rangey[i], image = self.img_file_overvieuw[i], anchor = NW)
                #ttk checkbutton per page already checked
                self.checkbox = ttk.Checkbutton(self.overvieuw_output, text = "page " + str(i + 1), onvalue = 1, offvalue = 0,  )
                self.checkbox.place(x = rangexcheckbox[i], y = rangeycheckbox[i])
        except:
            self.overvieuw_window.destroy()
            messagebox.showerror("Error", "No file is open")
            

        
    # a function to merge pdf files
    def merge_pdf(self):
        self.merge_window = Toplevel(self.master)
        self.merge_window.title("Merge PDF")
        self.merge_window.geometry("400x200")
        self.merge_window.resizable(False, False)
        self.merge_label = ttk.Label(self.merge_window, text = "Enter the name of the merged file")
        self.merge_label.grid(row = 0, column = 0, padx = 5, pady = 5)
        self.merge_entry = ttk.Entry(self.merge_window)
        self.merge_entry.grid(row = 1, column = 0, padx = 5, pady = 5)
        self.merge_button = ttk.Button(self.merge_window, text = "Merge", command = self.merge)
        self.merge_button.grid(row = 2, column = 0, padx = 5, pady = 5)

           

    def next_page(self, event=None):
        if self.fileisopen:
            if self.current_page <= self.numPages - 1:
                self.current_page += 1
                self.display_page()

    def previous_page(self, event=None):
        if self.fileisopen:
            if self.current_page > 0:
                self.current_page -= 1
                self.display_page()

    def merge(self):
        self.merge_name = self.merge_entry.get()
        if self.merge_name == "":
            messagebox.showerror("Error", "Please enter a name for the merged file")
        else:
            merger = PyPDF2.PdfMerger(strict=True)
            files = fd.askopenfilenames(title = "Select PDF files", initialdir = os.getcwd(), filetypes = (("PDF", "*.pdf"), ))
            if files:
                for file in files:
                    merger.append(file)
                merger.write(self.merge_name + ".pdf")
                merger.close()
                messagebox.showinfo("Success", "The PDF files have been merged successfully")
                self.merge_window.destroy()
                self.open_file(filepath=(self.merge_name + ".pdf"))

    def split_pdf(self):
        self.split_window = Toplevel(self.master)
        self.split_window.title("Split PDF")
        self.split_window.geometry("400x200")
        self.split_window.resizable(False, False)
        self.split_label = ttk.Label(self.split_window, text = "Enter the name of the split file")
        self.split_label.grid(row = 0, column = 0, padx = 5, pady = 5)
        self.split_entry = ttk.Entry(self.split_window)
        self.split_entry.grid(row = 1, column = 0, padx = 5, pady = 5)
        self.split_at = ttk.Label(self.split_window, text = "Split at")
        self.split_at.grid(row = 2, column = 0, padx = 1, pady = 1)
        self.split_pages = []
        for i in range(self.numPages):
            self.split_pages.append(str(i+1))
        self.split_combobox = ttk.Combobox(self.split_window, values = (self.split_pages))
        self.split_combobox.grid(row = 3, column = 0, padx = 5, pady = 5)
        self.split_upper_or_lower = ttk.Label(self.split_window, text = "Split upper or lower")
        self.split_upper_or_lower.grid(row = 4, column = 0, padx = 1, pady = 1)
        self.split_upper_or_lower = ttk.Combobox(self.split_window, values = ("Upper", "Lower"))
        self.split_upper_or_lower.grid(row = 5, column = 0, padx = 5, pady = 5)
        
        # split the upperhalf or the lowerhalf of the pdf file
        self.split_upper = ttk.Button(self.split_window, text = "Split", command = self.split)
        self.split_upper.grid(row = 6, column = 0, padx = 5, pady = 5)
        
    def split(self):
        self.split_name = self.split_entry.get()
        self.split_value = int(self.split_combobox.get())
        self.split_upper_or_lower = self.split_upper_or_lower.get()
        print(self.split_value)

        if self.split_name == "":
            messagebox.showerror("Error", "Please enter a name for the split file")
        elif self.split_value == "":
            messagebox.showerror("Error", "Please enter a page number to split at")
        elif self.split_upper_or_lower == "":
            messagebox.showerror("Error", "Please enter whether to split the upper or lower part")
        else:
            if self.split_upper_or_lower == "Upper":
                self.file = open(self.path, "rb")
                self.reader = PyPDF2.PdfReader(self.file)
                self.writer = PyPDF2.PdfWriter()
                for page in range(self.split_value):
                    self.current_page = self.reader.pages[page]
                    self.writer.add_page(self.current_page)
                self.writer.write(self.split_name + ".pdf")
                self.file.close()
            elif self.split_upper_or_lower == "Lower":
                self.file = open(self.path, "rb")
                self.reader = PyPDF2.PdfReader(self.file)
                self.writer = PyPDF2.PdfWriter()
                for page in range(self.split_value, self.numPages):
                    self.current_page = self.reader.pages[page]
                    self.writer.add_page(self.current_page)
                self.writer.write(self.split_name + ".pdf")
                self.file.close()
            messagebox.showinfo("Success", "The PDF file has been split successfully")
            self.split_window.destroy()
            self.open_file(filepath=(self.split_name + ".pdf"))

    def save_file(self):
        if self.fileisopen:
            self.name = self.miner.get_name()
            self.miner.save_file(self.name)
            messagebox.showinfo("Success", "The file has been saved successfully")

    def display_page(self, scale=1):
        if 0 <= self.current_page < self.numPages: # type: ignore
            self.img_file = self.miner.get_page(self.current_page, scale=scale)
            self.output.create_image(0, 0, anchor='nw', image=self.img_file)
            self.stringified_current_page = self.current_page + 1 # type: ignore
            self.page_label['text'] = str(self.stringified_current_page) + ' of ' + str(self.numPages)
            region = self.output.bbox(ALL)
            self.output.configure(scrollregion=region)

    def zoom_in(self, event=None):
        if self.fileisopen:
            self.zoom_level -= 1
            scale = 1.2 ** self.zoom_level
            self.display_page(scale=scale)

    def zoom_out(self, event=None):
        if self.fileisopen:
            self.zoom_level += 1
            scale = 1.2 ** self.zoom_level
            self.display_page(scale=scale)

    def zoom_reset(self, event=None):
        if self.fileisopen:
            self.zoom_level = 0
            scale = 1.2 ** self.zoom_level
            self.display_page(scale=scale)

    
root = Tk()
app = PDFViewer(root)
root.mainloop()
