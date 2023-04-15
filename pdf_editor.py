from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter import messagebox
import os
from miner import PDFMiner
from miner import PDFOvervieuwer
from tkinter import PhotoImage
import PyPDF2
import ctypes

class PDFEditor:
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
        self.overvieuwtoggle = False
        self.mergepdftoggle = False
        self.overvieuwopend = 0
        self.master.iconbitmap("img/pdf.ico")
         # the bindings
        PDFSettings.bindings(self)

        # the menu
        self.menu = Menu(self.master)
        self.master.config(menu = self.menu)
        self.filemenu = Menu(self.menu)
        self.menu.add_cascade(label = "File", menu = self.filemenu)
        self.filemenu.add_command(label="Open File", accelerator='Ctrl+O' , command=self.open_file)
        self.filemenu.add_command(label="Exit", command=self.master.destroy)
        self.filemenu.add_command(label = "Save", command = self.save_file)
        self.filemenu.add_command(label = "Save As", command = self.save_file_as_window)
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
        self.overvieuw_frame_main = ttk.Frame(self.top_frame, width=580, height=650)
        self.overvieuw_frame_main.grid(row=0, column=0)
        self.overvieuw_frame_main.grid_propagate(False)
        self.zoomin_icon = PhotoImage(file = "img/zoomin2.png", width = 20, height = 20)
        self.zoomout_icon = PhotoImage(file = "img/zoomout2.png", width = 20, height = 20)
        self.open_icon = PhotoImage(file = "img/folder.png", width = 20, height = 20)
        self.save_icon = PhotoImage(file = "img/save2.png", width = 20, height = 20)
        self.merge_icon = PhotoImage(file = "img/merge.png", width = 20, height = 20)
        self.split_icon = PhotoImage(file = "img/split.png", width = 20, height = 20)

        self.open_button = ttk.Button(self.toolbar_frame, image = self.open_icon, command = self.open_file)
        self.open_button.grid(row = 0, column = 0)
        self.save_button = ttk.Button(self.toolbar_frame, image = self.save_icon, command = self.save_file)
        self.save_button.grid(row = 0, column = 1)
        self.zoomin_button = ttk.Button(self.toolbar_frame, image = self.zoomin_icon, command = self.zoom_in)
        self.zoomin_button.grid(row = 0, column = 2)
        self.zoomout_button = ttk.Button(self.toolbar_frame, image = self.zoomout_icon, command = self.zoom_out)
        self.zoomout_button.grid(row = 0, column = 3)   
        self.merge_button = ttk.Button(self.toolbar_frame, image = self.merge_icon, command = self.merge_pdf)
        self.merge_button.grid(row = 0, column = 4)
        self.split_button = ttk.Button(self.toolbar_frame, image = self.split_icon, command = self.split_pdf)
        self.split_button.grid(row = 0, column = 5)

        self.scrolly = Scrollbar(self.top_frame, orient = VERTICAL)
        self.scrolly.grid(row = 0, column = 1, sticky = (N, S))
        self.scrollx = Scrollbar(self.top_frame, orient = HORIZONTAL)
        self.scrollx.grid(row = 1, column = 0, sticky = (W, E))
        self.output = Canvas(self.top_frame, bg='lightgray', width = 560, height = 630)
        self.output.configure(yscrollcommand = self.scrolly.set, xscrollcommand = self.scrollx.set)
        self.output.grid(row = 0, column = 0)
        self.scrolly.config(command = self.output.yview)
        self.scrollx.config(command = self.output.xview)
        self.output.bind('<MouseWheel>', self.on_mousewheel)
        self.output.bind('<Alt-MouseWheel>', self.ver_on_mousewheel)

        # the buttons
        self.uparrow_icon = PhotoImage(file='img/uparrow.png')
        self.downarrow_icon = PhotoImage(file='img/downarrow.png')
        self.uparrow = self.uparrow_icon.subsample(3, 3)
        self.downarrow = self.downarrow_icon.subsample(3, 3)
        self.upbutton = ttk.Button(self.bottom_frame, image=self.uparrow, command=self.previous_page)
        self.upbutton.grid(row=0, column=1, padx=(270, 5), pady=8)
        self.downbutton = ttk.Button(self.bottom_frame, image=self.downarrow, command=self.next_page)
        self.downbutton.grid(row=0, column=3, pady=8)
        self.page_label = ttk.Label(self.bottom_frame, text='page')
        self.page_label.grid(row=0, column=4, padx=5)

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

    def display_page(self, scale=1):
        if 0 <= self.current_page < self.numPages: # type: ignore
            if self.zoom_level == 0:
                scale = 1
            else:
                scale = 1.2 ** self.zoom_level
            self.miner = PDFMiner(self.path)
            self.img_file = self.miner.get_page(self.current_page, scale=scale)
            self.output.create_image(0, 0, anchor='nw', image=self.img_file)
            self.stringified_current_page = self.current_page + 1 # type: ignore
            self.page_label['text'] = str(self.stringified_current_page) + ' of ' + str(self.numPages)
            region = self.output.bbox(ALL)
            self.output.configure(scrollregion=region)

    def zoom_in(self, event=None):
        if self.fileisopen:
            self.zoom_level += 1
            scale = 1.2 ** self.zoom_level
            if self.overvieuwtoggle == True:
                self.overvieuw(scale=scale, change = "yes")
            else:
                self.display_page(scale=scale)

    def zoom_out(self, event=None):
        if self.fileisopen:
            self.zoom_level -= 1
            scale = 1.2 ** self.zoom_level
            if self.overvieuwtoggle == True:
                self.overvieuw(scale=scale, change = "yes")
            else:
                self.display_page(scale=scale)

    def zoom_reset(self, event=None):
        if self.fileisopen:
            self.zoom_level = 0
            scale = 1.2 ** self.zoom_level
            if self.overvieuwtoggle == True:
                self.overvieuw(scale=scale, change = "yes")
            else:
                self.display_page(scale=scale)

    def on_mousewheel(self, event):
        # Scroll the canvas and output widget together
        self.output.yview_scroll(-1 * (event.delta // 120), "units")

    def ver_on_mousewheel(self, event):
        # Scroll the canvas and output widget together
        self.output.xview_scroll(-1 * (event.delta // 120), "units")

    
    # a overvieuw that you can see 8 pages at the same time
    def overvieuw(self, event=None, scale=1, pagenumber=0, change = "no"):
        if change == "yes" and self.overvieuwtoggle == True:
            self.overvieuwtoggle = False

        if self.overvieuwtoggle == False:
            if not self.fileisopen:
                messagebox.showerror("Error", "No file is open")
                return
            if self.overvieuwopend != 0:
                self.checkbox.destroy()
            width = int(self.top_frame.winfo_width())
            height = int(self.top_frame.winfo_height())
            try:
                self.over = PDFOvervieuwer(self.path, scale=scale)
                totalpages = int(self.over.get_total_pages())
                self.img_file_overvieuw = []

                for i in range(totalpages):
                    self.img_file_overvieuw.append(self.over.pages(i))

                distance = (200 * scale)
                height_distance = (300 * scale)
                checkbox_distance = (250 * scale)
                checkboxx_distance = (50 * scale)
                total_images_per_row = int(width / distance) 
                total_images_per_column = int(height / height_distance)
                rangey = []
                rangex = []
                rangexcheckbox = []
                rangeycheckbox = []
                range_for_output = min(totalpages, total_images_per_row * total_images_per_column)

                for j in range(total_images_per_column):
                    for i in range(total_images_per_row):
                        rangey.append(j * height_distance)
                        rangex.append(i * distance)
                        rangeycheckbox.append(j * height_distance + checkbox_distance)
                        rangexcheckbox.append(i * distance + checkboxx_distance)



                
                    
                self.output.delete('all')
                self.output = Canvas(self.top_frame, bg='lightgray', width=width, height=height)
                self.output.grid(row=0, column=0)
                
                for i in range(range_for_output if range_for_output < 15 else 15):
                    self.output.create_image(rangex[i], rangey[i], image=self.img_file_overvieuw[i], anchor=NW)
                    self.checkbox = ttk.Checkbutton(self.output, text="page " + str(i + 1), onvalue=1, offvalue=0,)
                    self.checkbox.place(x=rangexcheckbox[i], y=rangeycheckbox[i])
                
            except Exception as e:
                print(e)
            
            self.overvieuwtoggle = True
            self.overvieuwopend += 1
        else:
            self.output.destroy()
            self.output = Canvas(self.top_frame, bg='lightgray', width=1150, height=730)
            self.output.grid(row=0, column=0)
            self.output.bind("<MouseWheel>", self.on_mousewheel)
            self.output.bind("<Button-4>", self.zoom_in)
            self.output.bind("<Button-5>", self.zoom_out)
            self.output.bind("<Button-3>", self.zoom_reset)
            self.output.bind("<Button-1>", self.zoom_reset)
            self.output.bind("<B1-Motion>", self.ver_on_mousewheel)
            self.display_page()
            self.overvieuwtoggle = False

    # a function to merge pdf files
    def merge_pdf(self):
        w = self.master.winfo_width()
        h = self.master.winfo_height()
        if self.fileisopen == False:
            messagebox.showerror("Error", "No file is open")
            return
        if self.mergepdftoggle == False:
            self.mergepdftoggle = True
            self.merge_window = ttk.Frame(self.top_frame, width = 80, height = 650)
            self.top_frame.config(width = w, height = h-100)
            self.output.config(width = w-100, height = h-150)
            self.merge_window.grid(row = 0, column = 0)
            self.output.grid(row = 0, column = 1, padx = 5, pady = 5)
            
            
            self.merge_label = ttk.Label(self.merge_window, text = "Enter the name of the merged file")
            self.merge_label.grid(row = 0, column = 0, padx = 5, pady = 5)
            self.merge_entry = ttk.Entry(self.merge_window)
            self.merge_entry.grid(row = 1, column = 0, padx = 5, pady = 5)
            self.merge_button = ttk.Button(self.merge_window, text = "Merge", command = self.merge)
            self.merge_button.grid(row = 2, column = 0, padx = 5, pady = 5)    
        else:
            self.merge_window.destroy()
            self.mergepdftoggle = False
            self.top_frame.config(width = 1150, height = 730)
            self.output.config(width = 1150, height = 730)
            self.output.grid(row = 0, column = 0)
            self.display_page()

    #a function to split pdf files
    def split_pdf(self):
        #if files is not open give error
        if not self.fileisopen:
            messagebox.showerror("Error", "No file is open")
            return
        

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

    def save_file_as_window(self):
        filename = fd.asksaveasfilename()
        if filename:
            self.save_file_as(filename)


    def resize(self, event=None):
        PDFSettings.resize(self, event)
    
    def open_file(self, event = None, file = None):
        PDFSettings.open_file(self, event, file)
    
    def merge(self):
        PDFSettings.merge(self)

    def split(self):
        PDFSettings.split(self)

    def save_file(self):
        PDFSettings.save_file(self)

    def save_file_as(self, new_doc_name):
        self.miner.save_as(new_doc_name)
        messagebox.showinfo("Success", "The file has been saved successfully")
        self.save_window.destroy()
        self.open_file(filepath=(new_doc_name + ".pdf"))
    

class PDFSettings(PDFEditor):
    def bindings(self):
        self.master.bind('<Control-Shift-O>', self.open_file)
        self.master.bind('<Control-o>', self.open_file)
        self.master.bind('<Up>', self.previous_page)
        self.master.bind('<Down>', self.next_page)
        self.master.bind('<Control-equal>', self.zoom_in)
        self.master.bind('<Control-minus>', self.zoom_out)
        self.master.bind('<Control-0>', self.zoom_reset)
        self.master.bind('<Control-s>', self.save_file)
        self.master.bind('<Configure>', self.resize)

    
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

    
            

    def resize(self, event=None):
        w = self.master.winfo_width()
        h = self.master.winfo_height()

        # update the size of the canvas and scrollbar
        if self.mergepdftoggle:
            self.top_frame.config(width=w, height=h-100)
            self.output.config(width=w-100, height=h-200)
            self.scrollx.config(width=w-40)
            self.scrolly.config(width=w-40)
        else:
            self.top_frame.config(width=w, height=h-100)
            self.output.config(width=w-20, height=h-150) 
            self.scrollx.config(width=w-40)
            self.scrolly.config(width=w-40)

        # update the size of the toolbar
        self.toolbar_frame.config(width=w, height=50)
        self.save_button.place(x=10, y=10)
        self.open_button.place(x=60, y=10)
        self.zoomin_button.place(x=110, y=10)
        self.zoomout_button.place(x=160, y=10)
        self.merge_button.place(x=210, y=10)
        self.split_button.place(x=260, y=10)

ctypes.windll.shcore.SetProcessDpiAwareness(1)    
root = Tk()
app = PDFEditor(root)
root.mainloop()