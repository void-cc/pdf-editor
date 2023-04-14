# this is for doing some math operations
import math
# this is for handling the PDF operations
import fitz
# importing PhotoImage from tkinter
from tkinter import PhotoImage



class PDFMiner:
    def __init__(self, filepath):
        # creating the file path
        self.filepath = filepath
        # opening the pdf document
        self.pdf = fitz.open(self.filepath)
        # loading the first page of the pdf document
        self.first_page = self.pdf.load_page(0)
        # getting the height and width of the first page
        self.width, self.height = self.first_page.rect.width, self.first_page.rect.height
        # initializing the zoom values of the page
        zoomdict = {800:0.8, 700:0.6, 600:1.0, 500:1.0}
        # getting the width value
        width = int(math.floor(self.width / 100.0) * 100)
        # zooming the page
        self.zoom = zoomdict[width]
        
    # this will get the metadata from the document like 
    # author, name of document, number of pages  
    def get_metadata(self):
        # getting metadata from the open PDF document
        metadata = self.pdf.metadata
        # getting number of pages from the open PDF document
        numPages = self.pdf.page_count
        # returning the metadata and the numPages
        return metadata, numPages
    
    # the function for getting the page
    def get_page(self, page_num, scale=1):
        # loading the page
        page = self.pdf.load_page(page_num)
        # checking if zoom is True
        if self.zoom:
            # creating a Matrix whose zoom factor is self.zoom
            mat = fitz.Matrix((self.zoom)* scale, (self.zoom)* scale)
            # gets the image of the page
            pix = page.get_pixmap(matrix=mat)
        # returns the image of the page  
        else:
            pix = page.get_pixmap()
        # a variable that holds a transparent image
        px1 = fitz.Pixmap(pix, 0) if pix.alpha else pix
        # converting the image to bytes
        imgdata = px1.tobytes("ppm")
        scale = 0.5
        # returning the image d aata
        return PhotoImage(data=imgdata)
    
    
    # function to get text from the current page
    def get_text(self, page_num):
        # loading the page
        page = self.pdf.load_page(page_num)
        # getting text from the loaded page
        text = page.getText('text')
        # returning text
        return text
    
    # function to get the text from the whole document
    def get_name(self):
        # getting the metadata
        filename = self.get_metadata()[0]['title']
        
        return filename
    
    def save_file(self, name):
        # saving the file
        self.pdf.save(self.filepath, garbage=4, deflate=True, clean=True)

    def save_as(self, name):
        # saving the file
        if name[-4:] != '.pdf':
            name += '.pdf'
        self.pdf.save(name, garbage=4, deflate=True, clean=True)

class PDFOvervieuwer(PDFMiner):
    def __init__(self, filepath, scale=1):
        self.filepath = filepath
        self.pdf = fitz.open(self.filepath)
        self.first_page = self.pdf.load_page(0)
        self.width, self.height = self.first_page.rect.width, self.first_page.rect.height
        zoomdict = {800: 0.3, 700: 0.2, 600: 0.3, 500: 0.3}
        zoomdict = zoomdict if scale == 1 else {k: v * scale for k, v in zoomdict.items()}
        width = int(math.floor(self.width / 100.0) * 100)
        self.zoom = zoomdict[width]


    def pages(self, page_num):
        page_num = int(page_num)
        page = self.pdf.load_page(page_num)
        if self.zoom:
            mat = fitz.Matrix(self.zoom, self.zoom)
            pix = page.get_pixmap(matrix=mat)
        else:
            pix = page.get_pixmap()
        px1 = fitz.Pixmap(pix, 0) if pix.alpha else pix
        imgdata = px1.tobytes("ppm")

        return PhotoImage(data=imgdata)
    
    def get_metadata(self):
        return super().get_metadata()
    
    def get_total_pages(self):
        return self.pdf.page_count