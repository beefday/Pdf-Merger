import os
import glob
from PyPDF2 import PdfFileMerger, PdfReader, PdfWriter
from tkinter import filedialog

# import aspose.slides as slides
path = str(os.path.realpath(os.path.dirname(__file__)) + "/")

# def ppt_to_pdf():
#     for file in glob.glob(path + "*.pptx"):
#         prs = slides.Presentation(file)
#         prs.save(str(file) + '.pdf', slides.export.SaveFormat.PDF)


def pdf_merge():
    # List of PDF files to be merged
    print(path)
    pdf_files = []
    
    file_path = filedialog.askopenfilename(initialdir = "/", title = "Select Tex file", filetypes = (("Tex files", "*.tex"), ("all files", "*.*")))
    
    
    
   # for file in glob.glob(path + "*.pdf"):
   #     pdf_files.append(file)

    # Create an instance of PdfFileMerger
    merger = PdfFileMerger()

    # Iterate through the list of PDF files and add them to the merger
    for file in pdf_files:
        print(file)
        merger.append(file)
    final = str(path + "merged.pdf")
    # Save the merged PDF to a file
    merger.write(final)

    # Close the merger
    merger.close()


def compress_pdf():
    reader = PdfReader(final)
    writer = PdfWriter()

    for page in reader.pages:
        page.compress_content_streams()
        writer.add_page(page)

    with open(path + "mergercompressed.pdf", "wb") as f:
        writer.write(f)


if __name__ == '__main__':
    pdf_merge()
