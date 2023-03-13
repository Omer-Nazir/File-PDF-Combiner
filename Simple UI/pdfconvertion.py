from fpdf import FPDF
from PyPDF2 import PdfReader, PdfMerger
import PIL.Image as Image
from win32com import client
import os
import tempfile
from shutil import rmtree

VALID_WORD_EXTENSIONS = ['.doc', '.docm', '.docx', '.dot', '.dotm', '.dotx', '.odt', '.rtf', '.txt', '.wps', '.xml', '.xps']
VALID_IMAGE_EXTENSIONS = ['.jpeg', '.jpg', '.png', '.gif']
VALID_EXCEL_EXTENSIONS = ['.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xls', '.xlt', '.xml', '.xlam', '.xla', '.xlw', '.xlr']
VALID_PPT_EXTENSIONS = ['.ppt','.pptx']
VALID_FILE_EXTENSIONS = VALID_EXCEL_EXTENSIONS + VALID_IMAGE_EXTENSIONS + VALID_WORD_EXTENSIONS + VALID_PPT_EXTENSIONS + ['.pdf']

class pdfmerging():
    def __init__(self):
        self.merger = PdfMerger()
        self.pptapp_bool = False
        self.tempfolder = os.path.join(tempfile.gettempdir(),'pdfconverter')

        try:
            rmtree(self.tempfolder)
        except:
            pass
        
        try:
            os.mkdir(self.tempfolder)
        except:
            pass

        self.MergedFile_path = os.path.join(self.tempfolder,'Merged.pdf')
        self.mergingExceptions = []
        self.closingException = ''

    def addfile(self,filepath,bookmark,addbm):
        filepath = os.path.abspath(filepath)
        try:
            fileext = self.file_ext(filepath)
            if fileext != '.pdf':
                filepath = self.convertfile(filepath,fileext)
            if addbm and bookmark != '':
                with open(filepath,'rb') as file:
                    self.merger.append(PdfReader(file),bookmark)
            else:
                with open(filepath,'rb') as file:
                    self.merger.append(PdfReader(file))
            return True
        except Exception as e:
            self.mergingExceptions.append((filepath,e))
            return False
        
    def convertfile(self,filepath,fileext):
        temp_filepath = os.path.abspath(os.path.join(self.tempfolder,os.path.splitext(os.path.basename(filepath))[0])+'.pdf')
        if fileext in VALID_EXCEL_EXTENSIONS:
            self.convert_excel(filepath,temp_filepath)
        elif fileext in VALID_IMAGE_EXTENSIONS:
            self.convert_image(filepath,temp_filepath)
        elif fileext in VALID_PPT_EXTENSIONS:
            self.convert_ppt(filepath,temp_filepath)
        elif fileext in VALID_WORD_EXTENSIONS:
            self.convert_word(filepath,temp_filepath)
        return temp_filepath

    def close_merger_task(self,destinationfile):
        if destinationfile == '':
            return True
        try:
            self.merger.write(destinationfile)
        except Exception as e:
            self.closingException = e
            return False
        try:
            self.merger.close()
            if self.pptapp_bool: self.openppt.Application.Quit()
            rmtree(self.tempfolder)
            return True
        except Exception as e:
            return 'Process Partially Success'

    def file_ext(self,filepath):
        return os.path.splitext(os.path.basename(filepath))[1]
        
    def convert_ppt(self, file_path: str, outputpdf_path: str):
        pptapp = client.Dispatch("Powerpoint.Application")
        # opens the ppt document
        openppt = pptapp.Presentations.Open(file_path,ReadOnly=1,WithWindow=0)

        # exports the ppt documet to pdf
        try:
            openppt.SaveAs(outputpdf_path, 32)
        except:
            openppt.PrintOut(PrintToFile=outputpdf_path)
            enable_exit = False
            while not enable_exit:
                try:
                    if len(PdfReader(open(outputpdf_path,'rb')).pages) > 0:
                        enable_exit = True
                    continue
                except: 
                    continue
        
        pptapp.Quit()


# Following Functions were Taken From Python pdfconverter:
# https://pypi.org/project/pdfconverter/#description

    def convert_image(self, filename: str, pdfname: str):
        # open the image
        image = Image.open(filename)

        # get the size of the image in px
        width, height = image.size

        # close the image
        image.close()

        # convert the size from px to mm
        relation_px_mm = 0.26
        mmWidth = width * relation_px_mm
        mmHeight = height * relation_px_mm
        relation = mmWidth / mmHeight

        # the size of a a4 pdf file, aka the max size for the image
        margins = 48
        max_mmWidth = 216 - margins     
        max_mmHeight = 279 - margins

        # changes the size of the image if it is bigger than the pdf, while keeping the realtion of the picture
        if mmWidth > max_mmWidth:
            mmWidth = max_mmWidth
            mmHeight = mmWidth / relation

        if mmHeight > max_mmHeight:
            mmHeight = max_mmHeight
            mmWidth = mmHeight * relation

        # convert the image to a pdf with the sizes we found in mm
        pdf = FPDF()
        pdf.add_page()
        pdf.image(filename, w=mmWidth, h=mmHeight)
        # pdf.image(filename) 
        pdf.output(pdfname)
        
    def convert_excel(self, file_path: str, outputpdf_path: str):  
        excel = client.Dispatch("Excel.Application")

        # opens the file
        sheets = excel.Workbooks.Open(file_path,ReadOnly=1)


        # Following set of instructions were casusing Problem
        # so skipped it

        # # # set the formatations for the excel file
        # # ## sets landscape orientation
        # # # sheets.ActiveSheet.PageSetup.Orientation = 2
        # # # ## makes the excel sheet fit on one pdf sheet       
        # # # sheets.ActiveSheet.PageSetup.Zoom = False
        # # # sheets.ActiveSheet.PageSetup.FitToPagesWide = 1
        # # # sheets.ActiveSheet.PageSetup.FitToPagesTall = 1
        # # # worksheets = sheets.Worksheets[0]

        # export the excel sheet to a pdf
        sheets.ExportAsFixedFormat(0, outputpdf_path)
        excel.Application.Quit()

    def convert_word(self, file_path: str, outputpdf_path: str):
        word = client.Dispatch("Word.Application")
        # opens the word document
        word.Documents.Open(file_path,ReadOnly=1)

        word.Documents(1).Activate

        word.ActiveDocument.ExportAsFixedFormat(outputpdf_path, 17)

        word.Application.Quit()


