import os
import sys

# pdf
from fpdf import FPDF

from datetime import date

# cover letter backend modules
from generation.cover_letter_utils import get_body_text, get_slash

class CoverLetter:
    
    def __init__(self):
        # internal variables
        self._curr_path = os.path.abspath(os.getcwd())
        os.chdir("./pdfs/")
        self._pdf_path = os.path.abspath(os.getcwd())
        # os.chdir("../docx_files")
        # self._docx_path = os.path.abspath(os.getcwd())
        os.chdir("..")
        self._MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
        
        # external variables
        self._full_name = ""
        self._phone_num = ""
        self._email = ""
        self._linkedin = ""
        self._company = ""
        self._cover_body_ptr = None
        self._cover_body_txt = ""
    
    def set_full_name(self, full_name:str):
        self._full_name = full_name

    def get_full_name(self):
        return self._full_name
    
    def set_phone_num(self, phone_num:str):
        self._phone_num = phone_num

    def get_phone_num(self):
        return self._phone_num
    
    def set_email(self, email:str):
        self._email = email
    
    def get_email(self):
        return self._email
    
    def set_linkedin(self, linkedin:str):
        self._linkedin = linkedin
    
    def get_linkedin(self):
        return self._linkedin
    
    def set_company(self, company:str):
        self._company = company
    
    def get_company(self):
        return self._company
    
    def set_cover_letter_ptr(self, cover_body_filename:str):
        try:
            cover_letter_body_ptr = open(cover_body_filename, "r")
            self._cover_body_ptr = cover_letter_body_ptr
            return True
        except FileNotFoundError:
            print("'{}' not found.".format(cover_body_filename))
            return False
        
    def set_cover_letter_body(self, cover_body_text:str):
        self._cover_body_txt = cover_body_text
        
    def get_cover_letter_body(self):
        return self._cover_body_txt
    
    def close_cover_letter_ptr(self):
        self._cover_body_ptr
    
    def generate_cover_letter(self, filename:str):

        # Add Header
        cover_letter_header_details = {
            "name": self._full_name,
            "phone_num": self._phone_num,
            "email": self._email,
            "linkedin": self._linkedin
        }
        PDF = self._cover_letter_header_section(cover_letter_header_details)

        # Add Body
        curr_date = date.today()
        curr_month = self._MONTHS[curr_date.month-1]
        curr_day = curr_date.day
        curr_year = curr_date.year
        cover_letter_body_details = {
            "curr_date": "{} {}, {}".format(curr_month, curr_day, curr_year),
            "company": self._company,
            "name": self._full_name
        }
        body_txt = get_body_text(self._cover_body_txt, cover_letter_body_details)
        pdf = self._cover_letter_body_section(PDF, body_txt)

        # Name or rename the docx
        file_no_type = os.path.splitext(" " + filename)[0].strip()
        file_rename = None
        if(file_no_type):
            file_rename = os.path.basename(file_no_type)
        else:
            file_rename = "auto_named_file"
            print("The given file name is invalid, so file has been automatically named '{}'.".format(file_rename))
            
        # Save the docx with new name
        if(file_rename):
            
            # Assign name to pdf
            pdf_filename = "{}{}".format(file_rename, ".pdf")    
            DIR_SLASH = get_slash(sys.platform)
            dest_pdf_path = "{}{}{}".format(self._pdf_path, DIR_SLASH, pdf_filename)
            
            # Generate the pdf file
            pdf.output(dest_pdf_path)
            
            return pdf_filename
        
        return file_rename+".pdf" # will never occur
    
    ''' Header Section '''
    def _cover_letter_header_section(self, cover_letter_details:dict):
        
        class PDF(FPDF):
            def header(self):
                
                self.add_font("Calibri", fname=r"./generation/calibri.ttf", uni=True)
                self.add_font("Calibri B", fname=r"./generation/calibrib.ttf", uni=True)
                
                # Line break
                self.ln(3.3)
                
                # Full Name
                self.set_font(family="Calibri B", size=13)
                self.cell(w=0, h=5, txt=cover_letter_details["name"], border=0, fill=0, align="C")
                
                # Line break
                self.ln(5.5)
                
                # Details
                cover_letter_details_txt = "{} | {} | {}".format(
                    cover_letter_details["phone_num"],
                    cover_letter_details["email"],
                    cover_letter_details["linkedin"]
                )
                self.set_font(family="calibri", size=10)
                self.cell(w=0, h=5, txt=cover_letter_details_txt, border=0, fill=0, align="C")
                
                self.link(x=71.9, y=21, w=24.5, h=1, link=cover_letter_details["email"])
                self.link(x=103.5, y=21, w=57, h=1, link=cover_letter_details["linkedin"])
                
                # Line break
                self.ln(8.5)
        return PDF

    ''' Body Content Section '''
    def _cover_letter_body_section(self, PDF, cover_letter_text:str):
        
        # save FPDF() class into a variable pdf
        pdf = PDF(format="Letter")

        # Styling
        pdf.l_margin = 24.4
        pdf.r_margin = 24.4
        
        # Add a page
        pdf.add_page()

        # set style and size of font that you want in the pdf
        pdf.set_font(family="calibri", size=11)

        # create a cell
        pdf.multi_cell(0, 5.5, txt=cover_letter_text , border=0, align = 'L')

        return pdf

    
# cv = CoverLetter()
# cv.set_full_name("Jan Carlos Rubio Sánchez")
# cv.set_phone_num("830-421-0344")
# cv.set_email("jcaj750@gmail.com")
# cv.set_linkedin("https://www.linkedin.com/in/jan-carlos-rubio-sanchez/")
# cv.set_company("TACC")
# cv.set_cover_letter_ptr("/mnt/c/Users/jan/Desktop/Files/Personal/Fun_Projects/Cover_Letter_Generator/cover_letters/cover_letter_body.txt")
# cv.generate_cover_letter()
# cv.close_cover_letter_ptr()