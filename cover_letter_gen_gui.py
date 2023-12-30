from tkinter import Tk, messagebox, filedialog, Menu, Label, Text, Button, Entry, END
from tkinter import *
from tkinter.simpledialog import askstring
from cover_letter import CoverLetter

class CoverLetterGenGUI:

    """ Display main menu of Cover Letter Generator GUI. """
    def __init__(self, columns:int):
        self.root = Tk()
        self._columns = columns

        ''' Cover Letter Initialization '''
        internal_current_cover_letter = CoverLetter()
        self._curr_cover_letter = internal_current_cover_letter

        ''' window and style '''

        self.root.geometry("800x900")
        self.root.title("Cover Letter Generator")
        self.root.configure(background="#222222")
        self.root.columnconfigure(tuple(range(self._columns)), weight=1)
        self.root.rowconfigure(tuple(range(20)), weight=1)
        self.root.grid()

        title_label = Label(self.root, text="Generate a cover letter pdf!", font="Arial 16 bold", fg="#fff", bg="#222222")
        title_label.grid(row=0, column=0, columnspan=self._columns, rowspan=2)
        # label.pack(side="top")

        # cover_empty_space = Label(self.root, text="", bg="#222222")
        # cover_empty_space.grid(row=1, column=0) # empty space

        ''' buttons and functionality '''     

        # menu bar
        self.menu_bar = Menu(self.root)
        self.file_menu = Menu(self.menu_bar, tearoff=0) # create bar

        self.file_menu.add_command(label="Close", command=self.on_closing)  # close button
        self.file_menu.add_command(label="Open File", command=self.open_file) # open user file
        self.file_menu.add_command(label="Save File as...", command=self.save_file) # save user file

        self.menu_bar.add_cascade(menu = self.file_menu, label="File...")  # name of bar option
        self.root.config(menu=self.menu_bar)

        # text area for company details

        self.cover_full_name_label = Label(self.root, text="Your Name:", font=("Arial", 13), fg="#fff", bg="#222222")
        self.cover_full_name_label.grid(row=2, column=0, padx=(50,0), sticky="W")
        full_name_text = StringVar()
        full_name_text.set("Jan Carlos Rubio Sánchez") # DEFAULT
        internal_current_cover_letter.set_full_name("Jan Carlos Rubio Sánchez")
        def set_full_name_text(self, var1, var2):
            new_full_name_value = full_name_text.get()
            internal_current_cover_letter.set_full_name(new_full_name_value)
        full_name_text.trace_add("write", set_full_name_text)
        self.cover_full_name_text_box = Entry(self.root, textvariable=full_name_text, font=("Times New Roman", 13))
        self.cover_full_name_text_box.grid(row=2, column=1, columnspan=self._columns-1, padx=(0,50), sticky="EW") # name
        
        self.cover_phone_label = Label(self.root, text="Phone Number:", font=("Arial", 13), fg="#fff", bg="#222222")
        self.cover_phone_label.grid(row=3, column=0, padx=(50,0), sticky="W")
        phone_text = StringVar()
        phone_text.set("830-421-0344") # DEFAULT
        internal_current_cover_letter.set_phone_num("830-421-0344")
        def set_phone_text(self, var1, var2):
            new_phone_value = phone_text.get()
            internal_current_cover_letter.set_phone_num(new_phone_value)
        phone_text.trace_add("write", set_phone_text)
        self.cover_phone_text_box = Entry(self.root, textvariable=phone_text, font=("Times New Roman", 13))
        self.cover_phone_text_box.grid(row=3, column=1, columnspan=self._columns-1, padx=(0,50), sticky="EW") # phone
        
        self.cover_email_label = Label(self.root, text="Email:", font=("Arial", 13), fg="#fff", bg="#222222")
        self.cover_email_label.grid(row=4, column=0, padx=(50,0), sticky="W")
        email_text = StringVar()
        email_text.set("jcaj750@gmail.com") # DEFAULT
        internal_current_cover_letter.set_email("jcaj750@gmail.com")
        def set_email_text(self, var1, var2):
            new_email_value = email_text.get()
            internal_current_cover_letter.set_email(new_email_value)
        email_text.trace_add("write", set_email_text)
        self.cover_email_text_box = Entry(self.root, textvariable=email_text, font=("Times New Roman", 13))
        self.cover_email_text_box.grid(row=4, column=1, columnspan=self._columns-1, padx=(0,50), sticky="EW") # email
        
        self.cover_linkedin_label = Label(self.root, text="Linkedin Link:", font=("Arial", 13), fg="#fff", bg="#222222")
        self.cover_linkedin_label.grid(row=5, column=0, padx=(50,0), sticky="W")
        linkedin_text = StringVar()
        linkedin_text.set("https://www.linkedin.com/in/jan-carlos-rubio-sanchez/") # DEFAULT
        internal_current_cover_letter.set_linkedin("https://www.linkedin.com/in/jan-carlos-rubio-sanchez/")
        def set_linkedin_text(self, var1, var2):
            new_linkedin_value = linkedin_text.get()
            internal_current_cover_letter.set_linkedin(new_linkedin_value)
        linkedin_text.trace_add("write", set_linkedin_text)
        self.cover_linkedin_text_box = Entry(self.root, textvariable=linkedin_text, font=("Times New Roman", 13))
        self.cover_linkedin_text_box.grid(row=5, column=1, columnspan=self._columns-1, padx=(0,50), sticky="EW") # linkedin

        self.cover_company_label = Label(self.root, text="Company you're applying for:", font=("Arial", 13), fg="#fff", bg="#222222")
        self.cover_company_label.grid(row=6, column=0, padx=(50,0), sticky="W")
        company_text = StringVar()
        def set_company_text(var, idx, mode):
            new_company_value = company_text.get()
            internal_current_cover_letter.set_company(new_company_value)
        company_text.trace_add("write", set_company_text)
        self.cover_company_text_box = Entry(self.root, textvariable=company_text, font=("Times New Roman", 13))
        self.cover_company_text_box.grid(row=6, column=1, columnspan=self._columns-1, padx=(0,50), sticky="EW") # company

        cover_empty_space = Label(self.root, text="", bg="#222222")
        cover_empty_space.grid(row=7, column=0) # empty space

        # text area for cover letter body
        def set_cover_body_text(new_cover_body):
            new_cover_body_value = new_cover_body.get("1.0", END).rstrip()
            internal_current_cover_letter.set_cover_letter_body(new_cover_body_value)
        self.cover_body_text_box = Text(self.root, font=("Times New Roman", 13))
        self.cover_body_text_box.bind("<KeyRelease>", lambda *args: set_cover_body_text(self.cover_body_text_box))
        self.cover_body_text_box.grid(row=8, columnspan=self._columns, rowspan=11, padx=(50), sticky="NESW")
        # self.cover_body_text_box.pack(side="top", fill="both", expand=True)

        # button to generate cover letter
        self.submit_btn = Button(self.root, text="generate cover letter", font=("Arial", 13), command=self.generate_cover_letter)
        self.submit_btn.grid(row=19, columnspan=self._columns)
        # self.submit_btn.pack(side="bottom")

        # exit
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        ''' display '''
        self.root.mainloop()
    
    """ Create a prompt in case user accidentally closes. """
    def on_closing(self):
        if((not self.cover_body_text_box.edit_modified()) and (len(self.cover_body_text_box.get("1.0", END)) == 1)):
            self._curr_cover_letter.close_cover_letter_ptr()
            self.root.destroy()
        else:
            if(self.cover_body_text_box.edit_modified()):
                if(messagebox.askyesno(title="Quit?", message="You haven't saved. Are you sure?")):
                    self._curr_cover_letter.close_cover_letter_ptr()
                    self.root.destroy()
            else:
                self._curr_cover_letter.close_cover_letter_ptr()
                self.root.destroy()

    """ Print user file from 'user_files' directory into text field. """
    def open_file(self):
        if(self.cover_body_text_box.edit_modified()):
            if(messagebox.askyesno(title="Text has not been saved...", message="Are you sure?")):
                self.root.filename = filedialog.askopenfile(initialdir="./content", title="Select a file...", filetypes=(("Text Files", "*.txt"),))
                if(self.root.filename):
                    file_dir = open(self.root.filename.name, "r")
                    file_txt = file_dir.read()
                    file_dir.close()
                    self.cover_body_text_box.delete("1.0", "end")
                    self.cover_body_text_box.insert("1.0", file_txt)
                    self.cover_body_text_box.edit_modified(False)

                    self._curr_cover_letter.set_cover_letter_ptr(self.root.filename.name)
                    self._curr_cover_letter.set_cover_letter_body(self.cover_body_text_box.get("1.0", END).rstrip())

        else:
            self.root.filename = filedialog.askopenfile(initialdir="./content", title="Select a file...", filetypes=(("Text Files", "*.txt"),))
            if(self.root.filename):
                file_dir = open(self.root.filename.name, "r")
                file_txt = file_dir.read()
                file_dir.close()
                if(len(self.cover_body_text_box.get("1.0", END)) == 1):
                    self.cover_body_text_box.insert("1.0", file_txt)
                else:
                    self.cover_body_text_box.delete("1.0", "end")
                    self.cover_body_text_box.insert("1.0", file_txt)
                self.cover_body_text_box.edit_modified(False)

                self._curr_cover_letter.set_cover_letter_ptr(self.root.filename.name)
                self._curr_cover_letter.set_cover_letter_body(self.cover_body_text_box.get("1.0", END).rstrip())
    
    def save_file(self):
        cover_body = self.cover_body_text_box.get("1.0", END)
        if(len(cover_body) != 1):
            self.root.filename = filedialog.asksaveasfile(initialdir="../cover_letters", title="Save file...", filetypes=(("Text Files", "*.txt"),))
            if(self.root.filename):
                file_dir = open(self.root.filename.name + ".txt", "w")
                file_dir.write(cover_body)
                file_dir.close()
                self.cover_body_text_box.edit_modified(False)

    def generate_cover_letter(self):
        # cover_letter_filename = askstring(title="File Name", prompt="What would you like to name your file?")
        
        new_pdf_name = filedialog.asksaveasfilename(initialdir="./pdfs", title="Save pdf as...", filetypes=(("PDF", "*.pdf"),))
        # filedialog.askopenfile(initialdir="./content", title="Select a file...", filetypes=(("Text Files", "*.txt"),))
        
        # self._curr_cover_letter.generate_cover_letter(cover_letter_filename)
        self._curr_cover_letter.generate_cover_letter(new_pdf_name)

CoverLetterGenGUI(10)