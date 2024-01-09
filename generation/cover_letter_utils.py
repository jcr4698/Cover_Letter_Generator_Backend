import docx
import os
import sys
import subprocess

def get_slash(operating_system):
    # Get the slash from environment
    ENV_BEING_USED = operating_system
    DIR_SLASH = "\\" if ("win" in ENV_BEING_USED) else ("/" if ("linux" in ENV_BEING_USED) else None)
    
    # stop program if this is undefined
    if(DIR_SLASH == None):
        print("slash symbol is undefined for '", sys.platform + "'")
        sys.exit()

    return DIR_SLASH

def get_soffice_cmd(operating_system):
    # Get the slash from environment
    ENV_BEING_USED = operating_system
    SOFFICE_CMD = "C:\Program Files\LibreOffice\program\soffice" if ("win" in ENV_BEING_USED) else ("soffice" if ("linux" in ENV_BEING_USED) else None)
    
    # stop program if this is undefined
    if(SOFFICE_CMD == None):
        print("soffice command is undefined for '", sys.platform + "'")
        sys.exit()

    return SOFFICE_CMD

def get_mv_cmd(operating_system, curr_path, dest_path):
    
    # Get the slash from environment
    ENV_BEING_USED = operating_system
    MV_CMD = None
    if("win" in ENV_BEING_USED):
        MV_CMD = [
            "Move-Item",
            "-Path",
            curr_path,
            "-Destination",
            dest_path
        ]
    elif("linux" in ENV_BEING_USED):
        MV_CMD = [
            "mv",
            curr_path,
            dest_path
        ]
    
    # stop program if this is undefined
    if(MV_CMD == None):
        print("move command is undefined for '", sys.platform + "'")
        sys.exit()

    return MV_CMD

def add_hyperlink(paragraph, text, url):
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a new run object (a wrapper over a 'w:r' element)
    new_run = docx.text.run.Run(docx.oxml.shared.OxmlElement('w:r'), paragraph)
    new_run.text = text

    # Set url font style
    new_run.font.underline = True

    # Join all the xml elements together
    hyperlink.append(new_run._element)
    paragraph._p.append(hyperlink)
    return hyperlink

def get_cover_letter_text(details:dict):
    cover_letter_txt = "{}\n\n{}\n\n{}\n\n".format(
        details["curr_date"],
        details["company"],
        "Dear recruiter:"
    )
    try:
        DIR_SLASH = get_slash(sys.platform)
        cover_letter_ptr = open("." + DIR_SLASH + "cover_letters" + DIR_SLASH + "cover_letter_body.txt", "r")
        cover_letter_txt += cover_letter_ptr.read()
    except FileNotFoundError:
        print("File '." + DIR_SLASH + "cover_letter_body.txt' could not be found.")
    cover_letter_txt += "\n\nSincerely,\n{}".format(details["name"])

    return cover_letter_txt

def get_body_text(cover_letter_ptr, details:dict):

    cover_letter_txt = "{}\n\n{}\n\n{}\n\n".format(
        details["curr_date"],
        details["company"],
        "Dear recruiter:"
    )
    print(type(cover_letter_ptr))
    try:
        cover_letter_txt += cover_letter_ptr.read()
    except Exception as e:
        return e
    cover_letter_txt += "\n\nSincerely,\n{}".format(details["name"])
    return cover_letter_txt

def get_body_text(cover_letter_body_txt:str, details:dict):

    cover_letter_txt = "{}\n\n{}\n\n{}\n\n".format(
        details["curr_date"],
        details["company"],
        "Dear recruiter:"
    )
    cover_letter_txt += cover_letter_body_txt
    cover_letter_txt += "\n\nSincerely,\n{}".format(details["name"])
    return cover_letter_txt

# def generate_pdf(curr_docx_path, dest_docx_path, curr_pdf_path, dest_pdf_path):
#     # generate pdf from docx
#     soffice = get_soffice_cmd(sys.platform)
#     try:
#         subprocess.call(
#             [soffice,
#             "--headless",
#             "--convert-to",
#             "pdf",
#             curr_docx_path]
#         )
#     except subprocess.CalledProcessError as e:
#         print("Error in line 105")
#         print(e)
#         return False

#     # rename/move pdf to appropriate folder
#     try:
#         os.rename(curr_pdf_path, dest_pdf_path)
#     except Exception as e:
#         print("Error: Could not move {} to {}".format(curr_pdf_path, dest_pdf_path))
#         print(e)
#         return False
    
#     # rename/move docx to appropriate folder
#     try:
#         os.rename(curr_docx_path, dest_docx_path)
#     except Exception as e:
#         print("Error: Could not move {} to {}".format(curr_docx_path, dest_docx_path))
#         print(e)
#         return False
    
#     # try:
#     #     # remove the docx
#     #     subprocess.call(
#     #         ["rm",
#     #         "{}/{}".format(docx_path, docx_filename)]
#     #     )
#     # except Exception as e:
#     #     print(e)
#     #     return False
    
#     return True

import aspose.words as aw

def generate_pdf(curr_docx_path, dest_docx_path, curr_pdf_path, dest_pdf_path):
    
    # rename/move docx to appropriate folder
    try:
        gen_docx = aw.Document(curr_docx_path)
        gen_docx.save(dest_pdf_path)
        os.rename(curr_docx_path, dest_docx_path)
    except Exception as e:
        print("Error: Could not move {} to {}".format(curr_docx_path, dest_docx_path))
        print(e)
        return False
    return True

# Saved code (just in case):
# # check if docx_name is valid
# docx_rename = docx_name[::]
# pdf_gen_name = ""
# if(docx_name):
#     if docx_name == ".docx":
#         return False
#     else:
#         if len(docx_name) > 5 and docx_name[len(docx_name)-5:] != ".docx":
#             docx_rename += ".docx"
# else:
#     return False

# # check if it contains ".pdf" at end
# pdf_rename = pdf_name[::]
# if(pdf_name):
#     if pdf_name == ".pdf":
#         return False
#     else:
#         if len(pdf_name) > 4 and pdf_name[len(pdf_name)-4:] != ".pdf":
#             pdf_rename += ".pdf"

# pdf_gen_name = "{}.pdf".format(docx_name[0:len(docx_name)-5])