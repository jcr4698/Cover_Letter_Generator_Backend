import sys

def get_slash(operating_system):
    # Get the slash from environment
    ENV_BEING_USED = operating_system
    DIR_SLASH = "\\" if ("win" in ENV_BEING_USED) else ("/" if ("linux" in ENV_BEING_USED) else None)
    
    # stop program if this is undefined
    if(DIR_SLASH == None):
        print("slash symbol is undefined for '", sys.platform + "'")
        sys.exit()

    return DIR_SLASH

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
