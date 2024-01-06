from flask import Flask, send_file, request, jsonify
import json
import os

from cover_letter import CoverLetter

app = Flask(__name__)

@app.route("/")
def hello_world():
	return "Hello, World!"

@app.route("/show_cover_letter/<file_name>", methods=["GET"])
def view_cover_letter(file_name:str):
    return send_file("./pdfs/"+file_name)

@app.route("/cover_letter", methods=["GET", "POST"])
def generate_cover_letter():

	try:
		cover_letter_info = request.get_json()
		if(cover_letter_info):
			print(cover_letter_info)

			curr_cov_let = CoverLetter()
			curr_cov_let.set_full_name(cover_letter_info["full_name"])
			curr_cov_let.set_phone_num(cover_letter_info["phone_num"])
			curr_cov_let.set_email(cover_letter_info["email"])
			curr_cov_let.set_linkedin(cover_letter_info["linkedin"])
			curr_cov_let.set_company(cover_letter_info["company"])
			curr_cov_let.set_cover_letter_body(cover_letter_info["cover_letter_body"])

			cov_let_filename = cover_letter_info["file_name"] + ".pdf"
			print(cov_let_filename)
			print(os.getcwd())
   
   
			curr_cov_let.generate_cover_letter(cov_let_filename)
			return send_file("./pdfs/"+cov_let_filename)

	except Exception as e:
		print(e)
		return "Nothing to show"