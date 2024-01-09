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
	try:
		return send_file("pdfs/" + file_name)
	except Exception as e:
		return file_name + "not found"

@app.route("/cover_letter/<file_name>", methods=["GET"])
def generate_cover_letter(file_name:str):

	try:
		cover_letter_info = request.get_json()
		if(cover_letter_info):
			curr_cov_let = CoverLetter()
			curr_cov_let.set_full_name(cover_letter_info["full_name"].strip())
			curr_cov_let.set_phone_num(cover_letter_info["phone_num"].strip())
			curr_cov_let.set_email(cover_letter_info["email"].strip())
			curr_cov_let.set_linkedin(cover_letter_info["linkedin"].strip())
			curr_cov_let.set_company(cover_letter_info["company"].strip())
			curr_cov_let.set_cover_letter_body(cover_letter_info["cover_letter_body"].strip())

			print(os.getcwd())
			filename_out = curr_cov_let.generate_cover_letter(file_name)
			return send_file("./pdfs/"+filename_out)

	except Exception as e:
		print(e)
		return "Nothing to show"