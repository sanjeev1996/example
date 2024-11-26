#from paragraph_line import app


import re
from flask import Flask, jsonify, render_template, request
import requests

#app = Flask(__name__, template_folder="template")
app = Flask(__name__)



@app.route('/')
def ticket_fields():
    #return "hello"
    return render_template('iframe2.html')

@app.route('/api', methods=["post"])
def ticket_fields1():
    text = str(request.get_json()["text"])
    text = ' '.join(text.split('\n'))
    print(text)
    # return "hello"
    #text = text.replace(r"\n", " ")
    return text


if __name__ == "__main__":
  app.run()
