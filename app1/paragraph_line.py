import re
from flask import Flask, jsonify, render_template
import requests

app = Flask(__name__, template_folder="template")

@app.route('/')
def ticket_fields():
    #return "hello"
    return render_template('iframe.html')