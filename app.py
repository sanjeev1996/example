#from paragraph_line import app


import re
from flask import Flask, jsonify, render_template, request
import requests

app = Flask(__name__)


print("a")

@app.route('/')
def ticket_fields():
    #return "hello"
    return {"1": "2"}


if __name__ == "__main__":
  app.run()
