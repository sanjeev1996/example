from urllib.parse import urlencode
import json
from flask import Flask, request, redirect
from flask import Flask
import requests
from bottle import route,  template, static_file, error, response

app = Flask(__name__)

@app.route('/home')
def show_home():
    return template('home')

@app.route('/')
def show_home1():
    return "this is good"

@app.route('/zendesk_profile')
def make_request():
    parameters = {
        'response_type': 'code',
        'redirect_uri': 'https://flaskserverauth.herokuapp.com/handle_user_decision',
        'client_id': 'my_app_sanjeev',
        'scope': 'read write'}
    url = 'https://sanjeev1996.zendesk.com/oauth/authorizations/new?' + urlencode(parameters)
    return redirect(url)


@app.route('/handle_user_decision')
def handle_decision():
    '''
    parameters = {
        'grant_type': 'authorization_code',
        'code': ,
        'client_id': 'my_app_sanjeev',
        'client_secret': '5cdbabf62299aed9716acd6973c9059ef5b733be3f5b9ce83548be0fa7c0dede',
        'redirect_uri': 'https://flaskserverauth.herokuapp.com/handle_user_decision',
        'scope': 'read'}
    payload = json.dumps(parameters)
    '''
    return json.dumps({'code': str(request.args.get('code')})


@app.route('/create_access_token', methods=['POST'])
def create_access_token():
    data = request.get_json()
    parameters = {
        'grant_type': 'authorization_code',
        'code': data["auth_code"],
        'client_id': 'my_app_sanjeev',
        'client_secret': '5cdbabf62299aed9716acd6973c9059ef5b733be3f5b9ce83548be0fa7c0dede',
        'redirect_uri': 'https://flaskserverauth.herokuapp.com/handle_user_decision',
        'scope': 'read'}
    payload = json.dumps(parameters)
    
    header = {'Content-Type': 'application/json'}
    url = 'https://sanjeev1996.zendesk.com/oauth/tokens'.format(data["auth_code"])
    r = requests.post(url, data=payload, headers=header)

    if r.status_code != 200:
        error_msg = 'Failed to get access token with error {}'.format(r.status_code)
        return template('error', error_msg=error_msg)
    else:
        data = r.json()
        #response.set_cookie('owat', data['access_token'])
        # error_msg = 'Failed to get access token with error {}'.format(data['access_token'])
        # return template('error', error_msg=error_msg)
        return template('auth', token=data['access_token'])
        #return redirect('/zendesk_profile')

@app.route('/ticket_fields', methods=['POST'])
def ticket_fields():
    data = request.get_json()
    for key, value in data.items():
        if int(value) == 1:    
            headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
            payload = json.dumps({"ticket_field": {"type": "text", "title": key}})
            url = 'https://sanjeev1996.zendesk.com/api/v2/ticket_fields.json'
            r = requests.post(url, data=payload, headers=headers)
            # if (r.status_code != 200)|(r.status_code != 201):
            #     error_msg = 'Failed to get access token with error {}'.format(r.status_code)
            #     return template('error', error_msg=error_msg)
            # else:
            #     pass
    return "1"

@app.route('/done')
def done():
    return template('auth', token=token)
    return template('apikey')