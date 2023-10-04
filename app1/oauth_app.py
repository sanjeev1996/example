from urllib.parse import urlencode
import json

import requests
from bottle import route, template, redirect, static_file, error, request, response, run


@route('/home')
def show_home():
    return template('home')


@route('/zendesk_profile')
def make_request():
    if request.get_cookie('owat'):
        # Get user data
        access_token = request.get_cookie('owat')
        bearer_token = 'Bearer ' + access_token
        header = {'Authorization': bearer_token}
        url = 'https://sanjeev1996.zendesk.com/api/v2/users/me.json'
        r = requests.get(url, headers=header)
        if r.status_code != 200:
            error_msg = 'Failed to get data with error {}'.format(r.status_code)
            return template('error', error_msg=error_msg)
        else:
            data = r.json()
            profile_data = {
                'name': data['user']['name'],
                'role': data['user']['role']}
            return template('details', data=profile_data)
    else:
        # Kick off authorization flow
        parameters = {
            'response_type': 'code',
            'redirect_uri': 'http://localhost:8080/handle_user_decision',
            'client_id': 'my_app_sanjeev_localhost',
            'scope': 'read write'}
        url = 'https://sanjeev1996.zendesk.com/oauth/authorizations/new?' + urlencode(parameters)
        redirect(url)


@route('/handle_user_decision')
def handle_decision():
    if 'error' in request.query_string:
        return template('error', error_msg=request.query.error_description)
    else:
        # Get access token
        parameters = {
            'grant_type': 'authorization_code',
            'code': request.query.code,
            'client_id': 'my_app_sanjeev_localhost',
            'client_secret': '1cb18a6909b01c36a0452ede68b474336e50878d1d97ce801faa5443bb2b9e04',
            'redirect_uri': 'http://localhost:8080/handle_user_decision',
            'scope': 'read'}
        payload = json.dumps(parameters)
        header = {'Content-Type': 'application/json'}
        url = 'https://sanjeev1996.zendesk.com/oauth/tokens'
        r = requests.post(url, data=payload, headers=header)
        if r.status_code != 200:
            error_msg = 'Failed to get access token with error {}'.format(r.status_code)
            return template('error', error_msg=error_msg)
        else:
            data = r.json()
            response.set_cookie('owat', data['access_token'])
            redirect('/zendesk_profile')


@route('/')
def handle_root_url():
    redirect('/home')


@route('/css/<filename>')
def send_css(filename):
    return static_file(filename, root='static/css')


@error(404)
def error404(error):
    return template('error', error_msg='404 error. Nothing to see here')


run(host='localhost', port=8080, debug=True)