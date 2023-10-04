from urllib.parse import urlencode
import json
import time
import re
from flask import Flask, request, redirect, url_for, jsonify
from flask import Flask
import requests
from bottle import route,  template, static_file, error, response
from app.model import Users, app, db

@app.route('/verify_apikey', methods=['POST'])
def verify_apikey():
    data = request.get_json()
    api_key = data["apikey"]
    headers = {"Content-Type":"application/json"}
    payload = json.dumps({"apikey": api_key})
    url = 'https://api.bytesview.com/verify/apikey'
    r = requests.post(url, data=payload, headers=headers)
    print(r.json()["0"])
    return r.json()

def db_add_user(token, subdomain, api_key):
    user = Users.query.filter_by(access_token=token).first()
    if not user:
        print("a")
        custom_ticket_field = check_custom_field(token, subdomain)
        available_model = ["sentiment", "keywords", "emotion"]
        new_token = Users(access_token=token, subdomain=subdomain, api_key=api_key, is_active=1, field_option=str(custom_ticket_field), model_option=str(available_model), selected_field = str({}))
        db.session.add(new_token)
        db.session.commit()
    else:
        user.access_token = token
        db.session.commit()

def create_trigger(token, subdomain, api_key):
    #api_key = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VybmFtZSI6Im5hbXJhdGFAYWxnb2RvbW1lZGlhLmNvbSJ9.ceDtxy7gbSSI1t3lrokTBMNajge7oPrmo07R7phKRI8"
    url = 'https://{}.zendesk.com/api/v2/triggers.json'.format(subdomain)
    headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
    x = requests.get(url, headers = headers)
    #print(x.json())
    url = 'https://{}.zendesk.com/api/v2/targets.json'.format(subdomain)
    headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
    x_target = requests.get(url, headers = headers)
    #print(x.json())
    target = x_target.json()["targets"]
    print(target)
    for i in target[::-1]:
        if i["title"] == "Bytesviewapi New ticket target":
            target_id = i["id"]

    url = 'https://{}.zendesk.com/api/v2/triggers.json'.format(subdomain)
    headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
    x_triggers = requests.get(url, headers = headers)
    trigger = x_triggers.json()["triggers"]
    for i in trigger:
        if i["title"] == "Notify assignee of assignment":
            category_id = i["category_id"]

    if not any([(i["title"] ==  "Bytesviewapi New Ticket Trigger")&(i["active"] ==  True) for i in x.json()["triggers"]]):
        url = 'https://{}.zendesk.com/api/v2/triggers.json'.format(subdomain)
        headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
        payload = json.dumps({"trigger": {"title": "Bytesviewapi New Ticket Trigger", "conditions": {"all": [{ "field": "update_type",
                "operator": "is", "value": "Create" }]},
                "actions": [{ "field": "notification_target", "value": [ target_id, "{\"text_list\": [\"{{ticket.title}}. {{ticket.description}}\"], \"username\": \"{{ticket.account}}\", \"ticket_id\": \"{{ticket.id}}\", \"access_token\": "+ "\"" + str(token) + "\"" +"}"] }],
                "category_id": category_id}})
        x = requests.post(url, headers = headers, data=payload)
        if (x.status_code != 200)&(x.status_code != 201):
            print(x.status_code)
        else:
            pass    
    else:
        for i in x.json()["triggers"]:
            if i["title"] ==  "Bytesviewapi New Ticket Trigger":
                trigger_id = i["id"]
        url = 'https://'+ subdomain +'.zendesk.com/api/v2/triggers/'+ str(trigger_id) +'.json'
        headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
        x = requests.delete(url, headers = headers)
        print(trigger_id, "  :  ", token)
        if (x.status_code != 200)&(x.status_code != 201)&(x.status_code != 204):
            print(x.status_code)
        else:
            pass

        url = 'https://'+ subdomain +'.zendesk.com/api/v2/triggers.json'
        headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
        payload = json.dumps({"trigger": {"title": "Bytesviewapi New Ticket Trigger", "conditions": {"all": [{ "field": "update_type",
                "operator": "is", "value": "Create" }]},
                #"actions": [{ "field": "notification_target", "value": [ target_id, "{\"text_list\": [\"{{ticket.title}}. {{ticket.description}}\"], \"username\": \"{{ticket.account}}\", \"ticket_id\": \"{{ticket.id}}\", \"access_token\": "+ "\"" + str(token) + "\"" +"}"] }],
                "actions": [{ "field": "notification_target", "value": [ target_id, "{\"text_list\": [\"{{ticket.title}}. {{ticket.description}}\"], \"username\": \"{{ticket.account}}\", \"ticket_id\": \"{{ticket.id}}\", \"access_token\": "+ "\"" + str(token) + "\"" +"}"] }],
                #"actions": [{ "field": "notification_target", "value": [ target_id, "{\"text_list\": [\"{{ticket.title}}. {{ticket.description}}\"], \"username\": \"{{ticket.account}}\", \"ticket_id\": \"{{ticket.id}}\", \"api_key\": "+ "\"" + str(api_key) + "\"" +"}"] }],
        "category_id": category_id}})
        x = requests.post(url, headers = headers, data=payload)
        if (x.status_code != 200)&(x.status_code != 201):
            print(x.status_code)
        else:
            pass
    return "Done"

def installationAPIkeyflag(access_token, subdomain, payload):
    header = {'Content-Type': 'application/json', "Authorization": "Bearer "+access_token}
    url = 'https://'+ subdomain +'.zendesk.com/api/v2/apps/installations.json'
    r = requests.get(url, headers=header)
    #print(r.json())
    for i in r.json()["installations"]:
        if "settings" in i:
            if "app_id" in i["settings"]:
                if i["settings"]["app_id"] == "123456789":
                    app_id = i["id"]
                    api_key = i["settings"]["api_key"]

    #print(str(app_id))
    payload = json.dumps({"settings": payload})
    header = {'Content-Type': 'application/json', "Authorization": "Bearer "+access_token}
    url = 'https://'+ subdomain +'.zendesk.com/api/v2/apps/installations/'+ str(app_id) +'.json'
    r = requests.put(url, data=payload, headers=header)
    return api_key
    #print(r.json())

@app.route('/handle_user_decision')
def handle_decision():
    #print("this is good")
    subdomain = str(request.args.get('state'))
    print(request.url)
    parameters = {
        'grant_type': 'authorization_code',
        'code': str(request.args.get('code')),
        'client_id': 'my_app_sanjeev',
        'client_secret': 'a3e18ecfcd22437c3f80305398e4571f6fe4c71bfa7f7985b0ccb5046a202543',
        'redirect_uri': 'https://flaskserverauth.herokuapp.com/handle_user_decision',
        'scope': 'read'}
    payload = json.dumps(parameters)
    header = {'Content-Type': 'application/json'}
    url = 'https://{}.zendesk.com/oauth/tokens'.format(subdomain)
    r = requests.post(url, data=payload, headers=header)
    #print(r.json())
    if r.status_code != 200:
        error_msg = 'Failed to get access token with error {}'.format(r.status_code)
        return template('error', error_msg=error_msg)
    else:
        data = r.json()
        payload = {"token": data['access_token'], "app_status" : "showFieldModel"}
        api_key = installationAPIkeyflag(data['access_token'], subdomain, payload)
        create_trigger(data['access_token'], subdomain, api_key)
        db_add_user(data['access_token'], subdomain, api_key)
        return redirect("https://"+subdomain+".zendesk.com/agent/apps/bytesview")
        #response.set_cookie('owat', data['access_token'])
        # error_msg = 'Failed to get access token with error {}'.format(data['access_token'])
        # return template('error', error_msg=error_msg)
        #return "this is good"
        #return template('auth', token=data['access_token'])
        #return redirect('/zendesk_profile')
        #redirect "https://sanjeevalgodom.zendesk.com/agent/apps/monkey-learn-ticket-tagging"

@app.route('/get_selected_field_option', methods=['POST'])
def get_selected_field_option():
    data = request.get_json()
    token = data['access_token'] 
    subdomain = data['subdomain']
    a = {}

    user = Users.query.filter_by(access_token=data['access_token']).first()
    field_option = check_custom_field(token, subdomain)

    a["custom_ticket_field"] = [i[0]+"("+str(i[1])+")" for i in field_option] 
    a["model_option"] = eval(user.model_option)
    a["selected_field"] = eval(user.selected_field)
    
    if user.field_option != str(field_option):
        user.field_option = str(field_option)
        db.session.commit()
    #print(a)
    return a

def check_custom_field(token, subdomain):
    headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
    url = 'https://{}.zendesk.com/api/v2/ticket_fields.json'.format(subdomain)
    r = requests.get(url, headers=headers)
    #print(r.status_code)
    if (r.status_code != 200)&(r.status_code != 201):
        error_msg = 'Server error {}'.format(r.status_code)
        return template('error', error_msg=error_msg)
    else:
        custom_ticket_field = []
        system_ticket = ["Type", "Priority"]
        for i in r.json()["ticket_fields"]:
            if (i["removable"] == True)&(i["active"] == True):
                if i["title"] not in  system_ticket:
                    #print([i["title"], i["id"]])
                    custom_ticket_field.append([i["title"], i["id"]])        
        if custom_ticket_field:
            return custom_ticket_field
        else:
            return []

@app.route('/ticket_fields', methods=['POST'])
def ticket_fields():
    data = request.get_json()
    print(data["dict"])

    user = Users.query.filter_by(access_token=data['access_token']).first()
    if eval(user.selected_field) != data["dict"]:
        selected_field = data["dict"]
        available_field = eval(user.field_option)

        for u,v in selected_field.items():
            print(v["field"], v["field"], [i[0]+"("+str(i[1])+")" for i in available_field])
            if v["field"] not in [i[0]+"("+str(i[1])+")" for i in available_field]:
                print("a")
                print(v["field"])
                headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(data["access_token"])}
                payload = json.dumps({"ticket_field": {"type": "text", "title": str(v["field"])}})
                url = 'https://{}.zendesk.com/api/v2/ticket_fields.json'.format(data["subdomain"])
                r = requests.post(url, data=payload, headers=headers)
                #print(r.status_code)
                print(r.status_code)
                if (r.status_code != 200)&(r.status_code != 201):
                    return jsonify({"error":"Error"}), r.status_code

                selected_field[u]["field"] = str(v["field"])+"("+str(r.json()["ticket_field"]["id"])+")" 
        print(selected_field)
        user.selected_field = str(selected_field)
        db.session.commit()
    payload = {"app_status" : "showfinalWindow"}
    installationAPIkeyflag(data['access_token'], data["subdomain"], payload)
    return "Done"

    '''
    print(data)    
    headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(data["access_token"])}
    payload = json.dumps({"ticket_field": {"type": "text", "title": data["dict"]['ticket_field']}})
    url = 'https://{}.zendesk.com/api/v2/ticket_fields.json'.format(data["subdomain"])
    r = requests.post(url, data=payload, headers=headers)
    #print(r.status_code)
    if (r.status_code != 200)&(r.status_code != 201):
        error_msg = 'Failed to get access token with error {}'.format(r.status_code)
        return template('error', error_msg=error_msg)
    else:
        payload = {"app_status" : "showfinalWindow"}
        installationAPIkeyflag(data['access_token'], data["subdomain"], payload)
        pass
    return "1"
    '''
    

@app.route('/trigger', methods=['POST'])
def trigger():
    data = request.get_json()
    #user = Users.query.filter_by(api_key=str(data["api_key"])).first()
    #access_token = str(user.access_token)
    user = Users.query.filter_by(access_token=str(data["access_token"])).first()
    api_key = str(user.api_key)
    text = re.sub("[-]+\n\n.*\n\n", " ", data["text_list"][0])
    

    if user.is_active == 1:
        for u, v in eval(user.selected_field).items():
            print(u)
            print(v)
            headers = {"Content-Type":"application/json", "X-ACCESS-TOKEN": str(api_key)}
            payload = json.dumps({"data":{"key1":text}, "lang":"en"})
            url = 'https://api.bytesview.com/1/static/'+v["model"]
            r = requests.post(url, data=payload, headers=headers)
            json_data = json.loads(r.text)
            if (v["model"] == "sentiment"):
                label = ["negative", "neutral", "positive"]
                output1 = label[json_data["success"]["key1"]["label"]]
            elif (v["model"] == "emotion"):
                label = ["anger", "fear", "happy", "love", "neutral", "sad"]
                output1 = label[json_data["success"]["key1"]["label"]] 
            elif v["model"] == "keywords":
                output1 = ','.join(json_data["success"]["key1"]["tags"])

            ticket_id = data["ticket_id"]
            #subdomain = data["username"]
            print(text)
            print(output1)
            print(ticket_id)
            url = 'https://'+ str(user.subdomain) +'.zendesk.com/api/v2/tickets/'+str(ticket_id)+'.json'
            #url = 'https://sanjeev1996.zendesk.com/api/v2/tickets/29.json'
            #headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(x.json()["access_token"])}
            headers = {"Content-Type":"application/json", "Authorization": "Bearer "+ str(data["access_token"])}
            #data = json.dumps({'ticket':{'custom_fields': [{'id': 360017592378, 'value': 'sanjeev'}]}})
            payload = json.dumps({"ticket": {"custom_fields":[{ "id": re.search(r"\(\d+\)$", v["field"]).group()[1:-1], "value": str(output1)}]}})
            #print()
            x = requests.put(url, headers = headers, data=payload)
            print(x.status_code)
            if (x.status_code != 200):
                print(x.status_code)
            else:
                pass
            #print(x.json())
    return "Done"


@app.route('/isactive', methods=['POST'])
def is_active():
    data = request.get_json()
    log = data["log"]
    token = data['access_token'] 
    user = Users.query.filter_by(access_token=data['access_token']).first()
    user.is_active = log
    db.session.commit()
    return "Done"

@app.route('/check_isactive', methods=['POST'])
def check_isactive():
    data = request.get_json()
    token = data['access_token'] 
    user = Users.query.filter_by(access_token=data['access_token']).first()
    return str(user.is_active)


# @app.route('/home')
# def show_home():
#     return template('home')

# @app.route('/')
# def show_home1():
#     return "this is good"

# @app.route('/zendesk_profile')
# def make_request():
#     subdomain = str(request.args.get('state'))
#     print(request.url)
#     parameters = {
#         'response_type': 'code',
#         'redirect_uri': 'https://flaskserverauth.herokuapp.com/handle_user_decision',
#         'client_id': 'my_app_sanjeev',
#         'scope': 'read write',
#         'state':subdomain,}
#     url = 'https://{}.zendesk.com/oauth/authorizations/new?'.format(subdomain) + urlencode(parameters)
#     return redirect(url)

# @app.route('/verify_access_token')
# def verify_access_token():
#     print(request.url)
#     #return template('auth', token='931c8d9cebf15776e41ce491b9751ea13d255e81c86240984bfafae523850df5')
#     return template('auth', token='undefined')
#     #return redirect('/zendesk_profile')
#     #redirect "https://sanjeevalgodom.zendesk.com/agent/apps/monkey-learn-ticket-tagging"

# @app.route('/wizard', methods=['POST'])
# def wizard():
#     data = request.get_json()
#     #data["subdomain"]
#     output = {"API_flag":False, "wizard5":True}
#     return output

# @app.route('/done')
# def done():
#     return template('auth', token=token)
#     return template('apikey')