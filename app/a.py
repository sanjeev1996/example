import requests
import json
from urllib.parse import urlencode



import pandas as pd


# url = 'https://{}.zendesk.com/api/v2/triggers.json'.format(subdomain)
# headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(token)}
# payload = json.dumps({"trigger": {"title": "Bytesviewapi New Ticket Trigger", "conditions": {"all": [{ "field": "update_type",
#         "operator": "is", "value": "Create" }]},
#         "actions": [{ "field": "notification_target", "value": [ target_id, "{\"text_list\": [\"{{ticket.title}}. {{ticket.description}}\"], \"username\": \"{{ticket.account}}\", \"ticket_id\": \"{{ticket.id}}\", \"api_key\": "+ "\"" + str(api_key) + "\"" +"}"] }],
#         "category_id": category_id}})
# x = requests.post(url, headers = headers, data=payload)



# curl -v -u {email}:{password} https://{subdomain}.zendesk.com/api/v2/triggers/{trigger_id}.json \
#   -H "Content-Type: application/json" -X PUT -d '{"trigger": {"title": "Roger Wilco II", "category_id": "10026"}}'




headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
payload = json.dumps({"app_id": "174203", "settings": {"name": "Helpful App", "api_token": "53xjt93n6tn4321p", "use_ssl": True}})
url = 'https://sanjeev1996.zendesk.com/api/v2/apps/installations.json'
r = requests.post(url, data=payload, headers=headers)
print(r.json())


# curl https://{subdomain}.zendesk.com/api/v2/apps/installations.json \
#   -d '{"app_id": "225", "settings": {"name": "Helpful App", "api_token": "53xjt93n6tn4321p", "use_ssl": true}}' \
#   -H "Content-Type: application/json" -X POST \
#   -u {email_address}:{password}


# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# payload = json.dumps({"app_id": "174203", "settings": {"name": "Helpful App", "api_token": "53xjt93n6tn4321p", "use_ssl": True}})
# url = 'https://sanjeev1996.zendesk.com/api/v2/apps/installations.json'
# r = requests.post(url, data=payload, headers=headers)
# print(r.json())



# url = 'https://sanjeev1996.zendesk.com/api/v2/apps/installations?include=app.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# x = requests.get(url, headers = headers)
# print(x.json())


# headers = {"Content-Type":"application/json"}
# payload = json.dumps({"apikey": "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VybmFtZSI6Im5hbXJhdGFAYWxnb2RvbW1lZGlhLmNvbSJ9.ceDtxy7gbSSI1t3lrokTBMNajge7oPrmo07R7phKRI8"})
# url = 'https://api.bytesview.com/verify/apikey'
# r = requests.post(url, data=payload, headers=headers)
# print(r.json()["0"])



# url = 'https://sanjeev1996.zendesk.com/api/v2/targets.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# x = requests.get(url, headers = headers)
# #print(x.json())
# target = x.json()["targets"]

# for i in target[::-1]:
#     if i["title"] == "Bytesviewapi New ticket target":
#         target_id = i["id"]

# print(target_id)





# {'targets': [{'url': 'https://sanjeev1996.zendesk.com/api/v2/targets/360001589897.json', 
#  'id': 360001589897, 'created_at': '2021-04-19T08:21:44Z', 'type': 'url_target_v2', 
#  'title': 'heroku', 'active': True, 'method': 'post', 'username': '', 
#  'content_type': 'application/json', 'password': None, 
#  'target_url': 'https://flaskserverauth.herokuapp.com/trigger'}, {'url': 'https://sanjeev1996.zendesk.com/api/v2/targets/360001607397.json', 'id': 360001607397, 'created_at': '2021-04-27T07:57:52Z', 'type': 'url_target_v2', 'title': 'Bytesviewapi New ticket target', 'active': True, 'method': 'post', 'username': None, 'content_type': 'application/json', 'password': None, 'target_url': 'https://flaskserverauth.herokuapp.com/trigger'}], 'next_page': None, 'previous_page': None, 'count': 2}



# url = 'https://sanjeev1996.zendesk.com/api/v2/targets.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# data = json.dumps({})

# {"target": {"type": "email_target", "title": "Test Email Target", "email": "hello@example.com", "subject": "Test Target"}}


# {"trigger": {"title": "test", "conditions": {"all": [{ "field": "status", \
#   "operator": "is", "value": "open" }, { "field": "priority", \
#   "operator": "less_than", "value": "high" }]}, \
#   "actions": [{ "field": "group_id", "value": "20455932" }], \
#   "category_id": "10026"}}


# x = requests.get(url, headers = headers)




# url = 'https://sanjeev1996.zendesk.com/api/v2/triggers.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# data = json.dumps({})

# {"trigger": {"title": "test", "conditions": {"all": [{ "field": "status", \
#   "operator": "is", "value": "open" }, { "field": "priority", \
#   "operator": "less_than", "value": "high" }]}, \
#   "actions": [{ "field": "group_id", "value": "20455932" }], \
#   "category_id": "10026"}}


# x = requests.get(url, headers = headers)
# print(x.status_code)
# print(x.content)
# if (x.status_code != 200):
#     print(x.status_code)
# else:
#     pass




# url = 'https://sanjeev1996.zendesk.com/api/v2/triggers.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# x = requests.get(url, headers = headers)
# print(x.status_code)
# print(x.content)
# if (x.status_code != 200):
#     print(x.status_code)
# else:
#     pass





# url = 'https://sanjeev1996.zendesk.com/api/v2/tickets/28.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# data = json.dumps({"ticket": {"custom_fields":[{ "id": 360018362718, "value": "positive"}]}})
# x = requests.put(url, headers = headers, data=data)
# print(x.status_code)
# if (x.status_code != 200):
#     print(x.status_code)
# else:
#     pass



# url = 'https://sanjeev1996.zendesk.com/api/v2/tickets/13.json'
# #headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(x.json()["access_token"])}
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# #data = json.dumps({'ticket':{'custom_fields': [{'id': 360017592378, 'value': 'sanjeev'}]}})
# data = json.dumps({"ticket": {"custom_fields":[{ "id": 360017592378, "value": "sanjeev"}]}})
# x = requests.put(url, headers = headers, data=data)
# print(x.json())



# url = 'https://sanjeev1996.zendesk.com/api/v2/tickets/13.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# x = requests.put(url, headers = headers)
# print(x.json())


'''
curl https://{subdomain}.zendesk.com/api/v2/tickets/{ticket_id}.json \
  -H "Content-Type: application/json" \
  -d '{"ticket": {"status": "open", "comment": { "body": "The smoke is very colorful.", "author_id": 494820284 }}}' \
  -v -u {email_address}:{password} -X PUT


url = 'https://sanjeev1996.zendesk.com/api/v2/ticket/13/tags.json'
#headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(x.json()["access_token"])}
headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
data = json.dumps({ "tags":["important"]})
x = requests.put(url, headers = headers, data=data)
print(x.json())
'''




# label = ["negative", "neutral", "positive"]
# headers = {"Content-Type":"application/json", "X-ACCESS-TOKEN": "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJ1c2VybmFtZSI6Im5hbXJhdGFAYWxnb2RvbW1lZGlhLmNvbSJ9.ceDtxy7gbSSI1t3lrokTBMNajge7oPrmo07R7phKRI8"}
# payload = json.dumps({"data":{"key1":"this is good"}, "lang":"en"})
# url = 'https://api.bytesview.com/1/static/sentiment'
# r = requests.post(url, data=payload, headers=headers)
# json_data = json.loads(r.text)
# print(label[json_data["success"]["key1"]["label"]]) 
# print(json_data)




# parameters = {
#         'response_type': 'code',
#         'redirect_uri': 'https://flaskserverauth.herokuapp.com/handle_user_decision',
#         'client_id': 'my_app_sanjeev',
#         'scope': 'read write'}
# url = 'https://sanjeev1996.zendesk.com/oauth/authorizations/new?' + urlencode(parameters)

# #return redirect(url)
# x = requests.get(url, headers = parameters)
# print(x.headers)

# url = 'https://sanjeev1996.zendesk.com/oauth/tokens'
# headers = {"Content-Type": "application/json"}
# #myobj = {"grant_type": "password", "client_id":"my_app_sanjeev", "client_secret":"8e2e80317be431efb7ae803fcc70cfadfeff7f1f059c2018cf26961aecbe5891", "scope": "read write", "username": "sanjeevamazingworld1996@gmail.com", "password": "Sanjeev@123"}
# myobj = "{\"grant_type\": \"password\", \"client_id\":\"my_app_sanjeev\", \"client_secret\":\"8e2e80317be431efb7ae803fcc70cfadfeff7f1f059c2018cf26961aecbe5891\", \"scope\": \"read write\", \"username\": \"sanjeevamazingworld1996@gmail.com\", \"password\": \"Sanjeev@123\"}"
# x = requests.post(url, headers = headers, data = myobj)

#print(x.json()["access_token"])



# url = 'https://sanjeev1996.zendesk.com/api/v2/ticket/13/tags.json'
# #headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(x.json()["access_token"])}
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# data = json.dumps({ "tags":["important"]})
# x = requests.put(url, headers = headers, data=data)
# print(x.json())




# curl 'https://sanjeev1996.zendesk.com/api/v2/tags.json' -v -u 'sanjeevamazingworld1996@gmail.com:Sanjeev@123'


# curl https://{subdomain}.zendesk.com/api/v2/ticket_fields.json \
#   -d '{"ticket_field": {"type": "text", "title": "Age"}}' \
#   -H "Content-Type: application/json" -X POST \
#   -v -u {email_address}:{password}




# url = 'https://sanjeev1996.zendesk.com/api/v2/ticket_fields.json'
# #headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(x.json()["access_token"])}
# headers = {"Content-Type":"application/json", "Authorization": "Bearer 5446cd8d62aa3e87032832ef9b038c220d4b357da22d7940b80074f60339d693"}
# data = json.dumps({"ticket_field": {"type": "text", "title": "Age1"}})
# x = requests.post(url, headers = headers, data=data)
# print(x.json())





# url = 'https://sanjeev1996.zendesk.com/api/v2/ticket_fields/360017592378.json'
# headers = {"Content-Type":"application/json", "Authorization": "Bearer {}".format(x.json()["access_token"])}
# data = '{"custom_field_options": {"Age": "37"}}'
# x = requests.put(url, headers = headers, data=data)

# print(x.json())

'''
async function getAllTicketFields() {
    const target = {
        url: '/api/v2/ticket_fields.json',
        type: 'GET',
        dataType: 'json'
    };
    try {
        return await client.request(target)
    }
    catch (e) {
        console.log(e)
    }
}

async function getTicketField(title) {
    const fields = await getAllTicketFields()
    const field = fields.ticket_fields.find(i => i.title == title)
    return field
}

async function updateTicketField(field) {
    const target = {
        url: `/api/v2/ticket_fields/${field.id}.json`,
        type: 'PUT',
        dataType: 'json',
        contentType: "application/json",
        data: JSON.stringify({ ticket_field: field })
    };
    try {
        return await client.request(target)
    }
    catch (e) {
        console.log(e)
    }
'''





# url = 'https://sanjeev1996.zendesk.com/api/v2/tickets.json'
# headers = {"Authorization": "Bearer {}".format(x.json()["access_token"])}

# x = requests.get(url, headers = headers)

# print(x.json())
