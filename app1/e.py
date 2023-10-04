import requests
import json

parameters = {
    'grant_type': 'authorization_code',
    'code': "2d5883e525058d9d7f9b344fedda989338f4310023b562cd9278d88504d7748a",
    'client_id': 'zdg-bytesview_zendesk_app',
    'client_secret': 'b4f1834ed2a398671ac3c646d3f146e957be72c5f508cdd6c21a4a7cf3a03f24',
    'redirect_uri': 'https://api.bytesview.com/zendesk/handle_user_decision',
    'scope': 'read'}
payload = json.dumps(parameters)
header = {'Content-Type': 'application/json'}
url = 'https://{}.zendesk.com/oauth/tokens'.format("d3v-sanjeev1999")
r = requests.post(url, data=payload, headers=header)
print(r.json())
if r.status_code != 200:
    print(r.status_code)