'''Microsoft graph API test program'''
import json
import os
import urllib.parse
import uuid

import adal
from bottle import app, get, post, redirect, request, route, run, view
from requests import Session
from json2html import *

# graph api constants
redirect_uri = 'http://localhost:5000/login/authorized'
resource_uri = 'https://graph.microsoft.com/'
authority_url = 'https://login.microsoftonline.com/common'
api_version = 'v1.0'

# Load graph app defaults
try:
    with open('graphconfig.json') as config_file:
        config_data = json.load(config_file)
except FileNotFoundError:
    sys.exit('Error: Expecting graphconfig.json in current folder')

client_id = config_data['appId']
client_secret = config_data['clientSecret']

os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'  # enable non-HTTPS for testing
SESSION = Session()

def display_payload(payload, apicall='/me'):
    '''Display JSON graph output in an HTML page'''
    header = '<h2>Graph output</h2><br/><br/>'
    form1 = '<br/><form action="/graphcall" method="post">' +\
        'api: <input name="apicall" type="text" size=50 value="' + apicall + '"/> ' +\
        '<input type="submit" value="Call"/>' +\
        ' [Examples: /me, /me/drive, /me/people/?$search=name, ' +\
        '/me/messages/?$select=subject,from&$search=Artificial]</form><br/>'
    form2 = '<br/><form action="/maildump" method="post">' +\
        'Folder: <input name="folder" type="text" size=50 value="Sent Items"/> ' +\
        'From: <input name="from_date" type="text" size=20 value="2017-11-01"/> ' +\
        'To: <input name="to_date" type="text" size=20 value="2017-11-08"/> ' +\
        '<input type="submit" value="Get"/></form><br/>'
    htmldata = json2html.convert(json=payload)
    footer = '<p><a href="/">Restart</a></p>'
    return header + form1 + form2 + htmldata + footer


def display_text(output):
    '''Display text output to page'''
    return output

@route('/')
@view('homepage.html')
def homepage():
    """Render the home page."""
    return {'sample': 'Microsoft Graph API'}


@route('/login')
def login():
    """Prompt user to authenticate."""
    auth_state = str(uuid.uuid4())
    SESSION.auth_state = auth_state

    prompt_behavior = 'none'
    #prompt_behavior = 'select_account'

    params = urllib.parse.urlencode({'response_type': 'code',
                                     'client_id': client_id,
                                     'redirect_uri': redirect_uri,
                                     'state': auth_state,
                                     'resource': resource_uri,
                                     'prompt': prompt_behavior})

    return redirect(authority_url + '/oauth2/authorize?' + params)


@route('/login/authorized')
def authorized():
    """Handler for the application's Redirect Uri."""
    code = request.query.code
    auth_state = request.query.state
    if auth_state != SESSION.auth_state:
        raise Exception('state returned to redirect URL does not match!')
    auth_context = adal.AuthenticationContext(authority_url, api_version=None)
    token_response = auth_context.acquire_token_with_authorization_code(
        code, redirect_uri, resource_uri, client_id, client_secret)
    SESSION.headers.update({'Authorization': f"Bearer {token_response['accessToken']}",
                            'User-Agent': 'adal-sample',
                            'Accept': 'application/json',
                            'Content-Type': 'application/json',
                            'SdkVersion': 'sample-python-adal',
                            'return-client-request-id': 'true'})
    return redirect('/maincall')


@get('/maincall')
def maincall():
    """Confirm user authentication by calling Graph and displaying data."""
    endpoint = resource_uri + api_version + '/me'
    http_headers = {'client-request-id': str(uuid.uuid4())}
    graphdata = SESSION.get(endpoint, headers=http_headers, stream=False).json()
    return display_payload(graphdata)


@post('/graphcall')
def graphcall():
    """Display custom graph API call results."""
    apicall = request.forms.get('apicall')
    endpoint = resource_uri + api_version + apicall
    http_headers = {'client-request-id': str(uuid.uuid4())}
    graphdata = SESSION.get(
        endpoint, headers=http_headers, stream=False).json()
    return display_payload(graphdata, apicall)

@post('/maildump')
def maildump():
    """Dump the contents of the specified mail folder."""
    folder = request.forms.get('folder')
    from_date = request.forms.get('from_date')
    to_date = request.forms.get('to_date')
    endpoint = resource_uri + api_version + '/me/mailFolders'
    http_headers = {'client-request-id': str(uuid.uuid4())}
    skip_num = 0
    mail_skip = 0

    # first get a list of folders
    graphdata = SESSION.get(endpoint, headers=http_headers, stream=False).json()
    while 'value' in graphdata:
        for folder_rec in graphdata['value']:
            if folder_rec['displayName'] == folder:
                # now dump mail for this folder
                print('Dumping mail for folder: ' + folder)
                mailtext = ''
                mailendpoint = resource_uri + api_version + '/me/mailFolders/' + folder_rec['id'] +\
                    '/messages?$select=subject,bodyPreview&$filter=sentDateTime ge ' + from_date +\
                     ' and sentDateTime le ' + to_date
                maildata = SESSION.get(mailendpoint, headers=http_headers, stream=False).json()
                while 'value' in maildata and len(maildata['value']) > 0:
                    for mail in maildata['value']:
                        mailtext += mail['subject'] + ' ' + mail['bodyPreview']
                    mail_skip += 10
                    maildata = SESSION.get(mailendpoint + '&$skip=' + str(mail_skip), headers=http_headers, stream=False).json()
                return display_text(mailtext)
        # if there was no match and there's a next link, call again
        skip_num += 10
        graphdata = SESSION.get(endpoint + '?$skip=' + str(skip_num), headers=http_headers, stream=False).json()
    return display_text(folder_rec)

if __name__ == '__main__':
    run(app=app(), server='wsgiref', host='localhost', port=5000)
