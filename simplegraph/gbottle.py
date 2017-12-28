'''Microsoft graph API test program'''
import os
import urllib.parse
import uuid
import json

import adal
import bottle
from json2html import *
import requests

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
SESSION = requests.Session()
bottle.TEMPLATE_PATH = ['./static/templates']


def display_payload(payload):
    '''Display JSON graph output in an HTML page'''
    header = '<h2>Graph output</h2><br/><br/>'
    form = '<br/><form action="/graphcall" method="post">' +\
        'api: <input name="apicall" type="text" /> ' +\
        '<input type="submit" value="Call"/>' +\
        ' [Examples: /me, /me/drive, /me/people/?$search=name]</form><br/>'
    htmldata = json2html.convert(json=payload)
    footer = '<p><a href="/">Restart</a></p>'
    return header + form + htmldata + footer


@bottle.route('/')
@bottle.view('homepage.html')
def homepage():
    """Render the home page."""
    return {'sample': 'Microsoft Graph API'}


@bottle.route('/login')
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

    return bottle.redirect(authority_url + '/oauth2/authorize?' + params)


@bottle.route('/login/authorized')
def authorized():
    """Handler for the application's Redirect Uri."""
    code = bottle.request.query.code
    auth_state = bottle.request.query.state
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
    return bottle.redirect('/maincall')


@bottle.get('/maincall')
def maincall():
    """Confirm user authentication by calling Graph and displaying data."""
    apicall = '/me'
    endpoint = resource_uri + api_version + apicall
    http_headers = {'client-request-id': str(uuid.uuid4())}
    graphdata = SESSION.get(
        endpoint, headers=http_headers, stream=False).json()
    return display_payload(graphdata)


@bottle.post('/graphcall')
def graphcall():
    """Confirm user authentication by calling Graph and displaying data."""
    apicall = bottle.request.forms.get('apicall')
    #print('apicall is ' + apicall)
    endpoint = resource_uri + api_version + apicall
    http_headers = {'client-request-id': str(uuid.uuid4())}
    graphdata = SESSION.get(
        endpoint, headers=http_headers, stream=False).json()
    return display_payload(graphdata)


if __name__ == '__main__':
    bottle.run(app=bottle.app(), server='wsgiref', host='localhost', port=5000)
