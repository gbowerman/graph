'''Microsoft graph API test program'''
from json2html import *
import json
import sys
import uuid

import requests
from flask import Flask, redirect, url_for, session, request, render_template, Markup
from flask_oauthlib.client import OAuth

# Load Azure app defaults
try:
    with open('graphconfig.json') as config_file:
        config_data = json.load(config_file)
except FileNotFoundError:
    sys.exit('Error: Expecting graphconfig.json in current folder')

client_id = config_data['appId']
client_secret = config_data['appSecret']

app = Flask(__name__)
app.debug = True
app.secret_key = 'development'
oauth = OAuth(app)

# local, no HTTPS, disable InsecureRequestWarning
requests.packages.urllib3.disable_warnings()

msgraphapi = oauth.remote_app(
    'microsoft',
    consumer_key=client_id,
    consumer_secret=client_secret,
    request_token_params={'scope': 'User.Read Mail.Read People.Read'},
    base_url='https://graph.microsoft.com/v1.0/',
    request_token_url=None,
    access_token_method='POST',
    access_token_url='https://login.microsoftonline.com/common/oauth2/v2.0/token',
    authorize_url='https://login.microsoftonline.com/common/oauth2/v2.0/authorize')


@app.route('/')
def index():
    """Handler for home page."""
    return render_template('connect.html')


@app.route('/login')
def login():
    """Handler for login route."""
    guid = uuid.uuid4()  # guid used to only accept initiated logins
    session['state'] = guid
    return msgraphapi.authorize(callback=url_for('authorized', _external=True), state=guid)


@app.route('/logout')
def logout():
    """Handler for logout route."""
    session.pop('microsoft_token', None)
    session.pop('state', None)
    return redirect(url_for('index'))


@app.route('/login/authorized')
def authorized():
    """Handler for login/authorized route."""
    response = msgraphapi.authorized_response()

    if response is None:
        return "Access Denied: Reason={0}\nError={1}".format(
            request.args['error'], request.args['error_description'])

    # Check response for state
    if str(session['state']) != str(request.args['state']):
        raise Exception('State has been messed with, end authentication')
    session['state'] = ''  # reset session state to prevent re-use

    # Okay to store this in a local variable, encrypt if it's going to client
    # machine or database. Treat as a password.
    session['microsoft_token'] = (response['access_token'], '')
    # Store the token in another session variable for easy access
    session['access_token'] = response['access_token']
    me_response = msgraphapi.get('me')
    me_data = me_response.data
    session['alias'] = me_data['displayName']
    session['givenName'] = me_data['givenName']
    session['userEmailAddress'] = me_data['userPrincipalName']
    return redirect('main')


@app.route('/main')
def main():
    """Handler for main route."""
    if session['alias']:
        graph_data = Markup('<br/>Hi ' + session['givenName'])
        return render_template('main.html', graphDump=graph_data)


@app.route('/me')
def me():
    response = call_users_endpoint(session['access_token'])
    graph_data = Markup(json2html.convert(json=response))
    return render_template('main.html', graphDump=graph_data)

@app.route('/user_lookup')
def user_lookup():
    """Handler for user_lookup route."""
    access_token = session['access_token']
    user_name = request.args.get('username') # get name from form
    users_url = 'https://graph.microsoft.com/v1.0/me/people/?$search="' + user_name + '"'
    response = call_endpoint(access_token, users_url)
    graph_data = Markup(json2html.convert(json=response))
    return render_template('main.html', graphDump=graph_data)

@app.route('/onedrive')
def onedrive():
    """Handler for user_lookup route."""
    access_token = session['access_token']
    users_url = 'https://graph.microsoft.com/v1.0/me/drive'
    response = call_endpoint(access_token, users_url)
    graph_data = Markup(json2html.convert(json=response))
    return render_template('main.html', graphDump=graph_data)

@msgraphapi.tokengetter
def get_token():
    """Return the Oauth token."""
    return session.get('microsoft_token')

def call_endpoint(access_token, url):
    # set request headers
    headers = {'User-Agent': 'python_gtest/1.0',
               'Authorization': 'Bearer {0}'.format(access_token),
               'Accept': 'application/json',
               'Content-Type': 'application/json'}

    # Headers to instrument calls
    request_id = str(uuid.uuid4())
    instrumentation = {'client-request-id': request_id, 'return-client-request-id': 'true'}
    headers.update(instrumentation)

    response = requests.get(url=url, headers=headers, verify=False, params=None)

    if response.ok:
        return response.text
    else:
        return '{0}: {1}'.format(response.status_code, response.text)

def call_users_endpoint(access_token):
    """Call the resource URL for the sendMail action."""
    users_url = 'https://graph.microsoft.com/v1.0/users/' + session['userEmailAddress']
    return call_endpoint(access_token, users_url)