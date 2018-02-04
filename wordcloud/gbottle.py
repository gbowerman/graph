'''Microsoft graph API test program'''
import base64
import datetime
import http.client
import json
import os
import urllib
import urllib.error
import urllib.parse
import urllib.request
import uuid
from string import punctuation

import adal
from bottle import (app, get, post, redirect, request, route, run, static_file,
                    view)
from json2html import *
from requests import Session
from wordcloud import WordCloud  # , STOPWORDS

# graph api constants
redirect_uri = 'http://localhost:5000/login/authorized'
resource_uri = 'https://graph.microsoft.com/'
authority_url = 'https://login.microsoftonline.com/common'
api_version = 'v1.0'

# output defaults
today = datetime.datetime.now()
to_date = str(today.date())
from_date = str((today - datetime.timedelta(days=7)).date())
search_str = 'wpa'
folder = 'Sent Items'
header = '<head><title>Graph output</title></head><body><h2>Graph output</h2><br/><br/>'
footer = '<p><a href="/">Restart</a></p></body>'

# Load graph app defaults
try:
    with open('graphconfig.json') as config_file:
        config_data = json.load(config_file)
except FileNotFoundError:
    sys.exit('Error: Expecting graphconfig.json in current folder')

client_id = config_data['appId']
client_secret = config_data['clientSecret']
text_analytics_uri = config_data['textAnalyticsURI']
text_key = config_data['textKey']

# use a custom stopwords file
with open('stopwords.txt', 'r') as stopfile:
    stopwords_str = stopfile.read()
stopwords = eval(stopwords_str)

os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'  # enable non-HTTPS for testing
SESSION = Session()


def search_form(mailfolder, fromdate, todate, searchstr):
    return '<br/><form action="/maildump" method="post">' +\
        'Folder: <input name="folder" type="text" size=50 value="' + mailfolder + '"/> ' +\
        'From: <input name="from_date" type="text" size=20 value="' + fromdate + '"/> ' +\
        'To: <input name="to_date" type="text" size=20 value="' + todate + '"/> ' +\
        'Search: <input name="search_str" type="text" size=20 value="' + searchstr + '"/> ' +\
        '<input type="submit" value="Get"/></form><br/>'


def display_payload(payload, apicall='/me'):
    '''Display JSON graph output in an HTML page'''
    form1 = '<br/><form action="/graphcall" method="post">' +\
        'api: <input name="apicall" type="text" size=50 value="' + apicall + '"/> ' +\
        '<input type="submit" value="Call"/>' +\
        ' [Examples: /me, /me/drive, /me/people/?$search=name, ' +\
        '/me/messages/?$select=subject,from&$search=Artificial]</form><br/>'
    form2 = search_form(folder, from_date, to_date, search_str)
    htmldata = json2html.convert(json=payload)
    return header + form1 + form2 + htmldata + footer


def show_analysis(output):
    '''Display text output to page'''
    form2 = search_form(folder, from_date, to_date, search_str)
    
    # check we're getting some output to analyze
    if len(output) < 4:
        errormsg = '{"Error": "No words found."}'
        return display_payload(errormsg)

    # call the Azure Text Analytics API
    # request headers
    headers = {
        'Content-Type': 'application/json',
        'Ocp-Apim-Subscription-Key': text_key,
        'Accept': 'application/json'
    }

    # define parameters
    params = urllib.parse.urlencode({})

    # request body
    body = {
        "documents": [
            {
                "language": "en",
                "id": "1",
                "text": output[:4096]
            }
        ]
    }
    keyphrase_html = '<h3>Key phrases</h3>'
    keyphrases = ""
    try:
        # call API
        conn = http.client.HTTPSConnection(text_analytics_uri)
        conn.request("POST", "/text/analytics/v2.0/keyPhrases?%s" % params, str(body).encode('UTF-8'), headers)
        response = conn.getresponse()
        data = response.read().decode('UTF-8')
        parsed = json.loads(data)
        print('Key phrases: ' + json.dumps(parsed))
        for document in parsed['documents']:
            for phrase in document['keyPhrases']:
                keyphrases += phrase + '<br/>'
        conn.close()
        keyphrase_html += keyphrases
    except Exception as e:
        keyphrase_html = 'Text analysis error: ' + str(e)

    sentiment = '<h3>Sentiment analysis</h3><table><tr><td>'
    try:
        # call API
        conn = http.client.HTTPSConnection(text_analytics_uri)
        conn.request("POST", "/text/analytics/v2.0/sentiment?%s" % params, str(body).encode('UTF-8'), headers)
        response = conn.getresponse()
        data = response.read().decode('UTF-8')
        parsed = json.loads(data)
        print('Sentiment: ' + json.dumps(parsed))
        for document in parsed['documents']:
            sentiment_val = document['score']
            sentiment += 'Score = ' + str(sentiment_val) + '</td><td>'
        conn.close()
        if sentiment_val < 0.4:
            sentiment += '<img src="/static/img/sad.png" height="100" width="100"/>'
        elif sentiment_val > 0.6:
            sentiment += '<img src="/static/img/happy.png" height="100" width="100"/>'
        else:
            sentiment += '<img src="/static/img/neutral.png" height="100" width="100"/>'
        sentiment += '</td></tr></table>'

    except Exception as e:
        sentiment = 'Sentiment analysis error: ' + str(e)

    # create a word cloud image file from the text payload
    wordcloud = WordCloud(width=1200, height=600, max_font_size=70, stopwords=stopwords).generate(output)
    wcimage = wordcloud.to_image()
    wcimage.save('static/img/wcimg.png')
    htmlimg = '<h3>Word cloud</h3><p><img src="/static/img/wcimg.png"/></p>' 

    return header + form2 + sentiment + htmlimg + keyphrase_html + footer


@route('/')
@view('homepage.html')
def homepage():
    """Render the home page."""
    return {'sample': 'Microsoft Graph'}


@route('/login')
def login():
    """Prompt user to authenticate."""
    auth_state = str(uuid.uuid4())
    SESSION.auth_state = auth_state

    #prompt_behavior = 'none'
    prompt_behavior = 'select_account'

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

@get("/static/img/<filepath:re:.*\.(jpg|png|gif|ico|svg)>")
def img(filepath):
    return static_file(filepath, root="static/img")

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
    global to_date, from_date, search_str, folder
    folder = request.forms.get('folder')
    from_date = request.forms.get('from_date')
    to_date = request.forms.get('to_date')
    search_str = request.forms.get('search_str')
    apply_search = False
    if len(search_str) > 0:
        apply_search = True
    else:
        search_str = ''
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
                if 'error' in maildata:
                    return display_payload(maildata)
                # loop through all mail data returned by query, and follow skip links
                while 'value' in maildata and len(maildata['value']) > 0:
                    for mail in maildata['value']:
                        # if there is a search string apply it manually here since graph doesn't
                        # let you mix filter and search query parameters
                        if apply_search is True:
                            slower = search_str.lower()
                            if slower in mail['subject'].lower() or slower in mail['bodyPreview'].lower():
                                mailtext += mail['subject'] + ' ' + mail['bodyPreview'] + ' '
                        else:
                            mailtext += mail['subject'] + ' ' + mail['bodyPreview']
                    mail_skip += 10
                    maildata = SESSION.get(mailendpoint + '&$skip=' + str(mail_skip), headers=http_headers, stream=False).json()

                # at this point mailtext has all the text we want
                # strip punctuation
                words = ''.join(c for c in mailtext if not c.isdigit()).lower()
                words = words.replace('microsoft.com', '')
                words = ''.join(c for c in words if c not in punctuation)
                return show_analysis(words)
        # if there was no match and there's a next link, call again
        skip_num += 10
        graphdata = SESSION.get(endpoint + '?$skip=' + str(skip_num), headers=http_headers, stream=False).json()
    return display_text(folder_rec)

if __name__ == '__main__':
    run(app=app(), server='wsgiref', host='localhost', port=5000)
