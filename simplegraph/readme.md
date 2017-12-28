# Simple Microsoft Graph API connection tool

This code is derived from the [Python authentication samples for Microsoft Graph](https://github.com/microsoftgraph/python-sample-auth). 

Like the Microsoft sample bottle app, it is based on Python 3 and bottle, with the addition of a free-form API call form.

To run this code, the setup steps are the same as in the above example, except that instead of a config.py file, it uses a file called graphconfig.json to store the application ID and application key.

To run the bottle server on localhost:5000, run:

python gbottle.py

![Graph app image](./img/graphapp.PNG)
