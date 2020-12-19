import msal
import jwt
import json 
import requests 
import pandas as pd 
from datetime import datetime
from configparser import ConfigParser

config = ConfigParser()
config.read('config.ini')

'''
print (config['apppermissiononly'] ['tenantID'])
print (config['apppermissiononly'] ['authority'])
print (config['apppermissiononly'] ['clientID'])
print (config['apppermissiononly'] ['clientSecret'])
'''

accessToken = None 
requestHeaders = None 
tokenExpiry = None 
queryResults = None 
graphURI = 'https://graph.microsoft.com'

def msgraph_auth():
    global accessToken
    global requestHeaders
    global tokenExpiry
    tenantID = config['apppermissiononly'] ['tenantID']
    authority = config['apppermissiononly'] ['authority'] + tenantID
    clientID = config['apppermissiononly'] ['clientID']
    clientSecret = config['apppermissiononly'] ['clientSecret']
    scope = ['https://graph.microsoft.com/.default']

    app = msal.ConfidentialClientApplication(clientID, authority=authority, client_credential = clientSecret)

    try:
        accessToken = app.acquire_token_silent(scope, account=None)
        if not accessToken:
            try:
                accessToken = app.acquire_token_for_client(scopes=scope)
                if accessToken['access_token']:
                    print('New access token retreived....')
                    requestHeaders = {'Authorization': 'Bearer ' + accessToken['access_token']}
                else:
                    print('Error Caught: Check the Config File to make sure tenantID, clientID and clientSecret is correct')
            except:
                pass 
        else:
            print('Token retreived from MSAL Cache....')
        return
    except Exception as err:
        print(err)

def msgraph_request(resource,requestHeaders):
    # Request
    results = requests.get(resource, headers=requestHeaders).json()
    return results

# Call the Authenticate against Graph Function
msgraph_auth()

#going after the V1 endpoint for users
queryResults = msgraph_request(graphURI +'/v1.0/users',requestHeaders)

# Send the results to a Pandas Dataframe
try:
    df = pd.read_json(json.dumps(queryResults['value']))
    # set ID column as index in Pandas Dataframe
    df = df.set_index('id')
    # print the entire datafarame from Pandas
    print(str(df))
    #print(str(df['displayName'] + " " + df['mail']))

except:
    print(json.dumps(queryResults, indent=2))