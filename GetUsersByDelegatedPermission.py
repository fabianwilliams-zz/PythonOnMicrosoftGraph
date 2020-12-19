import requests
import msal
import atexit
import os.path
import json 
import pandas as pd 
from configparser import ConfigParser

config = ConfigParser()
config.read('config.ini')

'''
print (config['delegatedpermissiononly'] ['tenantID'])
print (config['delegatedpermissiononly'] ['authority'])
print (config['delegatedpermissiononly'] ['clientID'])
print (config['delegatedpermissiononly'] ['clientSecret'])
'''

tenantID = config['delegatedpermissiononly'] ['tenantID']
authority = config['delegatedpermissiononly'] ['authority'] + tenantID
clientID = config['delegatedpermissiononly'] ['clientID']

ENDPOINT = 'https://graph.microsoft.com/v1.0'

SCOPES = [
    'User.Read.All',
    'User.Read'
]

app = msal.PublicClientApplication(clientID, authority=authority)

flow = app.initiate_device_flow(scopes=SCOPES)
if 'user_code' not in flow:
    raise Exception('Failed to create device flow')

print(flow['message'])

result = app.acquire_token_by_device_flow(flow)

if 'access_token' in result:
    result = requests.get(f'{ENDPOINT}/users', headers={'Authorization': 'Bearer ' + result['access_token']}).json()
    #Going after a specific user Diego S
    #result = requests.get(f'{ENDPOINT}/users/fa2480cc-103b-4b53-8ba6-da9720c81c2d/presence', headers={'Authorization': 'Bearer ' + result['access_token']})
    #result.raise_for_status()
    #print(result.json())
    df = pd.read_json(json.dumps(result['value']))
    # set ID column as index in Pandas Dataframe
    df = df.set_index('id')
    # print the entire datafarame from Pandas
    print(str(df))
    #print(str(df['displayName'] + " " + df['mail']))

else:
    raise Exception('no access token in result')