import requests
import json
import time
import random 

# Register the azure app first and make sure the app has the following permissions:
# files: Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
# user: User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
# mail: Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
# After registration, you must click on behalf of xxx to grant administrator consent, otherwise outlook api cannot be called
# Note: Added Directory.Read.All for subscription monitoring

calls = [
    'https://graph.microsoft.com/v1.0/me/drive/root',
    'https://graph.microsoft.com/v1.0/me/drive',
    'https://graph.microsoft.com/v1.0/drive/root',
    'https://graph.microsoft.com/v1.0/users',
    'https://graph.microsoft.com/v1.0/me/messages',
    'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',
    'https://graph.microsoft.com/v1.0/me/drive/root/children',
    'https://api.powerbi.com/v1.0/myorg/apps',
    'https://graph.microsoft.com/v1.0/me/mailFolders',
    'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',
    'https://graph.microsoft.com/v1.0/applications?$count=true',
    'https://graph.microsoft.com/v1.0/me/?$select=displayName,skills',
    'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',
    'https://graph.microsoft.com/beta/me/outlook/masterCategories',
    'https://graph.microsoft.com/beta/me/messages?$select=internetMessageHeaders&$top=1',
    'https://graph.microsoft.com/v1.0/sites/root/lists',
    'https://graph.microsoft.com/v1.0/sites/root',
    'https://graph.microsoft.com/v1.0/sites/root/drives'
]


def get_access_token(refresh_token, client_id, client_secret):
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    data = {
        'grant_type': 'refresh_token',
        'refresh_token': refresh_token,
        'client_id': client_id,
        'client_secret': client_secret,
        'redirect_uri': 'http://localhost:53682/'
    }
    html = requests.post('https://login.microsoftonline.com/common/oauth2/v2.0/token', data=data, headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    return access_token

def monitor_expiry(access_token):
    session = requests.Session()
    session.headers.update({
        'Authorization': access_token,
        'Content-Type': 'application/json'
    })
    endpoint = 'https://graph.microsoft.com/beta/directory/subscriptions'
    try:
        response = session.get(endpoint)
        if response.status_code == 200:
            subscriptions = response.json().get('value', [])
            developer_sku_id = 'c42b9cae-ea4f-4ab7-9717-81576235ccac'  # SKU ID for Microsoft 365 E5 Developer
            for sub in subscriptions:
                if sub.get('skuId') == developer_sku_id:
                    expiry_date = sub.get('nextLifecycleDateTime')
                    print(f"Subscription Expiry/Renewal Date: {expiry_date}")
                    return expiry_date
            print("Developer subscription not found.")
        else:
            print(f"Failed to fetch subscriptions: {response.status_code} - {response.text}")
    except requests.exceptions.RequestException as e:
        print(e)

def main():
    access_token = get_access_token(refresh_token, client_id, client_secret)
    
    # Monitor expiry before running calls
    monitor_expiry(access_token)
    
    session = requests.Session()
    session.headers.update({
        'Authorization': access_token,
        'Content-Type': 'application/json'
    })
    
    # Reduce aggression: Select only 4-6 random endpoints
    endpoints = random.sample(calls, random.randint(4, 6))
    
    num = 0
    for endpoint in endpoints:
        try:
            # Add variable sleep delay (5-10 seconds) to reduce aggression
            time.sleep(random.uniform(5, 10))
            response = session.get(endpoint)
            if response.status_code == 200:
                num += 1
                print(f'{num}th Call successful')
        except requests.exceptions.RequestException as e:
            print(e)
            pass
    
    localtime = time.asctime(time.localtime(time.time()))
    print('The end of this run is :', localtime)
    print('Number of calls is :', str(len(endpoints)))

# Reduced outer loop to 1 run for less aggression; run manually/infrequently (e.g., once a week)
main()
