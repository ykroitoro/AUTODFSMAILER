import requests

APP_KEY = 'xa655f8nqxpv813'
APP_SECRET = 'f8sqs8r38e6idhd'
AUTH_CODE = 'op-gAGXnjaMAAAAAAAACnvAvEPGmofGF4BicPx-6sBA'

res = requests.post("https://api.dropboxapi.com/oauth2/token", data={
    "code": AUTH_CODE,
    "grant_type": "authorization_code",
    "client_id": APP_KEY,
    "client_secret": APP_SECRET,
    "redirect_uri": "http://localhost:8080/"
})

print(res.json())  # Will contain refresh_token
