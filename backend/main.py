import requests
import logging
import msal
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Dict

# Configure logging
logging.basicConfig(level=logging.INFO)


# Azure AD App credentials (updated with provided values)
TENANT_ID =   # Tenant ID
CLIENT_ID = # Client ID (App ID)
CLIENT_ID = # SC ID value

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ['https://graph.microsoft.com/.default']

## No mock data or mappings, only real-time data

app = FastAPI()

# Allow CORS for local dev
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def get_access_token():
    app_ = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    token_result = app_.acquire_token_for_client(scopes=SCOPE)
    if "access_token" in token_result:
        return token_result["access_token"]
    else:
        raise Exception(f"Token acquisition failed: {token_result.get('error_description')}")


# Generic function to fetch live data from Microsoft Graph
def fetch_graph_data(endpoint: str, token: str, params: dict = None) -> dict:
    url = f"https://graph.microsoft.com/v1.0/{endpoint}"
    headers = {'Authorization': f'Bearer {token}'}
    response = requests.get(url, headers=headers, params=params)
    response.raise_for_status()
    return response.json()


# Example: Fetch all users with selected fields (with paging)
def get_users(token):
    url = "https://graph.microsoft.com/v1.0/users?$select=displayName,userPrincipalName,assignedLicenses"
    headers = {'Authorization': f'Bearer {token}'}
    users = []
    while url:
        res = requests.get(url, headers=headers)
        res.raise_for_status()
        data = res.json()
        users.extend(data.get("value", []))
        url = data.get("@odata.nextLink", None)
    return users


# Example endpoint: fetch user licenses (live data)
@app.get("/api/users-licenses")
def users_licenses():
    try:
        token = get_access_token()
        users = get_users(token)
        result = []
        for user in users:
            name = user.get("displayName", "N/A")
            email = user.get("userPrincipalName", "N/A")
            licenses = user.get("assignedLicenses", [])
            license_ids = [lic.get("skuId") for lic in licenses if lic.get("skuId")]
            result.append({
                "displayName": name,
                "userPrincipalName": email,
                "licenses": license_ids
            })
        return {"users": result}
    except Exception as e:
        import traceback
        logging.error('Error in /api/users-licenses: %s', str(e))
        logging.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"{str(e)}\n{traceback.format_exc()}")


# Generic endpoint to fetch live data from any Graph endpoint (for extensibility)
@app.get("/api/graph")
def graph_proxy(endpoint: str):
    """
    Example: /api/graph?endpoint=users
    """
    try:
        token = get_access_token()
        data = fetch_graph_data(endpoint, token)
        return data
    except Exception as e:
        import traceback
        logging.error('Error in /api/graph: %s', str(e))
        logging.error(traceback.format_exc())
        raise HTTPException(status_code=500, detail=f"{str(e)}\n{traceback.format_exc()}")
