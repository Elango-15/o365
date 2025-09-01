from flask import Flask, jsonify, request
from flask_cors import CORS
from msal import ConfidentialClientApplication
import os
from dotenv import load_dotenv
import json
from pathlib import Path
import uuid
from threading import Lock
from typing import Tuple
from cryptography.fernet import Fernet, InvalidToken
import requests

load_dotenv()

app = Flask(__name__)
# CORS configuration to handle any localhost port
CORS(app, origins=["http://localhost:8080", "http://localhost:8081", "http://localhost:8082", "http://localhost:8083", "http://localhost:8084", "http://localhost:8085", "http://localhost:8086", "http://localhost:8087", "http://localhost:8088", "http://localhost:8089", "http://localhost:8090", "http://localhost:5173", "http://localhost:3000", "http://127.0.0.1:8080", "http://127.0.0.1:8081", "http://127.0.0.1:8082", "http://127.0.0.1:8083", "http://127.0.0.1:8084", "http://127.0.0.1:8085", "http://127.0.0.1:8086", "http://127.0.0.1:8087", "http://127.0.0.1:8088", "http://127.0.0.1:8089", "http://127.0.0.1:8090", "http://127.0.0.1:5173", "http://127.0.0.1:3000"], supports_credentials=True)

# Simple file-backed storage for tenant configurations
_TENANTS_FILE = Path(__file__).parent / "tenants.json"
_TENANTS_LOCK = Lock()
_KEY_FILE = Path(__file__).parent / "tenant_key.key"

def _load_fernet() -> Fernet:
    key = os.environ.get("TENANT_SECRET_KEY", "").strip()
    if not key:
        # Use or create local key file for development
        if _KEY_FILE.exists():
            key = _KEY_FILE.read_text(encoding="utf-8").strip()
        else:
            key = Fernet.generate_key().decode("utf-8")
            _KEY_FILE.write_text(key, encoding="utf-8")
    return Fernet(key.encode("utf-8"))

_FERNET = _load_fernet()

def _encrypt_secret(plain: str) -> str:
    if plain is None:
        return ""
    token = _FERNET.encrypt(plain.encode("utf-8"))
    return token.decode("utf-8")

def _looks_encrypted(value: str) -> bool:
    return isinstance(value, str) and value.startswith("gAAAA") and len(value) > 20

def _read_tenants() -> list:
    with _TENANTS_LOCK:
        if not _TENANTS_FILE.exists():
            return []
        try:
            tenants = json.loads(_TENANTS_FILE.read_text(encoding="utf-8"))
            # Migrate any plaintext secrets to encrypted form
            changed = False
            for t in tenants:
                secret = t.get("clientSecret")
                if isinstance(secret, str) and secret and not _looks_encrypted(secret):
                    t["clientSecret"] = _encrypt_secret(secret)
                    changed = True
            if changed:
                _TENANTS_FILE.write_text(json.dumps(tenants, ensure_ascii=False, indent=2), encoding="utf-8")
            return tenants
        except Exception:
            return []

def _write_tenants(tenants: list) -> None:
    with _TENANTS_LOCK:
        _TENANTS_FILE.write_text(json.dumps(tenants, ensure_ascii=False, indent=2), encoding="utf-8")

# No hardcoded credentials - all tenant data comes from tenant management
SCOPE = ["https://graph.microsoft.com/.default"]

@app.route('/api/token', methods=['GET'])
def get_token():
    # This endpoint is deprecated - use tenant-specific endpoints instead
    return jsonify({
        "error": "No default credentials configured. Please configure tenants in the Tenant Management tab.",
        "message": "Use tenant-specific data endpoints instead."
    }), 400

@app.route('/api/users', methods=['GET'])
def get_users():
    # This endpoint is deprecated - use tenant-specific endpoints instead
    return jsonify({
        "error": "No default credentials configured. Please configure tenants in the Tenant Management tab.",
        "message": "Use tenant-specific data endpoints instead.",
        "value": []
    }), 400

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({"status": "healthy", "message": "Flask backend is running"})

@app.route('/api/metrics', methods=['GET'])
def get_metrics():
    # This endpoint is deprecated - use tenant-specific endpoints instead
    return jsonify({
        "error": "No default credentials configured. Please configure tenants in the Tenant Management tab.",
        "message": "Use tenant-specific data endpoints instead.",
        "totalUsers": 0,
        "activeUsers": 0,
        "disabledUsers": 0,
        "totalLicenses": 0,
        "usedLicenses": 0,
        "availableLicenses": 0,
        "userStatus": {"active": 0, "disabled": 0},
        "licenseStatus": {"used": 0, "available": 0}
    }), 400

# ----------------------
# Tenant CRUD Endpoints
# ----------------------

@app.route('/api/tenants', methods=['GET'])
def list_tenants():
    tenants = _read_tenants()
    # Do not return secrets; expose hasSecret flag
    redacted = []
    for t in tenants:
        item = dict(t)
        item.pop("clientSecret", None)
        item["hasSecret"] = bool(t.get("clientSecret"))
        redacted.append(item)
    return jsonify({"tenants": redacted})

@app.route('/api/tenants', methods=['POST'])
def create_tenant():
    try:
        payload = request.get_json(force=True) or {}
        required = ['name', 'tenantId', 'clientId', 'clientSecret']
        if not all(k in payload and str(payload[k]).strip() for k in required):
            return jsonify({"error": "Missing required fields"}), 400

        tenants = _read_tenants()
        new_item = {
            "id": str(uuid.uuid4()),
            "name": str(payload['name']).strip(),
            "tenantId": str(payload['tenantId']).strip(),
            "clientId": str(payload['clientId']).strip(),
            "clientSecret": _encrypt_secret(str(payload['clientSecret']).strip()),
            "isActive": bool(payload.get('isActive', True)),
            "lastSync": payload.get('lastSync', ''),
            "userCount": int(payload.get('userCount', 0) or 0),
            "licenseCount": int(payload.get('licenseCount', 0) or 0),
        }
        tenants.insert(0, new_item)
        _write_tenants(tenants)
        # return without secret
        resp = dict(new_item)
        resp.pop("clientSecret", None)
        resp["hasSecret"] = True
        return jsonify(resp), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/tenants/<tenant_id>', methods=['PUT'])
def update_tenant(tenant_id):
    try:
        payload = request.get_json(force=True) or {}
        tenants = _read_tenants()
        updated = None
        for idx, t in enumerate(tenants):
            if t.get('id') == tenant_id:
                # Update base fields
                t.update({
                    "name": str(payload.get('name', t.get('name', ''))).strip(),
                    "tenantId": str(payload.get('tenantId', t.get('tenantId', ''))).strip(),
                    "clientId": str(payload.get('clientId', t.get('clientId', ''))).strip(),
                    "isActive": bool(payload.get('isActive', t.get('isActive', True))),
                    "lastSync": payload.get('lastSync', t.get('lastSync', '')),
                    "userCount": int(payload.get('userCount', t.get('userCount', 0)) or 0),
                    "licenseCount": int(payload.get('licenseCount', t.get('licenseCount', 0)) or 0),
                })
                # Update secret only if provided and non-empty
                if 'clientSecret' in payload and str(payload['clientSecret']).strip():
                    t['clientSecret'] = _encrypt_secret(str(payload['clientSecret']).strip())
                updated = tenants[idx]
                break
        if not updated:
            return jsonify({"error": "Tenant not found"}), 404
        _write_tenants(tenants)
        resp = dict(updated)
        resp.pop("clientSecret", None)
        resp["hasSecret"] = bool(updated.get("clientSecret"))
        return jsonify(resp)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/tenants/<tenant_id>', methods=['DELETE'])
def delete_tenant(tenant_id):
    try:
        tenants = _read_tenants()
        new_list = [t for t in tenants if t.get('id') != tenant_id]
        if len(new_list) == len(tenants):
            return jsonify({"error": "Tenant not found"}), 404
        _write_tenants(new_list)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ----------------------
# Live Data Endpoints using Tenant Credentials
# ----------------------

def _decrypt_secret(encrypted: str) -> str:
    """Decrypt a tenant secret"""
    if not encrypted or not _looks_encrypted(encrypted):
        return encrypted  # Return as-is if not encrypted
    try:
        return _FERNET.decrypt(encrypted.encode("utf-8")).decode("utf-8")
    except (InvalidToken, Exception):
        return ""  # Return empty if decryption fails

def _get_tenant_credentials(tenant_id: str) -> Tuple[str, str, str]:
    """Get decrypted credentials for a specific tenant"""
    tenants = _read_tenants()
    for t in tenants:
        if t.get('id') == tenant_id and t.get('isActive', False):
            return (
                t.get('tenantId', ''),
                t.get('clientId', ''),
                _decrypt_secret(t.get('clientSecret', ''))
            )
    return ('', '', '')

def _collect_tenant_data(tenant_id: str):
    """Internal helper to collect tenant data; returns (payload_dict, status_code)."""
    try:
        tenant_id_val, client_id, client_secret = _get_tenant_credentials(tenant_id)

        if not all([tenant_id_val, client_id, client_secret]):
            return {"error": "Tenant not found or missing credentials"}, 404

        # Create MSAL app for this specific tenant
        authority = f"https://login.microsoftonline.com/{tenant_id_val}"
        app_msal = ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )

        # Get token for this tenant
        token = app_msal.acquire_token_for_client(scopes=SCOPE)
        if "access_token" not in token:
            return {"error": "Failed to acquire token for tenant", "details": token}, 500

        headers = {"Authorization": f"Bearer {token['access_token']}"}

        # Fetch all data in parallel
        import concurrent.futures

        def fetch_users():
            try:
                response = requests.get("https://graph.microsoft.com/v1.0/users", headers=headers)
                response.raise_for_status()
                return response.json()
            except Exception:
                return {"value": []}

        def fetch_licenses():
            try:
                response = requests.get("https://graph.microsoft.com/v1.0/subscribedSkus", headers=headers)
                response.raise_for_status()
                return response.json()
            except Exception:
                return {"value": []}

        def fetch_groups():
            try:
                response = requests.get("https://graph.microsoft.com/v1.0/groups", headers=headers)
                response.raise_for_status()
                return response.json()
            except Exception:
                return {"value": []}

        def fetch_sites():
            try:
                response = requests.get("https://graph.microsoft.com/v1.0/sites?search=*", headers=headers)
                response.raise_for_status()
                return response.json()
            except Exception:
                return {"value": []}

        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
            future_users = executor.submit(fetch_users)
            future_licenses = executor.submit(fetch_licenses)
            future_groups = executor.submit(fetch_groups)
            future_sites = executor.submit(fetch_sites)

            users_data = future_users.result()
            licenses_data = future_licenses.result()
            groups_data = future_groups.result()
            sites_data = future_sites.result()

        users = users_data.get('value', [])
        licenses = licenses_data.get('value', [])
        groups = groups_data.get('value', [])
        sites = sites_data.get('value', [])

        # Calculate metrics
        total_users = len(users)
        active_users = len([u for u in users if u.get('accountEnabled', True)])
        disabled_users = total_users - active_users

        total_licenses = sum(lic.get('prepaidUnits', {}).get('enabled', 0) for lic in licenses)
        used_licenses = sum(lic.get('consumedUnits', 0) for lic in licenses)
        available_licenses = total_licenses - used_licenses

        payload = {
            "users": users,
            "groups": groups,
            "sites": sites,
            "licenses": licenses,
            "metrics": {
                "totalUsers": total_users,
                "activeUsers": active_users,
                "disabledUsers": disabled_users,
                "totalLicenses": total_licenses,
                "usedLicenses": used_licenses,
                "availableLicenses": available_licenses,
                "userStatus": {
                    "active": active_users,
                    "disabled": disabled_users
                },
                "licenseStatus": {
                    "used": used_licenses,
                    "available": available_licenses
                }
            }
        }

        # Update tenant sync info
        tenants = _read_tenants()
        for t in tenants:
            if t.get('id') == tenant_id:
                from datetime import datetime
                t['lastSync'] = datetime.now().isoformat()
                t['userCount'] = payload['metrics']['totalUsers']
                t['licenseCount'] = payload['metrics']['totalLicenses']
                break
        _write_tenants(tenants)

        return payload, 200
    except Exception as e:
        return {"error": str(e)}, 500

@app.route('/api/tenants/<tenant_id>/data', methods=['GET'])
def get_tenant_data(tenant_id):
    """Fetch live data from a specific tenant"""
    payload, status = _collect_tenant_data(tenant_id)
    return jsonify(payload), status

@app.route('/api/tenants/<tenant_id>/sync', methods=['POST'])
def sync_tenant_data(tenant_id):
    """Sync and update tenant data"""
    try:
        # Get live data for this tenant via helper
        payload, status = _collect_tenant_data(tenant_id)
        if status != 200:
            return jsonify(payload), status

        # Update the tenant record (already updated inside helper as well)
        tenants = _read_tenants()
        updated_tenant = None
        for t in tenants:
            if t.get('id') == tenant_id:
                updated_tenant = dict(t)
                break

        if not updated_tenant:
            return jsonify({"error": "Tenant not found"}), 404

        # Return updated tenant without secret
        result = dict(updated_tenant)
        result.pop("clientSecret", None)
        result["hasSecret"] = bool(updated_tenant.get("clientSecret"))

        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    print("Starting Flask backend on http://127.0.0.1:5000")
    print("Frontend can run on any localhost port (8080-8090, 5173, 3000)")
    print("Make sure to use http://127.0.0.1:5000 (not localhost) for best compatibility")
    app.run(port=5000, debug=True, host='127.0.0.1')  # Use 127.0.0.1 instead of 0.0.0.0 for security 