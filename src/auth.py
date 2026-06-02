import json
import time
import configparser
import msal

CONFIG_PATH = "config.ini"

SCOPES = [
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Mail.ReadWrite",
]


def _load_config():
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH)
    return cfg


def _get_app(cfg):
    return msal.PublicClientApplication(
        client_id=cfg["azure"]["client_id"],
        authority=f"https://login.microsoftonline.com/common",
    )


def _save_token(token: dict, path: str):
    payload = {
        "access_token": token["access_token"],
        "refresh_token": token.get("refresh_token", ""),
        "expires_at": time.time() + token.get("expires_in", 3600) - 60,
    }
    with open(path, "w") as f:
        json.dump(payload, f, indent=2)


def _load_token(path: str):
    try:
        with open(path) as f:
            return json.load(f)
    except FileNotFoundError:
        return None


def get_access_token() -> str:
    cfg = _load_config()
    token_path = cfg["settings"]["token_cache_path"]
    app = _get_app(cfg)

    cached = _load_token(token_path)

    # Token still valid
    if cached and time.time() < cached["expires_at"]:
        return cached["access_token"]

    # Try silent refresh with refresh_token
    if cached and cached.get("refresh_token"):
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                _save_token(result, token_path)
                return result["access_token"]

        # Fallback: use refresh_token directly via acquire_token_by_refresh_token
        result = app.acquire_token_by_refresh_token(
            refresh_token=cached["refresh_token"],
            scopes=SCOPES,
        )
        if result and "access_token" in result:
            _save_token(result, token_path)
            return result["access_token"]

    # First-time login: device code flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Device flow failed: {flow.get('error_description')}")

    print("\n" + "=" * 60)
    print("ONE-TIME LOGIN REQUIRED")
    print("=" * 60)
    print(f"1. Open:  {flow['verification_uri']}")
    print(f"2. Enter code: {flow['user_code']}")
    print("=" * 60 + "\n")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description')}")

    _save_token(result, token_path)
    print("Login successful. Token cached.\n")
    return result["access_token"]
