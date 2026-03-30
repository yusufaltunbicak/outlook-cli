import os
from pathlib import Path

# API
BASE_URL = "https://outlook.office.com/api/v2.0/me"
OWA_URL = "https://outlook.office365.com/mail/"
OWA_SERVICE_URL = "https://outlook.office365.com/owa/service.svc"

# User-Agent mimics Outlook Web App
USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/120.0.0.0 Safari/537.36"
)

# Cache paths
CACHE_DIR = Path(os.environ.get("OUTLOOK_CLI_CACHE", Path.home() / ".cache" / "outlook-cli"))
TOKEN_FILE = CACHE_DIR / "token.json"
BROWSER_STATE_FILE = CACHE_DIR / "browser-state.json"
ID_MAP_FILE = CACHE_DIR / "id_map.json"
ACCOUNTS_CACHE_DIR = CACHE_DIR / "accounts"

# Config paths
CONFIG_DIR = Path(os.environ.get("OUTLOOK_CLI_CONFIG", Path.home() / ".config" / "outlook-cli"))
CONFIG_FILE = CONFIG_DIR / "config.yaml"
SIGNATURES_DIR = CONFIG_DIR / "signatures"
ACCOUNTS_FILE = CONFIG_DIR / "accounts.json"
ACCOUNTS_CONFIG_DIR = CONFIG_DIR / "accounts"

# Attachment size threshold: inline base64 for files under 3 MB,
# upload session for larger files.
ATTACHMENT_SIZE_THRESHOLD = 3 * 1024 * 1024  # 3 MB

# Extended property for deferred/scheduled send
DEFERRED_SEND_PROPERTY_ID = "SystemTime 0x3FEF"

# Scheduled messages tracking
SCHEDULED_FILE = CACHE_DIR / "scheduled.json"

# Secret storage
KEYRING_SERVICE_NAME = "outlook-cli"
