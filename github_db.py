"""
GitHub-backed JSON database for VER Data Portal.
Stores village records in data/ver_database.json and syncs to GitHub
so data persists across Streamlit Cloud redeployments.
"""
import json
import base64
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime
from collections import OrderedDict

DB_FILE = Path(__file__).parent / "data" / "ver_database.json"


def _read_db() -> dict:
    """Read the local JSON database file."""
    if not DB_FILE.exists():
        return {"villages": [], "metadata": {"version": "1.0", "last_updated": ""}}
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "villages" not in data:
            data["villages"] = []
        if "metadata" not in data:
            data["metadata"] = {"version": "1.0", "last_updated": ""}
        return data
    except (json.JSONDecodeError, IOError):
        return {"villages": [], "metadata": {"version": "1.0", "last_updated": ""}}


def _write_db(data: dict):
    """Write to the local JSON database file."""
    data["metadata"]["last_updated"] = datetime.now().isoformat()
    DB_FILE.parent.mkdir(parents=True, exist_ok=True)
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)


def _push_to_github(token: str, repo: str):
    """Commit the updated database file to GitHub via API."""
    if not token or not repo:
        return False, "No GitHub token or repo configured"

    file_path = "data/ver_database.json"
    api_url = f"https://api.github.com/repos/{repo}/contents/{file_path}"

    # Read current file content
    with open(DB_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    content_b64 = base64.b64encode(content.encode("utf-8")).decode("utf-8")

    # Get current file SHA (needed for update)
    sha = None
    req = urllib.request.Request(api_url, headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            existing = json.loads(resp.read().decode("utf-8"))
            sha = existing.get("sha")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            return False, f"GitHub API error: {e.code}"

    # Create or update file
    payload = {
        "message": f"Update VER database ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
        "content": content_b64,
        "branch": "main",
    }
    if sha:
        payload["sha"] = sha

    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(api_url, data=data_bytes, method="PUT", headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            if resp.status in (200, 201):
                # Also push to master branch for Streamlit Cloud deployment
                _sync_to_master(token, repo, file_path, content_b64)
                return True, "Saved to GitHub"
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        return False, f"GitHub push failed ({e.code}): {body}"

    return False, "Unknown error"


def _sync_to_master(token: str, repo: str, file_path: str, content_b64: str):
    """Sync the database file to master branch (Streamlit Cloud deploys from master)."""
    api_url = f"https://api.github.com/repos/{repo}/contents/{file_path}"

    # Get SHA on master
    sha = None
    req = urllib.request.Request(f"{api_url}?ref=master", headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            existing = json.loads(resp.read().decode("utf-8"))
            sha = existing.get("sha")
    except urllib.error.HTTPError:
        pass

    payload = {
        "message": f"Sync VER database to master ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
        "content": content_b64,
        "branch": "master",
    }
    if sha:
        payload["sha"] = sha

    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(api_url, data=data_bytes, method="PUT", headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    })
    try:
        urllib.request.urlopen(req)
    except urllib.error.HTTPError:
        pass  # Non-critical: master sync is best-effort


# ── Public API ──────────────────────────────────────────────

def load_all_villages() -> list[dict]:
    """Load all village records from the JSON database."""
    db = _read_db()
    return db.get("villages", [])


def save_village(record: dict, github_token: str = "", github_repo: str = "") -> str:
    """Save a village record. Returns a unique ID for the record."""
    db = _read_db()

    # Generate a simple unique ID
    vid = f"v_{datetime.now().strftime('%Y%m%d%H%M%S')}_{len(db['villages']) + 1}"
    record["_id"] = vid
    record["_created_at"] = datetime.now().isoformat()

    db["villages"].append(record)
    _write_db(db)

    # Push to GitHub if configured
    if github_token and github_repo:
        _push_to_github(github_token, github_repo)

    return vid


def delete_village(village_id: str, github_token: str = "", github_repo: str = ""):
    """Delete a village record by its ID."""
    db = _read_db()
    db["villages"] = [v for v in db["villages"] if v.get("_id") != village_id]
    _write_db(db)

    if github_token and github_repo:
        _push_to_github(github_token, github_repo)


def delete_all_villages(github_token: str = "", github_repo: str = ""):
    """Delete all village records."""
    db = _read_db()
    db["villages"] = []
    _write_db(db)

    if github_token and github_repo:
        _push_to_github(github_token, github_repo)


def get_village_count() -> int:
    """Return the number of stored villages."""
    db = _read_db()
    return len(db.get("villages", []))


def import_villages(records: list[dict], github_token: str = "", github_repo: str = ""):
    """Import a list of village records (replaces all existing data)."""
    db = _read_db()
    db["villages"] = []
    for i, rec in enumerate(records):
        if "_id" not in rec:
            rec["_id"] = f"v_import_{datetime.now().strftime('%Y%m%d%H%M%S')}_{i+1}"
        if "_created_at" not in rec:
            rec["_created_at"] = datetime.now().isoformat()
        db["villages"].append(rec)
    _write_db(db)

    if github_token and github_repo:
        _push_to_github(github_token, github_repo)


def sync_to_github(github_token: str, github_repo: str) -> tuple[bool, str]:
    """Manually trigger a sync to GitHub. Returns (success, message)."""
    return _push_to_github(github_token, github_repo)
