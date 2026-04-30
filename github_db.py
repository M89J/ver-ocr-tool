"""
GitHub-backed JSON database for VER Data Portal.
Stores village records in data/ver_database.json and syncs to GitHub
so data persists across Streamlit Cloud redeployments.
Data is append-only — no delete operations.
"""
import re
import json
import base64
import urllib.request
import urllib.error
from pathlib import Path
from datetime import datetime
from collections import OrderedDict


# Fields whose value is a "; "-separated list of species/items.
# When merging, take the UNION of old and new entries (case-insensitive dedup).
# The paired *_count field is recomputed after the union.
LIST_FIELDS = {
    "tree_diversity": "tree_diversity_count",
    "shrub_diversity": "shrub_diversity_count",
    "herb_grass_diversity": "herb_grass_diversity_count",
    "lower_plant_diversity": "lower_plant_count",
    "mammal_diversity": "mammal_count",
    "bird_diversity": "bird_count",
    "reptile_amphibian_diversity": "reptile_amphibian_count",
    "butterfly_diversity": "butterfly_count",
    "dragonfly_diversity": "dragonfly_count",
    "fish_insect_other_diversity": "fish_insect_other_count",
    "soil_macrofauna_diversity": "soil_macrofauna_count",
}

# Fields whose value is a "; "-separated list of "label:count" pairs.
# When merging, take max(old, new) per label.
COUNTED_PAIR_FIELDS = {
    "livestock_summary",
    "livestock_detailed",
    "drinking_water_sources",
    "livestock_water_sources",
}


def _split_list(s: str) -> list[str]:
    if not s:
        return []
    return [p.strip() for p in str(s).split(";") if p.strip()]


def _merge_list_field(old_str: str, new_str: str, max_items: int = 100) -> tuple[str, int]:
    """Union of "; "-separated lists, dedup case-insensitively, preserve insertion order."""
    seen = set()
    merged = []
    for item in _split_list(old_str) + _split_list(new_str):
        key = item.lower().strip()
        if key and key not in seen:
            seen.add(key)
            merged.append(item)
    merged = merged[:max_items]
    return "; ".join(merged), len(merged)


def _merge_counted_pairs(old_str: str, new_str: str) -> str:
    """Per label, keep max(old, new). Preserves order from old then appends new labels."""
    pat = re.compile(r'([A-Za-z][A-Za-z ]*?):\s*(\d+)')
    old_pairs = OrderedDict()
    for k, v in pat.findall(old_str or ""):
        old_pairs[k.strip()] = int(v)
    for k, v in pat.findall(new_str or ""):
        key = k.strip()
        old_pairs[key] = max(old_pairs.get(key, 0), int(v))
    return "; ".join(f"{k}:{v}" for k, v in old_pairs.items() if v > 0)

DB_FILE = Path(__file__).parent / "data" / "ver_database.json"


def _read_db() -> dict:
    """Read the local JSON database file."""
    if not DB_FILE.exists():
        return {"villages": [], "metadata": {"version": "2.0", "last_updated": ""}}
    try:
        with open(DB_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if "villages" not in data:
            data["villages"] = []
        if "metadata" not in data:
            data["metadata"] = {"version": "2.0", "last_updated": ""}
        return data
    except (json.JSONDecodeError, IOError):
        return {"villages": [], "metadata": {"version": "2.0", "last_updated": ""}}


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

    with open(DB_FILE, "r", encoding="utf-8") as f:
        content = f.read()
    content_b64 = base64.b64encode(content.encode("utf-8")).decode("utf-8")

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
                _sync_to_master(token, repo, file_path, content_b64)
                return True, "Saved to GitHub"
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        return False, f"GitHub push failed ({e.code}): {body}"

    return False, "Unknown error"


def _sync_to_master(token: str, repo: str, file_path: str, content_b64: str):
    """Sync the database file to master branch (Streamlit Cloud deploys from master)."""
    api_url = f"https://api.github.com/repos/{repo}/contents/{file_path}"

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
        pass


# ── Public API ──────────────────────────────────────────────

def load_all_villages() -> list[dict]:
    """Load all village records from the JSON database."""
    db = _read_db()
    return db.get("villages", [])


def find_village(village_name: str, state: str) -> tuple[int, dict | None]:
    """Find a village by name + state (case-insensitive). Returns (index, record) or (-1, None)."""
    db = _read_db()
    name_lower = village_name.strip().lower()
    state_lower = state.strip().lower()
    for i, v in enumerate(db["villages"]):
        if (v.get("village_name", "").strip().lower() == name_lower and
                v.get("state", "").strip().lower() == state_lower):
            return i, v
    return -1, None


def upsert_village(record: dict, github_token: str = "", github_repo: str = "") -> tuple[str, bool]:
    """Insert or update a village record. Returns (village_id, was_update).

    Matching is by village_name + state (case-insensitive).
    On update: merges new data into existing — prefers new non-empty values,
    keeps old values where new is empty/0.
    """
    db = _read_db()
    name = record.get("village_name", "").strip()
    state = record.get("state", "").strip()

    # Find existing record by name (+ state if available)
    existing_idx = -1
    if name:
        name_lower = name.lower()
        state_lower = state.lower() if state else ""
        for i, v in enumerate(db["villages"]):
            v_name = v.get("village_name", "").strip().lower()
            v_state = v.get("state", "").strip().lower()
            if v_name == name_lower and (not state_lower or not v_state or v_state == state_lower):
                existing_idx = i
                break

    if existing_idx >= 0:
        # UPDATE — merge new into existing
        existing = db["villages"][existing_idx]
        merged = dict(existing)  # start with old values

        for key, new_val in record.items():
            if key.startswith("_"):
                continue
            old_val = existing.get(key, "")

            # Multi-value species lists → union (preserves accumulated knowledge)
            if key in LIST_FIELDS:
                merged_str, count = _merge_list_field(old_val, new_val)
                merged[key] = merged_str
                merged[LIST_FIELDS[key]] = count
                continue

            # Counted-pair lists (livestock, water) → per-label max
            if key in COUNTED_PAIR_FIELDS:
                merged[key] = _merge_counted_pairs(old_val, new_val)
                continue

            # Skip *_count fields — recomputed alongside their *_diversity field above
            if key.endswith("_count") and key != "total_pages":
                continue

            # Default scalar/text behavior — prefer non-empty / non-zero new value
            if new_val and new_val != 0:
                merged[key] = new_val

        # Recompute total_species_count from the merged *_count fields
        merged["total_species_count"] = sum(
            int(merged.get(c, 0) or 0) for c in LIST_FIELDS.values()
        )

        merged["_updated_at"] = datetime.now().isoformat()
        db["villages"][existing_idx] = merged
        vid = merged.get("_id", "")
        was_update = True
    else:
        # INSERT — append new record
        vid = f"v_{datetime.now().strftime('%Y%m%d%H%M%S')}_{len(db['villages']) + 1}"
        record["_id"] = vid
        record["_created_at"] = datetime.now().isoformat()
        db["villages"].append(record)
        was_update = False

    _write_db(db)

    if github_token and github_repo:
        _push_to_github(github_token, github_repo)

    return vid, was_update


def save_village(record: dict, github_token: str = "", github_repo: str = "") -> str:
    """Save a village record (uses upsert). Returns the village ID."""
    vid, _ = upsert_village(record, github_token, github_repo)
    return vid


def get_village_count() -> int:
    """Return the number of stored villages."""
    db = _read_db()
    return len(db.get("villages", []))


def import_villages(records: list[dict], github_token: str = "", github_repo: str = ""):
    """Import village records using upsert (no duplicates, no data loss)."""
    for rec in records:
        upsert_village(rec, github_token=github_token, github_repo=github_repo)


def upload_pdf_to_github(pdf_bytes: bytes, filename: str, github_token: str, github_repo: str) -> tuple[bool, str]:
    """Upload a raw PDF to data/pdfs/ on GitHub. Returns (success, download_url)."""
    if not github_token or not github_repo:
        return False, "No GitHub token or repo configured"

    # Sanitize filename
    safe_name = "".join(c for c in filename if c.isalnum() or c in "._- ").strip()
    file_path = f"data/pdfs/{safe_name}"
    api_url = f"https://api.github.com/repos/{github_repo}/contents/{file_path}"

    content_b64 = base64.b64encode(pdf_bytes).decode("utf-8")

    # Check if file exists (get SHA)
    sha = None
    req = urllib.request.Request(api_url, headers={
        "Authorization": f"Bearer {github_token}",
        "Accept": "application/vnd.github.v3+json",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            existing = json.loads(resp.read().decode("utf-8"))
            sha = existing.get("sha")
    except urllib.error.HTTPError:
        pass

    payload = {
        "message": f"Upload VER PDF: {safe_name}",
        "content": content_b64,
        "branch": "main",
    }
    if sha:
        payload["sha"] = sha

    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(api_url, data=data_bytes, method="PUT", headers={
        "Authorization": f"Bearer {github_token}",
        "Accept": "application/vnd.github.v3+json",
        "Content-Type": "application/json",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            if resp.status in (200, 201):
                result = json.loads(resp.read().decode("utf-8"))
                download_url = result.get("content", {}).get("download_url", "")
                # Also sync to master
                _sync_file_to_master(github_token, github_repo, file_path, content_b64)
                return True, download_url
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        return False, f"Upload failed ({e.code}): {body[:200]}"

    return False, "Unknown error"


def _sync_file_to_master(token: str, repo: str, file_path: str, content_b64: str):
    """Sync any file to master branch."""
    api_url = f"https://api.github.com/repos/{repo}/contents/{file_path}"
    sha = None
    req = urllib.request.Request(f"{api_url}?ref=master", headers={
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github.v3+json",
    })
    try:
        with urllib.request.urlopen(req) as resp:
            sha = json.loads(resp.read().decode("utf-8")).get("sha")
    except urllib.error.HTTPError:
        pass
    payload = {"message": f"Sync PDF to master: {file_path}", "content": content_b64, "branch": "master"}
    if sha:
        payload["sha"] = sha
    req = urllib.request.Request(api_url, data=json.dumps(payload).encode("utf-8"), method="PUT", headers={
        "Authorization": f"Bearer {token}", "Accept": "application/vnd.github.v3+json", "Content-Type": "application/json",
    })
    try:
        urllib.request.urlopen(req)
    except urllib.error.HTTPError:
        pass


def get_pdf_download_url(filename: str, github_repo: str) -> str:
    """Get the raw download URL for a stored PDF."""
    safe_name = "".join(c for c in filename if c.isalnum() or c in "._- ").strip()
    return f"https://raw.githubusercontent.com/{github_repo}/main/data/pdfs/{safe_name}"


def sync_to_github(github_token: str, github_repo: str) -> tuple[bool, str]:
    """Manually trigger a sync to GitHub. Returns (success, message)."""
    return _push_to_github(github_token, github_repo)
