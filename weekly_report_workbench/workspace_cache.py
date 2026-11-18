from __future__ import annotations

from io import BytesIO
from pathlib import Path
import hashlib
import os
import pickle
import re
import tempfile
import time
import uuid


CACHE_ROOT = Path(tempfile.gettempdir()) / "weekly-report-workspaces"
TTL_SECONDS = 24 * 60 * 60
ID_PATTERN = re.compile(r"^[a-f0-9]{32}$")

FILE_KEYS = (
    "shared_registration_files",
    "shared_order_files",
    "shared_web_google_files",
    "shared_web_bing_files",
    "shared_app_google_files",
    "shared_app_meta_files",
    "shared_app_asa_files",
    "shared_app_appsflyer_files",
    "shared_ugc_app_reg_files",
    "shared_ugc_app_order_files",
)
VALUE_KEYS = (
    "web_manual_saved",
    "web_manual_current",
    "app_manual_saved",
)


class RestoredUpload(BytesIO):
    def __init__(self, data: bytes, name: str):
        super().__init__(data)
        self.name = name


def _cache_path(workspace_id: str) -> Path:
    return CACHE_ROOT / f"{workspace_id}.pkl"


def _cleanup_expired() -> None:
    CACHE_ROOT.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - TTL_SECONDS
    for path in CACHE_ROOT.glob("*.pkl"):
        try:
            if path.stat().st_mtime < cutoff:
                path.unlink()
        except OSError:
            continue


def _serialize_files(files) -> list[dict]:
    result = []
    for file in files or []:
        position = file.tell()
        file.seek(0)
        result.append({"name": str(file.name), "data": file.read()})
        file.seek(position)
    return result


def _payload_from_state(state) -> dict:
    payload = {"saved_at": time.time(), "files": {}, "values": {}}
    for key in FILE_KEYS:
        if state.get(key):
            payload["files"][key] = _serialize_files(state[key])
    for key in VALUE_KEYS:
        if state.get(key) is not None:
            payload["values"][key] = state[key]
    return payload


def initialize_workspace(query_params, state) -> str:
    _cleanup_expired()
    raw_id = str(query_params.get("workspace", "")).lower()
    workspace_id = raw_id if ID_PATTERN.fullmatch(raw_id) else uuid.uuid4().hex
    if raw_id != workspace_id:
        query_params["workspace"] = workspace_id
    if state.get("_workspace_cache_loaded_id") == workspace_id:
        return workspace_id
    path = _cache_path(workspace_id)
    if path.exists() and time.time() - path.stat().st_mtime <= TTL_SECONDS:
        try:
            payload = pickle.loads(path.read_bytes())
            for key, items in payload.get("files", {}).items():
                state[key] = [
                    RestoredUpload(item["data"], item["name"]) for item in items
                ]
            for key, value in payload.get("values", {}).items():
                state[key] = value
            state["_workspace_restored"] = True
            state["_workspace_cache_digest"] = hashlib.sha256(
                pickle.dumps(payload, protocol=pickle.HIGHEST_PROTOCOL)
            ).hexdigest()
        except (OSError, pickle.PickleError, EOFError, KeyError, TypeError):
            state["_workspace_restored"] = False
    state["_workspace_cache_loaded_id"] = workspace_id
    return workspace_id


def save_workspace(workspace_id: str, state) -> None:
    if not ID_PATTERN.fullmatch(str(workspace_id)):
        return
    CACHE_ROOT.mkdir(parents=True, exist_ok=True)
    payload = _payload_from_state(state)
    data = pickle.dumps(payload, protocol=pickle.HIGHEST_PROTOCOL)
    digest = hashlib.sha256(data).hexdigest()
    # saved_at changes on every call, so compare the meaningful content separately.
    comparable = dict(payload)
    comparable["saved_at"] = 0
    comparable_digest = hashlib.sha256(
        pickle.dumps(comparable, protocol=pickle.HIGHEST_PROTOCOL)
    ).hexdigest()
    if state.get("_workspace_content_digest") == comparable_digest:
        return
    path = _cache_path(workspace_id)
    temporary = path.with_suffix(".tmp")
    temporary.write_bytes(data)
    os.chmod(temporary, 0o600)
    temporary.replace(path)
    state["_workspace_cache_digest"] = digest
    state["_workspace_content_digest"] = comparable_digest
