"""Google Drive 共有ドライブ操作の薄いラッパー。

ローカル PC では Drive File Stream マウント（ファイルシステム経由）が使えるが、
Render などのクラウド環境では Drive API を使う必要がある。
両モードを統一インターフェースで隠蔽する。

モード切替:
- 環境変数 `DRIVE_MODE=api` を設定 → Google Drive API モード
- 未設定 or `local` → ローカルファイルシステム

Drive API モードの認証:
- 環境変数 `GOOGLE_SERVICE_ACCOUNT_JSON` に JSON 文字列を入れる（Render 推奨）
- または `GOOGLE_APPLICATION_CREDENTIALS` に JSON ファイルパスを入れる

Drive API モードでは「新着」「処理済み」「エラー」「過去　処理不要」フォルダの
Drive ID が必要。`DRIVE_FOLDER_NEW_ID` 等の環境変数で指定する。
"""
import os
import io
import json
import shutil
from typing import List, Optional, Tuple


class DriveFile:
    """ローカル/Drive API で共通のファイル抽象。"""
    def __init__(self, *, id: str, name: str, modified_time: str = "",
                 parent_id: Optional[str] = None, local_path: Optional[str] = None,
                 size: int = 0, mime_type: str = ""):
        self.id = id
        self.name = name
        self.modified_time = modified_time
        self.parent_id = parent_id
        self.local_path = local_path
        self.size = size
        self.mime_type = mime_type

    def __repr__(self):
        return f"DriveFile(name={self.name!r}, id={self.id!r}, size={self.size})"


# ─── ローカルモード ───────────────────────────────────────────

class LocalDriveClient:
    """ローカルファイルシステム経由で Drive ファイルを扱う。"""

    AUTO_BASE = r"G:\共有ドライブ\TWO\SCM\31_受発注確認\受注受信\_自動取込"

    FOLDERS = {
        "new":       os.path.join(AUTO_BASE, "新着"),
        "processed": os.path.join(AUTO_BASE, "処理済み"),
        "error":     os.path.join(AUTO_BASE, "エラー"),
        "archive":   os.path.join(AUTO_BASE, "過去　処理不要"),
    }

    def list_files(self, folder_key: str, extensions: Optional[Tuple[str, ...]] = None,
                   limit: Optional[int] = None) -> List[DriveFile]:
        path = self.FOLDERS[folder_key]
        if not os.path.exists(path):
            return []
        results = []
        for name in os.listdir(path):
            full = os.path.join(path, name)
            if not os.path.isfile(full):
                continue
            if extensions and not any(name.lower().endswith(ext.lower()) for ext in extensions):
                continue
            stat = os.stat(full)
            results.append(DriveFile(
                id=full,  # ローカルではフルパスを ID として扱う
                name=name,
                modified_time=str(stat.st_mtime),
                local_path=full,
                size=stat.st_size,
            ))
        results.sort(key=lambda f: f.modified_time)
        if limit:
            results = results[:limit]
        return results

    def download(self, file: DriveFile) -> bytes:
        with open(file.local_path, "rb") as f:
            return f.read()

    def read_text(self, file: DriveFile, encoding: str = "utf-8") -> str:
        with open(file.local_path, "r", encoding=encoding, errors="replace") as f:
            return f.read()

    def move(self, file: DriveFile, dest_folder_key: str) -> None:
        dest_dir = self.FOLDERS[dest_folder_key]
        os.makedirs(dest_dir, exist_ok=True)
        dest_path = os.path.join(dest_dir, file.name)
        shutil.move(file.local_path, dest_path)

    def find_sidecar(self, file: DriveFile,
                     all_files_in_folder: Optional[List[DriveFile]] = None) -> Optional[DriveFile]:
        """対応する .meta.json サイドカーを探す。"""
        base = os.path.splitext(file.name)[0]
        # 同じフォルダから検索
        folder = os.path.dirname(file.local_path)
        for candidate in (f"{base}.meta.json", f"{base}.json"):
            p = os.path.join(folder, candidate)
            if os.path.exists(p):
                return DriveFile(id=p, name=candidate, local_path=p)
        return None


# ─── Drive API モード ──────────────────────────────────────────

class GoogleDriveApiClient:
    """Google Drive API で共有ドライブを操作。Service Account 認証。"""

    SCOPES = ["https://www.googleapis.com/auth/drive"]

    def __init__(self):
        self._service = None
        # フォルダIDは環境変数 or デフォルト
        # GAS 設定と整合: 新着 = 107ZIFcg-u-eUq-gUlwkLqs7jf2oEbEaZ
        self.folder_ids = {
            "new":       os.environ.get("DRIVE_FOLDER_NEW_ID",
                                        "107ZIFcg-u-eUq-gUlwkLqs7jf2oEbEaZ"),
            "processed": os.environ.get("DRIVE_FOLDER_PROCESSED_ID", ""),
            "error":     os.environ.get("DRIVE_FOLDER_ERROR_ID", ""),
            "archive":   os.environ.get("DRIVE_FOLDER_ARCHIVE_ID", ""),
        }

    @property
    def service(self):
        if self._service is None:
            self._service = self._build_service()
        return self._service

    def _build_service(self):
        from google.oauth2 import service_account
        from googleapiclient.discovery import build

        sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
        sa_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()

        if sa_json:
            info = json.loads(sa_json)
            creds = service_account.Credentials.from_service_account_info(
                info, scopes=self.SCOPES)
        elif sa_path and os.path.exists(sa_path):
            creds = service_account.Credentials.from_service_account_file(
                sa_path, scopes=self.SCOPES)
        else:
            raise RuntimeError(
                "Service Account credentials not found. "
                "Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_APPLICATION_CREDENTIALS."
            )
        return build("drive", "v3", credentials=creds, cache_discovery=False)

    def list_files(self, folder_key: str, extensions: Optional[Tuple[str, ...]] = None,
                   limit: Optional[int] = None) -> List[DriveFile]:
        folder_id = self.folder_ids[folder_key]
        if not folder_id:
            raise RuntimeError(
                f"Folder ID for '{folder_key}' is not set. "
                f"Set DRIVE_FOLDER_{folder_key.upper()}_ID env var."
            )
        query = f"'{folder_id}' in parents and trashed=false"
        results = []
        page_token = None
        while True:
            response = self.service.files().list(
                q=query,
                pageSize=200,
                fields="nextPageToken, files(id, name, modifiedTime, parents, size, mimeType)",
                pageToken=page_token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                orderBy="modifiedTime",
            ).execute()
            for f in response.get("files", []):
                name = f.get("name", "")
                if extensions and not any(name.lower().endswith(ext.lower()) for ext in extensions):
                    continue
                results.append(DriveFile(
                    id=f["id"],
                    name=name,
                    modified_time=f.get("modifiedTime", ""),
                    parent_id=(f.get("parents") or [None])[0],
                    size=int(f.get("size") or 0),
                    mime_type=f.get("mimeType", ""),
                ))
            page_token = response.get("nextPageToken")
            if not page_token:
                break
        if limit:
            results = results[:limit]
        return results

    def download(self, file: DriveFile) -> bytes:
        from googleapiclient.http import MediaIoBaseDownload
        request = self.service.files().get_media(fileId=file.id, supportsAllDrives=True)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buf.getvalue()

    def read_text(self, file: DriveFile, encoding: str = "utf-8") -> str:
        return self.download(file).decode(encoding, errors="replace")

    def move(self, file: DriveFile, dest_folder_key: str) -> None:
        dest_id = self.folder_ids[dest_folder_key]
        if not dest_id:
            raise RuntimeError(f"Folder ID for '{dest_folder_key}' is not set.")
        prev_parents = file.parent_id or ""
        if not prev_parents:
            # 親IDを取り直す
            file_meta = self.service.files().get(
                fileId=file.id, fields="parents", supportsAllDrives=True).execute()
            prev_parents = ",".join(file_meta.get("parents") or [])
        self.service.files().update(
            fileId=file.id,
            addParents=dest_id,
            removeParents=prev_parents,
            supportsAllDrives=True,
            fields="id, parents",
        ).execute()

    def find_sidecar(self, file: DriveFile,
                     all_files_in_folder: Optional[List[DriveFile]] = None) -> Optional[DriveFile]:
        """対応する .meta.json サイドカーを探す。

        all_files_in_folder が指定されていればそこから検索（API回数を節約）、
        なければフォルダから直接検索する。
        """
        base = os.path.splitext(file.name)[0]
        candidates = (f"{base}.meta.json", f"{base}.json")
        if all_files_in_folder is not None:
            for f in all_files_in_folder:
                if f.name in candidates:
                    return f
            return None
        # フォルダから検索
        if not file.parent_id:
            return None
        query = (
            f"'{file.parent_id}' in parents and trashed=false and "
            f"(name='{candidates[0]}' or name='{candidates[1]}')"
        )
        response = self.service.files().list(
            q=query,
            pageSize=5,
            fields="files(id, name, parents)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        for f in response.get("files", []):
            return DriveFile(
                id=f["id"],
                name=f["name"],
                parent_id=(f.get("parents") or [None])[0],
            )
        return None


# ─── ファクトリ ─────────────────────────────────────────────

def get_drive_client():
    """環境変数 DRIVE_MODE で 'api' または 'local' を選択。

    デフォルトは 'local'。Render では DRIVE_MODE=api を設定する。
    """
    mode = (os.environ.get("DRIVE_MODE") or "local").lower()
    if mode == "api":
        return GoogleDriveApiClient()
    return LocalDriveClient()
