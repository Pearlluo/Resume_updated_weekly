import os
import re
import json
import math
import time
import tempfile
import requests
import pandas as pd
import fitz

from docx import Document
from base64 import b64encode
from datetime import datetime, timezone
from dotenv import load_dotenv
from azure.storage.blob import BlobServiceClient, ContentSettings

from Getonbroadingpeople import get_onboarding_people

load_dotenv()

TOKEN_URL = "https://auth.opms.com.au/api/authenticate/token"
TRAINING_SEARCH_URL = "https://api.opms.com.au/training/search"
TRAINING_DOCUMENT_URL = "https://api.opms.com.au/training/{id}/document"

CLIENT_ID = os.getenv("OPMS_CLIENT_ID")
CLIENT_SECRET = os.getenv("OPMS_CLIENT_SECRET")

AZURE_STORAGE_CONNECTION_STRING = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
AZURE_BLOB_CONTAINER = os.getenv("AZURE_BLOB_CONTAINER", "resumes")

RESUME_COMPETENCY_ID = 2437
STATUSES = ["pending", "completed", "archived"]

BATCH_SIZE = 50
PAGE_SIZE = 100
TEST_COUNT = None



def get_access_token(session):
    if not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError("Missing OPMS_CLIENT_ID or OPMS_CLIENT_SECRET")

    auth_str = f"{CLIENT_ID}:{CLIENT_SECRET}"
    b64_auth = b64encode(auth_str.encode()).decode()

    res = session.post(
        TOKEN_URL,
        headers={
            "Authorization": f"Basic {b64_auth}",
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        data={"grant_type": "client_credentials"},
        timeout=60,
    )

    print("🔐 Token status:", res.status_code)
    res.raise_for_status()

    token = res.json().get("access_token")
    if not token:
        raise RuntimeError("No access token returned")

    return token


def auth_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }


def chunk_list(values, chunk_size):
    for i in range(0, len(values), chunk_size):
        yield values[i:i + chunk_size]


def normalize_employee_id(value):
    if pd.isna(value):
        return None

    try:
        return int(value)
    except Exception:
        return str(value).strip()


def parse_datetime(value):
    if not value:
        return None

    try:
        return pd.to_datetime(value, utc=True)
    except Exception:
        return None


def get_file_extension(file_name, content_type):
    ext = os.path.splitext(file_name or "")[1].lower()

    if ext:
        return ext

    content_type = (content_type or "").lower()

    if "pdf" in content_type:
        return ".pdf"
    if "word" in content_type or "officedocument" in content_type:
        return ".docx"
    if "jpeg" in content_type or "jpg" in content_type:
        return ".jpg"
    if "png" in content_type:
        return ".png"

    return ".bin"


def get_content_type(ext):
    ext = (ext or "").lower()

    if ext == ".pdf":
        return "application/pdf"
    if ext == ".docx":
        return "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if ext == ".txt":
        return "text/plain; charset=utf-8"
    if ext == ".json":
        return "application/json"

    return "application/octet-stream"


def get_container_client():
    if not AZURE_STORAGE_CONNECTION_STRING:
        raise RuntimeError("Missing AZURE_STORAGE_CONNECTION_STRING")

    blob_service = BlobServiceClient.from_connection_string(
        AZURE_STORAGE_CONNECTION_STRING
    )

    container_client = blob_service.get_container_client(AZURE_BLOB_CONTAINER)

    try:
        container_client.create_container()
    except Exception:
        pass

    return container_client


def read_json_blob(container_client, blob_name):
    blob_client = container_client.get_blob_client(blob_name)

    if not blob_client.exists():
        return None

    try:
        data = blob_client.download_blob().readall()
        return json.loads(data.decode("utf-8"))
    except Exception as e:
        print(f"⚠️ Failed to read json blob {blob_name}: {e}")
        return None


def upload_bytes(container_client, blob_name, data, content_type):
    container_client.get_blob_client(blob_name).upload_blob(
        data,
        overwrite=True,
        content_settings=ContentSettings(content_type=content_type),
    )

    print("✅ Uploaded:", blob_name)
    return blob_name


def upload_text(container_client, blob_name, text):
    return upload_bytes(
        container_client=container_client,
        blob_name=blob_name,
        data=(text or "").encode("utf-8"),
        content_type="text/plain; charset=utf-8",
    )


def upload_json(container_client, blob_name, data):
    return upload_bytes(
        container_client=container_client,
        blob_name=blob_name,
        data=json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8"),
        content_type="application/json",
    )


def delete_old_employee_blobs(
    container_client,
    employee_id,
    keep_original_blob="",
    keep_text_blob="",
):
    prefixes = [
        f"original/{employee_id}/",
        f"text/{employee_id}/",
        f"classification/{employee_id}/",
    ]

    keep_set = {keep_original_blob, keep_text_blob}

    for prefix in prefixes:
        for blob in container_client.list_blobs(name_starts_with=prefix):
            if blob.name in keep_set:
                continue

            container_client.delete_blob(blob.name)
            print("🗑️ Deleted old blob:", blob.name)


def clean_resume_text(text):
    text = text or ""

    text = text.replace("\x00", " ")
    text = re.sub(r"\r\n|\r", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    patterns = [
        r"\n\s*references\s*\n",
        r"\n\s*referees\s*\n",
        r"\n\s*r\s*e\s*f\s*e\s*r\s*e\s*n\s*c\s*e\s*s\s*\n",
        r"\n\s*r\s*e\s*f\s*e\s*r\s*e\s*e\s*s\s*\n",
    ]

    lower_text = text.lower()
    cut_positions = []

    for pattern in patterns:
        match = re.search(pattern, lower_text, flags=re.IGNORECASE)
        if match:
            cut_positions.append(match.start())

    if cut_positions:
        text = text[:min(cut_positions)]

    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_text_from_pdf_bytes(file_bytes):
    text_parts = []

    try:
        with fitz.open(stream=file_bytes, filetype="pdf") as doc:
            for page in doc:
                text_parts.append(page.get_text("text"))
    except Exception as e:
        print("⚠️ PDF text extraction failed:", e)
        return ""

    return clean_resume_text("\n".join(text_parts))


def extract_text_from_docx_bytes(file_bytes):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    try:
        doc = Document(tmp_path)
        text = "\n".join(p.text for p in doc.paragraphs)
        return clean_resume_text(text)
    except Exception as e:
        print("⚠️ DOCX text extraction failed:", e)
        return ""
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


def extract_resume_text(file_bytes, file_ext):
    file_ext = (file_ext or "").lower()

    if file_ext == ".pdf":
        return extract_text_from_pdf_bytes(file_bytes)

    if file_ext == ".docx":
        return extract_text_from_docx_bytes(file_bytes)

    return ""


def search_resume_records(session, token, employee_ids, status):
    all_rows = []
    page = 1
    total_count = None

    while True:
        params = {
            "status": status,
            "employee_ids": ",".join(str(x) for x in employee_ids),
            "competency_ids": str(RESUME_COMPETENCY_ID),
            "page_size": PAGE_SIZE,
            "page": page,
        }

        try:
            res = session.get(
                TRAINING_SEARCH_URL,
                headers=auth_headers(token),
                params=params,
                timeout=120,
            )
        except Exception as e:
            print("❌ Search request error:", e)
            return None

        print("\n=================================================")
        print("📘 GET /training/search")
        print("status:", status)
        print("employee count:", len(employee_ids))
        print("page:", page)
        print("status code:", res.status_code)
        print("preview:", res.text[:300])

        if res.status_code == 500:
            return None

        if res.status_code >= 400:
            print("❌ API failed:", res.text[:1000])
            return []

        try:
            result = res.json()
        except Exception:
            print("❌ JSON parse failed")
            return []

        rows = result.get("data", [])
        total_count = result.get("count", total_count)

        if not rows:
            break

        all_rows.extend(rows)

        if len(rows) < PAGE_SIZE:
            break

        if total_count is not None:
            total_pages = math.ceil(total_count / PAGE_SIZE)
            if page >= total_pages:
                break

        page += 1
        time.sleep(0.3)

    return all_rows


def download_resume_bytes(session, token, document_id):
    url = TRAINING_DOCUMENT_URL.format(id=document_id)

    res = session.get(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "*/*",
        },
        params={"view_as": "USER"},
        timeout=180,
        allow_redirects=True,
    )

    print("\n📄 Download resume")
    print("document_id:", document_id)
    print("status:", res.status_code)
    print("content-type:", res.headers.get("Content-Type", ""))

    if res.status_code >= 400:
        print("❌ Download failed:", res.text[:500])
        return None, None

    return res.content, res.headers.get("Content-Type", "")


def get_employee_name(row):
    full_name = str(row.get("full_name", "") or "").strip()

    if full_name:
        return full_name

    first_name = str(row.get("first_name", "") or "").strip()
    last_name = str(row.get("last_name", "") or "").strip()

    return f"{first_name} {last_name}".strip()


def build_people_map(people_df):
    people_map = {}

    for _, row in people_df.iterrows():
        employee_id = normalize_employee_id(row.get("employee_id"))

        if employee_id is None:
            continue

        people_map[employee_id] = {
            "employee_id": employee_id,
            "employee_name": get_employee_name(row),
        }

    return people_map


def get_item_sort_key(item):
    date_fields = [
        item.get("last_modified_date"),
        item.get("created_date"),
        item.get("issue_date"),
    ]

    parsed_dates = [
        parse_datetime(value)
        for value in date_fields
        if value
    ]

    parsed_dates = [x for x in parsed_dates if x is not None]

    if parsed_dates:
        latest_date = max(parsed_dates)
        return latest_date.timestamp()

    try:
        return int(item.get("id", 0))
    except Exception:
        return 0


def choose_latest_resume_by_employee(all_resume_rows):
    latest_map = {}

    for item in all_resume_rows:
        worker = item.get("worker", {}) or {}
        employee_id = normalize_employee_id(worker.get("id"))

        if employee_id is None:
            continue

        current_score = get_item_sort_key(item)

        if employee_id not in latest_map:
            latest_map[employee_id] = item
            continue

        existing_score = get_item_sort_key(latest_map[employee_id])

        if current_score > existing_score:
            latest_map[employee_id] = item

    return latest_map


def process_resume_update(
    session,
    token,
    container_client,
    employee_id,
    employee_name,
    item,
):
    document_id = item.get("id")

    if not document_id:
        return {
            "employee_id": employee_id,
            "employee_name": employee_name,
            "action": "skip",
            "reason": "missing_document_id",
        }

    employee_id = str(employee_id)
    document_id = str(document_id)
    employee_document_key = f"{employee_id}_{document_id}"

    file_name = (
        item.get("file_name")
        or item.get("original_file_name")
        or item.get("originalFileName")
        or ""
    )

    latest_index_blob = f"index/latest/{employee_id}.json"
    history_index_blob = f"index/{employee_document_key}.json"

    existing = read_json_blob(container_client, latest_index_blob)

    existing_document_id = ""
    existing_original_blob = ""
    existing_text_blob = ""

    if existing:
        existing_document_id = str(existing.get("document_id", ""))
        existing_original_blob = existing.get("original_blob", "")
        existing_text_blob = existing.get("text_blob", "")

    if existing_document_id == document_id:
        print(f"⏭️ Skip employee {employee_id}, same document_id: {document_id}")

        return {
            "employee_id": employee_id,
            "employee_name": employee_name,
            "document_id": document_id,
            "employee_document_key": employee_document_key,
            "file_name": file_name,
            "action": "skip",
            "reason": "same_employee_id_and_document_id",
            "existing_document_id": existing_document_id,
            "original_blob": existing_original_blob,
            "text_blob": existing_text_blob,
            "latest_index_blob": latest_index_blob,
            "history_index_blob": existing.get("history_index_blob", "") if existing else "",
            "text_length": existing.get("text_length", 0) if existing else 0,
        }

    action = "new" if not existing_document_id else "replace"

    print("\n=================================================")
    print("🚀 Processing resume update")
    print("employee_id:", employee_id)
    print("employee_name:", employee_name)
    print("existing_document_id:", existing_document_id)
    print("new_document_id:", document_id)
    print("action:", action)

    file_bytes, content_type = download_resume_bytes(
        session=session,
        token=token,
        document_id=document_id,
    )

    if not file_bytes:
        return {
            "employee_id": employee_id,
            "employee_name": employee_name,
            "document_id": document_id,
            "employee_document_key": employee_document_key,
            "file_name": file_name,
            "action": "failed",
            "reason": "download_failed_keep_existing",
            "existing_document_id": existing_document_id,
            "original_blob": existing_original_blob,
            "text_blob": existing_text_blob,
            "latest_index_blob": latest_index_blob,
        }

    file_ext = get_file_extension(file_name, content_type)

    original_blob = f"original/{employee_id}/resume_{employee_document_key}{file_ext}"
    text_blob = f"text/{employee_id}/resume_{employee_document_key}.txt"

    text = extract_resume_text(file_bytes, file_ext)

    if not text:
        print("⚠️ No text extracted. Still uploading empty txt for tracking.")

    upload_bytes(
        container_client=container_client,
        blob_name=original_blob,
        data=file_bytes,
        content_type=get_content_type(file_ext),
    )

    upload_text(
        container_client=container_client,
        blob_name=text_blob,
        text=text or "",
    )

    if action == "replace":
        delete_old_employee_blobs(
            container_client=container_client,
            employee_id=employee_id,
            keep_original_blob=original_blob,
            keep_text_blob=text_blob,
        )

    index_data = {
        "employee_id": employee_id,
        "document_id": document_id,
        "employee_document_key": employee_document_key,

        "employee_name": employee_name,
        "file_name": file_name,

        "original_blob": original_blob,
        "text_blob": text_blob,
        "text_length": len(text or ""),

        "previous_document_id": existing_document_id,
        "previous_original_blob": existing_original_blob,
        "previous_text_blob": existing_text_blob,

        "last_modified_date": item.get("last_modified_date", ""),
        "created_date": item.get("created_date", ""),
        "issue_date": item.get("issue_date", ""),
        "expiry_date": item.get("expiry_date", ""),
        "status": item.get("status", ""),

        "history_index_blob": history_index_blob,
        "latest_index_blob": latest_index_blob,
        "synced_at_utc": datetime.now(timezone.utc).isoformat(),
    }

    upload_json(
        container_client=container_client,
        blob_name=history_index_blob,
        data=index_data,
    )

    upload_json(
        container_client=container_client,
        blob_name=latest_index_blob,
        data=index_data,
    )

    return {
        "employee_id": employee_id,
        "employee_name": employee_name,
        "document_id": document_id,
        "employee_document_key": employee_document_key,
        "file_name": file_name,
        "action": action,
        "reason": "new_or_changed_employee_id_document_id",
        "existing_document_id": existing_document_id,
        "new_document_id": document_id,
        "original_blob": original_blob,
        "text_blob": text_blob,
        "text_length": len(text or ""),
        "history_index_blob": history_index_blob,
        "latest_index_blob": latest_index_blob,
        "last_modified_date": item.get("last_modified_date", ""),
    }


def main():
    print("📥 Loading onboarding people...")

    people_df = get_onboarding_people(
        as_dataframe=True,
        include_raw=False,
    )

    if TEST_COUNT:
        people_df = people_df.head(TEST_COUNT)

    people_map = build_people_map(people_df)
    employee_ids = list(people_map.keys())

    print("✅ Total onboarding employees:", len(employee_ids))

    container_client = get_container_client()

    all_resume_rows = []


    with requests.Session() as session:
        token = get_access_token(session)

        for status in STATUSES:
            for batch_no, batch_employee_ids in enumerate(
                chunk_list(employee_ids, BATCH_SIZE),
                start=1,
            ):
                print("\n=================================================")
                print("🚀 Pulling resume records")
                print("batch:", batch_no)
                print("status:", status)
                print("employee count:", len(batch_employee_ids))

                rows = search_resume_records(
                    session=session,
                    token=token,
                    employee_ids=batch_employee_ids,
                    status=status,
                )

                if rows is None:
                    print("⚠️ Batch failed. Retrying one by one...")

                    rows = []

                    for single_employee_id in batch_employee_ids:
                        single_rows = search_resume_records(
                            session=session,
                            token=token,
                            employee_ids=[single_employee_id],
                            status=status,
                        )

                        if single_rows:
                            rows.extend(single_rows)

                if rows:
                    all_resume_rows.extend(rows)

                time.sleep(0.5)

        print("\n=================================================")
        print("✅ OPMS resume rows pulled:", len(all_resume_rows))

        latest_resume_map = choose_latest_resume_by_employee(all_resume_rows)

        print("✅ Latest resume employees:", len(latest_resume_map))

        for employee_id, item in latest_resume_map.items():
            worker = item.get("worker", {}) or {}

            employee_name = (
                people_map.get(employee_id, {}).get("employee_name")
                or f"{worker.get('first_name', '')} {worker.get('last_name', '')}".strip()
            )

            result = process_resume_update(
                session=session,
                token=token,
                container_client=container_client,
                employee_id=employee_id,
                employee_name=employee_name,
                item=item,
            )

            print("Result:", result)

    print("\n=================================================")
    print("✅ DONE")
    print("Actions:")
    print("- skip = same employee_id + document_id")
    print("- new = employee not found in latest index")
    print("- replace = employee found but document_id changed")
    print("- failed = download failed, old file kept")
    print("\nBlob structure:")
    print("original/{employee_id}/resume_{employee_id}_{document_id}.pdf")
    print("text/{employee_id}/resume_{employee_id}_{document_id}.txt")
    print("index/{employee_id}_{document_id}.json")
    print("index/latest/{employee_id}.json")


if __name__ == "__main__":
    main()