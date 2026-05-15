import os
import requests
from base64 import b64encode
from dotenv import load_dotenv
import pandas as pd

load_dotenv()

TOKEN_URL = "https://auth.opms.com.au/api/authenticate/token"
SITE_EMPLOYEES_URL = "https://api.opms.com.au/sites/employees"

CLIENT_ID = os.getenv("OPMS_CLIENT_ID")
CLIENT_SECRET = os.getenv("OPMS_CLIENT_SECRET")

ONBOARDING_SITE_ID = 11
PAGE_SIZE = 100


def get_access_token(session):
    auth_str = f"{CLIENT_ID}:{CLIENT_SECRET}"
    b64_auth = b64encode(auth_str.encode()).decode()

    headers = {
        "Authorization": f"Basic {b64_auth}",
        "Content-Type": "application/x-www-form-urlencoded",
        "Accept": "application/json",
    }

    res = session.post(
        TOKEN_URL,
        headers=headers,
        data={"grant_type": "client_credentials"},
        timeout=60
    )

    res.raise_for_status()
    return res.json()["access_token"]


def get_page(session, token, site_id=ONBOARDING_SITE_ID, page_size=PAGE_SIZE, after=None):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    params = {
        "site_ids": str(site_id),
        "page_size": page_size,
    }

    if after:
        params["after"] = after

    res = session.get(
        SITE_EMPLOYEES_URL,
        headers=headers,
        params=params,
        timeout=120
    )

    if res.status_code >= 400:
        print("OPMS response:", res.text[:2000])
        res.raise_for_status()

    return res.json()


def extract_rows(data):
    if isinstance(data, list):
        return data

    if isinstance(data, dict):
        for key in ["data", "employees", "results", "items"]:
            if isinstance(data.get(key), list):
                return data[key]

    return []


def extract_next_after(data):
    if not isinstance(data, dict):
        return None

    for key in ["after", "next_after", "next_cursor", "cursor"]:
        value = data.get(key)
        if value:
            return value

    for parent_key in ["pagination", "meta"]:
        parent = data.get(parent_key)
        if isinstance(parent, dict):
            for key in ["after", "next_after", "next_cursor", "cursor"]:
                value = parent.get(key)
                if value:
                    return value

    return None


def extract_employee(row):
    employee = row.get("employee") or row.get("Employee") or row

    first_name = employee.get("first_name") or employee.get("FirstName") or ""
    last_name = employee.get("last_name") or employee.get("LastName") or ""

    return {
        "employee_id": employee.get("id") or employee.get("employee_id"),
        "first_name": first_name,
        "last_name": last_name,
        "full_name": f"{first_name} {last_name}".strip(),
        "raw_row": str(row)[:1000],
    }


def get_onboarding_people(
    site_id=ONBOARDING_SITE_ID,
    page_size=PAGE_SIZE,
    as_dataframe=True,
    include_raw=True,
):
    """
    Pull all onboarding site employees from OPMS.

    Returns:
        pandas.DataFrame if as_dataframe=True
        list[dict] if as_dataframe=False
    """

    all_rows = []
    after = None

    with requests.Session() as session:
        token = get_access_token(session)

        while True:
            data = get_page(
                session=session,
                token=token,
                site_id=site_id,
                page_size=page_size,
                after=after
            )

            rows = extract_rows(data)

            if not rows:
                break

            all_rows.extend(rows)

            after = extract_next_after(data)

            if not after:
                break

    output_rows = [extract_employee(row) for row in all_rows]

    if not include_raw:
        for row in output_rows:
            row.pop("raw_row", None)

    df = pd.DataFrame(output_rows)

    if not df.empty and "employee_id" in df.columns:
        df = df.drop_duplicates(subset=["employee_id"])

    if as_dataframe:
        return df

    return df.to_dict("records")


def get_onboarding_people_lookup(site_id=ONBOARDING_SITE_ID):
    """
    Return dict lookup:
    {
        employee_id: {
            first_name,
            last_name,
            full_name
        }
    }
    """

    people = get_onboarding_people(
        site_id=site_id,
        as_dataframe=False,
        include_raw=False
    )

    return {
        person["employee_id"]: {
            "first_name": person.get("first_name"),
            "last_name": person.get("last_name"),
            "full_name": person.get("full_name"),
        }
        for person in people
        if person.get("employee_id") is not None
    }