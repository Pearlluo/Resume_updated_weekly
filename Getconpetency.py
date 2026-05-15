import os
import requests
from base64 import b64encode
from pprint import pprint
from dotenv import load_dotenv
import pandas as pd

# ===============================
# LOAD ENV
# ===============================
load_dotenv()

# ===============================
# CONFIG
# ===============================
TOKEN_URL = "https://auth.opms.com.au/api/authenticate/token"
COMPETENCIES_URL = "https://api.opms.com.au/competencies"

CLIENT_ID = os.getenv("OPMS_CLIENT_ID")
CLIENT_SECRET = os.getenv("OPMS_CLIENT_SECRET")


# ===============================
# VALIDATION
# ===============================
def validate_config():
    if not CLIENT_ID:
        raise ValueError("Missing environment variable: OPMS_CLIENT_ID")
    if not CLIENT_SECRET:
        raise ValueError("Missing environment variable: OPMS_CLIENT_SECRET")


# ===============================
# TOKEN
# ===============================
def get_access_token(session):
    validate_config()

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

    print("🔐 token status:", res.status_code)
    print("Token response:", res.text[:300])

    res.raise_for_status()

    data = res.json()
    token = data.get("access_token")

    if not token:
        raise RuntimeError(f"Token response missing access_token: {data}")

    return token


# ===============================
# NORMALISE COMPETENCY ITEM
# ===============================
def get_competency_object(item):
    """
    OPMS returns:
    {
        "competency": {
            "id": 2533,
            "name": "123"
        }
    }

    But some APIs may return:
    {
        "id": 2533,
        "name": "123"
    }

    This function supports both.
    """
    if isinstance(item, dict) and isinstance(item.get("competency"), dict):
        return item.get("competency")

    return item


# ===============================
# GET COMPETENCIES
# ===============================
def get_competencies(session, token, view_as):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }

    params = {
        "view_as": view_as
    }

    res = session.get(
        COMPETENCIES_URL,
        headers=headers,
        params=params,
        timeout=120
    )

    print("\n===============================")
    print(f"📘 Testing /competencies as {view_as}")
    print("===============================")
    print("Status Code:", res.status_code)
    print("Request URL:", res.url)
    print("Response Text:")
    print(res.text[:2000])

    if res.status_code == 403:
        print(f"\n❌ 403 Forbidden using view_as={view_as}")
        return None

    if res.status_code == 404:
        print(f"\n❌ 404 Not Found using view_as={view_as}")
        return None

    res.raise_for_status()
    return res.json()


# ===============================
# SAVE TO EXCEL
# ===============================
def save_competencies_to_excel(competencies, output_file):
    rows = []

    for item in competencies:
        comp = get_competency_object(item)

        if not isinstance(comp, dict):
            continue

        group = comp.get("group") or {}
        form = comp.get("form") or {}
        external_form = comp.get("external_form") or {}

        rows.append({
            "id": comp.get("id"),
            "name": comp.get("name"),

            "group_id": group.get("id"),
            "group_name": group.get("name"),

            "classified": comp.get("classified"),

            "reference1": comp.get("reference1"),
            "reference2": comp.get("reference2"),
            "reference3": comp.get("reference3"),
            "reference4": comp.get("reference4"),

            "validity_period": str(comp.get("validity_period")),
            "revalidation_reminder_period": str(comp.get("revalidation_reminder_period")),

            "form_id": form.get("id"),
            "form_type": form.get("type"),
            "form_can_complete": form.get("can_complete"),

            "external_form_id": external_form.get("id"),
            "external_form_type": external_form.get("type"),
        })

    df = pd.DataFrame(rows)

    output_file = os.path.abspath(output_file)
    df.to_excel(output_file, index=False)

    print(f"\n✅ Saved to Excel: {output_file}")
    print("Total competencies:", len(df))


# ===============================
# MAIN
# ===============================
def main():
    with requests.Session() as session:
        token = get_access_token(session)

        competencies = get_competencies(
            session=session,
            token=token,
            view_as="EMPLOYEE"
        )

        if competencies is None:
            competencies = get_competencies(
                session=session,
                token=token,
                view_as="USER"
            )

        if competencies is None:
            print("\n❌ Both EMPLOYEE and USER failed.")
            print("Token is valid, but this API client may not have permission for /competencies.")
            return

        print("\n✅ Competencies pulled successfully.")

        if isinstance(competencies, list):
            print("\nFirst 10 records:")

            for item in competencies[:10]:
                comp = get_competency_object(item)

                if not isinstance(comp, dict):
                    continue

                print(
                    comp.get("id"),
                    "-",
                    comp.get("name"),
                    "| group:",
                    (comp.get("group") or {}).get("name")
                )

            save_competencies_to_excel(
                competencies,
                "opms_competencies.xlsx"
            )

        else:
            print("\nResponse is not a list:")
            pprint(competencies)


if __name__ == "__main__":
    main()