import requests

def update_excel_cell(client_id, client_secret, filename, sheet, cell, value, storage_name=None, folder=None, value_type="string"):
    # 1. Get access token
    token_response = requests.post(
        "https://api.aspose.cloud/connect/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret
        }
    )
    token_data = token_response.json()
    access_token = token_data.get("access_token")
    if not access_token:
        raise RuntimeError(f"Failed to get access token: {token_data}")

    # 2. Build the correct API URL with storage parameter
    base_url = f"https://api.aspose.cloud/v3.0/cells/{filename}/worksheets/{sheet}/cells/{cell}"
    qparams = []
    if folder:
        qparams.append(f"folder={folder}")
    if storage_name:
        qparams.append(f"storageName={storage_name}")
    if qparams:
        base_url += "?" + "&".join(qparams)

    # 3. Set cell value by POST request
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    payload = {
        "value": value,
        "type": value_type
    }
    response = requests.post(base_url, headers=headers, json=payload)
    if response.status_code != 200:
        raise RuntimeError(f"Failed to update cell: HTTP {response.status_code}, {response.text}")
    print("Cell updated successfully!")

# <<<<<<<<<<<<<<<<<<< USE THESE EXACT VARIABLES >>>>>>>>>>>>>>>>>>>>
client_id = "d1b2557c-dee7-4b81-88e6-10f3b5c4425a"
client_secret = "7b6175ede5049835309f47259cc1733d"
filename = "Gudipati_Phani_Babu_Timesheet_Week_Ending_08152025.xlsx"
sheet = "Timesheet"
cell = "G2"
value = "Phani Babu"

storage_name = "Timesheet Backend Export"   # <-- this is your Aspose Storage name (not a folder)
folder = None                               # <-- your file is in the root, so leave this None or ""

if __name__ == "__main__":
    update_excel_cell(client_id, client_secret, filename, sheet, cell, value, storage_name, folder)
