"""Download the workbook from Yazeed's OneDrive via Microsoft Graph API."""

import os
import msal
import requests
from dotenv import load_dotenv
from config import EXCEL_FILE_ID, EXCEL_USER, ENV_PATH


def download_workbook(output_path='workbook.xlsx'):
    """Download latest Excel from Yazeed's OneDrive.

    Returns: (file_path, last_modified, modified_by)
    """
    load_dotenv(ENV_PATH)

    app = msal.ConfidentialClientApplication(
        os.getenv('AZURE_CLIENT_ID'),
        authority=f"https://login.microsoftonline.com/{os.getenv('AZURE_TENANT_ID')}",
        client_credential=os.getenv('AZURE_CLIENT_SECRET'),
    )
    token = app.acquire_token_for_client(
        scopes=['https://graph.microsoft.com/.default']
    )['access_token']
    headers = {'Authorization': f'Bearer {token}'}

    # Get file metadata
    resp = requests.get(
        f'https://graph.microsoft.com/v1.0/users/{EXCEL_USER}/drive/items/{EXCEL_FILE_ID}',
        headers=headers,
    )
    resp.raise_for_status()
    item = resp.json()

    last_modified = item.get('lastModifiedDateTime', '')
    modified_by = item.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'unknown')
    size = item.get('size', 0)

    # Download
    download_url = item.get('@microsoft.graph.downloadUrl')
    content = requests.get(download_url)
    content.raise_for_status()

    with open(output_path, 'wb') as f:
        f.write(content.content)

    print(f'Downloaded: {item["name"]} ({size:,} bytes)')
    print(f'  Last modified: {last_modified} by {modified_by}')
    print(f'  Saved to: {output_path}')

    return output_path, last_modified, modified_by
