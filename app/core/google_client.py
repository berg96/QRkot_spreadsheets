import json

from aiogoogle import Aiogoogle
from aiogoogle.auth.creds import ServiceAccountCreds

from app.core.config import settings

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

with open(settings.credentials_file, 'r') as file:
    service_account_info = json.load(file)
cred = ServiceAccountCreds(scopes=SCOPES, **service_account_info)


async def get_service():
    async with Aiogoogle(service_account_creds=cred) as aiogoogle:
        yield aiogoogle
