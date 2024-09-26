from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings

FORMAT = "%Y/%m/%d %H:%M:%S"
SPREADSHEET_TITLE = 'Отчёт на {}'
SHEETS_TITLE = 'Лист1'
SHEETS_ROW = 100
SHEETS_COLUMN = 10
TABLE_HEADER = [
    ['Отчёт от'],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора', 'Описание']
]


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheet_body = {
        'properties': {
            'title': SPREADSHEET_TITLE.format(now_date_time),
            'locale': 'ru_RU'
        },
        'sheets': [
            {
                'properties': {
                    'sheetType': 'GRID',
                    'sheetId': 0,
                    'title': SHEETS_TITLE,
                    'gridProperties': {
                        'rowCount': SHEETS_ROW,
                        'columnCount': SHEETS_COLUMN
                    }
                }
            }
        ]
    }
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    return response['spreadsheetId']


async def set_user_permissions(
        spreadsheetid: str,
        wrapper_services: Aiogoogle
) -> None:
    permissions_body = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': settings.email
    }
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheetid,
            json=permissions_body,
            fields="id"
        )
    )


async def spreadsheets_update_value(
        spreadsheetid: str,
        projects: list,
        wrapper_services: Aiogoogle
) -> None:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    TABLE_HEADER[0].append(now_date_time)
    table_values = TABLE_HEADER
    for project in projects:
        table_values.append(
            [
                project['name'],
                (
                    f'=INT({project["collection_time"]}/86400) & " days, " & '
                    f'TEXT({project["collection_time"]}/86400 - '
                    f'INT({project["collection_time"]}/86400); "hh:mm:ss")'
                ),
                project['description']
            ]
        )

    update_body = {
        'majorDimension': 'ROWS',
        'values': table_values
    }
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheetid,
            range='A1:E30',
            valueInputOption='USER_ENTERED',
            json=update_body
        )
    )
