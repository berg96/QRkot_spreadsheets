import copy
from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings

FORMAT = '%Y/%m/%d %H:%M:%S'
SPREADSHEET_TITLE = 'Отчёт на {}'
SHEETS_ROW = 100
SHEETS_COLUMN = 10
SPREADSHEET_BODY = dict(
    properties=dict(
        title='Отчет от ',
        locale='ru_RU',
    ),
    sheets=[
        dict(
            properties=dict(
                sheetType='GRID',
                sheetId=0,
                title='Лист1',
                gridProperties=dict(
                    rowCount=SHEETS_ROW,
                    columnCount=SHEETS_COLUMN,
                )
            )
        )
    ]
)
TABLE_HEADER = [
    ['Отчёт от'],
    ['Топ проектов по скорости закрытия'],
    ['Название проекта', 'Время сбора', 'Описание']
]
COLLECTION_TIME_IN_SHEETS = (
    '=INT({collection_time}/86400) & " days, " & '
    'TEXT({collection_time}/86400-INT({collection_time}/86400); "hh:mm:ss")'
)
INVALID_SIZE_ROW = (
    f'Количество строк {{}} обновляемых данных '
    f'больше размера пустой таблицы {SHEETS_ROW}'
)
INVALID_SIZE_COLUMN = (
    f'Количество колонок {{}} обновляемых данных '
    f'больше размера пустой таблицы {SHEETS_COLUMN}'
)


async def spreadsheets_create(wrapper_services: Aiogoogle) -> tuple[str, str]:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheet_body = copy.deepcopy(SPREADSHEET_BODY)
    spreadsheet_body['properties']['title'] = SPREADSHEET_TITLE.format(
        now_date_time
    )
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    return response['spreadsheetId'], response['spreadsheetUrl']


async def set_user_permissions(
        spreadsheet_id: str,
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
            fileId=spreadsheet_id,
            json=permissions_body,
            fields='id'
        )
    )


async def spreadsheets_update_value(
        spreadsheet_id: str,
        projects: list,
        wrapper_services: Aiogoogle
) -> None:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    table_header = copy.deepcopy(TABLE_HEADER)
    table_header[0].append(now_date_time)
    table_values = [
        *table_header,
        *[
            list(map(str, (
                project['name'],
                COLLECTION_TIME_IN_SHEETS.format(
                    collection_time=project['collection_time']
                ),
                project['description']
            )))
            for project in projects
        ]
    ]
    if len(table_values) > SHEETS_ROW:
        raise ValueError(
            INVALID_SIZE_ROW.format(len(table_values))
        )
    max_column = max(map(len, table_values))
    if max_column > SHEETS_COLUMN:
        raise ValueError(
            INVALID_SIZE_COLUMN.format(max_column)
        )
    update_body = {
        'majorDimension': 'ROWS',
        'values': table_values
    }
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheet_id,
            range=f'R1C1:R{len(table_values)}C{max_column}',
            valueInputOption='USER_ENTERED',
            json=update_body
        )
    )
