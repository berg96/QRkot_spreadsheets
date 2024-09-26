from datetime import datetime

from aiogoogle import Aiogoogle, ValidationError

from app.core.config import settings

FORMAT = '%Y/%m/%d %H:%M:%S'
SPREADSHEET_TITLE = 'Отчёт на {}'
SHEETS_TITLE = 'Лист1'
SHEETS_ROW = 100
SHEETS_COLUMN = 10
SPREADSHEET_BODY = dict(
    properties=dict(
        title=f'Отчет от {FORMAT}',
        locale='ru_RU',
    ),
    sheets=[
        dict(
            properties=dict(
                sheetType='GRID',
                sheetId=0,
                title='Лист1',
                gridProperties=dict(
                    rowCount=100,
                    columnCount=11,
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
SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/{}'
INVALID_SIZE = 'Объем данных не соответствует размеру пустой таблицы'


async def spreadsheets_create(wrapper_services: Aiogoogle) -> tuple[str, str]:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheet_body = SPREADSHEET_BODY.copy()
    spreadsheet_body['properties']['title'] = SPREADSHEET_TITLE.format(
        now_date_time
    )
    spreadsheet_body['sheets'][0]['properties']['title'] = SHEETS_TITLE
    spreadsheet_body['sheets'][0][
        'properties'
    ]['gridProperties']['rowCount'] = SHEETS_ROW
    spreadsheet_body['sheets'][0][
        'properties'
    ]['gridProperties']['columnCount'] = SHEETS_COLUMN
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    spreadsheet_id = response['spreadsheetId']
    return spreadsheet_id, SPREADSHEET_URL.format(spreadsheet_id)


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
    table_header = TABLE_HEADER.copy()
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
        raise ValidationError(INVALID_SIZE)
    update_body = {
        'majorDimension': 'ROWS',
        'values': table_values
    }
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheet_id,
            range=f'R1C1:R{SHEETS_ROW}C{SHEETS_COLUMN}',
            valueInputOption='USER_ENTERED',
            json=update_body
        )
    )
