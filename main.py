import httplib2 
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


def create_table():
    spreadsheet = service.spreadsheets().create(body = {
        'properties': {'title': 'Первый тестовый документ', 'locale': 'ru_RU'},
        'sheets': [{'properties': {'sheetType': 'GRID',
                                   'sheetId': 0,
                                   'title': 'Лист 1',
                                   'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
    }).execute()
    return spreadsheet['spreadsheetId'] # сохраняем идентификатор файла


def give_permissions():
    driveService = apiclient.discovery.build('drive', 'v3', http = httpAuth) # Выбираем работу с Google Drive и 3 версию API
    access = driveService.permissions().create(
        fileId = spreadsheetId,
        body = {'type': 'user', 'role': 'writer', 'emailAddress': 'test.sheets.api1@gmail.com'},  # Открываем доступ на редактирование
        fields = 'id'
    ).execute()


def set_data():
    values = {
        "majorDimension": "ROWS",
        "values": [
             ["Ячейка A1", '25', "=6*6"],  # Заполняем первую строку
             ["Ячейка A2", '25', "=6*6"]  # Заполняем вторую строку
         ]
    }
    range = "Лист 1!A1:C2"
    valueInputOption = "USER_ENTERED"
    insertDataOption = "OVERWRITE"

    results = service.spreadsheets().values().append(spreadsheetId=spreadsheetId, range=range, valueInputOption=valueInputOption, insertDataOption=insertDataOption, body=values).execute()


def update_data():
    updata_values = {
        "valueInputOption": "USER_ENTERED",
        "data": [
            {"range": "Лист 1!A1:C1",
             "majorDimension": "ROWS",
             "values": [
                 ["Новые данные", '45', "73"],  # Заполняем первую строку
             ]}
        ]
    }

    results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=updata_values).execute()


def clear_data():
    clear_values = {
          "ranges": "Лист 1!A2:C2"
    }
    results = service.spreadsheets().values().batchClear(spreadsheetId=spreadsheetId, body=clear_values).execute()


if __name__ == '__main__':
    CREDENTIALS_FILE = 'first-prokect-f1efb062b16e.json'
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])

    #авторизация
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

    #создание таблицы
    # spreadsheetId = create_table()
    # give_permissions()

    spreadsheetId = "1KoLw_9qx3rqrkSUBfQq76G9gGPKco-MHjZ0y8PuR0cQ"

    #заполнение данными
    set_data()

    #изменение данных
    update_data()
    #
    #удаление данных
    clear_data()

    print('https://docs.google.com/spreadsheets/d/' + spreadsheetId)