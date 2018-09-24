import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


CREDENTIALS_FILE = 'LSComponents.json'
spreadsheetId = '1QRRHkeS64Ln-yZgEbBaVtxYFi3rTLnOIr4pgkEX1iPE'
doubled = 'underline'
half = 'italic'
threat_begin = 2
threat_bonuses = 7
last = 'z'

def main():
    credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                      'https://www.googleapis.com/auth/drive'])
    httpAuth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

    answer = service.spreadsheets().get(spreadsheetId=spreadsheetId, includeGridData=True, ranges='i1:%s1' % last).execute()
    headers = answer['sheets'][0]['data'][0]['rowData'][0]['values']
    headers = list(filter(bool, headers))
    projects = {}
    for header in headers:
        weight = 1
        if header['effectiveFormat']['textFormat'][doubled]:
            weight = 2
        elif header['effectiveFormat']['textFormat'][half]:
            weight = 0.5
        projects[header['effectiveValue']['stringValue']] = weight
    count = sum(projects.values())
    threat_number = answer['sheets'][0]['properties']['gridProperties']['rowCount'] - threat_begin - threat_bonuses
    threats = {}
    for i in range(3, threat_number+3):
        answer = service.spreadsheets().get(spreadsheetId=spreadsheetId, includeGridData=True, ranges='d%i:%i%i' % (i, last, i).execute()
        threats_data = answer['sheets'][0]['data'][0]['rowData'][0]['values']
        threat = threats_data[0]['effectiveValue']['stringValue']
        threats[threat] =


    print(count)



if __name__ == '__main__':
    main()