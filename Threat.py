import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


CREDENTIALS_FILE = 'LSComponents.json'
spreadsheetId = '1QRRHkeS64Ln-yZgEbBaVtxYFi3rTLnOIr4pgkEX1iPE'
doubled = 'underline'
half = 'italic'
threat_begin = 3
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
    for i in range(threat_begin, threat_number+threat_begin):
        answer = service.spreadsheets().get(spreadsheetId=spreadsheetId, includeGridData=True, ranges='d%i:%s%i' % (i, last, i)).execute()
        threats_data = answer['sheets'][0]['data'][0]['rowData'][0]['values']
        try:
            threat = threats_data[0]['effectiveValue']['stringValue']
            threats[threat] = []
            for j in range(5, len(threats_data)):
                try:
                    color = threats_data[j]['userEnteredFormat']['backgroundColor']
                    risk = 0
                    if color == {'red': 1, 'green': 1}:
                        risk = 0.5
                    if color == {'red': 1, 'green': 0.6} or color == {'red': 1}:
                        risk = 1
                    threats[threat].append(risk)
                except KeyError:
                    pass
        except KeyError:
            pass
    risks = {}
    for (i, (project, weight)) in enumerate(projects.items()):
        risks[project] = {}
        risks[project]['weight'] = weight
        for threat in threats.keys():
            temp = threats[threat]
            risks[project][threat] = threats[threat][i]
    print(risks)
    rownumber = threat_begin
    for threat in threats.keys():
        total = 0
        for project in risks.keys():
            total += (risks[project]['weight'])*(risks[project][threat])
        percent = int(100*total/count)
        request_body = {"valueInputOption": "RAW",
                        "data": [{"range": "f%i:f%i" % (rownumber, rownumber), "values": [[str(percent)+"%"]]}]}
        request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId, body=request_body)
        _ = request.execute()
        rownumber += 1



if __name__ == '__main__':
    main()