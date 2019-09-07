import pygsheets
import boto3
from email.mime.text import MIMEText

aws_access_key_id = '' # 需要向行政組求 key
aws_secret_access_key = '' # 需要向行政組求 key
sender = 'program+noreply@coscup.org' # 或填上今年寄信用的信箱
bcc = [] # 填議程組負責議程軌事務的人的信箱

gc = pygsheets.authorize()
book = gc.open_by_key('') # 填上 spreadsheet 的 key
worksheet_proposal = book.sheet1 # 第一個 sheet 是稿件相關資料
worksheet_track = book.sheet2 # 第二個 sheet 是議程軌相關資料

acceptFile = open('accept.txt', 'r', encoding='UTF-8') # 錄取通知信文件
accept = acceptFile.read()
acceptFile.close()

denyFile = open('deny.txt', 'r', encoding='UTF-8') # 落選通知信文件
deny = denyFile.read()
denyFile.close()

def getTopicData(topic):
    global worksheet_track
    row = 2
    while worksheet_track.cell('A' + str(row)).value is not '':
        if topic == worksheet_track.cell('A' + str(row)).value:
            return {
                'track_name': worksheet_track.cell('A' + str(row)).value,
                'track_date': worksheet_track.cell('B' + str(row)).value,
                'track_time': worksheet_track.cell('C' + str(row)).value,
                'cc': worksheet_track.cell('D' + str(row)).value.split('\n')
            }
    return False

def generate_accept_mail(row):
    global worksheet_proposal
    global accept

    subject = 'COSCUP 2019 社群議程入選通知 / COSCUP 2019 program notification for author' # 填寫主旨
    topic = worksheet_proposal.cell('B' + str(row)).value
    data = getTopicData(topic)
    if data == False:
        return False

    data['name'] = worksheet_proposal.cell('D' + str(row)).value
    data['talk_name'] = worksheet_proposal.cell('C' + str(row)).value
    data['talk_date'] = worksheet_proposal.cell('F' + str(row)).value
    data['talk_time'] = worksheet_proposal.cell('G' + str(row)).value
    receiver = worksheet_proposal.cell('E' + str(row)).value
    content = accept.format(**data)
    print(content)

    msg = MIMEText(content)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = receiver
    msg['Cc'] = ', '.join(data['cc'])
    msg['Bcc'] = ', '.join(bcc)

    return msg

def generate_deny_mail(row):
    global worksheet_proposal
    global deny

    subject = 'COSCUP 2019 社群議程未入選通知 / COSCUP 2019 program notification for author' # 填寫主旨
    topic = worksheet_proposal.cell('C' + str(row)).value
    data = getTopicData(topic)
    if data == False:
        return False

    data['name'] = worksheet_proposal.cell('D' + str(row)).value
    data['talk_name'] = worksheet_proposal.cell('C' + str(row)).value
    receiver = worksheet_proposal.cell('E' + str(row)).value
    content = deny.format(**data)
    print(content)

    msg = MIMEText(content)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = receiver
    msg['Cc'] = ', '.join(data['cc'])
    msg['Bcc'] = ', '.join(bcc)

    return msg

def main():
    global worksheet_proposal
    row = 2

    while worksheet_proposal.cell('A' + str(row)).value is not '':
        isAccept = worksheet_proposal.cell('A' + str(row)).value
        server = boto3.client('ses', aws_access_key_id=aws_access_key_id, aws_secret_access_key=aws_secret_access_key, region_name='us-east-1')
        if isAccept == 'Y':
            msg = generate_accept_mail(row)
            if msg == False:
                print('No this topic {}'.format(worksheet_proposal.cell('B' + str(row)).value))
                return
            server.send_raw_email(RawMessage={'Data': msg.as_string()})
        elif isAccept == 'N':
            msg = generate_deny_mail(row)
            if msg == False:
                print('No this topic {}'.format(worksheet_proposal.cell('B' + str(row)).value))
                return
            server.send_raw_email(RawMessage={'Data': msg.as_string()})

        row += 1

main()

