import mail, excel, keys
# filename = '1차 후원자 이메일 설문.xlsx'
filename = '후원자명단1차.xlsx'
excel.init(filename)


sheet_list = [
    {
        'sheetname' : '선물#1 (5,000 원)',
        'name_column' : 'C',
        'email_column' : 'O',
        'discord_column' : 'P',
    },
    {
        'sheetname' : '선물#2 (20,000 원)',
        'name_column' : 'C',
        'email_column' : 'U',
        'discord_column' : 'V',
    },
    {
        'sheetname' : '선물#3 (21,000 원)',
        'name_column' : 'C',
        'email_column' : 'U',
        'discord_column' : 'V',
    },
    {
        'sheetname' : '선물#4 (30,000 원)',
        'name_column' : 'C',
        'email_column' : 'U',
        'discord_column' : 'V',
    },
    {
        'sheetname' : '선물#5 (60,000 원)',
        'name_column' : 'C',
        'email_column' : 'U',
        'discord_column' : 'V',
    },
]

sent_column = 'X'
key1_column = 'Y'
key2_column = 'Z'

for sheet in sheet_list:
    excel.write(sheet['sheetname'], 1, sheet['name_column'], '이름')
    excel.write(sheet['sheetname'], 1, sheet['email_column'], '이메일')
    excel.write(sheet['sheetname'], 1, sheet['discord_column'], '디스코드')
    excel.write(sheet['sheetname'], 1, sent_column, '발송 여부')
    excel.write(sheet['sheetname'], 1, key1_column, '스팀키1')
    excel.write(sheet['sheetname'], 1, key2_column, '스팀키2')

    for i in range(2, 100):
        name = excel.read(sheet['sheetname'], i, sheet['name_column'])
        if name == None:
            break
        email = excel.read(sheet['sheetname'], i, sheet['email_column'])
        # print(email)
        # print(type(email))
        if email == None:
            print(f'{name}: 이메일 없음')
            continue

        # TODO: 제품키 읽는 코드 삽입
        sent = excel.read(sheet['sheetname'], i, sent_column)
        if sent == 'O':
            print(f'{name}: 이미 발송됨')
            excel.read(sheet['sheetname'], i, key1_column)
            excel.read(sheet['sheetname'], i, key2_column)
            continue
        key1 = keys.get_key()
        key2 = keys.get_key()
        excel.write(sheet['sheetname'], i, sent_column, 'O')
        excel.write(sheet['sheetname'], i, key1_column, key1)
        excel.write(sheet['sheetname'], i, key2_column, key2)
        title = '[Sentinels] 스팀 제품 키 발송 안내'
        message = f'''{name}님 안녕하세요.
'Sentinels' 프로젝트를 후원해주셔서 감사합니다.

[스팀 키]
아래 스팀키를 사용해 스팀에서 베타 버전을 플레이 해보실 수 있습니다.
총 2개의 스팀 키를 발송해드렸습니다.
등록에 문제가 있다면 본 메일에 회신해주시면 감사하겠습니다.
(스팀 키 등록 방법: https://seogilang.tistory.com/324 )
스팀 키 1: {key1}
스팀 키 2: {key2}

[정식 출시 예정일]
특별한 문제가 없다면 11월 중순 즈음에 정식 출시될 예정입니다.

[디스코드 참여 링크]
https://discord.gg/9DjFyGQu6G

[버그 리포트]
디스코드의 bug-report 채널을 이용해주세요. 
영상 캡처이나 사진 캡처 그리고 자세한 설명을 첨부해주시면 큰 도움이 됩니다. 

[밸런스 설문]
https://forms.gle/KFVNtLYf2j4YAhyo6 
밸런스 설문 자료를 바탕으로 출시 전 캐릭터들의 성능이 조정될 수 있습니다.

{name}님 프로젝트를 후원해주셔서 다시 한 번 감사합니다.
'''
        mail.sendEmail(email, title, message)


excel.save()
keys.save()