from dotenv import load_dotenv
load_dotenv()
import os
import smtplib  # SMTP 사용을 위한 모듈
import re  # Regular Expression을 활용하기 위한 모듈
from email.mime.multipart import MIMEMultipart  # 메일의 Data 영역의 메시지를 만드는 모듈
from email.mime.text import MIMEText  # 메일의 본문 내용을 만드는 모듈
from email.mime.image import MIMEImage  # 메일의 이미지 파일을 base64 형식으로 변환하기 위한 모듈

# smpt 서버와 연결
gmail_smtp = "smtp.gmail.com"  # gmail smtp 주소
gmail_port = 465  # gmail smtp 포트번호. 고정(변경 불가)
smtp = smtplib.SMTP_SSL(gmail_smtp, gmail_port)
 
# 로그인
my_account = os.environ.get('email_account')  # .env 파일에 저장한 이메일
my_password = os.environ.get('email_password')  # .env 파일에 저장한 App Password
print(f'smtp login: ', my_account)
# print(my_account, my_password)
smtp.login(my_account, my_password)

def sendEmail(addr: str, title: str, message: str):
    # 메일을 받을 계정
    to_mail = addr
    
    # 메일 기본 정보 설정
    msg = MIMEMultipart()
    msg["Subject"] = title  # 메일 제목
    msg["From"] = my_account
    msg["To"] = to_mail
    
    # 메일 본문 내용
    content = message

    content_part = MIMEText(content, "plain")
    msg.attach(content_part)
    
    # 이미지 파일 추가
    # image_name = "test.png"
    # with open(image_name, 'rb') as file:
    #     img = MIMEImage(file.read())
    #     img.add_header('Content-Disposition', 'attachment', filename=image_name)
    #     msg.attach(img)
    
    # 받는 메일 유효성 검사 거친 후 메일 전송
    print(f'메일 전송: {addr}')
    reg = "^[a-zA-Z0-9.+_-]+@[a-zA-Z0-9]+\.[a-zA-Z]{2,3}$"  # 유효성 검사를 위한 정규표현식
    if re.match(reg, addr):
        smtp.sendmail(my_account, to_mail, msg.as_string())
        print("정상적으로 메일이 발송되었습니다.")
    else:
        print("받으실 메일 주소를 정확히 입력하십시오.")
 

def quit():
    smtp.quit()

