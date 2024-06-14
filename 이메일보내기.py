import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import getpass

# 이메일 서버 설정 (Gmail 예시)
smtp_server = 'smtp.naver.com'
smtp_port = 587

# 보내는 사람 정보
email = 'bough38@naver.com'
password = getpass.getpass('Please enter your email password: ')

# 받는 사람 정보와 메시지 내용
recipient_email = 'bough38@naver.com'

# 장문의 메시지 내용을 여기에 작성하세요.
message_content = '''
안녕하세요,

이것은 테스트 이메일입니다.

BadZipFile                                Traceback (most recent call last)
Cell In[1], line 21
     17     wb.save(filename=file_path)  # 변경 사항 저장
     19 file_path = 'D:\\공공기관 데이터모음\\인허가자료\\인허가전체지도만들기 하이퍼링크.xlsx'  # 예: 'D:\\공공기관 데이터모음\\디비리아자료.xlsx'
---> 21 add_naver_map_hyperlinks(file_path)  # 함수 실행
     23 print("네이버 지도 하이퍼링크 추가 작업 완료!")

Cell In[1], line 4, in add_naver_map_hyperlinks(file_path)
      3 def add_naver_map_hyperlinks(file_path):
----> 4     wb = load_workbook(filename=file_path)
      5     ws = wb.active
      7     for row in range(2, ws.max_row + 1):  # 첫 번째 행을 제외하고 시작

File ~\anaconda3\Lib\site-packages\openpyxl\reader\excel.py:344, in load_workbook(filename, read_only, keep_vba, data_only, keep_links, rich_text)
    314 def load_workbook(filename, read_only=False, keep_vba=KEEP_VBA,
    315                   data_only=False, keep_links=True, rich_text=False):
    316     """Open the given filename and return the workbook
    317 
    318     :param filename: the path to open or a file-like object
   (...)
    342 
    343     """
--> 344     reader = ExcelReader(filename, read_only, keep_vba,
    345                          data_only, keep_links, rich_text)
    346     reader.read()
    347     return reader.wb

File ~\anaconda3\Lib\site-packages\openpyxl\reader\excel.py:123, in ExcelReader.__init__(self, fn, read_only, keep_vba, data_only, keep_links, rich_text)
    121 def __init__(self, fn, read_only=False, keep_vba=KEEP_VBA,
    122              data_only=False, keep_links=True, rich_text=False):
--> 123     self.archive = _validate_archive(fn)
    124     self.valid_files = self.archive.namelist()
    125     self.read_only = read_only

File ~\anaconda3\Lib\site-packages\openpyxl\reader\excel.py:95, in _validate_archive(filename)
     88             msg = ('openpyxl does not support %s file format, '
     89                    'please check you can open '
     90                    'it with Excel first. '
     91                    'Supported formats are: %s') % (file_format,
     92                                                    ','.join(SUPPORTED_FORMATS))
     93         raise InvalidFileException(msg)
---> 95 archive = ZipFile(filename, 'r')
     96 return archive

File ~\anaconda3\Lib\zipfile.py:1302, in ZipFile.__init__(self, file, mode, compression, allowZip64, compresslevel, strict_timestamps, metadata_encoding)
   1300 try:
   1301     if mode == 'r':
-> 1302         self._RealGetContents()
   1303     elif mode in ('w', 'x'):
   1304         # set the modified flag so central directory gets written
   1305         # even if no files are added to the archive
   1306         self._didModify = True

File ~\anaconda3\Lib\zipfile.py:1369, in ZipFile._RealGetContents(self)
   1367     raise BadZipFile("File is not a zip file")
   1368 if not endrec:
-> 1369     raise BadZipFile("File is not a zip file")
   1370 if self.debug > 1:
   1371     print(endrec)

BadZipFile: File is not a zip file

​

감사합니다.
'''

# 이메일 서버에 로그인
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(email, password)

# 이메일 구성
msg = MIMEMultipart()
msg['From'] = email
msg['To'] = recipient_email
msg['Subject'] = '파이썬정답찾기 Email'
msg.attach(MIMEText(message_content, 'plain')) # 'plain'은 일반 텍스트를 의미합니다. HTML 이메일을 보내려면 'html'을 사용하세요.

# 이메일 보내기
server.send_message(msg)y

# 이메일 서버와 연결 종료
server.quit()
