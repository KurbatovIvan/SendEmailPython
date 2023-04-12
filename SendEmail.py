import sys
import os
import email, smtplib, ssl
import configparser  # импортируем библиотеку
import glob
import logging

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

logging.basicConfig(level=logging.INFO, filename="sendEmail_log.log", filemode="a",format="%(asctime)s %(levelname)s %(message)s")
#logging.debug("A DEBUG Message")
#logging.info("An INFO")
#logging.warning("A WARNING")
#logging.error("An ERROR")
#logging.critical("A message of CRITICAL severity")


config = configparser.ConfigParser()  # создаём объекта парсера
try:
 config.read("settings.ini")  # читаем конфиг
 subject =        config["Email"]["subject"]
 body =           config["Email"]["body"]
 sender_email =   config["Email"]["sender_email"]
 receiver_email = config["Email"]["receiver_email"]
 password =       config["Email"]["password"]
 smtp_server =    config["Email"]["smtp_server"]
 smtp_port =      config["Email"]["smtp_port"]
 path_to_dir =    config["Email"]["path_to_dir"]
 clear_dir   =    config["Email"]["clear_dir"]
 receiver_email_bcc = config["Email"]["receiver_email_bcc"]
except Exception as err:
 print(f"Unexpected {err=}, {type(err)=}")
 logging.error (f"Не определен ключ, добавьте его в файле settings.ini {err=}, {type(err)=}")
 raise   

logging.info("----")
logging.info("Начинаем отправку вложений ")

print(path_to_dir)
logging.info("Путь к папке с файлами которые надо отправить " + path_to_dir)

# Создание составного сообщения и установка заголовка
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject
if receiver_email_bcc!="":
  message["Bcc"] = receiver_email_bcc  # Если у вас несколько получателей , скрытая копия
print (receiver_email_bcc)

#sys.exit()

# Внесение тела письма
message.attach(MIMEText(body, "plain"))
 
# Получаем список файлов с маской
files = glob.glob(path_to_dir)

print (files)
for file in files:
            print(os.path.basename(file))

if len(files) == 0:
   logging.info ("Файлов в папке нет, поэтому ничего не шлем и уходим на отдых")
   sys.exit()

for filename in files:
 # Открытие XLS файла в бинарном режиме
 with open(filename, "rb") as attachment:
    # Заголовок письма application/octet-stream
    # Почтовый клиент обычно может загрузить это автоматически в виде вложения
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
 
 # Шифровка файла под ASCII символы для отправки по почте    
 encoders.encode_base64(part)
 
 # Внесение заголовка в виде пара/ключ к части вложения
 part.add_header(
    "Content-Disposition",
    f"attachment; filename= {os.path.basename(filename)}",)
 
 # Внесение вложения в сообщение и конвертация сообщения в строку
 message.attach(part)
# Цикл прикрепления файлов окончен

text = message.as_string()
 
# Подключение к серверу при помощи безопасного контекста и отправка письма
context = ssl.create_default_context()
with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
    server.set_debuglevel(0)
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, text)
    logging.info ("Отправлено письмо на адрес: " + receiver_email)

if clear_dir=="yes": 
   logging.info ("Очищаем папку")
   for f in glob.glob(path_to_dir):
      os.remove(f)

logging.info ("Закончили работу")
sys.exit()

