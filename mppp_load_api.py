import logging
import pathlib
import urllib.parse
import uuid
import warnings
from sys import platform
from time import sleep

import pandas as pd
import requests
import win32com.client
import yaml
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry

warnings.filterwarnings("ignore")

CADESCOM_CADES_BES = 1  # Тип подписи CAdES-BES
CAPICOM_ENCODE_BINARY = 0  # Кодировка DER
CERTSTORE = win32com.client.Dispatch("CAdESCOM.Store")
CERTSTORE.Open(2, "My", 0)

# Настройки для логера
if platform == "linux" or platform == "linux2":
    logging.basicConfig(
        filename="/var/log/log-execute/mppp_load.log.txt",
        level=logging.INFO,
        format=(
            "%(asctime)s - %(levelname)s - "
            "%(funcName)s: %(lineno)d - %(message)s"
        ),
    )
elif platform == "win32":
    logging.basicConfig(
        filename=(
            f"{pathlib.Path(__file__).parent.absolute()}/mppp_load.log.txt"
        ),
        level=logging.INFO,
        format=(
            "%(asctime)s - %(levelname)s - "
            "%(funcName)s: %(lineno)d - %(message)s"
        ),
    )

# Загружаем yaml файл с настройками
with open(
    f"{pathlib.Path(__file__).parent.absolute()}/settings.yaml",
    "r",
    encoding="utf8",
) as yaml_file:
    settings = yaml.safe_load(yaml_file)
telegram_settings = pd.DataFrame(settings["telegram"])
avsoltek_settings = pd.DataFrame(settings["avsoltek"])
greenrus_settings = pd.DataFrame(settings["greenrus"])
sunveter_settings = pd.DataFrame(settings["sunveter"])


def telegram(i, text):
    # Функция отправки уведомлений в telegram на любое количество каналов
    # (указать данные в yaml файле настроек)
    try:
        msg = urllib.parse.quote(str(text))
        bot_token = str(telegram_settings.bot_token[i])
        channel_id = str(telegram_settings.channel_id[i])

        retry_strategy = Retry(
            total=3,
            status_forcelist=[101, 429, 500, 502, 503, 504],
            method_whitelist=["GET", "POST"],
            backoff_factor=1,
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        http = requests.Session()
        http.mount("https://", adapter)
        http.mount("http://", adapter)

        http.post(
            f"https://api.telegram.org/bot{bot_token}/sendMessage?chat_id={channel_id}&text={msg}",
            verify=False,
            timeout=10,
        )
    except Exception as err:
        print(f"mppp_load: Ошибка при отправке в telegram - {err}")
        logging.error(f"mppp_load: Ошибка при отправке в telegram - {err}")


def br_login(signature_header):
    # функция авторизации на бр через api
    url_test_auth = "https://br.so-ups.ru:446/TestApi/TestAuth"
    while True:
        request_test_auth = SESSION.get(
            url_test_auth,
            headers={
                "Accept": "text/plain",
                "client-certificate": signature_header,
                "Host": "br.so-ups.ru:446",
                "Connection": "Keep-Alive",
            },
            verify=False,
        )
        if request_test_auth.status_code == 200:
            return request_test_auth.text
        print(request_test_auth.status_code)
        telegram(
            1,
            "mppp_load: Неуспешная авторизация на сайте br.so-ups.ru. Статус:"
            f" {request_test_auth.status_code}",
        )
        logging.info(
            "mppp_load: Неуспешная авторизация на сайте br.so-ups.ru. Статус:"
            f" {request_test_auth.status_code}"
        )
        sleep(5)


# загружаем макет из файла мппп
with open(
    f"{pathlib.Path(__file__).parent.absolute()}/AVSOLTEK.txt",
    "r",
    encoding="utf-8",
) as txt_file:
    maket = txt_file.read()
maket = maket.replace("\n", "\r\n")


# хардкод компания=avsoltek
# переделать на условие (возможно от названия файла)
company = "AVSOLTEK"

SESSION = requests.Session()

if company == "SUNVETER":
    SERIAL_NUMBER = str(sunveter_settings.x509id[0])
    SIGNOWNER = str(sunveter_settings.signowner[0])

if company == "AVSOLTEK":
    SERIAL_NUMBER = str(avsoltek_settings.x509id[0])
    SIGNOWNER = str(avsoltek_settings.signowner[0])

if company == "GREENRUS":
    SERIAL_NUMBER = str(greenrus_settings.x509id[0])
    SIGNOWNER = str(greenrus_settings.signowner[0])

# поиск сертификата в хранилище
for i in range(1, CERTSTORE.Certificates.count + 1):
    if CERTSTORE.Certificates.Item(i).SerialNumber == SERIAL_NUMBER:
        cert = CERTSTORE.Certificates.Item(i)

# для передачи подписания сертификатом передаем байтовую строку,
signer = win32com.client.Dispatch("CAdESCOM.CPSigner")
signer.Certificate = cert
signedData = win32com.client.Dispatch("CAdESCOM.CadesSignedData")
signedData.Content = bytes(maket, "utf-8")
signature = signedData.SignCades(signer, CADESCOM_CADES_BES, True, 0)
verify_cades = signedData.VerifyCades(signature, CADESCOM_CADES_BES, True)
signature = signature.replace("\r\n", "")


# для авторизации нужно передать в заголовке "client-certificate"
# подписанные сертификатом произвольные данные
# для этого используем генератор рандомного UUID
signedData.Content = bytes(str(uuid.uuid4()), "utf-8")
signature_header = signedData.SignCades(signer, CADESCOM_CADES_BES, False, 0)
verify_cades = signedData.VerifyCades(
    signature_header, CADESCOM_CADES_BES, False
)
signature_header = signature_header.replace("\r\n", "")


# формируем финальный json для post-запроса
json_to_post = {}
json_to_post["RequestTypeId"] = 1
json_to_post["CertificateSerialNumber"] = SERIAL_NUMBER
json_to_post["CertificateSubjectName"] = SIGNOWNER
json_to_post["Sign"] = signature
json_to_post["Maket"] = maket


# авторизация на бр
br_auth = br_login(signature_header)

# подписать и отправить
url_sign_save = "https://br.so-ups.ru:446/PersonalApi/SendMPOfferRequest"
request_sign_save = SESSION.post(
    url_sign_save,
    headers={
        "Accept": "text/plain",
        "Content-Type": "application/json",
        "client-certificate": signature_header,
        "Host": "br.so-ups.ru:446",
        "Expect": "100-continue",
    },
    json=json_to_post,
    verify=False,
)
