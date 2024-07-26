import requests
import json
import pyodbc
import smtplib 
from email.mime.text import MIMEText 
from email.mime.multipart import MIMEMultipart
import pandas as pd
from IPython.display import display
from datetime import datetime
from IPython.display import display
from datetime import datetime, timedelta
today = datetime.today() 
yesterday = today - timedelta(days=1)
tarih = yesterday.strftime('%Y-%m-%d')

def tokenAl():
    import requests
    import json
    url = "http://192.168.21.20:7034/api/v2/token"
    payload = "grant_type=password&branchcode=0&password={Sifre}&username={kullanıcıAdi}&dbname={DbName}&dbuser=TEMELSET&dbpassword={{}}&dbtype=0"
    headers = {
  'Content-Type': 'text/plain',
  'Cookie': 'EF_LOGINTOKEN=EF_LOGINTOKEN=3db8725a-2832-4c5f-96c5-9651f77d437b; EFlowCulture=tr'
}
    response = requests.request("GET", url, headers=headers, data=payload)
    data = json.loads(response.text)
    access_token = data['access_token']
   
    if(access_token):
        print("Token Bilgisi Alındı")
        return access_token
    elif access_token==None:
        print("Token alınamadı")

access_token=tokenAl()
irsaliyeNoBilgileri=[]  

class MailManager:
    def __init__(self) -> None:
        self.fromMail = "mailadresi@mail.com" 
        self.fromPassword ="123456" 
        self.smtpHost = "smtp.mail.com"
        self.smtpPort =465  
       

    def SendMail(self, subject, content, toMails):
        try:
            message = MIMEMultipart("alternative") 
            message["Subject"] = subject
            message["From"] = self.fromMail
           

            _content = MIMEText(content.encode('utf-8'), _charset='utf-8') 
            message.attach(_content) 

            with smtplib.SMTP_SSL(self.smtpHost, self.smtpPort) as server: 
                server.login(self.fromMail, self.fromPassword)
                for toMail in toMails:
                    message["To"]=toMail
                    server.sendmail(self.fromMail,toMail,message.as_string())
                    print(f"Mail başarıyla gönderildi: {toMail}") 
                
        except Exception as e:
            print(f"Mail gönderilemedi!\nHata Kodu: {e}\nMail başlığı : {subject} \nMail içeriği : {content} \nAlıcı : {toMail}")

mailManager = MailManager()

server = '192.168.21.20\\NETSIS' 
database = 'UYGULAMALAR'
username = 'kullaniciAdi'
password = 'sifre'
connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
conn = pyodbc.connect(connection_string)
cursor = conn.cursor()
sql_query = 'SELECT * FROM IRSALIYE_KES'  
cursor.execute(sql_query)
rows = cursor.fetchall()
df = pd.read_sql(sql_query, conn)
cursor.close()
conn.close()

def irsaliyeNoUret():
    url = "http://192.168.21.20:7034/api/v2/ItemSlips/NewEWaybillNumber"
    payload = {}

    headers = {
    'Authorization': 'Bearer ' +  access_token,
    'Content-Type': 'application/json',
    'Cookie': 'EF_LOGINTOKEN=EF_LOGINTOKEN=3db8725a-2832-4c5f-96c5-9651f77d437b; EFlowCulture=tr'
}

    json_data = {
    "Code": "IRS",
    "DocumentType": 'ftSIrs',
}

 
    json_string = json.dumps(json_data) 

    response = requests.post(url, headers=headers, data=json_string)
    FatuIRsNO=response.text
    response_data = response.json()
    FatuIRsNO = response_data["Data"] 
    prefix = FatuIRsNO[:-9]
    numeric_part = FatuIRsNO[-9:]
    updated_numeric_part = str(int(numeric_part) + 0).zfill(9)
    updated_fatuIRsNO = prefix + updated_numeric_part
  
    gibFaturs=updated_fatuIRsNO[:3]
    gibFaturs=gibFaturs+'2024'+'0000'+updated_fatuIRsNO[10:]
    return updated_fatuIRsNO
 
    
def gibIrsUret(): 
    gibFaturs=irsaliyeNoUret()[:3]
    gibFaturs=gibFaturs+'2024'+'0000'+irsaliyeNoUret()[10:]
    return gibFaturs

url = "http://192.168.21.20:7034/api/v2/ItemSlips"
headers = {
    'Content-Type': 'application/json', 
    'Authorization': 'Bearer ' +  access_token 
}
json_data_list = [] 
kalems_list = []
for index, row in df.iterrows():
    if row["TIPI"] == '2':
        kalem = {
            "StokKodu": row["STOK_KODU"],
            "Sira": index,
            "STra_FATIRSNO": irsaliyeNoUret(),
            "STra_GCMIK": int(row["MIKTAR"]),
            "STra_NF": float(row["FIYAT"]),
            "STra_BF": float(row["FIYAT"]),
            "STra_TAR": tarih,
            "Olcubr": 1,
            "STra_HTUR": "H",
            "STra_FTIRSIP": "A",
            "DEPO_KODU": row["DEPO_KODU"],
            "D_YEDEK10": tarih,
            "Stra_FiyatTar": tarih,
        }
        kalems_list.append(kalem)
         
    json_data = {
        "Seri": "IRS",
        "FatUst": {
            "Sube_Kodu": 0,
            "CariKod": "120-00-1315",
            "FATIRS_NO": irsaliyeNoUret(),
            "Tarih": tarih,
            "Tip": 2,
            "FiiliTarih": tarih,
            "DovBazTarihi":tarih,
            "TIPI": 2,
            "SIPARIS_TEST": tarih,
            "EFatOzelKod": 2,
            "EfaturaCarisiMi": True,
            "EIrsaliye": True,
            "GIB_FATIRS_NO": gibIrsUret(),
            "EXFIILITARIH": tarih,
            "KOD1": "2",
            "FIYATTARIHI": tarih,
        },
        "EIrsEkBilgi": {
     "PLAKA": "Yerinde Teslim",
    "TASIYICIVKN": "",
    "TASIYICIADI": "",
    "TASIYICIILCE": "",
    "TASIYICIIL": "",
    "TASIYICIULKE": "",
    "TASIYICIPOSTAKODU": "", 
    "SEVKTAR": tarih,
  },
        "Kalems": kalems_list,
    }
print(json_data)
json_data_list.append(json_data)
json_string = json.dumps(json_data)

kalems_list = []
response = requests.post(url, data=json_string, headers=headers)

if 'data' in response:
    data = response['data']
    if len(data) > 0:
        print(data[0])  
    else:
        print("Yanıt veri listesi boş.")
else:
    print("Yanıt beklenen yapıda değil.")


if response.status_code == 200:
    
    json_response = response.json()
   
    if json_response.get('ErrorCode') is not None:
        mailManager.SendMail("İrsaliye Aktarımı HATA", f"İrsaliye Aktarımında Hata Meydana Geldi : {json_response['ErrorDesc']} (ErrorCode: {json_response['ErrorCode']}", 'sistem@alindair.com.tr')
        print(f"Uyarı: ErrorCode Mevcut ({json_response['ErrorCode']})", json_response)
    else:
          
        import pyodbc
        import pandas as pd
        from IPython.display import display
        server = '192.168.21.20\\NETSIS'
        database = 'UYGULAMALAR'
        username = 'kullaniciAdi'
        password = 'sifre'
        connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        delete_query = "DELETE FROM IRSALIYE_KES WHERE TIPI = '2'" 
        cursor.execute(delete_query)
        conn.commit() 
        select_query = "SELECT * FROM IRSALIYE_KES"
        df = pd.read_sql(select_query, conn)
         
        cursor.close()
        conn.close()
        irsaliyeNoBilgileri.append(json_response['Data']['FatUst']['FATIRS_NO'])
        fatirs_no=json_response['Data']['FatUst']['FATIRS_NO']
        subject = "Otomatik İrsaliye Aktarımı"
        content = f"{fatirs_no} Nolu İrsaliye Aktarımı Başarılı Bir Şekilde Gerçekleştirilmiştir. Lütfen Erp üzerinden kontrol ediniz..!"
        recipients = ["mail1@mail.com.tr", "mail2@mail.com.tr","mail3@mail.com.tr"]

        mailManager.SendMail(subject, content, recipients)
        
        print("Kayıtlar silindi ve güncel veriler gösterildi.")
else:
    print("Hata:", response.status_code, response.text)
    mailManager.SendMail("Servis Hatası - UYARI", "İrsaliye Aktarım Servisinde Hata Meydana Geldi", 'mail1@mail.com.tr')
 
server = '192.168.21.20\\NETSIS' 
database = 'UYGULAMALAR'
username = 'kullaniciAdi'
password = 'sifre'
connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
conn = pyodbc.connect(connection_string)
cursor = conn.cursor()
sql_query = 'SELECT * FROM IRSALIYE_KES'  
cursor.execute(sql_query)
rows = cursor.fetchall()
df = pd.read_sql(sql_query, conn)
display(df)
cursor.close()
conn.close()

def irsaliyeNoUret2():
    url = "http://192.168.21.20:7034/api/v2/ItemSlips/NewEWaybillNumber"
    payload = {}

    headers = {
    'Authorization': 'Bearer ' +  access_token,
    'Content-Type': 'application/json',
    'Cookie': 'EF_LOGINTOKEN=EF_LOGINTOKEN=3db8725a-2832-4c5f-96c5-9651f77d437b; EFlowCulture=tr'
}

    json_data = {
    "Code": "GIR",
    "DocumentType": 'ftSIrs',
}

 
    json_string = json.dumps(json_data) 

    response = requests.post(url, headers=headers, data=json_string)
    FatuIRsNO=response.text
    response_data = response.json()
    FatuIRsNO = response_data["Data"] 
    prefix = FatuIRsNO[:-9]
    numeric_part = FatuIRsNO[-9:]
    updated_numeric_part = str(int(numeric_part) + 0).zfill(9)
    updated_fatuIRsNO = prefix + updated_numeric_part
  
    gibFaturs=updated_fatuIRsNO[:3]
    gibFaturs=gibFaturs+'2024'+'0000'+updated_fatuIRsNO[10:]
    return updated_fatuIRsNO
 
    
def gibIrsUret2(): 
    gibFaturs=irsaliyeNoUret2()[:3]
    gibFaturs=gibFaturs+'2024'+'0000'+irsaliyeNoUret2()[10:]
    return gibFaturs
gibIrsUret2()
irsaliyeNoUret2()

url = "http://192.168.21.20:7034/api/v2/ItemSlips"
headers = {
    'Content-Type': 'application/json', 
    'Authorization': 'Bearer ' +  access_token 
}
json_data_list = [] 
kalems_list = []
for index, row in df.iterrows():
    if row["TIPI"] == '6':
        kalem = {
            "StokKodu": row["STOK_KODU"],
            "Sira": index,
            "STra_FATIRSNO": irsaliyeNoUret2(),
            "STra_GCMIK": int(row["MIKTAR"]),
            "STra_NF": float(row["FIYAT"]),
            "STra_BF": float(row["FIYAT"]),
            "STra_TAR": tarih,
            "Olcubr": 1,
            "STra_HTUR": "H",
            "STra_FTIRSIP": "A",
            "DEPO_KODU": row["DEPO_KODU"],
            "D_YEDEK10": tarih,
            "Stra_FiyatTar": tarih,
        }
        kalems_list.append(kalem)  
    json_data = {
        "Seri": "IRS",
        "FatUst": {
            "Sube_Kodu": 0,
            "CariKod": "120-00-1315",
            "FATIRS_NO": irsaliyeNoUret2(),
            "Tarih": tarih,
            "Tip": 2,
            "FiiliTarih": tarih,
            "DovBazTarihi":tarih,
            "TIPI": 6,
            "SIPARIS_TEST": tarih,
            "EFatOzelKod": 2,
            "EfaturaCarisiMi": True,
            "EIrsaliye": True,
            "GIB_FATIRS_NO": gibIrsUret2(),
            "EXFIILITARIH": tarih,
            "KOD1": "5",
            "FIYATTARIHI": tarih,
             "EXPORTTYPE": 5,
            "EXPORTREFNO":irsaliyeNoUret2(),
            "PLA_KODU":'11',
        },
        "EIrsEkBilgi": {
     "PLAKA": "",
    "TASIYICIVKN": "",
    "TASIYICIADI": "",
    "TASIYICIILCE": "",
    "TASIYICIIL": "",
    "TASIYICIULKE": "",
    "TASIYICIPOSTAKODU": "", 
    "SEVKTAR": tarih,
  },
        "Kalems": kalems_list,
    }  
json_data_list.append(json_data)
json_string = json.dumps(json_data)
kalems_list = []


response = requests.post(url, data=json_string, headers=headers)
if response.status_code == 200:
    json_response = response.json()
    if json_response.get('ErrorCode') is not None:
        print("Başarılı:", response.json())
        print(f"Uyarı: ErrorCode mevcut ({json_response['ErrorCode']}) " )
    else:
        print("Başarılı:", json_response)
        server = '192.168.21.20\\NETSIS'
        database = 'UYGULAMALAR'
        username = 'kullaniciAdi'
        password = 'sifre'
        connection_string = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}'
        conn = pyodbc.connect(connection_string)
        cursor = conn.cursor()
        delete_query = "DELETE FROM IRSALIYE_KES WHERE TIPI = '6'" 
        cursor.execute(delete_query)
        conn.commit() 
        select_query = "SELECT * FROM IRSALIYE_KES"
        df = pd.read_sql(select_query, conn)
        cursor.close()
        conn.close()
        fatirs_no=json_response['Data']['FatUst']['FATIRS_NO']
        subject = "Otomatik İrsaliye Aktarımı"
        content = f"{fatirs_no} Nolu İrsaliye Aktarımı Başarılı Bir Şekilde Gerçekleştirilmiştir. Lütfen Erp üzerinden kontrol ediniz..!"
        recipients = ["mail1@mail.com.tr", "mail2@mail.com.tr","mail3@mail.com.tr"]

        mailManager.SendMail(subject, content, recipients)
        irsaliyeNoBilgileri.append(json_response['Data']['FatUst']['FATIRS_NO'])
        print("Kayıtlar silindi ve güncel veriler gösterildi.")
else:
    print("Hata:", response.status_code, response.text)
    mailManager.SendMail("Servis Hatası - UYARI", "İrsaliye Aktarım Servisinde Hata Meydana Geldi", 'mail1@mail.com.tr')
 


print(irsaliyeNoBilgileri)

 
