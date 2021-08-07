import gspread
from numpy import record
import pandas as pd
import smtplib, ssl
import time
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

'''
Change these to your credentials and name
'''

your_name = "Fes-te Jove Manlleu"
your_email = "festejovemanlleu@gmail.com"
your_password = ""

'''
# If you are using something other than gmail
# then change the 'smtp.gmail.com' and 465 in the line below
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)
'''


context = ssl.create_default_context()
server = smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context)
server.login(your_email, your_password)

# Gologolo stuff
# define the scope
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('formularis-322115-a44c9b460be4.json', scope)
# authorize the clientsheet 
client = gspread.authorize(creds)

while 1:

  # get the instance of the Spreadsheet
  sheet = client.open('Registre Xarrups de Fes-te Jove - Nit 07/08/21 (respostes)')

  # get the first sheet of the Spreadsheet
  sheet_instance = sheet.get_worksheet(0)

  # get all the records of the data
  records_data = sheet_instance.get_all_records()

  records_df = pd.DataFrame.from_dict(records_data)

  # Apanyem el nom de la columna per no tenir problemes a l'hora de llegir
  records_df['Correu electronic'] = records_df['Correu electrònic']
  records_df.drop('Correu electrònic',inplace=True,axis=1)

  #Convertim les dates a un format conegut
  #records_df["marca_temps_dt"] = datetime.strptime(records_df["Marca de temps"], "%d/%m/%Y %h:%M:%s")
  records_df["marca_temps_dt"] = pd.to_datetime(records_df["Marca de temps"], format="%d/%m/%Y %H:%M:%S")
  records_df_unfiltered = records_df

  # Ordenem per dt (no cal en principi)
  #records_df = records_df.sort_values(by = ['marca_temps_dt'])
  # Esborrar els duplicats de correu que ja tinguin 
  #records_df = records_df.drop_duplicates(subset = ["Correu electrònic"], keep = 'first')

  #Obtenir datetime de l'última fila a l'ultima execució (fitxer)

  with open('dt') as f:
      date_str = f.readline()

  last_dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
  # Obtenir files noves
  records_df = records_df[records_df["marca_temps_dt"] > last_dt]

  # Iterar i enviar correu
  for index, row in records_df.iterrows():

      #Checkem que només hi sigui 1 cop
      if len(records_df_unfiltered[records_df_unfiltered['Correu electronic'] == row['Correu electronic']]) > 1:
          subject = 'INSCRIPCIÓ FALLIDA XARRUPS FJ2021 - NIT 07/08/2021'
          
          html = f"""\
<html>
  <body>
    Hola, {row['Nom']},<br><br>
    Hem detectat que el correu {row['Correu electronic']} ja ha estat usat en aquest formulari. Si us plau, registra't amb un correu que no hagis fet servir.
    <br><br>
    Moltes gràcies,<br>Comissió Fes-te Jove Manlleu
  </body>
</html>
"""

      else:

          subject = 'INSCRIPCIÓ CONFIRMADA XARRUPS FJ2021 - NIT 07/08/2021'

          html = f"""\
<html>
  <body>
    Hola, {row['Nom']},<br><br>
    Et confirmem que les teves dades han estat enregistrades correctament i que, per tant, pots accedir al recinte dels "Xarrups de FJ" al Prat del Rocòdrom durant la nit del 7 d'agost de 2021. <strong>Et recordem que omplint aquest formulari NO ESTÀS RESERVANT LLOC a l'acte, només estàs cedint provisionalment les teves dades perquè la Comissió Fes-te Jove pugui posar-se en contacte amb tu si hi ha qualsevol incidència</strong>, sobretot relacionada amb el Coronavirus. La normativa vigent ens obliga a tenir un registre d'assistents.
    <br><br><div style='font-size:20px'><strong>Les dades que ens has facilitat són les següents:<br><br>
    Nom: {row['Nom']}<br>
    Cognoms: {row['Cognoms']}<br>
    Número de telèfon: {row['Telèfon mòbil']}<br>
    Correu electrònic: {row['Correu electronic']}<br></strong></div><br>
    <strong>Hauràs de mostrar aquest correu electrònic a l'entrada del recinte i confirmar que ets tu ensenyant el DNI al personal de seguretat. Tingues present que aquest tràmit només és vàlid per una nit.</strong><br><br>
    Aprofitem per recordar-te que és obligatori seguir totes les mesures de prevenció: ús de mascareta obligatori, respecte a les indicacions de l'organització, manteniment de les distàncies, etc. <br><br>
    Moltes gràcies per la teva col·laboració. Entre totes i tots fem Fes-te Jove.<br><br>
    Salut,<br>
    Comissió Fes-te Jove Manlleu
  </body>
</html>
"""


      message = MIMEMultipart("alternative")
      message["Subject"] = subject
      message["From"] = your_email
      message["To"] = row['Correu electronic']

      message.attach(MIMEText(html, 'html'))

      try:
        server.sendmail(your_email, row['Correu electronic'], message.as_string())
        #print('Email enviat correctament a {}'.format(row['Correu electronic']))
        with open('log', "w") as myfile:
          myfile.write('Email enviat correctament a {}'.format(row['Correu electronic']))
      except Exception as e:
        #print('Email a {} fallit!'.format(row['Correu electronic']))
        with open('log', "w") as myfile:
          myfile.write('Email a {} fallit!'.format(row['Correu electronic']))


  # Obtenir ultim dt
  if records_df.size > 0:
    dt = records_df['marca_temps_dt'].iloc[-1].strftime("%Y-%m-%d %H:%M:%S")
    with open('dt', "w") as myfile:
        myfile.write(dt)

  time.sleep(5)






