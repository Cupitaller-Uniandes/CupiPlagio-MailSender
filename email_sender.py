"""
Recieves an excel with all sections. Each section contains the student that may have cheated
Each section is sent as an independent excel via email to its respective professor
"""

import polars as pl
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import unicodedata
import re
from dotenv import load_dotenv

load_dotenv()

AWS_HOST = os.getenv('AWS_HOST')
AWS_PORT = os.getenv('AWS_PORT')
AWS_USER = os.getenv('AWS_USER')
AWS_PASS = os.getenv('AWS_PASS')
EMAIL_USERNAME = os.getenv('EMAIL_USERNAME')

def normalize_foreign_name(name):
    """
    Normalizes a foreign name by:
    1. Converting to lowercase.
    2. Decomposing Unicode characters to their base form (e.g., 'é' to 'e').
    3. Removing non-alphanumeric characters (except spaces).
    4. Stripping leading/trailing whitespace and reducing multiple spaces to single spaces.

    Args:
        name (str): The foreign name to normalize.

    Returns:
        str: The normalized name.
    """
    # Convert to lowercase
    name = name.lower()

    # Decompose Unicode characters (e.g., accents) and remove non-spacing marks
    name = ''.join(c for c in unicodedata.normalize('NFD', name) if unicodedata.category(c) != 'Mn')

    # Remove non-alphanumeric characters (keep spaces)
    name = re.sub(r'[^a-z0-9\s]', '', name)

    # Strip leading/trailing whitespace and replace multiple spaces with single spaces
    name = re.sub(r'\s+', ' ', name).strip()

    return name

def email_sender(to_emails, subject, body, attachment):
    msg = MIMEMultipart()
    msg['From'] = EMAIL_USERNAME
    msg['To'] = ", ".join(to_emails)
    msg['Subject'] = subject
    
    #Body
    msg.attach(MIMEText(body, 'html'))

    #Attachment
    filename = attachment
    file = open(filename, "rb")

    p = MIMEBase('application', 'octet-stream')

    p.set_payload((file).read())

    encoders.encode_base64(p)

    p.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msg.attach(p)

    s = smtplib.SMTP(AWS_HOST, AWS_PORT)

    s.starttls()

    s.login(AWS_USER, AWS_PASS)

    text = msg.as_string()

    s.sendmail(EMAIL_USERNAME, to_emails, text)

    s.quit()


#email_sender(["d.fuquen@uniandes.edu.co", "s.velasquezm2@uniandes.edu.co"], "Pruebita uno", "Enviando excel prueba1.xlsx", "prueba1.xlsx")

    




df = pl.read_excel(source="Prueba-Correos.xlsx", sheet_id=0)


for sheet in df.keys():
    #print(sheet)
    df_sheet = df[sheet]
    #print(df_sheet)
    #print(df_sheet["profesor"][0])
    prof_email = df_sheet["profesor"][0].replace('\xa0', '').strip()
    sheet = normalize_foreign_name(sheet)
    workbook=sheet+".xlsx"
    workbook = workbook.replace(" ", "_")
    df_sheet.write_excel(workbook=workbook,column_totals=True, autofit=True)  
    body = """
    <div style="font-family: Arial, sans-serif; color: #2c3e50; line-height: 1.6; max-width: 600px; margin: auto;">
      <p>Estimado/a <strong>$Profesor</strong>,</p>

      <p>Le informamos que ha recibido una <strong style="color: #e76c5f;">amonestación</strong> relacionada con su monitoría en el curso <strong>CUPI</strong>.</p>

      <div style="background-color: #f0f4f8; border-left: 4px solid #2980b9; padding: 12px 18px; margin: 20px 0; border-radius: 4px;">
        PLAGIO!!!!
      </div>

      <p><strong>Fecha de la amonestación:</strong> HOY</p>
      <p><strong>Registrada por:</strong> PEPITO PEREZ</p>

      <p style="margin-top: 20px;">
        Esta amonestación quedará registrada en el historial de actividad que mantiene la Coordinación de IP, como parte del seguimiento semanal para asegurar la calidad del servicio prestado.
      </p>

      <p><strong>Importante:</strong> Este mensaje no requiere respuesta, a menos que desee aportar una justificación válida.</p>

      <p>Las excusas por inasistencia deberán presentarse conforme a lo establecido en el <strong>Artículo 45</strong> del
        <a href="https://secretariageneral.uniandes.edu.co/images/documents/reglamento-pregrado-web-2025.pdf" 
          style="color: #2980b9; text-decoration: underline;" 
          target="_blank">
          Reglamento General de Estudiantes de Pregrado.
        </a>
      </p>

      <p>También puede revisar los criterios de excusas aceptadas en el siguiente 
        <a href="https://secretariageneral.uniandes.edu.co/images/documents/Reglamentacion-incapacidades-estudiantiles.pdf" 
          style="color: #2980b9; text-decoration: underline;" 
          target="_blank">
          enlace.
        </a>
      </p>

      <p>Por favor, tenga en cuenta que no se garantiza que otras respuestas o explicaciones no justificadas sean procesadas.</p>

      <p style="margin-top: 30px;">Atentamente,<br><strong>El equipo de CupiMonitores</strong></p>
    </div>

    """
    email_sender([prof_email, "cupitaller@uniandes.edu.co"], "Prueba 2, enviado desde el loop", body=body, attachment=workbook)
    print("Sent email to: ", sheet, prof_email)
    os.remove(workbook)
