import pandas as pd
from datetime import datetime, timedelta
import win32com.client as win32
from pathlib import Path
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.files.file import File
from io import BytesIO

# =========================
# SHAREPOINT SETTINGS
# =========================

username = "font.a@pg.com"  # Tu correo de SharePoint
password = "Cartag08411!"       # Tu contraseña de SharePoint
site_url = "https://pgone.sharepoint.com/sites/NATST"
file_url = "/sites/NATST/Ops/Shared Documents/PLATINUM/Change Management NA Platinum/TIM Change mgmt E2E source of truth.xlsx"
SHEET_NAME = "Changes tracker"   # Nombre de la hoja que deseas leer

# =========================
# FECHAS
# =========================

send_date = datetime.today().strftime("%B %d, %Y")
cab_date = (datetime.today() + timedelta(days=1)).date()

# =========================
# LEER EXCEL DE SHAREPOINT
# =========================

# Autenticación y creación del contexto
ctx = ClientContext(site_url).with_credentials(UserCredential(username, password))

# Descargar el archivo de SharePoint usando open_binary
response = File.open_binary(ctx, file_url)

# Leer el archivo de Excel en un DataFrame
excel_data = BytesIO(response.content)
df = pd.read_excel(excel_data, sheet_name=SHEET_NAME)

df.columns = df.columns.str.strip()
df["Planned TIM CAB"] = pd.to_datetime(df["Planned TIM CAB"], errors="coerce").dt.date
df = df[df["Planned TIM CAB"] == cab_date]

if df.empty:
    raise Exception("No hay cambios para el CAB de mañana.")

# =========================
# TABLA HTML (EXECUTIVE SUMMARY)
# =========================

table_html = """
<table width="100%" cellpadding="6" cellspacing="0"
       style="border-collapse: collapse;
              font-family: 'Aptos Narrow', Calibri, Arial, sans-serif;
              font-size: 14px;
              border: 1px solid #7A7A7A;">
<tr style="background-color:#80CEE1; font-weight:bold;">
<th style="border:1px solid #7A7A7A;">Application triggering change</th>
<th style="border:1px solid #7A7A7A;">Change number</th>
<th style="border:1px solid #7A7A7A;">Description</th>
<th style="border:1px solid #7A7A7A;">Link</th>
<th style="border:1px solid #7A7A7A;">Quality Review</th>
<th style="border:1px solid #7A7A7A;">Planned TIM CAB</th>
</tr>
"""

for _, row in df.iterrows():
    link = row.get("Link to the change Description(Confluence or SNOW)", "")
    link_html = f'<a href="{link}" target="_blank">Link</a>' if pd.notna(link) and link else ""
    table_html += f"""
<tr>
<td style="border:1px solid #7A7A7A;">{row.get('Application triggering change','')}</td>
<td style="border:1px solid #7A7A7A;">{row.get('Change number','')}</td>
<td style="border:1px solid #7A7A7A;">{row.get('Description','')}</td>
<td style="border:1px solid #7A7A7A; text-align:center;">{link_html}</td>
<td style="border:1px solid #7A7A7A;">{row.get('Quality Review','')}</td>
<td style="border:1px solid #7A7A7A;">{row['Planned TIM CAB'].strftime('%m-%d-%Y')}</td>
</tr>
"""

table_html += "</table>"

# =========================
# APPROVERS (TEXTO PLANO, ESTABLE)
# =========================

def get_approvers(app):
    app = app.upper().strip()
    if app == "PACE":
        return "@Humphreys, Erika; @Blochtchinski, Sasha; @Hanser, Travis"
    if app == "MACE":
        return "@Humphreys, Erika; @Blochtchinski, Sasha; @Hanser, Travis; @Bien, Adam"
    if app == "IDGTM":
        return "@Humphreys, Erika; @Blochtchinski, Sasha; @Hanser, Travis; @Bien, Adam"
    if app == "IDGTM ITALY":
        return "@Humphreys, Erika; @Bien, Adam; @Bendekovic, Leo"
    if app == "MACE ITALY":
        return "@Humphreys, Erika; @Bien, Adam; @Bendekovic, Leo; @Acuna, Victor"
    if app == "Launchpad":
        return "@Humphreys, Erika; @Blochtchinski, Sasha; @Hanser, Travis "
    if app == "PACE ITALY":
        return "@Humphreys, Erika; @Bien, Adam; @Bendekovic, Leo; @Piecek, Joanna"
    return ""

# =========================
# MINUTA HTML
# =========================

minutes_html = ""

for _, row in df.iterrows():
    app = row.get("Application triggering change", "")
    deploy_date = row["Planned Release Date"].strftime("%B %d, %Y")
    approvers = get_approvers(app)
    
    # Obtener el contenido de la columna "Notes"
    notes = row.get("Notes", "")
    
    # Verificar si las notas contienen "Conditional approval"
    if "Conditional approval" in notes:
        documentation_status = notes  # Solo muestra las notas
    else:
        documentation_status = f"Notes: {notes} <br> All the documentation is ready"
    
    minutes_html += f"""
<p style="font-family:'Aptos Narrow', Calibri, Arial, sans-serif;
          font-size:14px;
          margin-top:15px;">
<b>{app}</b> Triggering a change cross app impact, Deployment Day: <b>{deploy_date}</b>
<br><br>

{documentation_status}, approvals required: {approvers}

</p>
"""

# =========================
# LEER EL ARCHIVO HTML COMO PLANTILLA
# =========================

html_template_path = r"C:\Users\font.a\OneDrive - Procter and Gamble\Desktop\Automation\ChangeManagement\Plantilla.html"

with open(html_template_path, 'r', encoding='utf-16') as file:
    html_content = file.read()

# Reemplazar los marcadores de posición en el HTML
html_content = html_content.replace("{{SEND_DATE}}", send_date)
html_content = html_content.replace("{{CHANGE_ROWS}}", table_html)
html_content = html_content.replace("{{MINUTE_CONTENT}}", minutes_html)

# =========================
# CREAR Y ENVIAR CORREO
# =========================

outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

# Establecer el cuerpo del correo HTML
mail.HTMLBody = html_content
mail.Subject = f"[Upcoming changes cross application impact] Weekly Ecosystem Meeting -- {send_date}"
mail.To = "font.a@pg.com"
mail.Importance = 2
mail.Send()

print("✅ Correo enviado correctamente con tabla y minuta.")
