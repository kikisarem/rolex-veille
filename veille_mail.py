import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 🔐 Identifiants Google Sheet
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# 📄 Feuille Google Sheets
spreadsheet = client.open("Veille Auto Rolex")
sheet = spreadsheet.worksheet("Veille")

# 🧮 Récupérer les 10 dernières lignes
donnees = sheet.get_all_values()
dernieres_lignes = donnees[-10:]

# 🧱 Construction HTML dynamique
html_rows = ""
for ligne in dernieres_lignes:
    html_rows += f"""
    <tr>
        <td>{ligne[0]}</td>  <!-- Modèle -->
        <td>{ligne[1]}</td>  <!-- Prix -->
        <td>{ligne[2]}</td>  <!-- Localisation -->
        <td>{ligne[3]}</td>  <!-- Papiers -->
        <td>{ligne[4]}</td>  <!-- Boîte -->
        <td>{ligne[5]}</td>  <!-- Estim. revente -->
        <td>{ligne[6]}</td>  <!-- Marge -->
        <td>{ligne[7]}</td>  <!-- Tendance -->
        <td><a href="{ligne[8]}">Voir</a></td> <!-- Lien -->
    </tr>
    """

html_content = f"""
<html>
  <body>
    <h2>🕵️ Veille Rolex – Top 10 affaires du jour</h2>
    <table border="1" cellpadding="6" cellspacing="0">
      <tr>
        <th>Modèle</th><th>Prix €</th><th>Localisation</th><th>Papiers</th>
        <th>Boîte</th><th>Revente €</th><th>Marge €</th><th>Tendance</th><th>Lien</th>
      </tr>
      {html_rows}
    </table>
    <br>
    <p style="color:gray;">Email automatique – Rolex Tracker</p>
  </body>
</html>
"""

# 📧 Configuration email
from_email = "alexandre.bimbaud@gmail.com"
to_email = "alexandre.bimbaud@gmail.com"
password = "twzm pyvb yzlv ejhc"  # ⬅️ à modifier

message = MIMEMultipart("alternative")
message["Subject"] = "🔍 Veille Rolex – Top 10 du jour"
message["From"] = from_email
message["To"] = to_email
message.attach(MIMEText(html_content, "html"))

# 📤 Envoi via Gmail
with smtplib.SMTP("smtp.gmail.com", 587) as server:
    server.starttls()
    server.login(from_email, password)
    server.sendmail(from_email, to_email, message.as_string())

print("✅ Email des 10 dernières montres envoyé avec succès.")
