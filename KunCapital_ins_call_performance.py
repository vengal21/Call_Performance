import smtplib
import mysql.connector
import pandas as pd
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os

# Date config
Today_date = date.today()
date_str = Today_date.strftime('%Y-%m-%d')
filename = f"Insurance call performance-{date_str}.xlsx"
filepath = f"C:/Users/W2632/KUN Capital/Reports/{filename}"

# Zoho Mail config
zoho_username = 'vengal.rao@wyzmindz.com'
zoho_password = 'zFejjPCtEDdN'
zoho_server = 'smtp.zoho.com'

# Connect to Zoho Mail
server = smtplib.SMTP(zoho_server, 587)
server.starttls()
server.login(zoho_username, zoho_password)

# MySQL config
db_config = {
    'host': '192.168.1.45',
    'port': 10001,
    'user': 'kuncapital',
    'password': 'kuncapital@123',
    'database': 'kuncapital'
}

# SQL query
query = """
    SELECT 
        bm.brandName as 'Brand',
        ic.CreName AS 'CRE Name', 
        TIME_FORMAT(MIN(STR_TO_DATE(ic.callTime, '%H:%i:%s')), '%h:%i %p') AS 'First Call Time',
        TIME_FORMAT(MAX(STR_TO_DATE(ic.callTime, '%H:%i:%s')), '%h:%i %p') AS 'Last Call Time',
        COUNT(*) AS 'Total Calls',
        COUNT(CASE WHEN ic.isCallinitaited = 'initiated' THEN 1 END) AS 'Initiated Calls',
        COUNT(DISTINCT CASE WHEN ic.primarydisposition = 'Contacts' THEN insuranceAssignedInteraction_id END) AS 'Unique Contact',
        COUNT(CASE WHEN ic.primarydisposition = 'Contacts' THEN 1 END) AS 'Contacts',
        CONCAT(ROUND((SUM(IF(ic.calldispositiondata_id NOT IN (6, 7, 8, 9, 10, 43), 1, 0)) / COUNT(cicallinteraction_id)) * 100, 0), '%') AS 'Contact%'
    FROM 
        insurancecallhistorycube ic
    LEFT JOIN 
        brandmodel bm on ic.model = bm.modelname
    WHERE 
        callDate = CURDATE()
    GROUP BY 
        CreName
    ORDER BY 
        STR_TO_DATE(MIN(callTime), '%H:%i:%s') ASC;
"""

# Connect to DB and fetch data
conn = mysql.connector.connect(**db_config)
cursor = conn.cursor()
cursor.execute(query)
rows = cursor.fetchall()
columns = [i[0] for i in cursor.description]
df = pd.DataFrame(rows, columns=columns)

# Save to Excel
df.to_excel(filepath, index=False)

# Create HTML table
html_table = df.to_html(index=False)
html_table = html_table.replace('<table border="1" class="dataframe">',
                                '<table style="width:100%; border-collapse: collapse; text-align: center;">')
html_table = html_table.replace('<th>',
                                '<th style="background-color: #A3C1DA; color: black; text-align: center; padding: 8px; border: 1px solid black;">')
html_table = html_table.replace('<td>',
                                '<td style="background-color: #FFFFFF; text-align: center; padding: 8px; border: 1px solid black;">')

# Email body
body = f"""
<html>
<head>
    <style>
        table {{
            width: 100%;
            border-collapse: collapse;
            text-align: center;
        }}
        th, td {{
            border: 1px solid black;
            padding: 8px;
        }}
        th {{
            background-color: #A3C1DA;
            color: black;
            text-align: center;
        }}
        td {{
            background-color: #FFFFFF;
            text-align: center;
        }}
    </style>
</head>
<body>
    <p>Greetings!</p>
    <p>Kindly review the CRM reports for Tele caller Performance by brands.</p>
    <div style="overflow-x:auto;">
        {html_table}
    </div>
    <br>
    <p>PS - for detailed review, please log in to CRM Reports</p>
    <br>
    <p><b>Regards,</b><br>
    <b>Vengal Rao</b><br>
    <b>Business Analysis</b><br>
    <b>Autosherpa</b><br>
    <b>+91 9791154159</b></p>
</body>
</html>
"""

# Compose the email
msg = MIMEMultipart()
msg['Subject'] = f"Call Performance Report - KUN Capital - {date_str}"
msg['From'] = zoho_username
msg['To'] = "mis.insurance@kun-capital.com","shankar.g@kun-capital.com"
msg['CC'] = "heads.crm@wyzmindz.com","tharani@wyzmindz.com"

# msg['To'] = "vengal001star@gmail.com"
# msg['CC'] = "vengalraooffical@gmail.com"

# Attach HTML body
msg.attach(MIMEText(body, 'html'))

# Attach Excel file
with open(filepath, 'rb') as f:
    attachment = MIMEApplication(f.read(), _subtype="xlsx")
    attachment.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(attachment)

# Send the email
recipients = ["mis.insurance@kun-capital.com","shankar.g@kun-capital.com","heads.crm@wyzmindz.com","tharani@wyzmindz.com"]
server.sendmail(zoho_username, recipients, msg.as_string())
server.quit()
