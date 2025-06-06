# Importing modules
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import io

# ----- Configuration -----
client_name = "KUN Hyundai"
url = "http://103.249.80.116:9009"
user_name = "kun_inss_admin"
password = "kun@123"

# Email credentials
zoho_smtp = "smtppro.zoho.com"
zoho_port = 465
zoho_email = "vengal.rao@wyzmindz.com"
zoho_password = "CkUT4U8QBzag"
to_recipients = ["lakshmi.ins@kunhyundai.com", "insurancemanager@kunhyundai.com"]
cc_recipients = ["tharani@wyzmindz.com"]

# ----- Selenium Automation -----
firefox_options = webdriver.FirefoxOptions()
firefox_options.add_argument("--headless")
firefox_options.add_argument("--window-size=1920x1080")
firefox_options.add_argument("--disable-notifications")
firefox_options.add_argument("--no-sandbox")
firefox_options.add_argument("--disable-gpu")
firefox_options.add_argument("--disable-software-rasterizer")

driver = webdriver.Firefox(options=firefox_options)
time.sleep(1)
driver.get(url)

driver.find_element(By.XPATH, "//*[@title='Admin']").click()
driver.find_element(By.XPATH, "//*[@title='Enter UserName']").send_keys(user_name)
driver.find_element(By.XPATH, "//*[@title='Enter Password']").send_keys(password)
driver.find_element(By.XPATH, "//*[@title='Login']").click()

try:
    driver.find_element(By.XPATH, "//*[@title='Admin']").click()
except:
    pass

driver.find_element(By.XPATH, "//li[1]").click()
driver.find_element(By.CSS_SELECTOR, "li.insuranceDashboard a").click()
time.sleep(2)

Data_avail = driver.find_element(By.ID, "BoxInsuranceLiveDataAvail").text
CallScoreCard = driver.find_element(By.XPATH, "//span[@id='tblInsuranceCRECallLive_stl']/button")
driver.execute_script("arguments[0].click();", CallScoreCard)

WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.ID, "tblInsuranceCRECallLive"))
)

table_head = driver.find_elements(By.XPATH, "//*[@id='tblInsuranceCRECallLive']//thead//tr//th")
headers = [th.find_element(By.TAG_NAME, "div").get_attribute("innerText").strip() for th in table_head]
t_body = driver.find_elements(By.XPATH, "//*[@id='tblInsuranceCRECallLive']//tbody//tr")

data = []
for row in t_body:
    values = row.find_elements(By.TAG_NAME, "td")
    val = [v.text for v in values]
    data.append(val)

df = pd.DataFrame(data, columns=headers)
driver.quit()

try:
    dshdata = {
        'CRE': df['creName'].count(),
        'Data Avail': Data_avail,
        'Calls': df['Calls'].astype(int).sum(),
        'Calls/CRE': round(df['Calls'].astype(int).sum() / df['creName'].count()),
        'Contacts': df['Contacts'].astype(int).sum(),
        'Contacts %': round((df['Contacts'].astype(int).sum() / df['Calls'].astype(int).sum()) * 100),
        'FreshBookings': df['Fresh Appt'].astype(int).sum(),
        'Cancelled': df['Cancel'].astype(int).sum(),
        'Rescheduled': df['Re-Appt'].astype(int).sum()
    }
except:
    dshdata = {}

dashboard = pd.DataFrame([dshdata])

# Convert columns to integers before sorting
df['Fresh Appt'] = df['Fresh Appt'].astype(int)
df_top_html = df.nlargest(10, 'Fresh Appt').to_html(index=False)
df_bottom_html = df.nsmallest(10, 'Fresh Appt').to_html(index=False)
dashboard_html = dashboard.to_html(index=False)

# ----- Email Setup -----
msg = MIMEMultipart()
current_date = datetime.datetime.now().strftime("%d-%m-%Y")
msg["From"] = zoho_email
msg["To"] = ", ".join(to_recipients)
msg["CC"] = ", ".join(cc_recipients)
msg["Subject"] = f'Call Performance Report(INS) - {client_name} - {current_date}'

html = f"""
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
            text-align: center;
        }}
        th {{
            background-color: #008080;
            color: #FFFFFF;
        }}
    </style>
</head>
<body>
    <h4>Greetings!</h4>
    <h4>Kindly review the CRM Reports for TeleCaller Performance</h4>
    <div style="overflow-x:auto;">
        <h4>Dashboard</h4>
        {dashboard_html}
        <h4>Call ScoreCard</h4>
        <h4>Top 10 TeleCaller Performance</h4>
        {df_top_html}
        <h5>Bottom 10 TeleCaller Performance</h5>
        {df_bottom_html}
        <p>Please find attached file for complete report</p>
    </div>
    <br><br>
    <p>PS - For detailed review, please login to CRM Reports</p>
    <br><br>
    <p>Regards,<br><b>Team AUTOSherpas</b></p>
</body>
</html>
"""
msg.attach(MIMEText(html, "html"))

# ----- Attach Excel File from Memory -----
excel_buffer = io.BytesIO()
df.to_excel(excel_buffer, index=False, engine='openpyxl')
excel_buffer.seek(0)

part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
part.set_payload(excel_buffer.read())
encoders.encode_base64(part)
part.add_header("Content-Disposition", "attachment; filename=Call Performance.xlsx")
msg.attach(part)

# ----- Send Email -----
all_recipients = to_recipients + cc_recipients

try:
    server = smtplib.SMTP_SSL(zoho_smtp, zoho_port)
    server.login(zoho_email, zoho_password)
    server.sendmail(zoho_email, all_recipients, msg.as_string())
    server.quit()
    print(f"{client_name} INS Email sent successfully!")
except Exception as e:
    print(f"{client_name} INS Error sending email: {e}")
