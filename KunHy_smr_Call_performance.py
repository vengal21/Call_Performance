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
import os

# Credentials and Configurations
client_name = "KUN Hyundai"
url = "http://103.249.80.116:9009"
user_name = "kun_smr_admin"
password = "Network@123"

zoho_smtp = "smtppro.zoho.com"
zoho_port = 465
zoho_email = "vengal.rao@wyzmindz.com"
zoho_password = "CkUT4U8QBzag"

# to_recipients = ["edp@kunhyundai.com","jagadeesan@kunhyundai.com","autosherpa@kunhyundai.com"]
# cc_recipients = ["gmser@kunhyundai.com","crmservice@kunhyundai.com","gmser@kunhyundai.com","tharani@wyzmindz.com"]

to_recipients = ["vengalraooffical@gmail.com"]
cc_recipients = ["vengalrao001star@gmail.caom"]



# Set up the web driver
firefox_options = webdriver.FirefoxOptions()
firefox_options.add_argument("--headless")
firefox_options.add_argument("--window-size=1920x1080")
firefox_options.add_argument("--disable-notifications")
firefox_options.add_argument("--no-sandbox")
firefox_options.add_argument("--verbose")
firefox_options.add_argument("--disable-gpu")
firefox_options.add_argument("--disable-software-rasterizer")

driver = webdriver.Firefox(options=firefox_options)
time.sleep(1)

driver.get(url)
admin = driver.find_element(By.XPATH, "//*[@title='Admin']")
admin.click()

UserName = driver.find_element(By.XPATH, "//*[@title='Enter UserName']")
PassWord = driver.find_element(By.XPATH, "//*[@title='Enter Password']")
Login = driver.find_element(By.XPATH, "//*[@title='Login']")

UserName.send_keys(user_name)
PassWord.send_keys(password)
Login.click()

try:
    admin2 = driver.find_element(By.XPATH, "//*[@title='Admin']")
    admin2.click()
except:
    pass

smr_dashboard = driver.find_element(By.XPATH, "//li[1]")
smr_dashboard.click()
smr_live = driver.find_element(By.CSS_SELECTOR, ".nav.child_menu li:first-child")
smr_live.click()
time.sleep(2)

Data_avail = driver.find_element(By.ID, "BoxLiveDataAvail").text

CallScoreCard = driver.find_element(By.XPATH, "//span[@id='tblFordLocationSMRLive_stl']/button")
driver.execute_script("arguments[0].click();", CallScoreCard)

WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.ID, "tblFordLocationSMRLive"))
)

table_head = driver.find_elements(By.XPATH, "//*[@id='tblFordLocationSMRLive']//thead//tr//th")
headers = [th.find_element(By.TAG_NAME, "div").get_attribute("innerText").strip() for th in table_head]
t_body = driver.find_elements(By.XPATH, "//*[@id='tblFordLocationSMRLive']//tbody//tr")

data = []
for row in t_body:
    values = row.find_elements(By.TAG_NAME, "td")
    val = [v.text for v in values]
    data.append(val)

df = pd.DataFrame(data, columns=headers)

# Quiting the driver
driver.quit()

df['FreshBookings'] = df['FreshBookings'].astype(int)
df['Calls'] = df['Calls'].astype(int)

# Create top and bottom performer tables
df_top_html = df.nlargest(10, 'FreshBookings').to_html(index=False)
df_bottom_html = df.nsmallest(10, 'FreshBookings').to_html(index=False)

# Excel file
excel_filename = r'C:\Users\W2632\Evening mails\Call Performance.xlsx'
df.to_excel(excel_filename, index=False)

# Combine for sending
all_recipients = to_recipients + cc_recipients

msg = MIMEMultipart()
current_date = datetime.datetime.now().strftime("%d-%m-%Y")
msg["From"] = zoho_email
msg["To"] = ", ".join(to_recipients)
msg["CC"] = ", ".join(cc_recipients)
msg["Subject"] = f'Call Performance Report - {client_name} - {current_date}'

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
            text-align: center;
            padding: 8px;
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
    <p>Regards,<br>
    <b>Team AUTOSherpas</b></p>
</body>
</html>
"""

msg.attach(MIMEText(html, "html"))

# Attach Excel File
with open(excel_filename, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header(
        "Content-Disposition",
        f"attachment; filename={os.path.basename(excel_filename)}",
    )
    msg.attach(part)

# Send Email
try:
    server = smtplib.SMTP_SSL(zoho_smtp, zoho_port)
    server.login(zoho_email, zoho_password)
    server.sendmail(zoho_email, all_recipients, msg.as_string())
    server.quit()
    print(f"{client_name} SMR Email sent successfully!")
except Exception as e:
    print(f"{client_name} SMR Error sending email: {e}")

# Clean up
os.remove(excel_filename)
