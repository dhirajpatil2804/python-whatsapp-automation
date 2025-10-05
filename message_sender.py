import pandas as pd
import time
import urllib.parse
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# STEP 1: Load numbers from Excel
file_path = "contacts.xlsx"
df = pd.read_excel(file_path)
df.columns = df.columns.str.strip().str.lower()

possible_names = ['phone', 'phone number', 'contact', 'mobile', 'number']
phone_col = next((name for name in possible_names if name in df.columns), None)
if not phone_col:
    raise ValueError(f"No phone number column found! Please name it one of {possible_names}")

df = df[df[phone_col].astype(str).str.isdigit() & (df[phone_col].astype(str).str.len() == 10)]
numbers = [str(num) for num in df[phone_col]]
print(f"✅ Loaded {len(numbers)} valid numbers from Excel.")

# STEP 2: Setup Chrome driver options without user profile to avoid crash
options = webdriver.ChromeOptions()
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
# Uncomment below for headless run (no browser window)
# options.add_argument('--headless=new')

service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

# Open WhatsApp Web
driver.get("https://web.whatsapp.com")
input("📲 Scan QR code if not already logged in, then press Enter here...")

# STEP 3: Send message text
message = """
Disha Computer Institute Dehuroad Branch
🌟✨ DIWALI DHAMAKA OFFER ✨🌟
🎉 Celebrate this Diwali with Double Happiness! 🎉

📚 Join 1 Course & Get 1 Absolutely FREE!

🎓 Choose from our most popular courses:
✅ MS Excel
✅ TallyPrime
✅ Photoshop
✅ ChatGPT
✅ C++
✅ Web Design
& many more!

💥 Light up your future with new skills this Diwali!
👉 Limited Time Offer – Enroll Now!
Contact : 7620825396 
"""
encoded_msg = urllib.parse.quote(message)

for num in numbers:
    print(f"📨 Sending to {num}...")
    try:
        url = f"https://web.whatsapp.com/send?phone=91{num}&text={encoded_msg}"
        driver.get(url)

        # Wait for message box to load
        message_box = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )

        # Try clicking the send button
        try:
            send_btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-testid='send'][data-icon='send']"))
            )
            send_btn.click()
        except Exception:
            # If click fails, fallback to pressing Enter key
            message_box.send_keys(Keys.ENTER)

        print(f"✅ Sent to {num}")
        time.sleep(2)
    except Exception as e:
        print(f"❌ Failed to send to {num}: {e}")

print("🎯 Done sending all messages.")
driver.quit()
