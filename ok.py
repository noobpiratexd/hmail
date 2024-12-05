import threading
import time
import telebot
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Telegram bot token
TELEGRAM_BOT_TOKEN = 'YOUR_TELEGRAM_BOT_TOKEN'
bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)

# Path to your ChromeDriver
CHROME_DRIVER_PATH = '/path/to/chromedriver'

def check_inbox(email, password):
    """Logs in to the email account, retrieves inbox summary, and searches for specific keywords."""
    driver = webdriver.Chrome(executable_path=CHROME_DRIVER_PATH)
    wait = WebDriverWait(driver, 20)
    try:
        # Open Office.com login page
        driver.get("https://outlook.office.com/")
        
        # Enter email
        wait.until(EC.presence_of_element_located((By.NAME, "loginfmt"))).send_keys(email)
        driver.find_element(By.NAME, "loginfmt").send_keys(Keys.RETURN)
        time.sleep(2)
        
        # Enter password
        wait.until(EC.presence_of_element_located((By.NAME, "passwd"))).send_keys(password)
        driver.find_element(By.NAME, "passwd").send_keys(Keys.RETURN)
        time.sleep(2)
        
        # Handle "Stay signed in?" prompt if it appears
        try:
            stay_signed_in = wait.until(EC.presence_of_element_located((By.ID, "idSIButton9")))
            stay_signed_in.click()
        except:
            pass

        # Wait until inbox loads
        wait.until(EC.presence_of_element_located((By.ID, "app")))

        # Fetch inbox summary
        try:
            email_count = len(driver.find_elements(By.XPATH, "//div[@role='row']"))
            last_email = driver.find_element(By.XPATH, "//div[@role='row'][1]")
            last_subject = last_email.find_element(By.XPATH, ".//span[@class='subject']").text
            last_sender = last_email.find_element(By.XPATH, ".//span[@class='sender']").text
            last_date = last_email.find_element(By.XPATH, ".//span[@class='date']").text
        except:
            email_count = 0
            last_subject = "N/A"
            last_sender = "N/A"
            last_date = "N/A"

        # Prepare inbox summary
        inbox_summary = f"""
━━━━━━INBOX━━━━━━━━
📧 {email}:{password}

👤 Name: Example User
🌍 Country: Example Country

{last_sender}
- 📬 Emails: {email_count}
- 📅 Last Received: {last_date}
- ✉️ Last Subject: {last_subject}

━━━
        """

        # Search for "Subscription" keyword in emails
        search_box = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@aria-label='Search']")))
        search_box.send_keys("Subscription")
        search_box.send_keys(Keys.RETURN)
        
        time.sleep(5)

        emails = driver.find_elements(By.XPATH, "//div[@role='row']")
        email_details = []
        for email_row in emails:
            try:
                subject = email_row.find_element(By.XPATH, ".//span[@class='subject']").text
                sender = email_row.find_element(By.XPATH, ".//span[@class='sender']").text
                date = email_row.find_element(By.XPATH, ".//span[@class='date']").text
                email_details.append(f"- 📬 Subject: {subject}\n- 📨 Sender: {sender}\n- 📅 Date: {date}")
            except:
                continue

        # Prepare search results
        if email_details:
            search_results = f"*Emails matching 'Subscription':*\n{chr(10).join(email_details)}"
        else:
            search_results = "*No emails found matching 'Subscription'.*"

        # Combine results
        return inbox_summary + search_results + "\n━━━━━━━━━━━━━━"

    except Exception as e:
        return f"Error checking {email}: {e}"
    finally:
        driver.quit()


def process_accounts(file_path, chat_id):
    """Reads email-password pairs from a file and checks each inbox."""
    with open(file_path, 'r') as file:
        lines = file.readlines()

    def worker(email, password):
        result = check_inbox(email, password)
        bot.send_message(chat_id, result)

    threads = []
    for line in lines:
        try:
            email, password = line.strip().split(':')
            t = threading.Thread(target=worker, args=(email, password))
            threads.append(t)
            t.start()
        except ValueError:
            bot.send_message(chat_id, f"Invalid format in line: {line.strip()}")

    for t in threads:
        t.join()


@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    """Sends a welcome message."""
    bot.reply_to(message, "Welcome! Send me a file with email:password pairs to check inbox details.")


@bot.message_handler(content_types=['document'])
def handle_file(message):
    """Handles the uploaded file and processes it."""
    try:
        file_info = bot.get_file(message.document.file_id)
        file_path = bot.download_file(file_info.file_path)
        local_file_path = f"./{message.document.file_name}"
        with open(local_file_path, 'wb') as f:
            f.write(file_path)
        
        bot.reply_to(message, "Processing your file. This may take a while.")
        process_accounts(local_file_path, message.chat.id)
    except Exception as e:
        bot.reply_to(message, f"An error occurred: {e}")


if __name__ == '__main__':
    print("Bot is running...")
    bot.polling()
