#!/usr/bin/python3
# -*- coding: utf8 -*-
"""
说明: 
- 此脚本使用Selenium自动登录Microsoft账号。
- 登录成功后，导航到指定的OAuth URL以获取授权码。
- 使用 OneDriveUploader -a 处理授权，生成 auth.json。
- 将 auth.json 文件重命名为微软E5帐号的前缀部分（即 @ 之前的内容）+ `.json`。
- 使用 OneDriveUploader -c 上传重命名后的授权文件到 OneDrive 的目录 `wwwwww`。
"""
import os
import time
import random
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from urllib.parse import urlparse, parse_qs

# --- Optional Notification Setup ---
try:
    from sendNotify import send
except ImportError:
    print("通知文件 sendNotify.py 未找到，将仅打印到控制台。")
    def send(title, content):
        print(f"--- {title} ---")
        print(content)
        print("--- End Notification ---")
# --- End Notification Setup ---

List = []  # To store output messages

# --- Configuration ---
LOGIN_URL = 'https://admin.microsoft.com/'  # Use admin center for login context initially
OAUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=78d4dc35-7e46-42c6-9023-2d39314433a5&response_type=code&redirect_uri=http://localhost/onedrive-login&response_mode=query&scope=offline_access%20User.Read%20Files.ReadWrite.All"
REDIRECT_URI_START = "http://localhost/onedrive-login"
ONEDRIVE_UPLOADER = "/usr/local/bin/OneDriveUploader"
ONEDRIVE_AUTH_CONFIG = "auth1106.json"

# --- Helper Function ---
def setup_onedrive_uploader():
    """Download and configure OneDriveUploader and auth1106.json."""
    try:
        # Download OneDriveUploader
        List.append("正在下载 OneDriveUploader 到 /usr/local/bin/...")
        subprocess.run(
            ["wget", "https://raw.githubusercontent.com/MoeClub/OneList/master/OneDriveUploader/amd64/linux/OneDriveUploader", "-P", "/usr/local/bin/"],
            check=True
        )
        # Set execute permissions
        subprocess.run(["chmod", "+x", ONEDRIVE_UPLOADER], check=True)
        List.append("成功设置 OneDriveUploader 执行权限。")

        # Download auth1106.json
        List.append("正在下载 auth1106.json 文件...")
        subprocess.run(
            ["wget", "-O", ONEDRIVE_AUTH_CONFIG, "https://raw.githubusercontent.com/yghhbbuy/vvvioui/refs/heads/main/.github/workflows/auth.json"],
            check=True
        )
        List.append("成功下载 auth1106.json 文件。")
    except subprocess.CalledProcessError as e:
        List.append(f"!! 错误: 配置 OneDriveUploader 或下载 auth1106.json 时出错: {e}")
        exit(1)

def get_webdriver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36")
    options.binary_location = "/usr/bin/chromium-browser"

    try:
        driver = webdriver.Chrome(options=options)
        List.append("  - WebDriver 初始化成功 (使用 /usr/bin/chromium-browser)。")
        return driver
    except Exception as e:
        List.append(f"!! 错误：无法初始化 WebDriver: {e}")
        return None

def get_oauth_code(username, password):
    """Logs into Microsoft account and attempts to capture OAuth redirect URL."""
    List.append(f"开始处理账号: {username}")
    driver = get_webdriver()
    if not driver:
        List.append(f"!! 处理失败: {username} (WebDriver 初始化失败)")
        return

    try:
        # --- Login Steps ---
        driver.get(LOGIN_URL)
        wait = WebDriverWait(driver, 60)

        # Step 1: Enter Email
        email_field = wait.until(EC.visibility_of_element_located((By.ID, "i0116")))
        email_field.send_keys(username)
        next_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        driver.execute_script("arguments[0].click();", next_button)
        time.sleep(random.uniform(4, 6))

        # Step 2: Enter Password
        password_field = wait.until(EC.visibility_of_element_located((By.ID, "i0118")))
        password_field.send_keys(password)
        signin_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        driver.execute_script("arguments[0].click();", signin_button)

        # Step 3: Handle "Stay signed in?" (KMSI)
        try:
            kmsi_button_no = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.ID, "idBtn_Back"))
            )
            driver.execute_script("arguments[0].click();", kmsi_button_no)
        except TimeoutException:
            pass

        # Step 4: Navigate to OAuth URL
        driver.get(OAUTH_URL)
        time.sleep(3)
        WebDriverWait(driver, 30).until(lambda d: REDIRECT_URI_START in d.current_url)
        redirected_url = driver.current_url
        handle_one_drive_auth(username, redirected_url)
    except Exception as e:
        List.append(f"!! 处理账号 {username} 时发生意外错误: {e}")
    finally:
        driver.quit()

def handle_one_drive_auth(username, redirect_url):
    """Handles OneDriveUploader -a with the redirect URL and renames auth.json."""
    try:
        # Extract the prefix from the email (e.g., "example@domain.com" -> "example")
        prefix = username.split('@')[0]

        # Run the OneDriveUploader -a command
        List.append(f"  - 使用 OneDriveUploader 处理授权: {redirect_url}")
        auth_command = [ONEDRIVE_UPLOADER, "-a", redirect_url]
        result = subprocess.run(auth_command, capture_output=True, text=True)

        if result.returncode == 0:
            List.append("  - 授权成功，auth.json 文件已生成。")
            # Rename auth.json to {prefix}.json
            new_auth_file = f"{prefix}.json"
            os.rename("auth.json", new_auth_file)
            List.append(f"  - 已将 auth.json 重命名为 {new_auth_file}")

            # Upload {prefix}.json to OneDrive
            upload_to_onedrive(new_auth_file)
        else:
            List.append(f"!! 授权失败: {result.stderr}")
    except FileNotFoundError:
        List.append("!! 错误: auth.json 文件未生成，可能授权失败。")
    except Exception as e:
        List.append(f"!! 处理授权时发生意外错误: {e}")

def upload_to_onedrive(file_name):
    """Uploads the given file to OneDrive using OneDriveUploader."""
    try:
        List.append(f"  - 正在将 {file_name} 上传到 OneDrive 的目录 'wwwwww'...")
        upload_command = [ONEDRIVE_UPLOADER, "-c", ONEDRIVE_AUTH_CONFIG, "-s", file_name, "-r", "wwwwww"]
        result = subprocess.run(upload_command, capture_output=True, text=True)

        if result.returncode == 0:
            List.append(f"  - 成功上传文件到 OneDrive 的目录 'wwwwww': {file_name}")
        else:
            List.append(f"!! 上传到 OneDrive 目录 'wwwwww' 失败: {result.stderr}")
    except Exception as e:
        List.append(f"!! 上传文件到 OneDrive 目录 'wwwwww' 时发生意外错误: {e}")

# --- Main Function ---
if __name__ == "__main__":
    setup_onedrive_uploader()

    accounts = os.getenv('MS_E5_ACCOUNTS', '').split('&')
    if not accounts or accounts == ['']:
        List.append("!! 错误: 未找到环境变量 MS_E5_ACCOUNTS。")
        send("MS OAuth 登录自动化", '\n'.join(List))
        exit(1)

    for account in accounts:
        try:
            username, password = account.split('-')
            get_oauth_code(username, password)
        except ValueError:
            List.append(f"!! 错误: 无效账号配置: {account} (应为 email-password 格式)")

    send("MS OAuth 登录自动化", '\n'.join(List))
