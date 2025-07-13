import requests
# import json # Không được sử dụng trực tiếp, requests đã xử lý
import os
import random
from datetime import datetime, timedelta, timezone
import feedparser

# Lấy thông tin nhạy cảm từ GitHub Secrets (biến môi trường)
TENANT_ID = os.getenv('TENANT_ID')
CLIENT_ID = os.getenv('CLIENT_ID')
CLIENT_SECRET = os.getenv('CLIENT_SECRET')

# --- Thông tin repo GitHub ---
GITHUB_REPO_OWNER = 'ngqtuan'
GITHUB_REPO_NAME = 'renew_e5_images'
GITHUB_IMAGE_FOLDER = 'e5-images'

# Kiểm tra xem các secret đã được thiết lập chưa
if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
    print("❌ Lỗi: Vui lòng thiết lập các secret TENANT_ID, CLIENT_ID, CLIENT_SECRET trong cài đặt của repo GitHub.")
    exit()


def get_token():
    url = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()['access_token']

def get_users(token):
    headers = {'Authorization': f'Bearer {token}'}
    r = requests.get('https://graph.microsoft.com/v1.0/users', headers=headers)
    r.raise_for_status()
    users = r.json().get('value', [])
    print("✅ Danh sách user:")
    for u in users:
        print(" 👤", u['userPrincipalName'])
    return users

def get_calendar(token, user_id, email):
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/calendar/events'
    headers = {'Authorization': f'Bearer {token}'}
    r = requests.get(url, headers=headers)
    print(f"📅 Lịch của {email} – Status: {r.status_code}")
    if r.status_code == 200:
        events = r.json().get('value', [])
        print(f"📆 Số sự kiện: {len(events)}")
    else:
        print(r.text)

def create_daily_event(token, user_id):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/events"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    now = datetime.now(timezone.utc)
    start_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
    end_time = start_time + timedelta(minutes=30)
    payload = {
        "subject": "📌 Daily Auto Event",
        "body": {
            "contentType": "HTML",
            "content": "Tự động tạo để duy trì hoạt động lịch mỗi ngày."
        },
        "start": {
            "dateTime": start_time.isoformat(),
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": end_time.isoformat(),
            "timeZone": "UTC"
        }
    }
    r = requests.post(url, headers=headers, json=payload)
    print(f"📆 Tạo sự kiện lịch – Status: {r.status_code}")
    if r.status_code not in [200, 201]:
        print(r.text)

def get_news_rss():
    feed = feedparser.parse("https://vnexpress.net/rss/tin-moi-nhat.rss")
    news_list = []
    for entry in feed.entries[:5]:
        news_list.append(f"- {entry.title}")
    return "\n".join(news_list)

def generate_copilot_mock():
    samples = [
        "🧠 Hôm nay thời tiết tại Hà Nội nắng ráo, nhiệt độ cao nhất 35°C.",
        "📈 VN-Index tăng nhẹ, nhà đầu tư nước ngoài mua ròng mạnh.",
        "💡 Mẹo: Dùng Ctrl+Shift+V để dán văn bản không định dạng.",
        "🧘 Copilot gợi ý: Thử bài thở 4-7-8 để giảm căng thẳng.",
        "📅 Hôm nay có cuộc họp vào lúc 10h, đừng quên chuẩn bị.",
        "🌐 Trợ lý Copilot sẵn sàng giúp bạn viết email hoặc tạo bảng."
    ]
    return random.choice(samples)

def ensure_folder_exists(token, user_id, folder_name="E5Auto"):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root/children"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    data = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename"
    }
    requests.post(url, headers=headers, json=data)

def create_word_report(token, user_id, recipient, content):
    ensure_folder_exists(token, user_id, "CopilotReports")
    filename = f"report_{recipient.replace('@','_')}.docx"
    # Tạo file tạm thời để upload
    temp_file_path = os.path.join(os.getcwd(), filename)
    with open(temp_file_path, "w", encoding="utf-8") as f:
        f.write(content)
    
    upload_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/CopilotReports/{filename}:/content"
    with open(temp_file_path, "rb") as f:
        r = requests.put(upload_url, headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        }, data=f)
    print(f"📄 Upload Word – Status: {r.status_code}")
    os.remove(temp_file_path) # Xóa file tạm sau khi upload

def send_personalized_mails(token, sender_email, recipient_list, user_id):
    subject = "📌 MS365 – Bản tin & phản hồi Copilot"
    url = f'https://graph.microsoft.com/v1.0/users/{sender_email}/sendMail'
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    for recipient in recipient_list:
        today = datetime.now(timezone.utc).strftime('%Y-%m-%d')
        news_content = get_news_rss()
        copilot_msg = generate_copilot_mock()
        body = f"""📢 Bản tin cá nhân hóa ngày {today}:
{news_content}

🤖 Copilot nói:
{copilot_msg}

✅ Email tự động để duy trì hoạt động tài khoản."""

        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": recipient}}]
            }
        }
        print(f"📧 {sender_email} gửi đến: {recipient}")
        r = requests.post(url, headers=headers, json=payload)
        print(f"📨 Trạng thái: {r.status_code}")
        if r.status_code != 202:
            print(r.text)

        ensure_folder_exists(token, user_id, "CopilotChat")
        filename = f"copilot_{recipient.replace('@','_')}.txt"
        temp_file_path = os.path.join(os.getcwd(), filename)
        with open(temp_file_path, "w", encoding="utf-8") as f:
            f.write(body)
        
        upload_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/CopilotChat/{filename}:/content"
        with open(temp_file_path, "rb") as f:
            upload_res = requests.put(upload_url, headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "text/plain"
            }, data=f)
        print(f"☁️ Upload Copilot – Status: {upload_res.status_code}")
        os.remove(temp_file_path)

        create_word_report(token, user_id, recipient, body)

def check_onedrive_ready(token, user_id):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive"
    headers = { "Authorization": f"Bearer {token}" }
    r = requests.get(url, headers=headers)
    return r.status_code == 200

def upload_random_images(token, user_id):
    api_url = f'https://api.github.com/repos/{GITHUB_REPO_OWNER}/{GITHUB_REPO_NAME}/contents/{GITHUB_IMAGE_FOLDER}'
    try:
        res = requests.get(api_url)
        res.raise_for_status()
        files_data = res.json()
        
        image_files = [
            f for f in files_data 
            if isinstance(f, dict) and f.get('type') == 'file' and f['name'].lower().endswith(('.jpg', '.jpeg', '.png', '.gif'))
        ]

        if not image_files:
            print(f"❌ Không tìm thấy file ảnh nào trong thư mục '{GITHUB_IMAGE_FOLDER}' trên GitHub.")
            return

    except requests.exceptions.RequestException as e:
        print(f"❌ Lỗi khi truy cập GitHub API: {e}")
        print(f"   Vui lòng kiểm tra lại GITHUB_REPO_OWNER, GITHUB_REPO_NAME và đảm bảo repo là public.")
        return

    ensure_folder_exists(token, user_id, "E5Auto")
    
    selected_files = random.sample(image_files, min(3, len(image_files)))
    
    for file_info in selected_files:
        filename = file_info['name']
        download_url = file_info['download_url']
        
        print(f"🖼️ Đang tải file: {filename} từ GitHub...")
        
        try:
            image_response = requests.get(download_url)
            image_response.raise_for_status()
            image_content = image_response.content

            upload_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/drive/root:/E5Auto/{filename}:/content"
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/octet-stream'
            }
            
            r = requests.put(upload_url, headers=headers, data=image_content)
            
            print(f"🖼️ Upload {filename} – Status: {r.status_code}")
            if r.status_code not in [200, 201]:
                print(r.text)

        except requests.exceptions.RequestException as e:
            print(f"❌ Lỗi khi tải hoặc upload file {filename}: {e}")

def create_daily_task(token, user_id):
    url = f"https://graph.microsoft.com/v1.0/users/{user_id}/todo/lists"
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(url, headers=headers)
    if r.status_code != 200:
        print("❌ Không lấy được danh sách To Do.")
        return
    lists = r.json().get('value', [])
    if not lists:
        print("⚠️ Không có danh sách To Do nào.")
        return
    
    default_list_id = lists[0]['id']

    task_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/todo/lists/{default_list_id}/tasks"
    headers["Content-Type"] = "application/json"
    payload = {
        "title": "📝 Daily Reminder",
        "body": {"content": "Tự động tạo để duy trì hoạt động hàng ngày.", "contentType": "text"},
        "dueDateTime": {
            "dateTime": (datetime.now(timezone.utc) + timedelta(days=1)).isoformat(),
            "timeZone": "UTC"
        }
    }
    r = requests.post(task_url, headers=headers, json=payload)
    print(f"📋 Tạo task – Status: {r.status_code}")
    if r.status_code not in [200, 201]:
        print(r.text)

if __name__ == '__main__':
    try:
        token = get_token()
        users = get_users(token)
        for sender in users:
            sender_email = sender['userPrincipalName']
            sender_id = sender['id']
            print(f"\n🔄 Đang xử lý cho: {sender_email}")
            
            # 💡 FIX: Thêm lại các lệnh gọi hàm đã bị thiếu
            get_calendar(token, sender_id, sender_email)
            create_daily_event(token, sender_id)
            create_daily_task(token, sender_id)

            # Gửi mail cho những người khác
            others = [u['userPrincipalName'] for u in users if u['userPrincipalName'] != sender_email]
            if others:
                # Chọn ngẫu nhiên số người nhận từ 2 đến 5
                num_recipients = random.randint(2, 5)
                recipients = random.sample(others, min(num_recipients, len(others)))
                send_personalized_mails(token, sender_email, recipients, sender_id)

            # Tải ảnh lên OneDrive
            if check_onedrive_ready(token, sender_id):
                upload_random_images(token, sender_id)
            else:
                print("⚠️ OneDrive của người dùng này chưa sẵn sàng.")

    except Exception as e:
        print("❌ Lỗi toàn cục:", e)