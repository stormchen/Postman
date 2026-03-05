"""
測試郵件發送腳本

發送一封簡單的測試郵件給自己，確認SMTP設定是否正常
"""

import smtplib
from email.message import EmailMessage
from dotenv import load_dotenv
import os

def send_test_email():
    """發送測試郵件"""

    # 載入環境變數
    load_dotenv()

    # SMTP 設定
    smtp_server = os.getenv('SMTP_SERVER')
    smtp_port = os.getenv('SMTP_PORT')
    sender_email = os.getenv('SENDER_EMAIL')
    recipient_email = os.getenv('TEAMS_CHANNEL_EMAIL')  # 發送給自己

    print("📧 準備發送測試郵件...")
    print(f"SMTP 伺服器：{smtp_server}:{smtp_port}")
    print(f"寄件人：{sender_email}")
    print(f"收件人：{recipient_email}")
    print()

    # 建立郵件
    msg = EmailMessage()
    msg['Subject'] = '測試郵件 - 確認SMTP設定正常'
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg.set_content('''這是一封測試郵件，用於確認SMTP設定是否正常工作。

如果您收到此郵件，表示：
✅ SMTP 伺服器連線正常
✅ 郵件發送功能正常
✅ 您的郵箱設定正確

此郵件由自動化測試腳本發送，請勿回覆。
''')

    try:
        # 建立SMTP連線
        print("🔄 連接到SMTP伺服器...")
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        print("✓ SMTP連線成功")

        # 發送郵件
        print("📤 發送郵件中...")
        server.send_message(msg)
        server.quit()

        print("✅ 測試郵件已成功發送！")
        print(f"📬 請檢查您的郵箱：{recipient_email}")
        print("如果收到郵件，表示SMTP設定完全正常。")

    except Exception as e:
        print(f"❌ 發送測試郵件失敗：{e}")
        print("請檢查SMTP設定或網路連線。")

if __name__ == "__main__":
    send_test_email()