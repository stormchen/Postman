"""
SMTP 連線測試工具

測試 SMTP 伺服器是否可連線以及基本通訊功能
"""

import smtplib
from dotenv import load_dotenv
import os
import sys

def test_smtp_connection():
    """測試 SMTP 連線"""
    
    print("=" * 60)
    print("SMTP 連線測試工具")
    print("=" * 60)
    print()
    
    # 硬編碼 SMTP 設定 (用於測試)
    smtp_server = "10.210.2.21"
    smtp_port = "25"
    sender_email = "storm.chen@gigacomputing.com"
    teams_channel_email = "ac75feb8.gigacomputing.com@apac.teams.ms"
    
    # 顯示讀取的設定
    print("📋 硬編碼的 SMTP 設定：")
    print(f"  SMTP_SERVER：{smtp_server}")
    print(f"  SMTP_PORT：{smtp_port}")
    print(f"  SENDER_EMAIL：{sender_email}")
    print(f"  TEAMS_CHANNEL_EMAIL：{teams_channel_email}")
    print()
    
    # 嘗試連接 SMTP 伺服器
    print("🔄 嘗試連接 SMTP 伺服器...")
    try:
        # 建立 SMTP 連線
        server = smtplib.SMTP(smtp_server, int(smtp_port), timeout=10)
        print(f"✓ 成功連接到 {smtp_server}:{smtp_port}")
        
        # 嘗試 STARTTLS
        print("🔄 嘗試啟動 TLS 加密...")
        server.starttls()
        print("✓ 成功啟動 TLS")
        
        # 顯示稍微詳細的連線資訊
        print()
        print("✓ SMTP 連線測試成功！")
        print()
        print("✅ 你的 SMTP 伺服器可以正常連接")
        print(f"   伺服器：{smtp_server}:{smtp_port}")
        print(f"   寄件者：{sender_email}")
        print(f"   收件者：{teams_channel_email}")
        
        server.quit()
        return True
    
    except smtplib.SMTPServerDisconnected as e:
        print(f"❌ SMTP 連線中斷：{e}")
        print("   可能原因：伺服器主動關閉連線或超時")
        return False
    
    except smtplib.SMTPAuthenticationError as e:
        print(f"❌ 認證失敗：{e}")
        print("   可能原因：寄件者帳號或密碼不正確")
        return False
    
    except smtplib.SMTPException as e:
        print(f"❌ SMTP 錯誤：{e}")
        return False
    
    except TimeoutError:
        print(f"❌ 連線逾時：無法在規定時間內連接到伺服器")
        print(f"   伺服器：{smtp_server}:{smtp_port}")
        print("   可能原因：")
        print("      1. 伺服器位址或埠號不正確")
        print("      2. 防火牆阻擋了連線")
        print("      3. 伺服器離線或無法連接")
        return False
    
    except ConnectionRefusedError:
        print(f"❌ 連線被拒絕：無法與伺服器建立連線")
        print(f"   伺服器：{smtp_server}:{smtp_port}")
        print("   可能原因：")
        print("      1. 伺服器埠號不正確")
        print("      2. 伺服器未啟動")
        print("      3. 防火牆阻擋了連線")
        return False
    
    except OSError as e:
        print(f"❌ 網路錯誤：{e}")
        print("   可能原因：")
        print("      1. 伺服器位址不存在 (DNS 解析失敗)")
        print("      2. 網路連線有問題")
        print("      3. 防火牆或代理設定問題")
        return False
    
    except Exception as e:
        print(f"❌ 未預期的錯誤：{e}")
        return False
    
    finally:
        print()
        print("=" * 60)

if __name__ == "__main__":
    success = test_smtp_connection()
    sys.exit(0 if success else 1)
