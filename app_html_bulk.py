# ============================================================
#   IMPORT
# ============================================================
import os
import time
import pandas as pd
import pythoncom
import win32com.client as win32

from flask import Flask, render_template, request

# ============================================================
#   CONFIG
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

EMAIL_TEMPLATE_PATH = os.path.join(BASE_DIR, "email_template.html")
EXCEL_DIR = os.path.join(BASE_DIR, "excel_uploads")

os.makedirs(EXCEL_DIR, exist_ok=True)

SENDER_MAILBOX = "customerservice@ocbs.com.vn"   # đổi nếu cần


# ============================================================
#   LOAD HTML TEMPLATE
# ============================================================
def load_email_template():
    try:
        with open(EMAIL_TEMPLATE_PATH, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        print("❌ Không đọc được email_template.html:", e)
        return ""


# ============================================================
#   RENDER HTML TỪ 1 HÀNG EXCEL
# ============================================================
def render_html_from_row(base_html, row):
    html = base_html
    for col_name, value in row.items():
        placeholder = "{{" + str(col_name).strip() + "}}"
        val = "" if pd.isna(value) else str(value).strip()
        html = html.replace(placeholder, val)
    return html


# ============================================================
#   GỬI EMAIL HTML BẰNG OUTLOOK — HỖ TRỢ CC
# ============================================================
def send_email_html_only(to_email, subject, content_html, cc_emails=""):
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.SentOnBehalfOfName = SENDER_MAILBOX
        mail.To = to_email
        mail.Subject = subject

        # ====== SET CC ======
        if cc_emails:
            mail.CC = cc_emails
            print("✔ CC:", cc_emails)

        # ====== ATTACH LOGO ======
        logo_path = r"C:\Python\HTML_BULK_MAIL\logo.png"

        print(">>> LOGO PATH:", logo_path)
        print(">>> EXISTS:", os.path.exists(logo_path))

        if os.path.exists(logo_path):
            attachment = mail.Attachments.Add(logo_path)
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                "ocbslogo"
            )

        # ====== SET HTML BODY ======
        mail.HTMLBody = content_html

        mail.Send()
        print("✔ Đã gửi có logo:", to_email)

    except Exception as e:
        print("❌ Lỗi gửi email:", e)

    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass


# ============================================================
#   FLASK APP
# ============================================================
app = Flask(__name__)


@app.route("/bulk", methods=["GET", "POST"])
def bulk():

    if request.method == "POST":

        subject = request.form.get("subject", "OCBS – Thông báo")
        cc_list = request.form.get("cc_list", "").strip()   # <<== THÊM CC

        excel_file = request.files.get("excel")
        if not excel_file:
            return "❌ Bạn chưa chọn file Excel", 400

        excel_path = os.path.join(EXCEL_DIR, excel_file.filename)
        excel_file.save(excel_path)

        try:
            df = pd.read_excel(excel_path, dtype=str).fillna("")
        except:
            return "❌ Không đọc được file Excel", 400

        df.columns = df.columns.str.strip()

        if "EMAIL" not in df.columns:
            return "❌ Excel bắt buộc có cột EMAIL", 400

        base_html = load_email_template()
        if not base_html:
            return "❌ Không đọc được email_template.html", 500

        sent_count = 0

        for _, row in df.iterrows():
            email = str(row["EMAIL"]).strip()
            if not email:
                continue

            content_html = render_html_from_row(base_html, row)

            # Truyền CC vào đây
            send_email_html_only(email, subject, content_html, cc_list)

            sent_count += 1
            time.sleep(1)

        return f"✔ Đã gửi {sent_count} email (CC: {cc_list})!"

    return render_template("bulk.html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5009, debug=False)
