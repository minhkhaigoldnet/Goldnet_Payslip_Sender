import streamlit as st
import pandas as pd
from fpdf import FPDF
from PyPDF2 import PdfReader, PdfWriter
import smtplib
import ssl
from email.message import EmailMessage
import os
import shutil


def create_pdf(employee, logo_path, output_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.image(logo_path, x=10, y=8, w=33)
    pdf.set_font("Arial", size=12)
    pdf.ln(40)
    pdf.cell(200, 10, txt=f"Phiếu lương tháng {employee['Tháng']}", ln=True, align='C')
    pdf.ln(10)
    for col in employee.index:
        pdf.cell(50, 10, txt=f"{col}", ln=0)
        pdf.cell(100, 10, txt=f": {employee[col]}", ln=1)
    pdf.output(output_path)


def encrypt_pdf(input_path, output_path, password):
    reader = PdfReader(input_path)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(password)
    with open(output_path, "wb") as f:
        writer.write(f)


def send_email(receiver_email, subject, body, attachment_path, sender_email, app_password):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender_email
    msg["To"] = receiver_email
    msg.set_content(body)
    with open(attachment_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
    msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=file_name)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, app_password)
        server.send_message(msg)


st.set_page_config(page_title="Gửi Phiếu Lương - GOLDNET", layout="centered")
st.title("📤 Gửi Phiếu Lương TỰ ĐỘNG")
st.write("Ứng dụng giúp gửi phiếu lương cá nhân hóa qua Gmail. PDF được bảo mật bằng mã nhân viên.")

file = st.file_uploader("📁 Tải lên file lương (.xlsx)", type=["xlsx"])
logo = st.file_uploader("🖼️ Tải lên logo công ty (.png)", type=["png"])
sender_email = st.text_input("📧 Gmail người gửi", placeholder="yourname@gmail.com")
app_password = st.text_input("🔑 Mật khẩu ứng dụng Gmail", type="password")
email_body = st.text_area("📄 Nội dung email", value="Xin chào, đây là phiếu lương bảo mật của bạn. Mật khẩu mở file là mã nhân viên.")

if st.button("🚀 Gửi phiếu lương"):
    if not file or not logo:
        st.warning("Vui lòng tải lên file lương và logo.")
    else:
        df = pd.read_excel(file)
        required_columns = {'MaNV', 'Email', 'HoTen', 'Tháng'}
        if not required_columns.issubset(set(df.columns)):
            st.error("❌ File Excel thiếu một số cột bắt buộc: MaNV, Email, HoTen, Tháng")
            st.stop()

        with open("temp_logo.png", "wb") as f:
            f.write(logo.read())

        temp_folder = "temp_pdfs"
        os.makedirs(temp_folder, exist_ok=True)

        for index, row in df.iterrows():
            filename = f"{row['MaNV']}_luong.pdf"
            temp_pdf = os.path.join(temp_folder, f"temp_{filename}")
            final_pdf = os.path.join(temp_folder, filename)
            create_pdf(row, logo_path="temp_logo.png", output_path=temp_pdf)
            encrypt_pdf(temp_pdf, final_pdf, password=str(row["MaNV"]))
            try:
                send_email(
                    receiver_email=row["Email"],
                    subject="Phiếu lương tháng " + str(row["Tháng"]),
                    body=email_body,
                    attachment_path=final_pdf,
                    sender_email=sender_email,
                    app_password=app_password,
                )
                st.success(f"✅ Đã gửi cho {row['HoTen']} ({row['Email']})")
            except Exception as e:
                st.error(f"❌ Lỗi gửi cho {row['HoTen']}: {e}")

        shutil.rmtree(temp_folder)
        os.remove("temp_logo.png")

