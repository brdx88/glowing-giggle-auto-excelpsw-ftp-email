import os
import win32com.client as win32
from ftplib import FTP
import pandas as pd
from datetime import datetime

log_file = r"path_to_log\testing_excel_process_log.txt"
first_log = True

# Define needs
file_path = r"path_to_excel\test_excel_password.xlsx"
file_name = os.path.basename(file_path)
input_file = os.path.abspath(file_path)  # Use absolute path
password = "securepassword123"

df = pd.read_excel(input_file)
n_rows = len(df)

ftp_host = "192.168.0.0"  # FTP server address
ftp_username = "FTPUSER"  # FTP username
ftp_password = "FTPPASSWORD"  # FTP password
ftp_dir = "/FOLDER/SUB_FOLDER/%Y-%m/"  # Dynamic FTP directory to upload file
dynamic_folder = ""  # Global variable for FTP directory

recipient_email = "recipient@email.com"
cc_recipients = ["recipient2@email.com", "recipient3@email.com"]
email_subject = "Encrypted Excel File Uploaded to File Transfer Protocol"
email_body = """Dear User,<br><br>
The encrypted Excel file has been successfully uploaded to the FTP server <i><strong>{ftp_folder}{file_name}</strong></i> .<br><br>
The data is: {n_rows:,} rows of data.<br><br>
Thank you.
"""

# 00-log message
def log_message(message):
    global first_log
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}"
    print(log_entry)  # Print to console
    with open(log_file, "a") as file:  # Append mode
        if first_log:
            file.write("\n" + "=" * 50 + "\n")  # Add a separator for new sessions
            first_log = False
        file.write(log_entry + "\n")

# 01-Encrypt Excel file
def encrypt_excel_file(input_file, password):
    # Use pywin32 to set a password on the Excel file
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False  # Run Excel in the background

    # Check if the file exists
    if not os.path.exists(input_file):
        log_message(f"Error: The file '{input_file}' does not exist.")
        return None

    workbook = excel.Workbooks.Open(input_file)
    encrypted_file = input_file.replace(".xlsx", "_protected.xlsx")

    workbook.Password = password  # Set the password
    workbook.SaveAs(encrypted_file, FileFormat=51)  # FileFormat=51 for .xlsx files
    workbook.Close()
    excel.Quit()
    log_message(f"Encrypted file saved at '{encrypted_file}'")
    return encrypted_file

# 02-Upload to FTP
def upload_to_ftp(local_file, dynamic_ftp_dir, host, username, password):
    global dynamic_folder       # Declare the variable as global
    
    # Connect to FTP
    ftp = FTP(host)
    ftp.login(username, password)

    # Create dynamic folder based on current month and year
    dynamic_folder = datetime.now().strftime(dynamic_ftp_dir)
    try:
        ftp.cwd(dynamic_folder)
    except Exception:
        ftp.mkd(dynamic_folder)
        ftp.cwd(dynamic_folder)
        log_message(f"Created directory: '{dynamic_folder}'")

    # Upload the file
    with open(local_file, "rb") as file:
        ftp.storbinary(f"STOR {os.path.basename(local_file)}", file)
        log_message(f"File '{local_file}' uploaded to '{dynamic_folder}'.")

    ftp.quit()
    return dynamic_folder

# 03-Send Email with Outlook
def send_email_with_outlook(recipient, subject, body, attachment_path=None, cc=None):
    try:
        # Create an instance of Outlook
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # Create a new email

        # Set email properties
        mail.To = recipient
        mail.Subject = subject
        
        # Add CC recipients
        if cc:
            mail.CC = ";".join(cc)  # Combine multiple CC emails with semicolons

        # mail.Body = body
        disclaimer = """
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="2" color="#0F243E"><span lang="en-US" style="font-size:9pt"><i>This email was generated automatically.</i></span></font></span></font></div>
        """

        signature = """
        <div name="divtagdefaultwrapper" style="font-family:Calibri,Arial,Helvetica,sans-serif; font-size:; margin:0">
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="2"><span lang="en-US" style="font-size:10pt"><b>Best regards,</b></span></font></span></font></div>
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font color="#17365D"><b>&nbsp;</b></font></span></font></div>
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font color="#17365D"><b>Brian Ivan Cusuanto</b></font></span></font></div>
        <div align="justify" style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="2" color="#17365D"><span lang="en-US" style="font-size:9pt">Business Analytics Department</span></font></span></font></div>
        <div align="justify" style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="2" color="#17365D"><span lang="en-US" style="font-size:10pt">Data Management &amp; Analytics (DMA)</span></font></span></font></div>
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="1" color="#365F91"><span lang="en-US" style="font-size:7pt"><b>&nbsp;</b></span></font></span></font></div>
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="2" color="#0D0D0D"><span lang="en-US" style="font-size:10pt"><b>PT. Bank Negara Indonesia (Persero) Tbk.</b></span></font></span></font></div>
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="2" color="#0F243E"><span lang="en-US" style="font-size:9pt">Menara BNI 15th floor, Jl. Pejompongan Raya no.7</span></font></span></font></div>
        <div style="font-family:Tahoma; font-size:13px; background-color:rgb(255,255,255); margin:0px">
        <font face="Calibri,sans-serif" size="2"><span style="font-size:11pt"><font size="2" color="#0F243E"><span lang="en-US" style="font-size:9pt">Jakarta 10210, Indonesia</span></font></span></font></div>
        </div>
        """
        
        # Combine body, disclaimer, and signature
        formatted_body = f"{body}<br><br><br>{disclaimer}<br><br>{signature}"
        mail.HTMLBody = formatted_body

        # Add attachment if provided
        if attachment_path and os.path.exists(attachment_path):
            mail.Attachments.Add(attachment_path)

        # Send the email
        mail.Send()
        log_message(f"Email sent successfully to '{recipient_email}'!")
    except Exception as e:
        log_message(f"Error: {e}")

if __name__ == "__main__":
    log_message("Script execution started.")

    # Encrypt the file
    encrypted_file = encrypt_excel_file(input_file, password)

    # # Clean up the original file (optional)
    # if encrypted_file:
    #     os.remove(input_file)

    upload_to_ftp(
        host = ftp_host,
        username = ftp_username,
        password = ftp_password,
        local_file = encrypted_file,
        dynamic_ftp_dir = ftp_dir
    )

    send_email_with_outlook(
        recipient = recipient_email, 
        subject = email_subject, 
        body = email_body.format(ftp_folder = dynamic_folder, n_rows = n_rows, file_name = file_name),
        cc = cc_recipients
    )

    log_message("Script execution completed.")
