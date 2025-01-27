# Automated Data Processing with Encryption and FTP Upload

## Why Use This Code?
This script automates the process of encrypting an Excel file, uploading it to an FTP server, and notifying recipients via email. It is designed to simplify repetitive tasks, improve data security, and ensure efficient communication.

---

## Problems Addressed
1. **Data Security:** Encrypts sensitive Excel files with a password to ensure confidentiality.
2. **Manual Workload:** Automates file uploads to a dynamic FTP directory.
3. **Communication Gap:** Sends automated email notifications with details of the upload.
4. **Repetitive Logging:** Provides comprehensive logs for better monitoring.

---

## Features
- Encrypts Excel files with a specified password.
- Dynamically generates FTP directories based on the current date.
- Uploads files securely to an FTP server.
- Sends HTML emails with custom body, signature, and attachments using Outlook.
- Logs all major events with timestamps for traceability.

---

## Dependencies
- Python Libraries: `os`, `ftplib`, `pandas`, `datetime`
- External Modules: `pywin32` (for Outlook and Excel operations)

Install dependencies using:
```bash
pip install pandas pywin32
```

---

## Setup Instructions
1. Clone the repository:
   ```bash
   git clone https://github.com/brdx88/glowing-giggle-auto-excelpsw-ftp-email.git
   ```

2. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Update the following placeholders in the script:
   - **File Paths:** Replace `path_to_excel` and `path_to_log` with your local paths.
   - **FTP Details:** Update `ftp_host`, `ftp_username`, and `ftp_password`.
   - **Email Details:** Add recipient emails in `recipient_email` and `cc_recipients`.

4. Execute the script:
   ```bash
   auto_excelpsw_ftp_email.py
   ```

---

## How It Works
### Step 1: Encryption
The `encrypt_excel_file` function uses `pywin32` to apply a password to the input Excel file, creating a secure copy.

### Step 2: FTP Upload
The `upload_to_ftp` function connects to the FTP server, creates a dynamic folder, and uploads the encrypted file.

### Step 3: Email Notification
The `send_email_with_outlook` function sends a detailed email, including the file's upload path and row count.

### Step 4: Logging
All actions are logged in a text file for tracking.

---

## Key Code Sections
1. **Encryption:**
   ```python
   encrypted_file = encrypt_excel_file(input_file, password)
   ```
2. **FTP Upload:**
   ```python
   upload_to_ftp(
       host=ftp_host,
       username=ftp_username,
       password=ftp_password,
       local_file=encrypted_file,
       dynamic_ftp_dir=ftp_dir
   )
   ```
3. **Email Sending:**
   ```python
   send_email_with_outlook(
       recipient=recipient_email,
       subject=email_subject,
       body=email_body.format(ftp_folder=dynamic_folder, n_rows=n_rows, file_name=file_name),
       cc=cc_recipients
   )
   ```

---

## Portfolio Repository Structure
```
project-folder/
├── data/
│   └── sample_file.xlsx
├── logs/
│   └── testing_excel_process_log.txt
├── src/
│   ├── encrypt_and_upload.py
│   └── utilities.py
├── README.md
└── requirements.txt
```
- **data/**: Contains sample input files.
- **logs/**: Stores logs for debugging and record-keeping.
- **src/**: Holds Python scripts.
- **README.md**: Provides project details.
- **requirements.txt**: Lists dependencies.

---

## Notes
- Ensure the FTP server is accessible from your environment.
- Outlook must be installed and configured on your machine.
- Log file location: `path_to_log/testing_excel_process_log.txt`

---

## Conclusions
This script provides an efficient, secure, and automated solution for handling Excel files, ensuring data integrity and seamless communication. Customize and expand it to meet your specific workflow requirements.

For any issues or contributions, feel free to raise an [issue](https://github.com/brdx88/glowing-giggle-auto-excelpsw-ftp-email/issues) or submit a pull request.
