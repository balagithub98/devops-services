import os
import requests
import pandas as pd
import smtplib
from email.message import EmailMessage
from requests.exceptions import HTTPError, RequestException

# Load configuration from environment variables
GITHUB_REPOS = os.getenv('MY_GITHUB_REPOS', '').split(',')
GITHUB_TOKEN = os.getenv('MY_GITHUB_TOKEN')  # Updated secret name
EMAIL_SENDER = os.getenv('EMAIL_SENDER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
TEAM_EMAILS = os.getenv('TEAM_EMAILS', '')
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))

# Parse team emails
RECIPIENTS = [email.strip() for email in TEAM_EMAILS.split(',') if email.strip()]

def get_open_prs(repo):
    """
    Fetch open pull requests for a given repository.
    """
    url = f'https://api.github.com/repos/{repo}/pulls?state=open'
    headers = {'Authorization': f'token {GITHUB_TOKEN}'}

    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except HTTPError as http_err:
        print(f"[ERROR] HTTP error for repo '{repo}': {http_err}")
    except RequestException as req_err:
        print(f"[ERROR] Request error for repo '{repo}': {req_err}")
    except Exception as e:
        print(f"[ERROR] Unexpected error for repo '{repo}': {e}")
    return []

def create_excel_report(pr_list, filename='open_prs_report.xlsx'):
    """
    Generate an Excel file from list of PRs.
    """
    try:
        data = []
        for pr in pr_list:
            data.append({
                'Repository': pr['repo'],
                'PR Number': pr['number'],
                'Title': pr['title'],
                'Author': pr['user']['login'],
                'URL': pr['html_url'],
                'Created At': pr['created_at'],
                'Updated At': pr['updated_at'],
                'State': pr['state']
            })
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False)
        print(f"[INFO] Excel report created: {filename}")
        return filename
    except Exception as e:
        print(f"[ERROR] Failed to create Excel report: {e}")
        raise

def send_email(subject, body, attachment_path):
    """
    Send email with Excel file attached.
    """
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = EMAIL_SENDER
        msg['To'] = ', '.join(RECIPIENTS)
        msg.set_content(body)

        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data,
                           maintype='application',
                           subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           filename=file_name)

        print(f"[INFO] Sending email FROM: {EMAIL_SENDER}")
        print(f"[INFO] Sending email TO: {', '.join(RECIPIENTS)}")

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)

        print("[SUCCESS] Email sent successfully.")
    except FileNotFoundError:
        print(f"[ERROR] Report file not found: {attachment_path}")
    except smtplib.SMTPException as smtp_err:
        print(f"[ERROR] SMTP error: {smtp_err}")
    except Exception as e:
        print(f"[ERROR] Unexpected error while sending email: {e}")

def main():
    all_open_prs = []

    for repo in GITHUB_REPOS:
        repo = repo.strip()
        if not repo:
            continue
        print(f"[INFO] Fetching open PRs from: {repo}")
        prs = get_open_prs(repo)
        for pr in prs:
            pr['repo'] = repo  # Add repo name to each PR dict
        all_open_prs.extend(prs)

    if not all_open_prs:
        print("[INFO] No open PRs found.")
        return

    try:
        report_file = create_excel_report(all_open_prs)
    except Exception:
        print("[ERROR] Skipping email due to report creation failure.")
        return

    subject = "Daily GitHub Open PR Report"
    body = f"Attached is the daily open PR report for:\n\n{', '.join(GITHUB_REPOS)}"
    send_email(subject, body, report_file)

if __name__ == "__main__":
    main()
