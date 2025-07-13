import os
import requests
import pandas as pd
import smtplib
from email.message import EmailMessage
from requests.exceptions import HTTPError, RequestException

# Read environment variables
GITHUB_REPOS = os.getenv('GITHUB_REPOS', '').split(',')
GITHUB_TOKEN = os.getenv('GITHUB_TOKEN')
EMAIL_SENDER = os.getenv('EMAIL_SENDER')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD')
TEAM_EMAILS = os.getenv('TEAM_EMAILS', '')
SMTP_SERVER = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))

# Parse teammates' emails into a list and remove empty entries
RECIPIENTS = [email.strip() for email in TEAM_EMAILS.split(',') if email.strip()]

def get_open_prs(repo):
    """
    Fetch open pull requests from a GitHub repository using the GitHub API.
    Returns a list of PR JSON objects.
    Raises HTTPError on bad response.
    """
    url = f'https://api.github.com/repos/{repo}/pulls?state=open'
    headers = {'Authorization': f'token {GITHUB_TOKEN}'}
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except HTTPError as http_err:
        print(f"HTTP error occurred while fetching PRs for repo '{repo}': {http_err}")
    except RequestException as req_err:
        print(f"Request error occurred while fetching PRs for repo '{repo}': {req_err}")
    except Exception as err:
        print(f"Unexpected error occurred while fetching PRs for repo '{repo}': {err}")
    return []  # Return empty list if any error occurs

def create_excel(all_prs, filename='open_prs_report.xlsx'):
    """
    Create an Excel file summarizing all PRs.
    The Excel contains a table with columns including repository, PR number, title, author, URL, etc.
    Returns the filename of the created Excel file.
    """
    try:
        data = []
        for pr in all_prs:
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
        print(f"Excel report created: {filename}")
        return filename
    except Exception as e:
        print(f"Error creating Excel file: {e}")
        raise  # Let the caller handle this

def send_email_with_attachment(subject, body, attachment_path):
    """
    Sends an email with the specified subject and body, attaching the file at attachment_path.
    The email is sent to all recipients in RECIPIENTS.
    """
    try:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = EMAIL_SENDER
        msg['To'] = ', '.join(RECIPIENTS)
        msg.set_content(body)

        # Read the Excel file and attach it
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)

        msg.add_attachment(file_data, maintype='application',
                           subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           filename=file_name)

        # Connect to SMTP server and send email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.starttls()
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
        print(f"Email successfully sent to: {', '.join(RECIPIENTS)}")
    except FileNotFoundError:
        print(f"Attachment file not found: {attachment_path}")
    except smtplib.SMTPException as smtp_err:
        print(f"SMTP error occurred: {smtp_err}")
    except Exception as e:
        print(f"Unexpected error occurred while sending email: {e}")

def main():
    # Collect PRs from all repos into one list
    all_prs = []
    for repo in GITHUB_REPOS:
        repo = repo.strip()
        if not repo:
            continue
        print(f"Fetching open PRs for repository: {repo}")
        prs = get_open_prs(repo)
        for pr in prs:
            pr['repo'] = repo  # Tag PR with repo name
        all_prs.extend(prs)

    if not all_prs:
        # No open PRs found - send simple notification email or skip?
        body = "No open pull requests today across monitored repositories."
        subject = "Daily Open PR Report - No Open PRs"
        print(body)
        # Optionally, you can send an email even if no PRs, or skip sending
        return

    # Create Excel report file
    try:
        filename = create_excel(all_prs)
    except Exception:
        print("Failed to create Excel report, aborting email send.")
        return

    # Email the Excel report to teammates
    subject = "Daily Open PR Report for Multiple Repositories"
    body = f"Please find attached the daily report of open PRs for repos: {', '.join(GITHUB_REPOS)}."
    send_email_with_attachment(subject, body, filename)

if __name__ == "__main__":
    main()
