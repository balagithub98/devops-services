import os
import smtplib
import pathlib
import traceback
from github import Github, GithubException
from openpyxl import Workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# === CONFIGURATION ===
REPOSITORIES = [
    "your-org/repo-1",
    "your-org/repo-2",
    # Add more repositories as needed
]
RECIPIENTS = [
    "dev-team@example.com",
    "qa-team@example.com"
]
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 465  # Use 587 for TLS if not using SSL

# === ENVIRONMENT VARIABLES ===
GITHUB_TOKEN = os.getenv("GH_TOKEN")
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = os.getenv("EMAIL_PASS")

if not all([GITHUB_TOKEN, EMAIL_USER, EMAIL_PASS]):
    raise EnvironmentError("Missing required environment variables: GH_TOKEN, EMAIL_USER, EMAIL_PASS")

# === BLOCK 1: FETCH OPEN PULL REQUESTS ===
def fetch_pull_requests():
    pr_data = []
    try:
        github_client = Github(GITHUB_TOKEN)
        for repo_name in REPOSITORIES:
            try:
                repo = github_client.get_repo(repo_name)
                prs = repo.get_pulls(state="open")
                for pr in prs:
                    pr_data.append([repo_name, pr.title, pr.user.login, pr.html_url])
            except GithubException as e:
                print(f"[ERROR] Failed to fetch PRs for {repo_name}: {e}")
    except Exception as e:
        print("[CRITICAL] Error initializing GitHub client")
        traceback.print_exc()
    return pr_data

# === BLOCK 2: GENERATE EXCEL REPORT ===
def generate_excel(pr_data, report_path):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Open PRs"
        ws.append(["Repository", "PR Title", "Author", "URL"])
        for row in pr_data:
            ws.append(row)
        report_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(report_path)
        print(f"[INFO] Excel report saved to {report_path}")
    except Exception as e:
        print("[ERROR] Failed to generate Excel report")
        traceback.print_exc()

# === BLOCK 3: SEND EMAIL WITH ATTACHMENT ===
def send_email(report_path):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_USER
        msg['To'] = ", ".join(RECIPIENTS)
        msg['Subject'] = "Automated Open PR Report"

        body = (
            "Hi team,\n\n"
            "Attached is the latest automated report of open pull requests across our repositories.\n\n"
            "Best regards,\nDevOps Bot"
        )
        msg.attach(MIMEText(body, "plain"))

        with open(report_path, "rb") as file:
            part = MIMEApplication(file.read(), Name="pr_report.xlsx")
            part['Content-Disposition'] = 'attachment; filename="pr_report.xlsx"'
            msg.attach(part)

        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(EMAIL_USER, EMAIL_PASS)
        server.sendmail(EMAIL_USER, RECIPIENTS, msg.as_string())
        server.quit()
        print("[INFO] Email sent successfully.")
    except Exception as e:
        print("[ERROR] Failed to send email")
        traceback.print_exc()

# === MAIN WORKFLOW ===
def main():
    try:
        report_path = pathlib.Path(__file__).resolve().parent.parent / "reports" / "pr_report.xlsx"
        pr_data = fetch_pull_requests()
        if pr_data:
            generate_excel(pr_data, report_path)
            send_email(report_path)
        else:
            print("[INFO] No open pull requests found. Skipping report and email.")
    except Exception as e:
        print("[FATAL] Unexpected error in main workflow")
        traceback.print_exc()

if __name__ == "__main__":
    main()
