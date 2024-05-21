import os
import requests
from adal import AuthenticationContext
import json
import time
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

def load_config(config_file):
    with open(config_file) as file:
        return json.load(file)

def get_access_token(authority, resource, username, password, client_id):
    context = AuthenticationContext(authority)
    token = context.acquire_token_with_username_password(resource, username, password, client_id)
    if 'accessToken' in token:
        return token['accessToken']
    else:
        raise Exception("Failed to retrieve access token: " + json.dumps(token))

def setup_headers(access_token):
    return {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }

def get_script_directory():
    return os.getcwd()

def export_report_to_pdf(workspace_id, report_id, output_path, filter_expression, headers, save_locally):
    # Prepare export request URL and body
    export_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports/{report_id}/ExportTo"
    export_body = {
        "format": "PDF",
        "powerBIReportConfiguration": {
            "reportLevelFilters": [
                {
                    "filter": filter_expression
                }
            ]
        }
    }

    # Invoke export API
    export_response = requests.post(export_url, headers=headers, json=export_body)
    if not export_response.ok:
        print(f"Failed export initiation response: {export_response.text}")
        raise Exception(f"Error initiating export: {export_response.text}")

    export_id = export_response.json().get('id')
    print(f"Export initiated. Export ID: {export_id}")

    # Poll for export status
    status_url = f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/reports/{report_id}/exports/{export_id}"
    start_time = time.time()
    timeout = 600  # 10 minutes

    while True:
        time.sleep(5)
        status_response = requests.get(status_url, headers=headers)
        if not status_response.ok:
            print(f"Failed status check response: {status_response.text}")
            raise Exception(f"Error getting export status: {status_response.text}, Status code: {status_response.status_code}")

        status = status_response.json()
        export_status = status['status']
        percent_complete = status.get('percentComplete', 0)
        print(f"Export status: {export_status}, Percent complete: {percent_complete}%")

        if export_status == "Succeeded":
            file_url = status['resourceLocation']
            print(f"Export succeeded. Download URL: {file_url}")
            break
        elif export_status == "Failed":
            raise Exception("Export failed")

        if time.time() - start_time > timeout:
            raise Exception("Export process timed out")

    # Download the exported report
    download_response = requests.get(file_url, headers=headers)
    if download_response.status_code // 100 == 2:
        if save_locally:
            with open(output_path, 'wb') as file:
                file.write(download_response.content)
            print(f"Report saved to {output_path}")
        else:
            print("Report downloaded but not saved locally as per configuration.")
        return file_url, output_path
    else:
        raise Exception(f"Error downloading report: {download_response.text}")

def send_email(subject, body, attachment_path, file_url, headers, save_locally):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ", ".join(to_email) 
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    attachment_name = os.path.basename(attachment_path)
    
    if save_locally:
        with open(attachment_path, 'rb') as attachment:
            part = MIMEApplication(attachment.read(), Name=attachment_name)
            part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
            msg.attach(part)
    else:
        download_response = requests.get(file_url, headers=headers)
        if download_response.status_code // 100 == 2:
            part = MIMEApplication(download_response.content, Name=attachment_name)
            part['Content-Disposition'] = f'attachment; filename="{attachment_name}"'
            msg.attach(part)
        else:
            raise Exception(f"Error downloading report for email: {download_response.text}")

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            server.send_message(msg)
            print(f"Email sent to {', '.join(to_email)} with attachment {attachment_name}")
    except smtplib.SMTPAuthenticationError as e:
        print(f"SMTP Authentication Error: {e}")
        print("Please ensure that SMTP client authentication is enabled for your tenant.")
        print("Visit https://aka.ms/smtp_auth_disabled for more information.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Main function
def main():
    # Load configuration
    config = load_config('PBIconfig.json')

    # Authentication parameters
    authority = f"https://login.microsoftonline.com/{config['tenant_id']}"
    resource = 'https://analysis.windows.net/powerbi/api'

    # Get access token
    access_token = get_access_token(authority, resource, config['username'], config['password'], config['client_id'])

    # Setup headers
    headers = setup_headers(access_token)

    # Determine script directory
    script_dir = get_script_directory()

    # Ensure the PDF Reports folder exists
    pdf_reports_folder = os.path.join(script_dir, "PDF Reports")
    if not os.path.exists(pdf_reports_folder):
        os.makedirs(pdf_reports_folder)

    # Define base output path
    base_output_path = os.path.join(pdf_reports_folder, "Report_{}_{}.pdf")

    # Email configuration
    global smtp_server, smtp_port, smtp_username, smtp_password, from_email, to_email, save_locally
    smtp_server = config['smtp_server']
    smtp_port = config['smtp_port']
    smtp_username = config['smtp_username']
    smtp_password = config['smtp_password']
    from_email = config['from_email']
    to_email = config['to_email']  # Expecting this to be a list of email addresses
    save_locally = config.get('save_locally', True)  # Default to True if not specified

    # List of employees to loop through
    employees = [
        {"id": 123, "name": "LastName, FirstName"},
        {"id": 456, "name": "LastName, FirstName"},
        {"id": 789, "name": "LastName, FirstName"}
    ]

    # Get current date in YYMMDD format
    current_date = datetime.now().strftime("%y%m%d")

    for employee in employees:
        filter_expression = f"Employee/Employee_x0020_Last_x0020_First_x0020_Name eq '{employee['name']}'"
        output_path = base_output_path.format(employee['name'].replace(", ", "_"), current_date)
        print(f"Generating report for {employee['name']} with Employee ID: {employee['id']}")
        file_url, attachment_path = export_report_to_pdf(config['workspace_id'], config['report_id'], output_path, filter_expression, headers, save_locally)
        
        # Email details
        subject = f"Report - {employee['name']} - {current_date}"
        body = f"Please find attached the report for {employee['name']} dated {current_date}."

        # Send email with the generated PDF report
        send_email(subject, body, attachment_path, file_url, headers, save_locally)

# Run the main function
if __name__ == "__main__":
    main()