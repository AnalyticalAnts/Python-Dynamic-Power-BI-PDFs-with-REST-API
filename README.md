# Python: Dynamic Power BI PDFs with REST API

Welcome to the repository for generating dynamic Power BI PDFs using Python and the Power BI REST API. This repository contains scripts to automate the export of Power BI reports to PDF and send them via email.

For a detailed write-up on how everything works, please visit: [Python: Dynamic Power BI PDFs with REST API](https://analyticalants.co/data-engineering/python-dynamic-power-bi-pdfs-with-rest-api/)

## Overview

This project aims to automate the process of exporting Power BI reports to PDF and emailing them to specified recipients. The automation is achieved using Python scripts, with support from a PowerShell script to schedule the automation.

## Files in the Repository

- **PBIconfig.json**: Configuration file containing all necessary credentials and parameters.
- **AutomatedReports.py**: Main Python script to export Power BI reports to PDF.
- **PS_Start_Automation.ps1**: PowerShell script to schedule and run the Python script.
- **SupportingScripts.ipynb**: Jupyter notebook with supporting scripts for additional tasks.

## Getting Started

### Prerequisites

- Python 3.x
- Required Python libraries: `requests`, `adal`, `json`, `time`, `smtplib`, `email`
- PowerShell (for automation)

### Configuration

1. **PBIconfig.json**: Configure your credentials and parameters in this file.
    ```json
    {
        "client_id": "your-client-id",
        "client_secret": "your-client-secret",
        "tenant_id": "your-tenant-id",
        "username": "your-username",
        "password": "your-password",
        "workspace_id": "your-workspace-id",
        "report_id": "your-report-id",
        "smtp_server": "smtp-server-address",
        "smtp_port": 587,
        "smtp_username": "your-smtp-username",
        "smtp_password": "your-smtp-password",
        "from_email": "your-email@example.com",
        "to_email": ["recipient1@example.com", "recipient2@example.com"],
        "save_locally": true  // Set to true or false as needed
    }
    ```

### Running the Scripts

1. **Python Script**: Run the `AutomatedReports.py` to export the report to PDF and send emails.
    ```bash
    python AutomatedReports.py
    ```

2. **PowerShell Script**: Use `PS_Start_Automation.ps1` to schedule the Python script to run at specified intervals.
    ```powershell
    # PowerShell script to run Python script
    $pythonExecutable = "C:\path\to\python.exe"
    $scriptPath = "C:\path\to\AutomatedReports.py"
    & $pythonExecutable $scriptPath
    ```

### Supporting Scripts

- **SupportingScripts.ipynb**: This Jupyter notebook contains additional supporting scripts for various tasks related to the project.

## License

This project is licensed under the MIT License.

## Contact

For any questions or issues, please refer to the detailed write-up or open an issue in this repository.
