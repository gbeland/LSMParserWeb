# PythonAnywhere Deployment Guide

This guide describes how to deploy the LSMParser Web App to a free PythonAnywhere account.

## Prerequisites
1.  A [PythonAnywhere](https://www.pythonanywhere.com/) account (Beginner/Free tier is fine).
2.  A [Gmail](https://gmail.com/) account for sending emails.
3.  An **App Password** for that Gmail account (Regular passwords won't work).
    *   Go to [Google Account Security](https://myaccount.google.com/security).
    *   Enable 2-Step Verification if not already on.
    *   Search for "App Passwords" and create one named "LSMParser".
    *   **Keep this password key safe.**

## Step 1: Upload Code
1.  Log in to PythonAnywhere.
2.  Go to the **Consoles** tab and start a **Bash** console.
3.  Clone your repository (if public) or upload your files.
    *   *If uploading manually*: Go to **Files** tab, create a directory `mysite`, and upload all project files (`web_app.py`, `main.py`, `requirements.txt`, `templates/`, `utils/`, etc.).
    *   *Using Git (Recommended)*:
        ```bash
        git clone https://github.com/gbeland/LSMParserWeb.git mysite
        ```

## Step 2: Virtual Environment
In the Bash console, run:
```bash
cd mysite
mkvirtualenv --python=/usr/bin/python3.10 myenv
pip install -r requirements.txt
```
*Note: This might take a few minutes.*

## Step 3: Configure Web App
1.  Go to the **Web** tab.
2.  Click **Add a new web app**.
3.  Click **Next**, select **Flask**, then select **Python 3.10**.
4.  **Path**: It will ask for the path. Enter: `/home/seabeland/mysite/web_app.py` (Verify this path matches where you put the code).
5.  **Virtualenv**: In the Virtualenv section, enter the name of the env you created: `myenv` (or full path `/home/seabeland/.virtualenvs/myenv`).

## Step 4: WSGI Configuration
1.  In the **Web** tab, look for the **WSGI configuration file** link (e.g., `/var/www/seabeland_pythonanywhere_com_wsgi.py`) and click it.
2.  Delete everything and paste this (adjusting paths if needed):
    ```python
    import sys
    import os

    # Add your project directory to the sys.path
    project_home = '/home/seabeland/mysite'
    if project_home not in sys.path:
        sys.path = [project_home] + sys.path

    # Set environment variables for Email
    os.environ['SMTP_SERVER'] = 'smtp.gmail.com'
    os.environ['SMTP_PORT'] = '587'
    os.environ['SMTP_USER'] = 'your-email@gmail.com'  # <--- UPDATE THIS
    os.environ['SMTP_PASSWORD'] = 'your-app-password'   # <--- UPDATE THIS

    from web_app import app as application
    ```
3.  **IMPORTANT**: Replace `your-email@gmail.com`, and `your-app-password` with your actual details.
4.  Click **Save**.

## Step 5: Static Files (Optional but Recommended)
In the **Web** tab, under **Static files**:
-   **URL**: `/static/`
-   **Directory**: `/home/seabeland/mysite/static`

## Step 6: HTTPS
PythonAnywhere provides HTTPS automatically for `seabeland.pythonanywhere.com`.
-   In the **Web** tab, enable **Force HTTPS**.

## Step 7: Reload
Click the big green **Reload** button at the top of the Web tab.

## verification
Visit `https://seabeland.pythonanywhere.com`.
1.  Upload a log file.
2.  Check if parsing works.
3.  Try "Email Report" (Check your inbox).

## Troubleshooting
-   **Email Fails?** Check the **Error Log** in the Web tab. Ensure you used a Google App Password, not your login password.
-   **PDF Fails?** The free tier might struggle with PDF generation libraries. If it crashes, the app is designed to still allow HTML viewing and Excel downloading.
