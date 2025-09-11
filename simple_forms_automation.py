#!/usr/bin/env python3
"""
Debug version of forms automation script
Provides detailed error information
"""

import os
import sys
import logging

# Set up logging first
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)

def check_environment():
    """Check all required environment variables"""
    required_vars = [
        'DATA_SOURCE',
        'EMAIL_USER', 
        'EMAIL_PASSWORD',
        'EMAIL_RECIPIENTS',
        'GOOGLE_SHEET_ID',
        'GOOGLE_CREDENTIALS_JSON'
    ]
    
    missing_vars = []
    for var in required_vars:
        value = os.getenv(var)
        if not value:
            missing_vars.append(var)
        else:
            if var == 'GOOGLE_CREDENTIALS_JSON':
                logging.info(f"✅ {var}: Found (length: {len(value)} chars)")
            elif var == 'EMAIL_PASSWORD':
                logging.info(f"✅ {var}: Found (length: {len(value)} chars)")
            else:
                logging.info(f"✅ {var}: {value}")
    
    if missing_vars:
        logging.error(f"❌ Missing required environment variables: {missing_vars}")
        return False
    
    return True

def check_imports():
    """Check if all required packages can be imported"""
    required_packages = [
        ('pandas', 'pd'),
        ('matplotlib.pyplot', 'plt'),
        ('gspread', 'gspread'),
        ('google.oauth2.service_account', 'Credentials'),
        ('smtplib', 'smtplib'),
        ('json', 'json')
    ]
    
    for package, alias in required_packages:
        try:
            if package == 'matplotlib.pyplot':
                import matplotlib.pyplot as plt
                logging.info(f"✅ {package}: imported successfully")
            elif package == 'google.oauth2.service_account':
                from google.oauth2.service_account import Credentials
                logging.info(f"✅ {package}: imported successfully")
            elif package == 'pandas':
                import pandas as pd
                logging.info(f"✅ {package}: imported successfully (version: {pd.__version__})")
            else:
                exec(f"import {package}")
                logging.info(f"✅ {package}: imported successfully")
        except ImportError as e:
            logging.error(f"❌ Failed to import {package}: {e}")
            return False
    
    return True

def test_google_sheets_connection():
    """Test Google Sheets connection"""
    try:
        import gspread
        import json
        from google.oauth2.service_account import Credentials
        
        # Get credentials
        creds_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        
        if not creds_json:
            logging.error("❌ No Google credentials found")
            return False
        
        # Parse credentials
        try:
            creds_dict = json.loads(creds_json)
            logging.info("✅ Google credentials JSON parsed successfully")
        except json.JSONDecodeError as e:
            logging.error(f"❌ Invalid JSON in Google credentials: {e}")
            return False
        
        # Create credentials object
        credentials = Credentials.from_service_account_info(
            creds_dict,
            scopes=['https://spreadsheets.google.com/feeds',
                   'https://www.googleapis.com/auth/drive']
        )
        
        # Authorize gspread
        gc = gspread.authorize(credentials)
        logging.info("✅ Google Sheets authorization successful")
        
        # Try to open the sheet
        sheet = gc.open_by_key(sheet_id).sheet1
        logging.info(f"✅ Connected to sheet: {sheet.title}")
        
        # Try to read data
        values = sheet.get_all_values()
        logging.info(f"✅ Read {len(values)} rows from sheet")
        
        if len(values) > 0:
            logging.info(f"📋 Headers: {values[0]}")
        if len(values) > 1:
            logging.info(f"📝 Sample row: {values[1]}")
        
        return True
        
    except Exception as e:
        logging.error(f"❌ Google Sheets connection failed: {e}")
        return False

def test_email_configuration():
    """Test email configuration"""
    try:
        import smtplib
        
        email_user = os.getenv('EMAIL_USER')
        email_password = os.getenv('EMAIL_PASSWORD')
        
        if not email_user or not email_password:
            logging.error("❌ Email credentials missing")
            return False
        
        # Detect SMTP server
        domain = email_user.split('@')[1].lower()
        smtp_map = {
            'gmail.com': ('smtp.gmail.com', 587),
            'outlook.com': ('smtp-mail.outlook.com', 587),
            'hotmail.com': ('smtp-mail.outlook.com', 587),
            'live.com': ('smtp-mail.outlook.com', 587),
            'yahoo.com': ('smtp.mail.yahoo.com', 587),
            'icloud.com': ('smtp.mail.me.com', 587)
        }
        
        smtp_server, smtp_port = smtp_map.get(domain, (f'smtp.{domain}', 587))
        
        # Override with environment variables if provided
        smtp_server = os.getenv('SMTP_SERVER', smtp_server)
        smtp_port = int(os.getenv('SMTP_PORT', smtp_port))
        
        logging.info(f"📧 Testing email: {email_user}")
        logging.info(f"🔗 SMTP: {smtp_server}:{smtp_port}")
        
        # Test SMTP connection
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email_user, email_password)
        server.quit()
        
        logging.info("✅ Email connection successful")
        return True
        
    except Exception as e:
        logging.error(f"❌ Email connection failed: {e}")
        
        # Provide helpful suggestions
        if "authentication failed" in str(e).lower():
            logging.error("💡 Suggestion: Check if you need an app password instead of regular password")
        elif "connection refused" in str(e).lower():
            logging.error("💡 Suggestion: Check SMTP server and port settings")
        
        return False

def main():
    """Main debug function"""
    logging.info("🔍 Starting debug diagnostics...")
    
    # Check 1: Environment variables
    logging.info("\n📋 Step 1: Checking environment variables...")
    env_ok = check_environment()
    
    # Check 2: Package imports
    logging.info("\n📦 Step 2: Checking package imports...")
    imports_ok = check_imports()
    
    # Check 3: Google Sheets connection
    logging.info("\n📊 Step 3: Testing Google Sheets connection...")
    sheets_ok = test_google_sheets_connection()
    
    # Check 4: Email configuration
    logging.info("\n📧 Step 4: Testing email configuration...")
    email_ok = test_email_configuration()
    
    # Summary
    logging.info("\n🎯 Summary:")
    logging.info(f"Environment Variables: {'✅' if env_ok else '❌'}")
    logging.info(f"Package Imports: {'✅' if imports_ok else '❌'}")
    logging.info(f"Google Sheets: {'✅' if sheets_ok else '❌'}")
    logging.info(f"Email Config: {'✅' if email_ok else '❌'}")
    
    if all([env_ok, imports_ok, sheets_ok, email_ok]):
        logging.info("\n🎉 All checks passed! The system should work correctly.")
        
        # Try to run a simple report
        try:
            logging.info("\n📊 Running simplified report...")
            from datetime import datetime
            
            # Simple success report
            recipients = os.getenv('EMAIL_RECIPIENTS', '').split(',')
            logging.info(f"✅ Debug completed successfully!")
            logging.info(f"📧 Would send report to: {len(recipients)} recipients")
            logging.info(f"🕐 Report time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
            return True
            
        except Exception as e:
            logging.error(f"❌ Error running simplified report: {e}")
            return False
    else:
        logging.error("\n❌ Some checks failed. Please fix the issues above.")
        return False

if __name__ == "__main__":
    try:
        success = main()
        exit_code = 0 if success else 1
        logging.info(f"\n🏁 Debug completed with exit code: {exit_code}")
        sys.exit(exit_code)
    except Exception as e:
        logging.error(f"💥 Unexpected error: {e}")
        sys.exit(1)
