#!/usr/bin/env python3
"""
Forms automation script for meeting data
Fixed for separate date/time columns and email configuration
"""

import os
import gspread
import json
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class MeetingFormsReporter:
    def __init__(self):
        # Email configuration with better defaults
        self.email_user = os.getenv('EMAIL_USER')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        self.smtp_server = os.getenv('SMTP_SERVER') or self.detect_smtp_server()
        
        # Handle SMTP_PORT more safely
        smtp_port_str = os.getenv('SMTP_PORT', '').strip()
        if smtp_port_str and smtp_port_str.isdigit():
            self.smtp_port = int(smtp_port_str)
        else:
            self.smtp_port = 587  # Default port
        
        # Handle USE_TLS safely
        use_tls_str = os.getenv('USE_TLS', 'true').lower().strip()
        self.use_tls = use_tls_str in ['true', 'yes', '1']
        
        self.recipients = [email.strip() for email in os.getenv('EMAIL_RECIPIENTS', '').split(',') if email.strip()]
        
        # Google Sheets configuration
        self.sheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
        
        # Column mapping for your specific data structure
        self.column_mapping = {
            'date': 'date',
            'time': 'time',
            'first_name': 'Student First Name',
            'last_name': 'Student Last Name', 
            'course': 'Course Section',
            'meeting_person': 'Meeting person',
            'meeting_type': 'Meeting Type',
            'topic': 'Topic',
            'powerapps_id': '__PowerAppsId__'
        }
    
    def detect_smtp_server(self):
        """Auto-detect SMTP server based on email address"""
        if not self.email_user:
            return 'smtp.gmail.com'
        
        domain = self.email_user.split('@')[1].lower()
        
        smtp_map = {
            'gmail.com': 'smtp.gmail.com',
            'googlemail.com': 'smtp.gmail.com',
            'outlook.com': 'smtp-mail.outlook.com',
            'hotmail.com': 'smtp-mail.outlook.com',
            'live.com': 'smtp-mail.outlook.com',
            'msn.com': 'smtp-mail.outlook.com',
            'yahoo.com': 'smtp.mail.yahoo.com',
            'yahoo.co.uk': 'smtp.mail.yahoo.com',
            'icloud.com': 'smtp.mail.me.com',
            'me.com': 'smtp.mail.me.com',
            'mac.com': 'smtp.mail.me.com'
        }
        
        detected_server = smtp_map.get(domain, f'smtp.{domain}')
        logging.info(f"📧 Auto-detected SMTP server: {detected_server} for {domain}")
        return detected_server
    
    def connect_to_sheets(self):
        """Connect to Google Sheets"""
        try:
            if not self.credentials_json:
                logging.error("No Google credentials found")
                return None
            
            creds_dict = json.loads(self.credentials_json)
            credentials = Credentials.from_service_account_info(
                creds_dict,
                scopes=['https://spreadsheets.google.com/feeds',
                       'https://www.googleapis.com/auth/drive']
            )
            
            gc = gspread.authorize(credentials)
            sheet = gc.open_by_key(self.sheet_id).sheet1
            
            logging.info(f"✅ Connected to sheet: {sheet.title}")
            return sheet
            
        except Exception as e:
            logging.error(f"❌ Failed to connect to Google Sheets: {e}")
            return None
    
    def load_data(self):
        """Load data from Google Sheets"""
        sheet = self.connect_to_sheets()
        if not sheet:
            return None
        
        try:
            values = sheet.get_all_values()
            
            if len(values) < 2:
                logging.warning("Sheet has no data rows")
                return pd.DataFrame()
            
            # Create DataFrame
            df = pd.DataFrame(values[1:], columns=values[0])
            
            # Combine date and time columns into datetime
            date_col = self.column_mapping['date']
            time_col = self.column_mapping['time']
            
            if date_col in df.columns and time_col in df.columns:
                # Combine date and time
                df['datetime'] = df[date_col] + ' ' + df[time_col]
                df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
                # Remove rows with invalid dates
                df = df.dropna(subset=['datetime'])
            elif date_col in df.columns:
                # Use date only
                df['datetime'] = pd.to_datetime(df[date_col], errors='coerce')
                df = df.dropna(subset=['datetime'])
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            logging.info(f"📊 Loaded {len(df)} meeting records")
            return df
            
        except Exception as e:
            logging.error(f"❌ Error loading data: {e}")
            return None
    
    def generate_statistics(self, df):
        """Generate meeting statistics"""
        if df is None or df.empty:
            return {
                'total_meetings': 0,
                'today_meetings': 0,
                'yesterday_meetings': 0,
                'this_week_meetings': 0,
                'avg_daily_meetings': 0,
                'peak_hour': 'N/A',
                'most_recent': 'No data',
                'popular_meeting_type': 'N/A',
                'active_courses': 0,
                'unique_students': 0
            }
        
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)
        week_ago = datetime.now() - timedelta(days=7)
        
        # Basic statistics
        stats = {
            'total_meetings': len(df),
            'today_meetings': len(df[df['datetime'].dt.date == today]),
            'yesterday_meetings': len(df[df['datetime'].dt.date == yesterday]),
            'this_week_meetings': len(df[df['datetime'] >= week_ago]),
            'avg_daily_meetings': df.groupby(df['datetime'].dt.date).size().mean() if len(df) > 0 else 0,
            'peak_hour': df.groupby(df['datetime'].dt.hour).size().idxmax() if len(df) > 0 else 'N/A',
            'most_recent': df['datetime'].max() if len(df) > 0 else 'No data'
        }
        
        # Meeting-specific statistics
        meeting_type_col = self.column_mapping['meeting_type']
        course_col = self.column_mapping['course']
        first_name_col = self.column_mapping['first_name']
        last_name_col = self.column_mapping['last_name']
        
        if meeting_type_col in df.columns and not df[meeting_type_col].empty:
            meeting_types = df[meeting_type_col].value_counts()
            stats['popular_meeting_type'] = meeting_types.index[0] if len(meeting_types) > 0 else 'N/A'
        else:
            stats['popular_meeting_type'] = 'N/A'
        
        if course_col in df.columns:
            stats['active_courses'] = df[course_col].nunique()
        else:
            stats['active_courses'] = 0
        
        if first_name_col in df.columns and last_name_col in df.columns:
            df['full_name'] = df[first_name_col].astype(str) + ' ' + df[last_name_col].astype(str)
            stats['unique_students'] = df['full_name'].nunique()
        else:
            stats['unique_students'] = 0
        
        return stats
    
    def create_visualizations(self, df):
        """Create meeting analytics visualizations"""
        if df is None or df.empty:
            plt.figure(figsize=(10, 6))
            plt.text(0.5, 0.5, 'No Meeting Data Available', ha='center', va='center', fontsize=16)
            plt.title('Meeting Analytics Dashboard')
            plt.axis('off')
            plt.savefig('meeting_analytics.png', dpi=300, bbox_inches='tight')
            plt.close()
            return
        
        # Create multi-panel analytics dashboard
        fig, axes = plt.subplots(2, 2, figsize=(15, 10))
        fig.suptitle('Meeting Analytics Dashboard', fontsize=16, fontweight='bold')
        
        # Daily meetings (last 30 days)
        last_30_days = df[df['datetime'] >= datetime.now() - timedelta(days=30)]
        if len(last_30_days) > 0:
            daily_counts = last_30_days.groupby(last_30_days['datetime'].dt.date).size()
            daily_counts.plot(kind='line', ax=axes[0,0], marker='o', color='#0078d4')
            axes[0,0].set_title('Daily Meetings (Last 30 Days)')
            axes[0,0].tick_params(axis='x', rotation=45)
            axes[0,0].set_ylabel('Number of Meetings')
        else:
            axes[0,0].text(0.5, 0.5, 'No recent data', ha='center', va='center')
            axes[0,0].set_title('Daily Meetings - No Data')
        
        # Meeting types distribution
        meeting_type_col = self.column_mapping['meeting_type']
        if meeting_type_col in df.columns and not df[meeting_type_col].empty:
            meeting_counts = df[meeting_type_col].value_counts()
            if len(meeting_counts) > 0:
                meeting_counts.head(5).plot(kind='bar', ax=axes[0,1], color='#00bcf2')
                axes[0,1].set_title('Meeting Types Distribution')
                axes[0,1].tick_params(axis='x', rotation=45)
                axes[0,1].set_ylabel('Number of Meetings')
            else:
                axes[0,1].text(0.5, 0.5, 'No meeting type data', ha='center', va='center')
                axes[0,1].set_title('Meeting Types - No Data')
        else:
            axes[0,1].text(0.5, 0.5, 'No meeting type data', ha='center', va='center')
            axes[0,1].set_title('Meeting Types - No Data')
        
        # Course sections activity
        course_col = self.column_mapping['course']
        if course_col in df.columns and not df[course_col].empty:
            course_counts = df[course_col].value_counts()
            if len(course_counts) > 0:
                course_counts.head(5).plot(kind='bar', ax=axes[1,0], color='#40e0d0')
                axes[1,0].set_title('Active Course Sections')
                axes[1,0].tick_params(axis='x', rotation=45)
                axes[1,0].set_ylabel('Number of Meetings')
            else:
                axes[1,0].text(0.5, 0.5, 'No course data', ha='center', va='center')
                axes[1,0].set_title('Course Activity - No Data')
        else:
            axes[1,0].text(0.5, 0.5, 'No course data', ha='center', va='center')
            axes[1,0].set_title('Course Activity - No Data')
        
        # Hourly meeting pattern
        if len(df) > 0:
            hourly_counts = df.groupby(df['datetime'].dt.hour).size()
            if len(hourly_counts) > 0:
                hourly_counts.plot(kind='bar', ax=axes[1,1], color='#ff6b6b')
                axes[1,1].set_title('Meeting Times by Hour')
                axes[1,1].set_xlabel('Hour of Day')
                axes[1,1].set_ylabel('Number of Meetings')
            else:
                axes[1,1].text(0.5, 0.5, 'No hourly data', ha='center', va='center')
                axes[1,1].set_title('Hourly Pattern - No Data')
        else:
            axes[1,1].text(0.5, 0.5, 'No data', ha='center', va='center')
            axes[1,1].set_title('Hourly Pattern - No Data')
        
        plt.tight_layout()
        plt.savefig('meeting_analytics.png', dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        logging.info("📊 Meeting analytics visualizations created")
    
    def create_email_body(self, stats):
        """Create HTML email body for meeting reports"""
        today = datetime.now().strftime('%B %d, %Y')
        email_provider = self.email_user.split('@')[1] if self.email_user else 'Email'
        
        # Trend analysis
        today_count = stats.get('today_meetings', 0)
        yesterday_count = stats.get('yesterday_meetings', 0)
        
        if today_count > yesterday_count:
            trend = "📈 Increasing"
        elif today_count < yesterday_count:
            trend = "📉 Decreasing"
        else:
            trend = "➡️ Stable"
        
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; color: #333; }}
                .container {{ max-width: 600px; margin: 0 auto; background: #f5f5f5; }}
                .header {{ background: linear-gradient(135deg, #0078d4, #00bcf2); color: white; padding: 30px 20px; text-align: center; }}
                .header h1 {{ margin: 0; font-size: 24px; }}
                .email-badge {{ background: rgba(255,255,255,0.2); padding: 5px 15px; border-radius: 15px; font-size: 12px; margin-top: 10px; display: inline-block; }}
                .content {{ padding: 30px 20px; background: white; }}
                .stats-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin: 20px 0; }}
                .stat-card {{ background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; border: 1px solid #e9ecef; }}
                .stat-number {{ font-size: 28px; font-weight: bold; color: #0078d4; margin: 0; }}
                .stat-label {{ color: #666; font-size: 14px; margin-top: 5px; }}
                .highlight {{ background: linear-gradient(135deg, #e7f3ff, #cce7ff); border-left: 4px solid #0078d4; }}
                .summary-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                .summary-table th {{ background: #f1f3f4; padding: 12px; text-align: left; }}
                .summary-table td {{ padding: 12px; border-bottom: 1px solid #eee; }}
                .footer {{ text-align: center; padding: 20px; background: #f8f9fa; color: #666; font-size: 12px; }}
                .success {{ color: #28a745; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📅 Meeting Analytics Report</h1>
                    <div class="email-badge">📧 via {email_provider}</div>
                    <p>Student Meeting Tracking - {today}</p>
                </div>
                
                <div class="content">
                    <div class="stats-grid">
                        <div class="stat-card highlight">
                            <div class="stat-number">{stats.get('today_meetings', 0)}</div>
                            <div class="stat-label">Meetings Today</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('total_meetings', 0)}</div>
                            <div class="stat-label">Total Meetings</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('unique_students', 0)}</div>
                            <div class="stat-label">Unique Students</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('active_courses', 0)}</div>
                            <div class="stat-label">Active Courses</div>
                        </div>
                    </div>
                    
                    <table class="summary-table">
                        <tr>
                            <th>📊 Meeting Analytics</th>
                            <th>Value</th>
                        </tr>
                        <tr>
                            <td>Daily Trend</td>
                            <td>{trend}</td>
                        </tr>
                        <tr>
                            <td>This Week's Meetings</td>
                            <td>{stats.get('this_week_meetings', 0)}</td>
                        </tr>
                        <tr>
                            <td>Most Popular Meeting Type</td>
                            <td>{stats.get('popular_meeting_type', 'N/A')}</td>
                        </tr>
                        <tr>
                            <td>Peak Meeting Hour</td>
                            <td>{stats.get('peak_hour', 'N/A')}:00</td>
                        </tr>
                        <tr>
                            <td>Average Daily Meetings</td>
                            <td>{stats.get('avg_daily_meetings', 0):.1f}</td>
                        </tr>
                        <tr>
                            <td>Email System Status</td>
                            <td class="success">✅ {self.smtp_server}:{self.smtp_port}</td>
                        </tr>
                    </table>
                    
                    <div style="background: #e8f5e8; padding: 15px; border-radius: 6px; border-left: 4px solid #28a745;">
                        <strong>📎 Attachments Included:</strong><br>
                        • Meeting analytics dashboard with trend visualizations<br>
                        • Complete meeting data export in CSV format<br>
                        • Real-time data from Power Automate integration
                    </div>
                </div>
                
                <div class="footer">
                    <p>🤖 Automated meeting analytics powered by Power Automate & Python</p>
                    <p>Email provider: {email_provider} | SMTP: {self.smtp_server}:{self.smtp_port} | Next report: Tomorrow at 9:00 AM</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html_body
    
    def send_email_report(self, stats, df):
        """Send email report"""
        try:
            logging.info(f"📧 Preparing email report...")
            logging.info(f"🔗 SMTP: {self.smtp_server}:{self.smtp_port}")
            logging.info(f"👥 Recipients: {len(self.recipients)}")
            
            msg = MIMEMultipart()
            msg['From'] = self.email_user
            msg['To'] = ', '.join(self.recipients)
            msg['Subject'] = f'📅 Meeting Analytics Report - {datetime.now().strftime("%B %d, %Y")}'
            
            # Create email body
            html_body = self.create_email_body(stats)
            msg.attach(MIMEText(html_body, 'html'))
            
            # Create and attach CSV
            if df is not None and not df.empty:
                csv_filename = f"meeting_data_{datetime.now().strftime('%Y%m%d')}.csv"
                df.to_csv(csv_filename, index=False)
                
                with open(csv_filename, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={csv_filename}')
                    msg.attach(part)
                
                os.remove(csv_filename)
                logging.info(f"📎 CSV attachment created: {csv_filename}")
            
            # Attach analytics chart
            if os.path.exists('meeting_analytics.png'):
                with open('meeting_analytics.png', 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment; filename=meeting_analytics.png')
                    msg.attach(part)
                
                os.remove('meeting_analytics.png')
                logging.info(f"📊 Chart attachment created")
            
            # Send email
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            
            if self.use_tls:
                server.starttls()
                logging.info("🔒 TLS enabled")
            
            server.login(self.email_user, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"✅ Email sent successfully to {len(self.recipients)} recipients")
            
        except Exception as e:
            logging.error(f"❌ Failed to send email: {e}")
            
            # Provide helpful error messages
            if "authentication failed" in str(e).lower():
                logging.error("💡 Email auth failed - check if you need an app password")
                if "gmail" in self.smtp_server:
                    logging.error("💡 Gmail requires app password: https://support.google.com/accounts/answer/185833")
            elif "connection refused" in str(e).lower():
                logging.error(f"💡 Connection failed - check SMTP server: {self.smtp_server}:{self.smtp_port}")
    
    def run_daily_report(self):
        """Main function to run daily meeting report"""
        try:
            logging.info("🚀 Starting meeting analytics report...")
            logging.info(f"📧 Email: {self.email_user}")
            logging.info(f"🔗 SMTP: {self.smtp_server}:{self.smtp_port}")
            
            # Load data
            df = self.load_data()
            
            if df is None:
                logging.error("❌ Failed to load data")
                return
            
            # Generate statistics
            stats = self.generate_statistics(df)
            logging.info(f"📊 Generated stats: {stats['total_meetings']} total meetings")
            
            # Create visualizations
            self.create_visualizations(df)
            
            # Send email report
            self.send_email_report(stats, df)
            
            logging.info("✅ Daily meeting report completed successfully!")
            
        except Exception as e:
            logging.error(f"💥 Error in daily report: {e}")

def main():
    reporter = MeetingFormsReporter()
    
    if os.getenv('GITHUB_ACTIONS'):
        logging.info("🤖 Running in GitHub Actions")
        reporter.run_daily_report()
    else:
        logging.info("🖥️ Running locally")
        reporter.run_daily_report()

if __name__ == "__main__":
    main()