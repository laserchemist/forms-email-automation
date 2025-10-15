#!/usr/bin/env python3
"""
Enhanced daily meeting analytics with semester-wide statistics
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
        self.email_user = os.getenv('EMAIL_USER')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        self.recipients = [email.strip() for email in os.getenv('EMAIL_RECIPIENTS', '').split(',') if email.strip()]
        
        self.sheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
        
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
        """Load ALL data from Google Sheets with enhanced debugging"""
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
            logging.info(f"📊 Loaded {len(df)} total rows from sheet")
            logging.info(f"📋 Columns found: {df.columns.tolist()}")
            
            # Get column names
            date_col = self.column_mapping['date']
            time_col = self.column_mapping['time']
            
            # Check if columns exist
            if date_col not in df.columns:
                logging.error(f"❌ Date column '{date_col}' not found! Available: {df.columns.tolist()}")
                return pd.DataFrame()
            
            if time_col not in df.columns:
                logging.error(f"❌ Time column '{time_col}' not found! Available: {df.columns.tolist()}")
                return pd.DataFrame()
            
            # Clean the data - strip whitespace
            df[date_col] = df[date_col].astype(str).str.strip()
            df[time_col] = df[time_col].astype(str).str.strip()
            
            # Remove empty rows (where date is empty or 'nan')
            df = df[df[date_col].str.len() > 0]
            df = df[~df[date_col].isin(['', 'nan', 'None', 'NaN'])]
            
            logging.info(f"📊 After filtering empty dates: {len(df)} rows")
            
            # Show sample of what we're trying to parse
            logging.info(f"📅 Sample dates: {df[date_col].head(3).tolist()}")
            logging.info(f"⏰ Sample times: {df[time_col].head(3).tolist()}")
            
            # Try to parse datetime
            # Handle cases where time might be empty
            df['date_time_str'] = df[date_col] + ' ' + df[time_col].replace('', '00:00')
            df['datetime'] = pd.to_datetime(df['date_time_str'], format='%m/%d/%Y %H:%M', errors='coerce')
            
            # Alternative: try without specifying format
            null_mask = df['datetime'].isna()
            if null_mask.sum() > 0:
                logging.info(f"⚠️ Retrying {null_mask.sum()} failed parses with flexible format")
                df.loc[null_mask, 'datetime'] = pd.to_datetime(
                    df.loc[null_mask, 'date_time_str'], 
                    errors='coerce'
                )
            
            # Check parsing results
            valid_dates = df['datetime'].notna().sum()
            invalid_dates = df['datetime'].isna().sum()
            
            logging.info(f"📅 Successfully parsed {valid_dates}/{len(df)} dates")
            
            if invalid_dates > 0:
                logging.warning(f"⚠️ Failed to parse {invalid_dates} dates")
                # Show examples of failed parses
                failed_examples = df[df['datetime'].isna()][['date_time_str']].head(5)
                logging.warning(f"Failed examples:\n{failed_examples}")
            
            if valid_dates == 0:
                logging.error("❌ No valid dates parsed! Check date format in sheet")
                logging.error(f"Sample date values: {df[date_col].head().tolist()}")
                logging.error(f"Sample time values: {df[time_col].head().tolist()}")
                return pd.DataFrame()
            
            # Remove rows with invalid dates
            df = df.dropna(subset=['datetime'])
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            # Drop the temporary column
            df = df.drop('date_time_str', axis=1)
            
            # Show date range
            if len(df) > 0:
                logging.info(f"📆 Date range: {df['datetime'].min()} to {df['datetime'].max()}")
                logging.info(f"✅ Successfully loaded {len(df)} meetings")
            
            return df
            
        except Exception as e:
            logging.error(f"❌ Error loading data: {e}")
            import traceback
            logging.error(traceback.format_exc())
            return None
    
    def generate_statistics(self, df):
        """Generate comprehensive meeting statistics"""
        if df is None or df.empty:
            return {
                'total_meetings': 0,
                'semester_total': 0,
                'today_meetings': 0,
                'yesterday_meetings': 0,
                'this_week_meetings': 0,
                'last_7_days': 0,
                'avg_daily_meetings': 0,
                'avg_weekly_meetings': 0,
                'peak_hour': 'N/A',
                'most_recent': 'No data',
                'popular_meeting_type': 'N/A',
                'active_courses': 0,
                'unique_students': 0,
                'semester_start': 'N/A',
                'days_active': 0
            }
        
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)
        week_ago = datetime.now() - timedelta(days=7)
        
        # Semester-wide stats
        semester_start = df['datetime'].min()
        semester_end = df['datetime'].max()
        days_in_semester = (semester_end - semester_start).days + 1
        
        stats = {
            # Semester totals
            'semester_total': len(df),
            'semester_start': semester_start.strftime('%B %d, %Y'),
            'semester_end': semester_end.strftime('%B %d, %Y'),
            'days_active': days_in_semester,
            
            # Recent activity
            'total_meetings': len(df),
            'today_meetings': len(df[df['datetime'].dt.date == today]),
            'yesterday_meetings': len(df[df['datetime'].dt.date == yesterday]),
            'this_week_meetings': len(df[df['datetime'] >= week_ago]),
            'last_7_days': len(df[df['datetime'] >= week_ago]),
            
            # Averages
            'avg_daily_meetings': len(df) / days_in_semester if days_in_semester > 0 else 0,
            'avg_weekly_meetings': len(df) / (days_in_semester / 7) if days_in_semester > 0 else 0,
            
            # Other stats
            'peak_hour': df.groupby(df['datetime'].dt.hour).size().idxmax() if len(df) > 0 else 'N/A',
            'most_recent': df['datetime'].max() if len(df) > 0 else 'No data'
        }
        
        # Meeting type analysis
        meeting_type_col = self.column_mapping['meeting_type']
        if meeting_type_col in df.columns and not df[meeting_type_col].empty:
            meeting_types = df[meeting_type_col].value_counts()
            stats['popular_meeting_type'] = meeting_types.index[0] if len(meeting_types) > 0 else 'N/A'
            stats['meeting_type_breakdown'] = meeting_types.to_dict()
        else:
            stats['popular_meeting_type'] = 'N/A'
            stats['meeting_type_breakdown'] = {}
        
        # Course activity
        course_col = self.column_mapping['course']
        if course_col in df.columns:
            stats['active_courses'] = df[course_col].nunique()
            stats['course_breakdown'] = df[course_col].value_counts().to_dict()
        else:
            stats['active_courses'] = 0
            stats['course_breakdown'] = {}
        
        # Student activity
        first_name_col = self.column_mapping['first_name']
        last_name_col = self.column_mapping['last_name']
        if first_name_col in df.columns and last_name_col in df.columns:
            df['full_name'] = df[first_name_col].astype(str) + ' ' + df[last_name_col].astype(str)
            stats['unique_students'] = df['full_name'].nunique()
        else:
            stats['unique_students'] = 0
        
        return stats
    
    def create_visualizations(self, df):
        """Create comprehensive analytics visualizations"""
        if df is None or df.empty:
            plt.figure(figsize=(10, 6))
            plt.text(0.5, 0.5, 'No Meeting Data Available', ha='center', va='center', fontsize=16)
            plt.title('Meeting Analytics Dashboard')
            plt.axis('off')
            plt.savefig('meeting_analytics.png', dpi=300, bbox_inches='tight')
            plt.close()
            return
        
        fig, axes = plt.subplots(2, 2, figsize=(15, 10))
        fig.suptitle('Meeting Analytics Dashboard - All Semester', fontsize=16, fontweight='bold')
        
        # 1. Last 7 days trend
        week_ago = datetime.now() - timedelta(days=7)
        last_week = df[df['datetime'] >= week_ago]
        if len(last_week) > 0:
            daily_counts = last_week.groupby(last_week['datetime'].dt.date).size()
            daily_counts.plot(kind='line', ax=axes[0,0], marker='o', color='#0078d4', linewidth=2)
            axes[0,0].set_title('Last 7 Days Trend')
            axes[0,0].set_ylabel('Meetings per Day')
            axes[0,0].tick_params(axis='x', rotation=45)
            axes[0,0].grid(True, alpha=0.3)
        else:
            axes[0,0].text(0.5, 0.5, 'No meetings in last 7 days', ha='center', va='center')
            axes[0,0].set_title('Last 7 Days - No Data')
        
        # 2. Semester cumulative trend
        df_sorted = df.sort_values('datetime')
        df_sorted['cumulative'] = range(1, len(df_sorted) + 1)
        axes[0,1].plot(df_sorted['datetime'], df_sorted['cumulative'], color='#00bcf2', linewidth=2)
        axes[0,1].set_title('Cumulative Meetings - All Semester')
        axes[0,1].set_ylabel('Total Meetings')
        axes[0,1].tick_params(axis='x', rotation=45)
        axes[0,1].grid(True, alpha=0.3)
        
        # 3. Course section breakdown (all semester)
        course_col = self.column_mapping['course']
        if course_col in df.columns and not df[course_col].empty:
            course_counts = df[course_col].value_counts().head(10)
            course_counts.plot(kind='barh', ax=axes[1,0], color='#40e0d0')
            axes[1,0].set_title('Meetings by Course Section (All Semester)')
            axes[1,0].set_xlabel('Number of Meetings')
        else:
            axes[1,0].text(0.5, 0.5, 'No course data', ha='center', va='center')
            axes[1,0].set_title('Course Sections - No Data')
        
        # 4. Meeting type distribution (all semester)
        meeting_type_col = self.column_mapping['meeting_type']
        if meeting_type_col in df.columns and not df[meeting_type_col].empty:
            meeting_counts = df[meeting_type_col].value_counts()
            meeting_counts.plot(kind='pie', ax=axes[1,1], autopct='%1.1f%%', startangle=90)
            axes[1,1].set_title('Meeting Types (All Semester)')
            axes[1,1].set_ylabel('')
        else:
            axes[1,1].text(0.5, 0.5, 'No meeting type data', ha='center', va='center')
            axes[1,1].set_title('Meeting Types - No Data')
        
        plt.tight_layout()
        plt.savefig('meeting_analytics.png', dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        logging.info("📊 Created comprehensive analytics visualizations")
    
    def create_email_body(self, stats):
        """Create enhanced HTML email body with semester statistics"""
        today = datetime.now().strftime('%B %d, %Y')
        
        # Trend analysis
        today_count = stats.get('today_meetings', 0)
        yesterday_count = stats.get('yesterday_meetings', 0)
        
        if today_count > yesterday_count:
            trend = "📈 Increasing"
        elif today_count < yesterday_count:
            trend = "📉 Decreasing"
        else:
            trend = "➡️ Stable"
        
        # Course breakdown HTML
        course_breakdown_html = ""
        course_breakdown = stats.get('course_breakdown', {})
        if course_breakdown:
            course_breakdown_html = "<ul style='list-style: none; padding: 0; margin: 10px 0;'>"
            for course, count in sorted(course_breakdown.items(), key=lambda x: x[1], reverse=True)[:10]:
                course_breakdown_html += f"<li style='padding: 4px 0;'><strong>{course}:</strong> {count} meetings</li>"
            course_breakdown_html += "</ul>"
        
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; color: #333; }}
                .container {{ max-width: 700px; margin: 0 auto; background: #f5f5f5; }}
                .header {{ background: #2c5282; color: #dc2626; padding: 30px 20px; text-align: center; }}
                .header h1 {{ margin: 0; font-size: 24px; }}
                .content {{ padding: 30px 20px; background: white; }}
                .stats-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 15px; margin: 20px 0; }}
                .stat-card {{ background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; border: 1px solid #e9ecef; }}
                .stat-number {{ font-size: 28px; font-weight: bold; color: #0078d4; margin: 0; }}
                .stat-label {{ color: #666; font-size: 14px; margin-top: 5px; }}
                .highlight {{ background: linear-gradient(135deg, #e7f3ff, #cce7ff); border-left: 4px solid #0078d4; }}
                .semester-section {{ background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 6px; }}
                .section {{ background: #f8f9fa; padding: 15px; margin: 15px 0; border-radius: 6px; border-left: 4px solid #28a745; }}
                .summary-table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                .summary-table th {{ background: #f1f3f4; padding: 12px; text-align: left; }}
                .summary-table td {{ padding: 12px; border-bottom: 1px solid #eee; }}
                .footer {{ text-align: center; padding: 20px; background: #f8f9fa; color: #666; font-size: 12px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📅 Daily Meeting Analytics</h1>
                    <p>Student Meeting Tracking - {today}</p>
                </div>
                
                <div class="content">
                    <div class="semester-section">
                        <h3 style="margin: 0 0 10px 0;">🎓 Semester Overview</h3>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
                            <div><strong>Total Meetings:</strong> {stats.get('semester_total', 0)}</div>
                            <div><strong>Unique Students:</strong> {stats.get('unique_students', 0)}</div>
                            <div><strong>Semester Start:</strong> {stats.get('semester_start', 'N/A')}</div>
                            <div><strong>Days Active:</strong> {stats.get('days_active', 0)}</div>
                            <div><strong>Avg/Day:</strong> {stats.get('avg_daily_meetings', 0):.1f}</div>
                            <div><strong>Avg/Week:</strong> {stats.get('avg_weekly_meetings', 0):.1f}</div>
                        </div>
                    </div>
                    
                    <div class="stats-grid">
                        <div class="stat-card highlight">
                            <div class="stat-number">{stats.get('today_meetings', 0)}</div>
                            <div class="stat-label">Meetings Today</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('last_7_days', 0)}</div>
                            <div class="stat-label">Last 7 Days</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('active_courses', 0)}</div>
                            <div class="stat-label">Active Sections</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{trend}</div>
                            <div class="stat-label">Trend</div>
                        </div>
                    </div>
                    
                    <div class="section">
                        <h3 style="margin: 0 0 10px 0;">📊 Course Section Attendance</h3>
                        {course_breakdown_html}
                    </div>
                    
                    <table class="summary-table">
                        <tr>
                            <th>Metric</th>
                            <th>Value</th>
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
                            <td>Yesterday's Meetings</td>
                            <td>{stats.get('yesterday_meetings', 0)}</td>
                        </tr>
                    </table>
                    
                    <div style="background: #e8f5e8; padding: 15px; border-radius: 6px; border-left: 4px solid #28a745;">
                        <strong>🔎 Attachments:</strong><br>
                        • Semester-wide analytics dashboard<br>
                        • 7-day trend analysis<br>
                        • Complete meeting data (CSV)
                    </div>
                </div>
                
                <div class="footer">
                    <p>🤖 Automated daily report | Generated {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html_body
    
    def send_email_report(self, stats, df):
        """Send email report"""
        try:
            logging.info(f"📧 Sending daily report via Gmail...")
            
            msg = MIMEMultipart()
            msg['From'] = self.email_user
            msg['To'] = ', '.join(self.recipients)
            msg['Subject'] = f'📅 Daily Meeting Analytics - {stats.get("semester_total", 0)} Total Meetings - {datetime.now().strftime("%B %d, %Y")}'
            
            html_body = self.create_email_body(stats)
            msg.attach(MIMEText(html_body, 'html'))
            
            # Attach CSV
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
            
            # Attach chart
            if os.path.exists('meeting_analytics.png'):
                with open('meeting_analytics.png', 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment; filename=meeting_analytics.png')
                    msg.attach(part)
                
                os.remove('meeting_analytics.png')
            
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.email_user, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"✅ Email sent to {len(self.recipients)} recipients")
            
        except Exception as e:
            logging.error(f"❌ Failed to send email: {e}")
    
    def run_daily_report(self):
        """Main function"""
        try:
            logging.info("🚀 Starting enhanced daily meeting analytics...")
            
            df = self.load_data()
            
            if df is None or df.empty:
                logging.error("❌ No data loaded")
                return
            
            stats = self.generate_statistics(df)
            logging.info(f"📊 Stats: {stats['semester_total']} total, {stats['today_meetings']} today, {stats['last_7_days']} this week")
            
            self.create_visualizations(df)
            self.send_email_report(stats, df)
            
            logging.info("✅ Daily report completed!")
            
        except Exception as e:
            logging.error(f"💥 Error: {e}")
            import traceback
            logging.error(traceback.format_exc())

def main():
    reporter = MeetingFormsReporter()
    
    if os.getenv('GITHUB_ACTIONS'):
        logging.info("🤖 Running in GitHub Actions")
    else:
        logging.info("🖥️ Running locally")
    
    reporter.run_daily_report()

if __name__ == "__main__":
    main()
