#!/usr/bin/env python3
"""
Simple Microsoft Forms Email Automation
No API required - works with webhook data or manual exports
"""

import os
import smtplib
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import json
import logging
from pathlib import Path

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class SimpleFormsReporter:
    def __init__(self):
        # Email configuration
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
        self.email_user = os.getenv('EMAIL_USER')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        
        # Recipients
        self.recipients = self.load_recipients()
        
        # Sample data for demonstration
        self.sample_data = self.create_sample_data()
    
    def load_recipients(self):
        """Load email recipients"""
        recipients_env = os.getenv('EMAIL_RECIPIENTS')
        if recipients_env:
            return [email.strip() for email in recipients_env.split(',')]
        return ['example@email.com']
    
    def create_sample_data(self):
        """Create sample data for demonstration"""
        # This simulates form responses - replace with your actual data source
        dates = pd.date_range(start='2024-08-01', end=datetime.now(), freq='D')
        
        sample_responses = []
        for date in dates:
            # Random number of responses per day (0-10)
            num_responses = min(10, max(0, int(abs(hash(str(date)) % 10))))
            
            for i in range(num_responses):
                response_time = date + timedelta(hours=hash(f"{date}{i}") % 24)
                sample_responses.append({
                    'Timestamp': response_time,
                    'Question1': f'Response {i+1}',
                    'Question2': f'Answer {i+1}',
                    'Rating': (hash(f"{date}{i}") % 5) + 1,
                    'Category': ['A', 'B', 'C'][hash(f"{date}{i}") % 3]
                })
        
        return pd.DataFrame(sample_responses)
    
    def load_data_from_csv(self, csv_path):
        """Load data from CSV file (for manual exports)"""
        try:
            if os.path.exists(csv_path):
                df = pd.read_csv(csv_path)
                df['Timestamp'] = pd.to_datetime(df['Timestamp'])
                return df
        except Exception as e:
            logging.error(f"Error loading CSV: {e}")
        
        return self.sample_data
    
    def generate_statistics(self, df):
        """Generate summary statistics"""
        today = datetime.now().date()
        yesterday = today - timedelta(days=1)
        week_ago = datetime.now() - timedelta(days=7)
        
        stats = {
            'total_responses': len(df),
            'today_responses': len(df[df['Timestamp'].dt.date == today]),
            'yesterday_responses': len(df[df['Timestamp'].dt.date == yesterday]),
            'this_week_responses': len(df[df['Timestamp'] >= week_ago]),
            'avg_daily_responses': df.groupby(df['Timestamp'].dt.date).size().mean() if len(df) > 0 else 0,
            'peak_hour': df.groupby(df['Timestamp'].dt.hour).size().idxmax() if len(df) > 0 else 'N/A',
            'most_recent': df['Timestamp'].max() if len(df) > 0 else 'No data'
        }
        
        return stats
    
    def create_visualizations(self, df):
        """Create charts and save as image"""
        plt.figure(figsize=(15, 10))
        plt.style.use('default')
        
        # Daily responses over last 30 days
        plt.subplot(2, 2, 1)
        last_30_days = df[df['Timestamp'] >= datetime.now() - timedelta(days=30)]
        if len(last_30_days) > 0:
            daily_counts = last_30_days.groupby(last_30_days['Timestamp'].dt.date).size()
            daily_counts.plot(kind='line', marker='o', color='#0078d4')
            plt.title('Daily Responses (Last 30 Days)', fontsize=12, fontweight='bold')
            plt.xticks(rotation=45)
            plt.ylabel('Number of Responses')
        else:
            plt.text(0.5, 0.5, 'No data available', ha='center', va='center')
            plt.title('Daily Responses - No Data')
        
        # Hourly distribution
        plt.subplot(2, 2, 2)
        if len(df) > 0:
            hourly_counts = df.groupby(df['Timestamp'].dt.hour).size()
            bars = plt.bar(hourly_counts.index, hourly_counts.values, color='#00bcf2')
            plt.title('Response Distribution by Hour', fontsize=12, fontweight='bold')
            plt.xlabel('Hour of Day')
            plt.ylabel('Number of Responses')
            plt.xticks(range(0, 24, 2))
        else:
            plt.text(0.5, 0.5, 'No data available', ha='center', va='center')
            plt.title('Hourly Distribution - No Data')
        
        # Weekly pattern
        plt.subplot(2, 2, 3)
        if len(df) > 0:
            df['day_of_week'] = df['Timestamp'].dt.day_name()
            weekly_counts = df.groupby('day_of_week').size()
            # Reorder days
            day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            weekly_counts = weekly_counts.reindex([day for day in day_order if day in weekly_counts.index])
            
            bars = plt.bar(range(len(weekly_counts)), weekly_counts.values, color='#40e0d0')
            plt.title('Response Distribution by Day of Week', fontsize=12, fontweight='bold')
            plt.xticks(range(len(weekly_counts)), [day[:3] for day in weekly_counts.index], rotation=45)
            plt.ylabel('Number of Responses')
        else:
            plt.text(0.5, 0.5, 'No data available', ha='center', va='center')
            plt.title('Weekly Pattern - No Data')
        
        # Response trend (last 7 days)
        plt.subplot(2, 2, 4)
        recent_data = df[df['Timestamp'] >= datetime.now() - timedelta(days=7)]
        if len(recent_data) > 0:
            recent_daily = recent_data.groupby(recent_data['Timestamp'].dt.date).size()
            bars = plt.bar(range(len(recent_daily)), recent_daily.values, color='#ff6b6b')
            plt.title('Last 7 Days Responses', fontsize=12, fontweight='bold')
            plt.xticks(range(len(recent_daily)), [d.strftime('%m/%d') for d in recent_daily.index], rotation=45)
            plt.ylabel('Number of Responses')
        else:
            plt.text(0.5, 0.5, 'No data available', ha='center', va='center')
            plt.title('Last 7 Days - No Data')
        
        plt.tight_layout()
        plt.savefig('forms_report.png', dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        logging.info("Visualizations created successfully")
    
    def create_csv_export(self, df):
        """Create CSV export file"""
        filename = f"forms_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(filename, index=False)
        return filename
    
    def create_email_body(self, stats):
        """Create HTML email body"""
        today = datetime.now().strftime('%B %d, %Y')
        
        # Determine trend emoji
        today_count = stats.get('today_responses', 0)
        yesterday_count = stats.get('yesterday_responses', 0)
        
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
                body {{ 
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                    line-height: 1.6; 
                    color: #333;
                    margin: 0;
                    padding: 0;
                }}
                .container {{ max-width: 600px; margin: 0 auto; }}
                .header {{ 
                    background: linear-gradient(135deg, #0078d4, #00bcf2); 
                    color: white; 
                    padding: 30px 20px; 
                    text-align: center; 
                    border-radius: 10px 10px 0 0;
                }}
                .header h1 {{ margin: 0; font-size: 24px; }}
                .header p {{ margin: 10px 0 0 0; opacity: 0.9; }}
                .content {{ padding: 20px; background: #f8f9fa; }}
                .stats-grid {{ 
                    display: grid; 
                    grid-template-columns: 1fr 1fr; 
                    gap: 15px; 
                    margin: 20px 0; 
                }}
                .stat-card {{ 
                    background: white; 
                    padding: 20px; 
                    border-radius: 8px; 
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    text-align: center;
                }}
                .stat-number {{ 
                    font-size: 32px; 
                    font-weight: bold; 
                    color: #0078d4; 
                    margin: 0;
                }}
                .stat-label {{ 
                    color: #666; 
                    font-size: 14px; 
                    margin: 5px 0 0 0;
                }}
                .highlight {{ 
                    background: linear-gradient(135deg, #e7f3ff, #cce7ff); 
                    border-left: 4px solid #0078d4;
                }}
                .summary-table {{ 
                    width: 100%; 
                    border-collapse: collapse; 
                    margin: 20px 0; 
                    background: white;
                    border-radius: 8px;
                    overflow: hidden;
                    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                }}
                .summary-table th {{ 
                    background: #f1f3f4; 
                    padding: 15px; 
                    text-align: left;
                    font-weight: 600;
                }}
                .summary-table td {{ 
                    padding: 12px 15px; 
                    border-bottom: 1px solid #eee;
                }}
                .summary-table tr:last-child td {{ border-bottom: none; }}
                .trend {{ 
                    font-weight: bold; 
                    padding: 5px 10px; 
                    border-radius: 20px; 
                    background: #e7f3ff;
                    color: #0078d4;
                    display: inline-block;
                }}
                .footer {{ 
                    text-align: center;
                    padding: 20px; 
                    font-size: 12px; 
                    color: #666; 
                    background: #f8f9fa;
                    border-radius: 0 0 10px 10px;
                }}
                .attachment-note {{
                    background: #fff3cd;
                    border: 1px solid #ffeaa7;
                    border-radius: 6px;
                    padding: 15px;
                    margin: 20px 0;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📊 Daily Forms Report</h1>
                    <p>Generated on {today}</p>
                </div>
                
                <div class="content">
                    <div class="stats-grid">
                        <div class="stat-card highlight">
                            <div class="stat-number">{stats.get('today_responses', 0)}</div>
                            <div class="stat-label">Responses Today</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('total_responses', 0)}</div>
                            <div class="stat-label">Total Responses</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('this_week_responses', 0)}</div>
                            <div class="stat-label">This Week</div>
                        </div>
                        <div class="stat-card">
                            <div class="stat-number">{stats.get('avg_daily_responses', 0):.1f}</div>
                            <div class="stat-label">Daily Average</div>
                        </div>
                    </div>
                    
                    <table class="summary-table">
                        <tr>
                            <th>📈 Trend Analysis</th>
                            <th>Value</th>
                        </tr>
                        <tr>
                            <td>Daily Trend</td>
                            <td><span class="trend">{trend}</span></td>
                        </tr>
                        <tr>
                            <td>Yesterday's Responses</td>
                            <td>{stats.get('yesterday_responses', 0)}</td>
                        </tr>
                        <tr>
                            <td>Peak Response Hour</td>
                            <td>{stats.get('peak_hour', 'N/A')}:00</td>
                        </tr>
                        <tr>
                            <td>Most Recent Response</td>
                            <td>{stats.get('most_recent', 'No data')}</td>
                        </tr>
                    </table>
                    
                    <div class="attachment-note">
                        <strong>📎 Attachments Included:</strong><br>
                        • Visual charts showing response trends and patterns<br>
                        • Complete data export in CSV format for detailed analysis
                    </div>
                </div>
                
                <div class="footer">
                    <p>🤖 This report was generated automatically by your Forms Automation System</p>
                    <p>For questions or support, contact your system administrator</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html_body
    
    def send_email_report(self, stats, csv_filename):
        """Send email with report and attachments"""
        try:
            msg = MIMEMultipart('alternative')
            msg['From'] = self.email_user
            msg['To'] = ', '.join(self.recipients)
            msg['Subject'] = f'📊 Daily Forms Report - {datetime.now().strftime("%B %d, %Y")}'
            
            # Add HTML body
            html_body = self.create_email_body(stats)
            msg.attach(MIMEText(html_body, 'html'))
            
            # Attach CSV file
            if csv_filename and os.path.exists(csv_filename):
                with open(csv_filename, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename={csv_filename}'
                    )
                    msg.attach(part)
            
            # Attach chart image
            if os.path.exists('forms_report.png'):
                with open('forms_report.png', 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        'attachment; filename=forms_report.png'
                    )
                    msg.attach(part)
            
            # Send email
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            server.login(self.email_user, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"📧 Email report sent successfully to {len(self.recipients)} recipients")
            return True
            
        except Exception as e:
            logging.error(f"❌ Failed to send email: {e}")
            return False
    
    def run_daily_report(self):
        """Main function to generate and send daily report"""
        try:
            logging.info("🚀 Starting daily report generation...")
            
            # Try to load data from CSV file (you can manually export and upload)
            # For GitHub Actions, you could upload the CSV file to the repo
            csv_file = 'forms_data.csv'  # You can upload this file manually
            df = self.load_data_from_csv(csv_file)
            
            # Generate statistics
            stats = self.generate_statistics(df)
            logging.info(f"📊 Generated statistics for {stats['total_responses']} total responses")
            
            # Create visualizations
            self.create_visualizations(df)
            
            # Create CSV export
            csv_filename = self.create_csv_export(df)
            
            # Send email report
            success = self.send_email_report(stats, csv_filename)
            
            # Clean up files
            if csv_filename and os.path.exists(csv_filename):
                os.remove(csv_filename)
            if os.path.exists('forms_report.png'):
                os.remove('forms_report.png')
            
            if success:
                logging.info("✅ Daily report completed successfully!")
            else:
                logging.error("❌ Daily report failed to send")
                
        except Exception as e:
            logging.error(f"💥 Error in daily report generation: {e}")

def main():
    """Main function"""
    reporter = SimpleFormsReporter()
    
    if os.getenv('GITHUB_ACTIONS'):
        logging.info("🤖 Running in GitHub Actions environment")
        reporter.run_daily_report()
    else:
        logging.info("🖥️ Running locally")
        reporter.run_daily_report()

if __name__ == "__main__":
    main()