#!/usr/bin/env python3
"""
Enhanced weekly instructor reports with cumulative semester statistics
Includes detailed section attendance and semester-wide analytics
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
from pathlib import Path
from wordcloud import WordCloud
import re

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class EnhancedInstructorReporter:
    def __init__(self):
        self.email_user = os.getenv('EMAIL_USER')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        
        self.sheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
        
        self.instructor_mappings = self.load_instructor_config()
        self.config_metadata = {}
        
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
    
    def load_instructor_config(self):
        """Load instructor mappings from CSV"""
        mappings = {}
        csv_file = Path('instructor_mappings.csv')
        
        if csv_file.exists():
            try:
                df = pd.read_csv(csv_file)
                logging.info(f"📋 Loading instructor mappings from {csv_file}")
                
                for _, row in df.iterrows():
                    if row.get('active', True) and str(row.get('active', True)).lower() != 'false':
                        course_section = row['course_section'].strip()
                        mappings[course_section] = {
                            'instructor': row['instructor_name'].strip(),
                            'email': row['email'].strip(),
                            'course_name': row.get('course_name', '').strip(),
                            'active': True
                        }
                
                logging.info(f"✅ Loaded {len(mappings)} active instructor mappings")
                return mappings
                
            except Exception as e:
                logging.warning(f"⚠️ Error reading CSV config: {e}")
        else:
            self.create_sample_config()
            logging.info("📝 Created sample instructor_mappings.csv")
            return {}
    
    def create_sample_config(self):
        """Create sample configuration file"""
        sample_data = [
            ['course_section', 'instructor_name', 'email', 'course_name', 'active'],
            ['M/W/F 12:00', 'Dr. Smith', 'dr.smith@university.edu', 'General Chemistry I', 'true'],
            ['T/R 9:30 AM', 'Prof. Johnson', 'prof.johnson@university.edu', 'Organic Chemistry', 'true'],
        ]
        
        with open('instructor_mappings.csv', 'w', newline='') as f:
            import csv
            writer = csv.writer(f)
            writer.writerows(sample_data)
    
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
        """Load ALL semester data from Google Sheets"""
        sheet = self.connect_to_sheets()
        if not sheet:
            return None
        
        try:
            values = sheet.get_all_values()
            
            if len(values) < 2:
                logging.warning("Sheet has no data rows")
                return pd.DataFrame()
            
            df = pd.DataFrame(values[1:], columns=values[0])
            logging.info(f"📊 Loaded {len(df)} total rows")
            
            date_col = self.column_mapping['date']
            time_col = self.column_mapping['time']
            
            if date_col in df.columns and time_col in df.columns:
                df['datetime'] = pd.to_datetime(df[date_col] + ' ' + df[time_col], errors='coerce')
            elif date_col in df.columns:
                df['datetime'] = pd.to_datetime(df[date_col], errors='coerce')
            
            df = df.dropna(subset=['datetime'])
            df = df.dropna(how='all')
            
            if len(df) > 0:
                logging.info(f"📅 Date range: {df['datetime'].min()} to {df['datetime'].max()}")
            
            return df
            
        except Exception as e:
            logging.error(f"❌ Error loading data: {e}")
            return None
    
    def get_weekly_data(self, df):
        """Filter data for the past week"""
        if df is None or df.empty:
            return df
        
        week_ago = datetime.now() - timedelta(days=7)
        weekly_df = df[df['datetime'] >= week_ago].copy()
        
        logging.info(f"📅 Weekly data: {len(weekly_df)} meetings in last 7 days")
        return weekly_df
    
    def group_by_instructor(self, df_all, df_weekly):
        """Group meetings by instructor with semester and weekly stats"""
        if df_all is None or df_all.empty:
            return {}
        
        course_col = self.column_mapping['course']
        instructor_groups = {}
        
        for course_section in df_all[course_col].unique():
            if pd.isna(course_section) or course_section == '':
                continue
            
            instructor_info = self.instructor_mappings.get(course_section.strip())
            if not instructor_info:
                logging.warning(f"⚠️ No instructor for section: '{course_section}'")
                continue
            
            instructor_name = instructor_info['instructor']
            
            # Get ALL semester data for this section
            section_all_data = df_all[df_all[course_col] == course_section].copy()
            
            # Get weekly data for this section
            section_weekly_data = df_weekly[df_weekly[course_col] == course_section].copy() if df_weekly is not None and not df_weekly.empty else pd.DataFrame()
            
            if instructor_name not in instructor_groups:
                instructor_groups[instructor_name] = {
                    'email': instructor_info['email'],
                    'course_name': instructor_info.get('course_name', ''),
                    'sections': [],
                    'semester_data': pd.DataFrame(),
                    'weekly_data': pd.DataFrame(),
                    'section_stats': {}
                }
            
            # Add section
            instructor_groups[instructor_name]['sections'].append(course_section)
            
            # Combine semester data
            instructor_groups[instructor_name]['semester_data'] = pd.concat([
                instructor_groups[instructor_name]['semester_data'], 
                section_all_data
            ], ignore_index=True)
            
            # Combine weekly data
            instructor_groups[instructor_name]['weekly_data'] = pd.concat([
                instructor_groups[instructor_name]['weekly_data'], 
                section_weekly_data
            ], ignore_index=True)
            
            # Section-specific statistics
            instructor_groups[instructor_name]['section_stats'][course_section] = {
                'semester_meetings': len(section_all_data),
                'semester_students': len(section_all_data.groupby([
                    self.column_mapping['first_name'], 
                    self.column_mapping['last_name']
                ])),
                'weekly_meetings': len(section_weekly_data),
                'weekly_students': len(section_weekly_data.groupby([
                    self.column_mapping['first_name'], 
                    self.column_mapping['last_name']
                ])) if len(section_weekly_data) > 0 else 0
            }
        
        # Calculate totals for each instructor
        for instructor_name in instructor_groups:
            data = instructor_groups[instructor_name]
            data['total_semester_meetings'] = len(data['semester_data'])
            data['total_weekly_meetings'] = len(data['weekly_data'])
            data['unique_semester_students'] = len(data['semester_data'].groupby([
                self.column_mapping['first_name'], 
                self.column_mapping['last_name']
            ])) if len(data['semester_data']) > 0 else 0
            data['unique_weekly_students'] = len(data['weekly_data'].groupby([
                self.column_mapping['first_name'], 
                self.column_mapping['last_name']
            ])) if len(data['weekly_data']) > 0 else 0
        
        logging.info(f"👨‍🏫 Grouped data for {len(instructor_groups)} instructors")
        return instructor_groups
    
    def generate_instructor_statistics(self, instructor_data):
        """Generate comprehensive statistics"""
        semester_df = instructor_data['semester_data']
        weekly_df = instructor_data['weekly_data']
        
        if semester_df.empty:
            return {}
        
        # Calculate semester range
        semester_start = semester_df['datetime'].min()
        semester_end = semester_df['datetime'].max()
        days_in_semester = (semester_end - semester_start).days + 1
        
        stats = {
            # Semester totals
            'total_semester_meetings': instructor_data['total_semester_meetings'],
            'unique_semester_students': instructor_data['unique_semester_students'],
            'semester_start': semester_start.strftime('%B %d, %Y'),
            'semester_end': semester_end.strftime('%B %d, %Y'),
            'days_active': days_in_semester,
            'avg_daily_semester': instructor_data['total_semester_meetings'] / days_in_semester if days_in_semester > 0 else 0,
            
            # Weekly totals
            'total_weekly_meetings': instructor_data['total_weekly_meetings'],
            'unique_weekly_students': instructor_data['unique_weekly_students'],
            
            # Section breakdown
            'sections': instructor_data['sections'],
            'section_count': len(instructor_data['sections']),
            'section_breakdown': instructor_data['section_stats'],
            'course_name': instructor_data.get('course_name', '')
        }
        
        # Weekly meeting type breakdown
        if not weekly_df.empty:
            meeting_type_col = self.column_mapping['meeting_type']
            if meeting_type_col in weekly_df.columns:
                stats['weekly_meeting_types'] = weekly_df[meeting_type_col].value_counts().to_dict()
            
            # Daily breakdown for this week
            daily_counts = weekly_df.groupby(weekly_df['datetime'].dt.date).size()
            stats['daily_counts'] = daily_counts.to_dict()
            stats['busiest_day'] = daily_counts.idxmax() if len(daily_counts) > 0 else 'N/A'
            stats['avg_daily_week'] = daily_counts.mean() if len(daily_counts) > 0 else 0
        else:
            stats['weekly_meeting_types'] = {}
            stats['daily_counts'] = {}
            stats['busiest_day'] = 'No meetings this week'
            stats['avg_daily_week'] = 0
        
        return stats
    
    def create_instructor_visualization(self, instructor_name, instructor_data, stats):
        """Create enhanced visualization with semester data"""
        semester_df = instructor_data['semester_data']
        weekly_df = instructor_data['weekly_data']
        
        if semester_df.empty:
            return None
        
        fig, axes = plt.subplots(2, 2, figsize=(14, 10))
        fig.suptitle(f'{instructor_name} - Weekly Report + Semester Overview', 
                     fontsize=14, fontweight='bold')
        
        # 1. Semester cumulative trend
        semester_sorted = semester_df.sort_values('datetime')
        semester_sorted['cumulative'] = range(1, len(semester_sorted) + 1)
        axes[0,0].plot(semester_sorted['datetime'], semester_sorted['cumulative'], 
                      color='#0078d4', linewidth=2)
        axes[0,0].set_title(f'Cumulative Meetings - All Semester ({len(semester_df)} total)')
        axes[0,0].set_ylabel('Total Meetings')
        axes[0,0].tick_params(axis='x', rotation=45)
        axes[0,0].grid(True, alpha=0.3)
        
        # 2. Section breakdown (semester totals)
        section_breakdown = stats.get('section_breakdown', {})
        if section_breakdown:
            sections = list(section_breakdown.keys())
            semester_counts = [section_breakdown[s]['semester_meetings'] for s in sections]
            weekly_counts = [section_breakdown[s]['weekly_meetings'] for s in sections]
            
            x = range(len(sections))
            width = 0.35
            axes[0,1].bar([i - width/2 for i in x], semester_counts, width, 
                         label='Semester Total', color='#00bcf2', alpha=0.7)
            axes[0,1].bar([i + width/2 for i in x], weekly_counts, width, 
                         label='This Week', color='#40e0d0')
            axes[0,1].set_title('Section Attendance: Semester vs This Week')
            axes[0,1].set_ylabel('Number of Meetings')
            axes[0,1].set_xticks(x)
            axes[0,1].set_xticklabels([s.split()[0][:10] for s in sections], rotation=45)
            axes[0,1].legend()
        else:
            axes[0,1].text(0.5, 0.5, 'No section data', ha='center', va='center')
        
        # 3. Daily meetings this week
        if not weekly_df.empty:
            daily_counts = pd.Series(stats.get('daily_counts', {}))
            if len(daily_counts) > 0:
                daily_counts.plot(kind='bar', ax=axes[1,0], color='#667eea')
                axes[1,0].set_title('Daily Meetings This Week')
                axes[1,0].set_xlabel('Date')
                axes[1,0].set_ylabel('Number of Meetings')
                axes[1,0].tick_params(axis='x', rotation=45)
            else:
                axes[1,0].text(0.5, 0.5, 'No meetings this week', ha='center', va='center')
        else:
            axes[1,0].text(0.5, 0.5, 'No meetings this week', ha='center', va='center')
        
        # 4. Student attendance by section (semester)
        if section_breakdown:
            student_counts = [section_breakdown[s]['semester_students'] for s in sections]
            axes[1,1].barh(range(len(sections)), student_counts, color='#ff6b6b')
            axes[1,1].set_title('Unique Students per Section (Semester)')
            axes[1,1].set_xlabel('Number of Students')
            axes[1,1].set_yticks(range(len(sections)))
            axes[1,1].set_yticklabels([s.split()[0][:10] for s in sections])
        else:
            axes[1,1].text(0.5, 0.5, 'No student data', ha='center', va='center')
        
        plt.tight_layout()
        
        filename = f'weekly_report_{instructor_name.replace(" ", "_").replace(".", "")}.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        return filename
    
    def create_wordcloud_from_topics(self, instructor_name, df):
        """Create wordcloud from meeting topics"""
        try:
            if df.empty:
                return None
            
            topic_col = self.column_mapping['topic']
            topics = df[topic_col].dropna().astype(str)
            topics = topics[topics != '']
            topics = topics[topics.str.lower() != 'not specified']
            
            if len(topics) == 0:
                return None
            
            all_topics_text = ' '.join(topics)
            all_topics_text = re.sub(r'[^\w\s]', ' ', all_topics_text)
            all_topics_text = ' '.join(all_topics_text.split())
            
            if len(all_topics_text.strip()) == 0:
                return None
            
            wordcloud = WordCloud(
                width=800, height=400, background_color='white',
                max_words=50, colormap='viridis', relative_scaling=0.5,
                min_font_size=12, max_font_size=72
            ).generate(all_topics_text)
            
            filename = f'wordcloud_{instructor_name.replace(" ", "_").replace(".", "")}.png'
            
            plt.figure(figsize=(12, 6))
            plt.imshow(wordcloud, interpolation='bilinear')
            plt.axis('off')
            plt.title(f'Meeting Topics - {instructor_name}', fontsize=16, fontweight='bold', pad=20)
            plt.tight_layout(pad=1)
            plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
            plt.close()
            
            logging.info(f"✅ Created wordcloud for {instructor_name}")
            return filename
            
        except Exception as e:
            logging.error(f"❌ Error creating wordcloud: {e}")
            return None
    
    def create_instructor_email_body(self, instructor_name, stats):
        """Create enhanced email with semester statistics"""
        week_start = (datetime.now() - timedelta(days=7)).strftime('%B %d')
        week_end = datetime.now().strftime('%B %d, %Y')
        
        # Section breakdown table
        sections_html = ""
        section_breakdown = stats.get('section_breakdown', {})
        if section_breakdown:
            sections_html = """
            <table style='width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px;'>
                <tr style='background: #f1f3f4;'>
                    <th style='padding: 10px; border: 1px solid #ddd; text-align: left;'>Section</th>
                    <th style='padding: 10px; border: 1px solid #ddd;'>Semester Total</th>
                    <th style='padding: 10px; border: 1px solid #ddd;'>Semester Students</th>
                    <th style='padding: 10px; border: 1px solid #ddd;'>This Week</th>
                    <th style='padding: 10px; border: 1px solid #ddd;'>Week Students</th>
                </tr>
            """
            for section, section_stats in section_breakdown.items():
                sections_html += f"""
                <tr>
                    <td style='padding: 8px; border: 1px solid #ddd;'><strong>{section}</strong></td>
                    <td style='padding: 8px; border: 1px solid #ddd; text-align: center;'>{section_stats['semester_meetings']}</td>
                    <td style='padding: 8px; border: 1px solid #ddd; text-align: center;'>{section_stats['semester_students']}</td>
                    <td style='padding: 8px; border: 1px solid #ddd; text-align: center; background: #e8f5e8;'>{section_stats['weekly_meetings']}</td>
                    <td style='padding: 8px; border: 1px solid #ddd; text-align: center; background: #e8f5e8;'>{section_stats['weekly_students']}</td>
                </tr>
                """
            sections_html += "</table>"
        
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; color: #333; }}
                .container {{ max-width: 800px; margin: 0 auto; background: #f5f5f5; }}
                .header {{ background: #2c3e50; color: white; padding: 30px 20px; text-align: center; }}
                .content {{ padding: 30px 20px; background: white; }}
                .stats-grid {{ display: table; width: 100%; margin: 20px 0; }}
                .stats-row {{ display: table-row; }}
                .stat-card {{ display: table-cell; background: #f8f9fa; padding: 15px; text-align: center; border: 1px solid #e9ecef; }}
                .stat-number {{ font-size: 24px; font-weight: bold; color: #2c3e50; }}
                .stat-label {{ color: #666; font-size: 14px; margin-top: 5px; }}
                .semester-section {{ background: #fff3cd; border-left: 4px solid #ffc107; padding: 20px; margin: 20px 0; border-radius: 6px; }}
                .section {{ margin: 20px 0; padding: 20px; border-left: 4px solid #3498db; background: #f8f9ff; border-radius: 6px; }}
                .footer {{ text-align: center; padding: 20px; background: #f8f9fa; color: #666; font-size: 12px; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📚 Weekly Meeting Report</h1>
                    <p><strong>{instructor_name}</strong></p>
                    <p>{stats.get('course_name', 'N/A')} | Week of {week_start} - {week_end}</p>
                </div>
                
                <div class="content">
                    <div class="semester-section">
                        <h3 style="margin: 0 0 15px 0;">🎓 Semester Overview</h3>
                        <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px;">
                            <div>
                                <div style="font-size: 28px; font-weight: bold; color: #2c3e50;">{stats.get('total_semester_meetings', 0)}</div>
                                <div style="color: #666; font-size: 14px;">Total Meetings</div>
                            </div>
                            <div>
                                <div style="font-size: 28px; font-weight: bold; color: #2c3e50;">{stats.get('unique_semester_students', 0)}</div>
                                <div style="color: #666; font-size: 14px;">Unique Students</div>
                            </div>
                            <div>
                                <div style="font-size: 28px; font-weight: bold; color: #2c3e50;">{stats.get('avg_daily_semester', 0):.1f}</div>
                                <div style="color: #666; font-size: 14px;">Avg/Day</div>
                            </div>
                        </div>
                        <p style="margin: 10px 0 0 0; font-size: 13px; color: #666;">
                            <strong>Period:</strong> {stats.get('semester_start', 'N/A')} to {stats.get('semester_end', 'N/A')} 
                            ({stats.get('days_active', 0)} days)
                        </p>
                    </div>
                    
                    <div class="stats-grid">
                        <div class="stats-row">
                            <div class="stat-card" style="background: #e8f5e8;">
                                <div class="stat-number" style="color: #27ae60;">{stats.get('total_weekly_meetings', 0)}</div>
                                <div class="stat-label">This Week's Meetings</div>
                            </div>
                            <div class="stat-card" style="background: #e8f5e8;">
                                <div class="stat-number" style="color: #27ae60;">{stats.get('unique_weekly_students', 0)}</div>
                                <div class="stat-label">Students This Week</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-number">{stats.get('section_count', 0)}</div>
                                <div class="stat-label">Your Sections</div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="section">
                        <h3 style="margin: 0 0 15px 0;">📊 Section Attendance Breakdown</h3>
                        {sections_html}
                    </div>
                    
                    <div class="section">
                        <h3 style="margin: 0 0 10px 0;">📈 This Week's Activity</h3>
                        <p><strong>Busiest day:</strong> {stats.get('busiest_day', 'N/A')}</p>
                        <p><strong>Average daily meetings:</strong> {stats.get('avg_daily_week', 0):.1f}</p>
                    </div>
                    
                    <div style="background: #e8f5e8; padding: 15px; border-radius: 6px; border-left: 4px solid #27ae60; margin: 20px 0;">
                        <strong>📎 Attachments:</strong><br>
                        • Semester-wide analytics dashboard<br>
                        • Meeting topics wordcloud<br>
                        • Section attendance breakdown<br>
                        • Complete meeting data (CSV)
                    </div>
                </div>
                
                <div class="footer">
                    <p>🤖 Automated weekly report | Generated {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html_body
    
    def send_instructor_email(self, instructor_name, instructor_email, instructor_data, stats, 
                             chart_filename, wordcloud_filename=None):
        """Send email to instructor"""
        try:
            logging.info(f"📧 Sending email to {instructor_name}")
            
            msg = MIMEMultipart()
            msg['From'] = self.email_user
            msg['To'] = instructor_email
            msg['Subject'] = f'📚 Weekly Report - {instructor_name} - {stats.get("total_semester_meetings", 0)} Semester Meetings'
            
            html_body = self.create_instructor_email_body(instructor_name, stats)
            msg.attach(MIMEText(html_body, 'html'))
            
            # Attach chart
            if chart_filename and os.path.exists(chart_filename):
                with open(chart_filename, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={chart_filename}')
                    msg.attach(part)
            
            # Attach wordcloud
            if wordcloud_filename and os.path.exists(wordcloud_filename):
                with open(wordcloud_filename, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={wordcloud_filename}')
                    msg.attach(part)
            
            # Create CSV with meeting details
            csv_filename = f"meetings_{instructor_name.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.csv"
            if not instructor_data['weekly_data'].empty:
                instructor_data['weekly_data'].to_csv(csv_filename, index=False)
                
                with open(csv_filename, 'rb') as f:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={csv_filename}')
                    msg.attach(part)
            
            # Send email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.email_user, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"✅ Email sent to {instructor_name}")
            
            # Clean up files
            for filename in [chart_filename, wordcloud_filename, csv_filename]:
                if filename and os.path.exists(filename):
                    os.remove(filename)
            
            return True
            
        except Exception as e:
            logging.error(f"❌ Failed to send email to {instructor_name}: {e}")
            return False
    
    def run_weekly_report(self):
        """Main function"""
        try:
            logging.info("🚀 Starting enhanced weekly instructor reports...")
            
            if not self.instructor_mappings:
                logging.error("❌ No instructor mappings!")
                return
            
            # Load ALL semester data
            df_all = self.load_data()
            if df_all is None or df_all.empty:
                logging.error("❌ No data loaded")
                return
            
            # Get weekly subset
            df_weekly = self.get_weekly_data(df_all)
            
            # Group by instructor
            instructor_groups = self.group_by_instructor(df_all, df_weekly)
            if not instructor_groups:
                logging.warning("⚠️ No instructor groups found")
                return
            
            success_count = 0
            
            # Process each instructor
            for instructor_name, instructor_data in instructor_groups.items():
                logging.info(f"👨‍🏫 Processing {instructor_name}...")
                
                stats = self.generate_instructor_statistics(instructor_data)
                
                chart_filename = self.create_instructor_visualization(
                    instructor_name, instructor_data, stats
                )
                
                wordcloud_filename = self.create_wordcloud_from_topics(
                    instructor_name, instructor_data['weekly_data']
                )
                
                if self.send_instructor_email(
                    instructor_name, 
                    instructor_data['email'],
                    instructor_data,
                    stats, 
                    chart_filename, 
                    wordcloud_filename
                ):
                    success_count += 1
            
            logging.info(f"✅ Weekly reports completed! Sent {success_count}/{len(instructor_groups)} emails")
            
        except Exception as e:
            logging.error(f"💥 Error: {e}")
            import traceback
            logging.error(traceback.format_exc())

def main():
    reporter = EnhancedInstructorReporter()
    
    if os.getenv('GITHUB_ACTIONS'):
        logging.info("🤖 Running in GitHub Actions")
    else:
        logging.info("🖥️ Running locally")
    
    reporter.run_weekly_report()

if __name__ == "__main__":
    main()
