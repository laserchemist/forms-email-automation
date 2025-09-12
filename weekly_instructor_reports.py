#!/usr/bin/env python3
"""
Configuration-based weekly instructor summary system
Reads instructor mappings from editable config files
Enhanced with wordcloud generation and Outlook-compatible styling
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
from collections import Counter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class ConfigBasedInstructorReporter:
    def __init__(self):
        # Email configuration
        self.email_user = os.getenv('EMAIL_USER')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        
        # Google Sheets configuration
        self.sheet_id = os.getenv('GOOGLE_SHEET_ID')
        self.credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
        
        # Load instructor mappings from config files
        self.instructor_mappings = self.load_instructor_config()
        self.config_metadata = {}
        
        # Column mapping
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
        """Load instructor mappings from config files (CSV preferred, JSON fallback)"""
        mappings = {}
        
        # Try CSV first (easiest to edit)
        csv_file = Path('instructor_mappings.csv')
        json_file = Path('instructor_mappings.json')
        
        if csv_file.exists():
            try:
                df = pd.read_csv(csv_file)
                logging.info(f"📋 Loading instructor mappings from {csv_file}")
                
                for _, row in df.iterrows():
                    # Only include active instructors
                    if row.get('active', True) and str(row.get('active', True)).lower() != 'false':
                        course_section = row['course_section'].strip()
                        mappings[course_section] = {
                            'instructor': row['instructor_name'].strip(),
                            'email': row['email'].strip(),
                            'course_name': row.get('course_name', '').strip(),
                            'active': True
                        }
                
                logging.info(f"✅ Loaded {len(mappings)} active instructor mappings from CSV")
                return mappings
                
            except Exception as e:
                logging.warning(f"⚠️ Error reading CSV config: {e}")
        
        # Try JSON as fallback
        elif json_file.exists():
            try:
                with open(json_file, 'r') as f:
                    config = json.load(f)
                
                logging.info(f"📋 Loading instructor mappings from {json_file}")
                
                self.config_metadata = {
                    'semester': config.get('semester', 'Unknown'),
                    'last_updated': config.get('last_updated', 'Unknown')
                }
                
                mappings = config.get('course_mappings', {})
                logging.info(f"✅ Loaded {len(mappings)} instructor mappings from JSON")
                logging.info(f"📅 Configuration: {self.config_metadata['semester']} (Updated: {self.config_metadata['last_updated']})")
                
                return mappings
                
            except Exception as e:
                logging.warning(f"⚠️ Error reading JSON config: {e}")
        
        # If no config files found, create a sample CSV
        else:
            logging.warning("⚠️ No instructor mapping config file found!")
            self.create_sample_config()
            logging.info("📝 Created sample instructor_mappings.csv file")
            logging.info("🔧 Please edit instructor_mappings.csv with your actual instructor information")
            
            # Return empty mappings - user needs to configure
            return {}
    
    def create_sample_config(self):
        """Create sample configuration file"""
        sample_data = [
            ['course_section', 'instructor_name', 'email', 'course_name', 'active'],
            ['M/W/F 12:00', 'Dr. Smith', 'dr.smith@university.edu', 'General Chemistry I', 'true'],
            ['M/W 5:30 PM', 'Prof. Johnson', 'prof.johnson@university.edu', 'Organic Chemistry', 'true'],
            ['T/R 8:00 AM', 'Dr. Wilson', 'dr.wilson@university.edu', 'Physical Chemistry', 'true'],
            ['T/R 9:30 AM Nyquist', 'Prof. Nyquist', 'prof.nyquist@university.edu', 'Analytical Chemistry', 'true'],
            ['T/R 9:30 Stefanile', 'Dr. Stefanile', 'dr.stefanile@university.edu', 'Biochemistry', 'true'],
            ['T/R 11:00 AM', 'Dr. Martinez', 'dr.martinez@university.edu', 'Inorganic Chemistry', 'true'],
            ['T/R 3:30 PM', 'Prof. Chen', 'prof.chen@university.edu', 'Advanced Chemistry', 'true'],
            ['T/R 5:30 PM', 'Dr. Brown', 'dr.brown@university.edu', 'Chemistry Lab', 'true']
        ]
        
        with open('instructor_mappings.csv', 'w', newline='') as f:
            import csv
            writer = csv.writer(f)
            writer.writerows(sample_data)
    
    def get_instructor_for_section(self, course_section):
        """Get instructor information for a course section"""
        return self.instructor_mappings.get(course_section.strip(), None)
    
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
                df['datetime'] = df[date_col] + ' ' + df[time_col]
                df['datetime'] = pd.to_datetime(df['datetime'], errors='coerce')
                df = df.dropna(subset=['datetime'])
            elif date_col in df.columns:
                df['datetime'] = pd.to_datetime(df[date_col], errors='coerce')
                df = df.dropna(subset=['datetime'])
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            logging.info(f"📊 Loaded {len(df)} meeting records")
            return df
            
        except Exception as e:
            logging.error(f"❌ Error loading data: {e}")
            return None
    
    def get_weekly_data(self, df):
        """Filter data for the past week"""
        if df is None or df.empty:
            return df
        
        # Get last 7 days
        week_ago = datetime.now() - timedelta(days=7)
        weekly_df = df[df['datetime'] >= week_ago].copy()
        
        logging.info(f"📅 Weekly data: {len(weekly_df)} meetings in the last 7 days")
        return weekly_df
    
    def group_by_course_section(self, df):
        """Group meetings by course section and map to instructors"""
        if df is None or df.empty:
            return {}
        
        course_col = self.column_mapping['course']
        instructor_groups = {}
        
        # Group by course section
        for course_section in df[course_col].unique():
            if pd.isna(course_section) or course_section == '':
                continue
            
            # Get instructor info from config
            instructor_info = self.get_instructor_for_section(course_section)
            if not instructor_info:
                logging.warning(f"⚠️ No instructor configured for section: '{course_section}'")
                continue
            
            instructor_name = instructor_info['instructor']
            course_data = df[df[course_col] == course_section].copy()
            
            if instructor_name not in instructor_groups:
                instructor_groups[instructor_name] = {
                    'email': instructor_info['email'],
                    'course_name': instructor_info.get('course_name', ''),
                    'sections': [],
                    'data': pd.DataFrame(),
                    'total_meetings': 0,
                    'unique_students': 0
                }
            
            # Add this section's data to the instructor
            instructor_groups[instructor_name]['sections'].append(course_section)
            instructor_groups[instructor_name]['data'] = pd.concat([
                instructor_groups[instructor_name]['data'], 
                course_data
            ], ignore_index=True)
            
            instructor_groups[instructor_name]['total_meetings'] = len(instructor_groups[instructor_name]['data'])
            instructor_groups[instructor_name]['unique_students'] = len(instructor_groups[instructor_name]['data'].groupby([
                self.column_mapping['first_name'], 
                self.column_mapping['last_name']
            ]))
        
        logging.info(f"👨‍🏫 Found {len(instructor_groups)} instructors with meetings")
        
        # Log section coverage
        configured_sections = set(self.instructor_mappings.keys())
        actual_sections = set(df[course_col].unique())
        missing_config = actual_sections - configured_sections
        
        if missing_config:
            logging.warning(f"⚠️ Sections in data but not configured: {missing_config}")
            logging.info("💡 Add these sections to instructor_mappings.csv")
        
        return instructor_groups
    
    def generate_instructor_statistics(self, instructor_data):
        """Generate statistics for a specific instructor"""
        df = instructor_data['data']
        
        if df.empty:
            return {}
        
        # Basic stats
        stats = {
            'total_meetings': len(df),
            'unique_students': instructor_data['unique_students'],
            'sections': instructor_data['sections'],
            'section_count': len(instructor_data['sections']),
            'course_name': instructor_data.get('course_name', '')
        }
        
        # Meeting type breakdown
        meeting_type_col = self.column_mapping['meeting_type']
        if meeting_type_col in df.columns:
            meeting_types = df[meeting_type_col].value_counts().to_dict()
            stats['meeting_types'] = meeting_types
        else:
            stats['meeting_types'] = {}
        
        # Daily breakdown
        daily_counts = df.groupby(df['datetime'].dt.date).size()
        stats['daily_counts'] = daily_counts.to_dict()
        stats['busiest_day'] = daily_counts.idxmax() if len(daily_counts) > 0 else 'N/A'
        stats['avg_daily'] = daily_counts.mean() if len(daily_counts) > 0 else 0
        
        # Student meeting frequency
        student_counts = df.groupby([
            self.column_mapping['first_name'], 
            self.column_mapping['last_name']
        ]).size().sort_values(ascending=False)
        stats['top_students'] = student_counts.head(5).to_dict()
        
        # Section-specific stats
        section_stats = {}
        course_col = self.column_mapping['course']
        for section in instructor_data['sections']:
            section_df = df[df[course_col] == section]
            section_stats[section] = {
                'meetings': len(section_df),
                'students': len(section_df.groupby([
                    self.column_mapping['first_name'], 
                    self.column_mapping['last_name']
                ]))
            }
        stats['section_breakdown'] = section_stats
        
        return stats
    
    def create_student_meeting_list(self, df):
        """Create a detailed list of student meetings"""
        if df.empty:
            return []
        
        meetings = []
        for _, row in df.iterrows():
            meeting = {
                'student_name': f"{row[self.column_mapping['first_name']]} {row[self.column_mapping['last_name']]}",
                'course': row[self.column_mapping['course']],
                'date': row['datetime'].strftime('%Y-%m-%d'),
                'time': row['datetime'].strftime('%I:%M %p'),
                'meeting_type': row[self.column_mapping['meeting_type']],
                'topic': row[self.column_mapping['topic']] if row[self.column_mapping['topic']] else 'Not specified'
            }
            meetings.append(meeting)
        
        # Sort by date (most recent first)
        meetings.sort(key=lambda x: x['date'], reverse=True)
        return meetings
    
    def create_wordcloud_from_topics(self, instructor_name, df):
        """Create a wordcloud from meeting topics"""
        try:
            if df.empty:
                return None
            
            topic_col = self.column_mapping['topic']
            
            # Extract all topics and clean them
            topics = df[topic_col].dropna().astype(str)
            topics = topics[topics != '']  # Remove empty strings
            topics = topics[topics.str.lower() != 'not specified']  # Remove "not specified"
            
            if len(topics) == 0:
                logging.warning(f"⚠️ No valid topics found for {instructor_name}")
                return None
            
            # Combine all topics into one text
            all_topics_text = ' '.join(topics)
            
            # Clean the text - remove special characters, extra spaces
            all_topics_text = re.sub(r'[^\w\s]', ' ', all_topics_text)
            all_topics_text = ' '.join(all_topics_text.split())  # Remove extra whitespace
            
            if len(all_topics_text.strip()) == 0:
                logging.warning(f"⚠️ No text content after cleaning for {instructor_name}")
                return None
            
            # Create wordcloud with professional styling
            wordcloud = WordCloud(
                width=800, 
                height=400, 
                background_color='white',
                max_words=50,
                colormap='viridis',
                relative_scaling=0.5,
                min_font_size=12,
                max_font_size=72,
                prefer_horizontal=0.7,
                normalize_plurals=False
            ).generate(all_topics_text)
            
            # Save wordcloud with instructor-specific filename
            filename = f'wordcloud_{instructor_name.replace(" ", "_").replace(".", "")}.png'
            
            # Create the plot
            plt.figure(figsize=(12, 6))
            plt.imshow(wordcloud, interpolation='bilinear')
            plt.axis('off')
            plt.title(f'Meeting Topics - {instructor_name}', fontsize=16, fontweight='bold', pad=20)
            plt.tight_layout(pad=1)
            
            # Save with high quality
            plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white', edgecolor='none')
            plt.close()
            
            logging.info(f"✅ Created wordcloud for {instructor_name}: {filename}")
            return filename
            
        except Exception as e:
            logging.error(f"❌ Error creating wordcloud for {instructor_name}: {e}")
            return None
    
    def create_instructor_visualization(self, instructor_name, instructor_data, stats):
        """Create visualization for specific instructor"""
        df = instructor_data['data']
        
        if df.empty:
            return None
        
        fig, axes = plt.subplots(2, 2, figsize=(12, 8))
        fig.suptitle(f'Weekly Meeting Summary - {instructor_name}', fontsize=14, fontweight='bold')
        
        # Daily meetings this week
        daily_counts = pd.Series(stats['daily_counts'])
        if len(daily_counts) > 0:
            daily_counts.plot(kind='bar', ax=axes[0,0], color='#0078d4')
            axes[0,0].set_title('Daily Meetings This Week')
            axes[0,0].set_xlabel('Date')
            axes[0,0].set_ylabel('Number of Meetings')
            axes[0,0].tick_params(axis='x', rotation=45)
        else:
            axes[0,0].text(0.5, 0.5, 'No meetings this week', ha='center', va='center')
            axes[0,0].set_title('Daily Meetings - No Data')
        
        # Meeting types
        meeting_types = stats.get('meeting_types', {})
        if meeting_types:
            pd.Series(meeting_types).plot(kind='pie', ax=axes[0,1], autopct='%1.1f%%')
            axes[0,1].set_title('Meeting Types Distribution')
            axes[0,1].set_ylabel('')
        else:
            axes[0,1].text(0.5, 0.5, 'No meeting type data', ha='center', va='center')
            axes[0,1].set_title('Meeting Types - No Data')
        
        # Section breakdown
        section_breakdown = stats.get('section_breakdown', {})
        if section_breakdown:
            sections = list(section_breakdown.keys())
            meeting_counts = [section_breakdown[s]['meetings'] for s in sections]
            
            axes[1,0].bar(range(len(sections)), meeting_counts, color='#00bcf2')
            axes[1,0].set_title('Meetings by Section')
            axes[1,0].set_xlabel('Course Section')
            axes[1,0].set_ylabel('Number of Meetings')
            axes[1,0].set_xticks(range(len(sections)))
            axes[1,0].set_xticklabels([s.split()[0] for s in sections], rotation=45)
        else:
            axes[1,0].text(0.5, 0.5, 'No section data', ha='center', va='center')
            axes[1,0].set_title('Course Sections - No Data')
        
        # Top students (meeting frequency)
        top_students = stats.get('top_students', {})
        if top_students:
            # Convert the tuple keys to readable names
            student_names = [f"{name[0]} {name[1]}" for name in top_students.keys()]
            meeting_counts = list(top_students.values())
            
            axes[1,1].bar(range(len(student_names)), meeting_counts, color='#40e0d0')
            axes[1,1].set_title('Most Active Students')
            axes[1,1].set_xlabel('Students')
            axes[1,1].set_ylabel('Number of Meetings')
            axes[1,1].set_xticks(range(len(student_names)))
            axes[1,1].set_xticklabels([name.split()[0] for name in student_names], rotation=45)
        else:
            axes[1,1].text(0.5, 0.5, 'No student data', ha='center', va='center')
            axes[1,1].set_title('Student Activity - No Data')
        
        plt.tight_layout()
        
        # Save with instructor-specific filename
        filename = f'weekly_report_{instructor_name.replace(" ", "_").replace(".", "")}.png'
        plt.savefig(filename, dpi=300, bbox_inches='tight', facecolor='white')
        plt.close()
        
        return filename
    
    def create_instructor_email_body(self, instructor_name, instructor_data, stats, meetings_list):
        """Create personalized email body for instructor with Outlook-compatible styling"""
        week_start = (datetime.now() - timedelta(days=7)).strftime('%B %d')
        week_end = datetime.now().strftime('%B %d, %Y')
        
        # Section details
        sections_info = ""
        section_breakdown = stats.get('section_breakdown', {})
        if section_breakdown:
            sections_info = "<ul>"
            for section, section_stats in section_breakdown.items():
                sections_info += f"<li><strong>{section}:</strong> {section_stats['meetings']} meetings, {section_stats['students']} students</li>"
            sections_info += "</ul>"
        else:
            sections_info = "<p>No section data available.</p>"
        
        # Create student meetings table
        meetings_html = ""
        if meetings_list:
            meetings_html = "<table style='width: 100%; border-collapse: collapse; margin: 10px 0; font-size: 13px;'>"
            meetings_html += "<tr style='background: #f1f3f4;'><th style='padding: 8px; border: 1px solid #ddd;'>Student</th><th style='padding: 8px; border: 1px solid #ddd;'>Section</th><th style='padding: 8px; border: 1px solid #ddd;'>Date</th><th style='padding: 8px; border: 1px solid #ddd;'>Time</th><th style='padding: 8px; border: 1px solid #ddd;'>Type</th><th style='padding: 8px; border: 1px solid #ddd;'>Topic</th></tr>"
            
            for meeting in meetings_list[:20]:  # Limit to 20 most recent
                meetings_html += f"<tr><td style='padding: 6px; border: 1px solid #ddd;'>{meeting['student_name']}</td><td style='padding: 6px; border: 1px solid #ddd;'>{meeting['course']}</td><td style='padding: 6px; border: 1px solid #ddd;'>{meeting['date']}</td><td style='padding: 6px; border: 1px solid #ddd;'>{meeting['time']}</td><td style='padding: 6px; border: 1px solid #ddd;'>{meeting['meeting_type']}</td><td style='padding: 6px; border: 1px solid #ddd;'>{meeting['topic']}</td></tr>"
            
            meetings_html += "</table>"
            
            if len(meetings_list) > 20:
                meetings_html += f"<p style='font-style: italic; color: #666;'>Showing 20 most recent meetings. Total: {len(meetings_list)} meetings this week.</p>"
        else:
            meetings_html = "<p>No meetings recorded this week.</p>"
        
        # Meeting type summary
        meeting_types_html = ""
        meeting_types = stats.get('meeting_types', {})
        if meeting_types:
            meeting_types_html = "<ul>"
            for meeting_type, count in meeting_types.items():
                meeting_types_html += f"<li><strong>{meeting_type}:</strong> {count} meetings</li>"
            meeting_types_html += "</ul>"
        else:
            meeting_types_html = "<p>No meeting type data available.</p>"
        
        # Semester info from config
        semester_info = self.config_metadata.get('semester', 'Current Semester')
        
        # Outlook-compatible HTML with solid colors instead of gradients
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 0; color: #333; }}
                .container {{ max-width: 800px; margin: 0 auto; background: #f5f5f5; }}
                .header {{ background: #2c3e50; color: white; padding: 30px 20px; text-align: center; }}
                .header h1 {{ margin: 0; font-size: 24px; }}
                .content {{ padding: 30px 20px; background: white; }}
                .stats-grid {{ display: table; width: 100%; margin: 20px 0; }}
                .stats-row {{ display: table-row; }}
                .stat-card {{ display: table-cell; background: #f8f9fa; padding: 15px; text-align: center; border: 1px solid #e9ecef; vertical-align: middle; width: 33.33%; }}
                .stat-number {{ font-size: 24px; font-weight: bold; color: #2c3e50; margin: 0; }}
                .stat-label {{ color: #666; font-size: 14px; margin-top: 5px; }}
                .section {{ margin: 30px 0; padding: 20px; border-left: 4px solid #3498db; background: #f8f9ff; }}
                .section h3 {{ margin-top: 0; color: #2c3e50; }}
                .footer {{ text-align: center; padding: 20px; background: #f8f9fa; color: #666; font-size: 12px; }}
                table {{ font-size: 13px; }}
                .attachment-note {{ background: #e8f5e8; padding: 15px; border-radius: 6px; border-left: 4px solid #27ae60; margin: 20px 0; }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>📚 Weekly Meeting Summary</h1>
                    <p><strong>{instructor_name}</strong></p>
                    <p>{semester_info} | Week of {week_start} - {week_end}</p>
                </div>
                
                <div class="content">
                    <div class="stats-grid">
                        <div class="stats-row">
                            <div class="stat-card">
                                <div class="stat-number">{stats.get('total_meetings', 0)}</div>
                                <div class="stat-label">Total Meetings</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-number">{stats.get('unique_students', 0)}</div>
                                <div class="stat-label">Unique Students</div>
                            </div>
                            <div class="stat-card">
                                <div class="stat-number">{stats.get('section_count', 0)}</div>
                                <div class="stat-label">Course Sections</div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="section">
                        <h3>📋 Your Course Sections</h3>
                        {sections_info}
                        <p><strong>Course:</strong> {stats.get('course_name', 'N/A')}</p>
                        <p><strong>Busiest day:</strong> {stats.get('busiest_day', 'N/A')}</p>
                        <p><strong>Average daily meetings:</strong> {stats.get('avg_daily', 0):.1f}</p>
                    </div>
                    
                    <div class="section">
                        <h3>📊 Meeting Types This Week</h3>
                        {meeting_types_html}
                    </div>
                    
                    <div class="section">
                        <h3>👥 Student Meeting Details</h3>
                        {meetings_html}
                    </div>
                    
                    <div class="attachment-note">
                        <strong>📎 Attachments:</strong><br>
                        • Visual analytics dashboard for your sections<br>
                        • Meeting topics wordcloud visualization<br>
                        • Complete meeting data export (CSV format)<br>
                        • Weekly trends and student activity patterns
                    </div>
                </div>
                
                <div class="footer">
                    <p>🤖 Automated weekly summary | Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
                    <p>Configuration: {semester_info} | Questions about this report? Contact your system administrator.</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return html_body
    
    def send_instructor_email(self, instructor_name, instructor_email, stats, meetings_list, chart_filename, wordcloud_filename=None):
        """Send personalized email to instructor with optional wordcloud"""
        try:
            logging.info(f"📧 Preparing email for {instructor_name} ({instructor_email})")
            
            msg = MIMEMultipart()
            msg['From'] = self.email_user
            msg['To'] = instructor_email
            msg['Subject'] = f'📚 Weekly Meeting Summary - {instructor_name} - Week of {datetime.now().strftime("%B %d, %Y")}'
            
            # Create email body
            html_body = self.create_instructor_email_body(instructor_name, {}, stats, meetings_list)
            msg.attach(MIMEText(html_body, 'html'))
            
            # Attach chart if it exists
            if chart_filename and os.path.exists(chart_filename):
                with open(chart_filename, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={chart_filename}')
                    msg.attach(part)
            
            # Attach wordcloud if it exists
            if wordcloud_filename and os.path.exists(wordcloud_filename):
                with open(wordcloud_filename, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={wordcloud_filename}')
                    msg.attach(part)
            
            # Create and attach CSV with instructor's data
            csv_filename = f"meetings_{instructor_name.replace(' ', '_').replace('.', '')}_{datetime.now().strftime('%Y%m%d')}.csv"
            if meetings_list:
                df_meetings = pd.DataFrame(meetings_list)
                df_meetings.to_csv(csv_filename, index=False)
                
                with open(csv_filename, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename={csv_filename}')
                    msg.attach(part)
            
            # Send email
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.email_user, self.email_password)
            server.send_message(msg)
            server.quit()
            
            logging.info(f"✅ Email sent successfully to {instructor_name}")
            
            # Clean up files
            for filename in [chart_filename, wordcloud_filename, csv_filename]:
                if filename and os.path.exists(filename):
                    os.remove(filename)
            
            return True
            
        except Exception as e:
            logging.error(f"❌ Failed to send email to {instructor_name}: {e}")
            return False
    
    def run_weekly_report(self):
        """Main function to run weekly instructor reports"""
        try:
            logging.info("🚀 Starting configuration-based weekly instructor reports...")
            
            # Check if we have instructor mappings
            if not self.instructor_mappings:
                logging.error("❌ No instructor mappings configured!")
                logging.info("💡 Please edit instructor_mappings.csv with your instructor information")
                return
            
            # Load data
            df = self.load_data()
            if df is None:
                logging.error("❌ Failed to load data")
                return
            
            # Get weekly data
            weekly_df = self.get_weekly_data(df)
            if weekly_df.empty:
                logging.warning("⚠️ No meetings in the past week")
                return
            
            # Group by course section and map to instructors
            instructor_groups = self.group_by_course_section(weekly_df)
            if not instructor_groups:
                logging.warning("⚠️ No instructors found with configured sections")
                return
            
            success_count = 0
            week_start = (datetime.now() - timedelta(days=7)).strftime('%B %d')
            week_end = datetime.now().strftime('%B %d, %Y')
            
            # Process each instructor
            for instructor_name, instructor_data in instructor_groups.items():
                logging.info(f"👨‍🏫 Processing {instructor_name}...")
                
                # Generate statistics
                stats = self.generate_instructor_statistics(instructor_data)
                
                # Create student meeting list
                meetings_list = self.create_student_meeting_list(instructor_data['data'])
                
                # Create visualization
                chart_filename = self.create_instructor_visualization(instructor_name, instructor_data, stats)
                
                # Create wordcloud from topics
                wordcloud_filename = self.create_wordcloud_from_topics(instructor_name, instructor_data['data'])
                
                # Send email with both attachments
                if self.send_instructor_email(instructor_name, instructor_data['email'], stats, meetings_list, chart_filename, wordcloud_filename):
                    success_count += 1
            
            # Send administrator summary email
            logging.info("📊 Preparing administrator summary...")
            admin_stats = self.create_admin_summary_statistics(instructor_groups, weekly_df)
            admin_sent = self.send_admin_summary_email(admin_stats, week_start, week_end)
            
            if admin_sent:
                logging.info(f"✅ Weekly reports completed! Sent {success_count}/{len(instructor_groups)} instructor emails + admin summary")
            else:
                logging.info(f"✅ Weekly reports completed! Sent {success_count}/{len(instructor_groups)} instructor emails (admin summary failed)")
            
            
        except Exception as e:
            logging.error(f"💥 Error in weekly report generation: {e}")

def main():
    reporter = ConfigBasedInstructorReporter()
    
    if os.getenv('GITHUB_ACTIONS'):
        logging.info("🤖 Running weekly reports in GitHub Actions")
        reporter.run_weekly_report()
    else:
        logging.info("🖥️ Running weekly reports locally")
        reporter.run_weekly_report()

if __name__ == "__main__":
    main()