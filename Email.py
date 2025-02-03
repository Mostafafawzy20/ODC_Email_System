import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
import time
from datetime import datetime
from typing import Tuple, Dict, Any, Optional, List
from dataclasses import dataclass
try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

# Load the image as bytes
with open('static/image001.png', 'rb') as f:
    IMAGE_DATA = f.read()

def create_email_body(prop_code: str, group_data: pd.DataFrame, additional_info: str = "") -> str:
    try:
        hotel_name = group_data['hotel'].iloc[0] if not group_data.empty else prop_code
        sorted_data = group_data.sort_values('file_name')
        
        email_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; font-size: 14px; }}
                table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .important {{ font-weight: bold; }}
                .instruction-image {{ max-width: 800px; width: 100%; margin: 20px 0; }}
            </style>
        </head>
        <body>
            <h2>Missing ODC Files Report for {prop_code} - {hotel_name}</h2>
            <p>Dear Team,</p>
            <br>
            <p>As per the Missing ODC report for your hotel, the following files have not been received by One Yield. 
            Missing files might create discrepancies in OY group data impacting your rate recommendation, hurdle output and inventory optimization.</p>
            <br>
            <table>
                <tr>
                    <th>Prop Code</th>
                    <th>File Name</th>
                    <th>Business Date</th>
                    <th>Days Missing</th>
                </tr>
        """
        
        for _, row in sorted_data.iterrows():
            business_date = pd.to_datetime(row['business_date']).strftime('%Y-%m-%d') if pd.notnull(row['business_date']) else ''
            email_body += f"""
                <tr>
                    <td>{row['prop_code']}</td>
                    <td>{row['file_name']}</td>
                    <td>{business_date}</td>
                    <td><strong>{int(row['days_missing'])}</strong></td>
                </tr>
            """
        
        email_body += """
            </table>
            <br>
            <p>Please resend the files by close of business today. You can find all required documentation at the following link:</p>
            <p><a href="https://servicenow.marriott.com/kb_view.do?sysparm_article=KB0484970">Property Software and Applications – Opera v5.6 Data Capture guide</a> (Service Now KB article: KB0484970)</p>
            <p>Select: ODC Export – Opera v5.6 Data Capture Guide – pages 43 to 48</p>
            <br>
            <p class='important'><strong>Important Reminder:</strong></p>
            <p>This is a night audit process and it is their responsibility to ensure files are generated and sent. 
            Please refer to the above article and the screen shot below to see how to check this.</p>
            <br>
            <img src="cid:instructions_image" alt="ODC Export Instructions" class="instruction-image">
            <br>
            <p class='important'>Please note, files missing for more than 28 days can not be re-sent from Opera.</p>
            <br>
            <p>Thank you,<br>EMEA Revenue Management Operations</p>
        </body>
        </html>
        """
        
        return email_body
    except Exception as e:
        raise Exception(f"Error creating email body: {str(e)}")

def create_email_body_over_28(prop_code: str, group_data: pd.DataFrame, additional_info: str = "") -> str:
    """Create email body for hotels with files missing for more than 28 days"""
    try:
        hotel_name = group_data['hotel'].iloc[0] if not group_data.empty else prop_code
        sorted_data = group_data.sort_values('file_name')
        
        email_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; font-size: 14px; }}
                table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; }}
                th {{ background-color: #f2f2f2; }}
                .important {{ font-weight: bold; color: #ff0000; }}
                .instruction-image {{ max-width: 800px; width: 100%; margin: 20px 0; }}
            </style>
        </head>
        <body>
            <h2>Critical: Long-term Missing ODC Files Report for {prop_code} - {hotel_name}</h2>
            <p>Dear Team,</p>
            <br>
            <p class='important'>URGENT ACTION REQUIRED: Your hotel has files missing for more than 28 days.</p>
            <p>This situation requires immediate attention as these files can no longer be recovered through the standard Opera export process.</p>
            <br>
            <table>
                <tr>
                    <th>Prop Code</th>
                    <th>File Name</th>
                    <th>Business Date</th>
                    <th>Days Missing</th>
                </tr>
        """
        
        for _, row in sorted_data.iterrows():
            business_date = pd.to_datetime(row['business_date']).strftime('%Y-%m-%d') if pd.notnull(row['business_date']) else ''
            email_body += f"""
                <tr>
                    <td>{row['prop_code']}</td>
                    <td>{row['file_name']}</td>
                    <td>{business_date}</td>
                    <td><strong>{int(row['days_missing'])}</strong></td>
                </tr>
            """
        
        email_body += """
            </table>
            <br>
            <p class='important'>Recovery Plan Required:</p>
            <ol>
                <li>Please escalate this issue to your property leadership team immediately.</li>
                <li>Contact the Revenue Management Operations team to discuss data recovery options.</li>
                <li>A manual data recovery plan will need to be implemented.</li>
                <li>Future preventive measures must be put in place to avoid such situations.</li>
            </ol>
            <br>
            <p class='important'>Please note: Standard Opera export process cannot recover files older than 28 days.</p>
            <br>
            <p>For immediate assistance, please contact:</p>
            <ul>
                <li>Revenue Management Operations Team</li>
                <li>Your Regional Revenue Management Leader</li>
            </ul>
            <br>
            <p>Thank you for your immediate attention to this critical matter.</p>
            <br>
            <p>Best regards,<br>EMEA Revenue Management Operations</p>
        </body>
        </html>
        """
        
        return email_body
    except Exception as e:
        raise Exception(f"Error creating email body: {str(e)}")

def select_email_template(prop_code: str, group_data: pd.DataFrame, additional_info: str = "", force_template: str = None) -> str:
    """
    Select appropriate email template based on days missing or forced template type
    
    Args:
        prop_code: Hotel property code
        group_data: DataFrame containing hotel data
        additional_info: Additional information to include in email
        force_template: Force specific template ("standard" or "critical")
    """
    if force_template == "critical":
        return create_email_body_over_28(prop_code, group_data, additional_info)
    elif force_template == "standard":
        return create_email_body(prop_code, group_data, additional_info)
    else:
        # Automatic selection based on days missing
        max_days_missing = group_data['days_missing'].max()
        if max_days_missing > 28:
            return create_email_body_over_28(prop_code, group_data, additional_info)
        else:
            return create_email_body(prop_code, group_data, additional_info)

@dataclass
class EmailStatus:
    """Class to track email sending status"""
    hotel_code: str
    email_address: str
    timestamp: datetime
    status: str
    message: str

@dataclass
class EmailProcessingResult:
    """Class to hold email processing results"""
    success: bool
    message: str
    data: Optional[Dict[str, Any]] = None
    status: Optional[EmailStatus] = None

class EmailReport:
    """Class to generate email sending reports"""
    def __init__(self):
        self.statuses: List[EmailStatus] = []

    def add_status(self, status: EmailStatus):
        self.statuses.append(status)

    def generate_pdf(self, filename: str):
        """Generate PDF report of email sending status"""
        if not FPDF_AVAILABLE:
            raise ImportError("FPDF module is not installed. Please install it using: pip install fpdf")
            
        pdf = FPDF()
        pdf.add_page()
        
        # Set up header
        pdf.set_font('Arial', 'B', 16)
        pdf.cell(0, 10, 'Email Notification Report', 0, 1, 'C')
        pdf.ln(10)
        
        # Add timestamp
        pdf.set_font('Arial', '', 10)
        pdf.cell(0, 10, f'Report Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1)
        pdf.ln(5)
        
        # Table header
        pdf.set_font('Arial', 'B', 10)
        headers = ['Hotel Code', 'Email Address', 'Timestamp', 'Status', 'Message']
        col_widths = [30, 60, 40, 20, 40]
        
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], 10, header, 1)
        pdf.ln()
        
        # Table content
        pdf.set_font('Arial', '', 8)
        for status in self.statuses:
            pdf.cell(col_widths[0], 10, status.hotel_code, 1)
            pdf.cell(col_widths[1], 10, status.email_address[:35], 1)
            pdf.cell(col_widths[2], 10, status.timestamp.strftime("%Y-%m-%d %H:%M:%S"), 1)
            pdf.cell(col_widths[3], 10, status.status, 1)
            pdf.cell(col_widths[4], 10, status.message[:30], 1)
            pdf.ln()
        
        # Save PDF
        pdf.output(filename)

    def get_summary(self) -> Dict[str, int]:
        """Get summary of email sending results"""
        return {
            'total': len(self.statuses),
            'success': sum(1 for s in self.statuses if s.status == 'Success'),
            'failed': sum(1 for s in self.statuses if s.status == 'Failed')
        }

class EmailSender:
    def __init__(self, smtp_server: str, smtp_port: int):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self._smtp = None

    def connect(self, username: str, password: str):
        """Connect to SMTP server"""
        try:
            self._smtp = smtplib.SMTP(self.smtp_server, self.smtp_port)
            self._smtp.starttls()
            self._smtp.login(username, password)
            return True, None
        except Exception as e:
            return False, str(e)

    def send_email(self, 
                  to_address: str, 
                  subject: str, 
                  html_body: str, 
                  from_address: str,
                  bcc_address: str = None) -> Tuple[bool, str]:
        """Send email using SMTP"""
        try:
            msg = MIMEMultipart('related')
            msg['Subject'] = subject
            msg['From'] = from_address
            msg['To'] = to_address
            if bcc_address:
                msg['Bcc'] = bcc_address

            # Create HTML part
            html_part = MIMEText(html_body, 'html')
            msg.attach(html_part)

            # Attach image
            img = MIMEImage(IMAGE_DATA)
            img.add_header('Content-ID', '<instructions_image>')
            msg.attach(img)

            # Send email
            self._smtp.send_message(msg)
            return True, "Email sent successfully"
        except Exception as e:
            return False, f"Failed to send email: {str(e)}"

    def close(self):
        """Close SMTP connection"""
        if self._smtp:
            try:
                self._smtp.quit()
            except:
                pass

def load_and_prepare_data(missing_file: Any, emails_file: Any) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Load and prepare the missing files and emails data"""
    try:
        # Process missing files data
        missing = pd.read_excel(missing_file, sheet_name='Missing', header=None)
        missing = missing.iloc[3:].reset_index(drop=True)
        column_names = ['prop_code', 'hotel', 'file_name', 'business_date', 'dow', 'days_missing']
        missing.columns = column_names[:len(missing.columns)]
        missing['business_date'] = pd.to_datetime(missing['business_date'], errors='coerce')
        missing = missing.dropna(subset=['prop_code'])

        # Process emails data
        emails = pd.read_excel(emails_file, header=0)
        emails.rename(columns={emails.columns[0]: 'Hotels', emails.columns[1]: 'Email'}, inplace=True)
        
        # Ensure prop_code is string type for consistent matching
        missing['prop_code'] = missing['prop_code'].astype(str)
        emails['Hotels'] = emails['Hotels'].astype(str)

        return missing, emails
    except Exception as e:
        raise Exception(f"Error in data preparation: {str(e)}")

def validate_hotel_data(hotel_code: str, missing_df: pd.DataFrame, emails_df: pd.DataFrame) -> Tuple[bool, str, Optional[pd.DataFrame], Optional[str]]:
    """Validate hotel data before sending email"""
    try:
        hotel_data = missing_df[missing_df['prop_code'] == str(hotel_code)]
        if hotel_data.empty:
            return False, f"No missing files data found for hotel {hotel_code}", None, None
        
        hotel_email_data = emails_df[emails_df['Hotels'] == str(hotel_code)]
        if hotel_email_data.empty:
            return False, f"No email address found for hotel {hotel_code}", None, None
        
        email_address = hotel_email_data['Email'].iloc[0]
        if not isinstance(email_address, str) or not email_address.strip():
            return False, f"Invalid email address for hotel {hotel_code}", None, None
        
        return True, "", hotel_data, email_address
    except Exception as e:
        return False, f"Validation error for hotel {hotel_code}: {str(e)}", None, None

def process_single_notification(
    missing_df: pd.DataFrame, 
    emails_df: pd.DataFrame, 
    hotel_code: str, 
    smtp_settings: Dict[str, Any],
    additional_info: str = "",
    delay: float = 2.0,
    from_email: str = "",
    bcc_email: str = None,
    force_template: str = None
) -> EmailProcessingResult:
    try:
        # Validate hotel data
        is_valid, error_message, hotel_data, email_address = validate_hotel_data(hotel_code, missing_df, emails_df)
        if not is_valid:
            return EmailProcessingResult(False, error_message)
        
        hotel_name = hotel_data['hotel'].iloc[0] if not hotel_data.empty else hotel_code
        
        # Create email content
        email_body = select_email_template(hotel_code, hotel_data, additional_info, force_template)
        
        # Set subject based on template type
        if force_template == "critical" or hotel_data['days_missing'].max() > 28:
            subject = f"CRITICAL: Long-term Missing ODC Files - {hotel_code} - {hotel_name}"
        else:
            subject = f"Action required: Missing ODC Files Report - {hotel_code} - {hotel_name}"

        # Initialize email sender
        sender = EmailSender(smtp_settings['server'], smtp_settings['port'])
        success, error = sender.connect(smtp_settings['username'], smtp_settings['password'])
        
        if not success:
            return EmailProcessingResult(False, f"Failed to connect to SMTP server: {error}")

        try:
            # Send email
            success, message = sender.send_email(
                to_address=email_address,
                subject=subject,
                html_body=email_body,
                from_address=from_email,
                bcc_address=bcc_email
            )

            if success:
                return EmailProcessingResult(
                    True,
                    f"Notification sent successfully to {hotel_code} ({email_address})",
                    {'hotel_code': hotel_code, 'email': email_address}
                )
            else:
                return EmailProcessingResult(False, message)
        finally:
            sender.close()

    except Exception as e:
        return EmailProcessingResult(False, f"Error processing notification for {hotel_code}: {str(e)}")

def process_bulk_notifications(
    missing_df: pd.DataFrame,
    emails_df: pd.DataFrame,
    hotel_codes: List[str],
    smtp_settings: Dict[str, Any],
    additional_info: str = "",
    batch_size: int = 10,
    delay_between_emails: float = 2.0,
    progress_callback: Optional[callable] = None,
    from_email: str = "",
    bcc_email: str = None,
    force_template: str = None
) -> Tuple[bool, EmailReport]:
    """Process and send notifications for multiple hotels"""
    report = EmailReport()
    sender = None
    
    try:
        # Initialize email sender
        sender = EmailSender(smtp_settings['server'], smtp_settings['port'])
        success, error = sender.connect(smtp_settings['username'], smtp_settings['password'])
        
        if not success:
            return False, f"Failed to connect to SMTP server: {error}"

        total_hotels = len(hotel_codes)
        
        for i, hotel_code in enumerate(hotel_codes):
            result = process_single_notification(
                missing_df=missing_df,
                emails_df=emails_df,
                hotel_code=hotel_code,
                smtp_settings=smtp_settings,
                additional_info=additional_info,
                delay=delay_between_emails,
                from_email=from_email,
                bcc_email=bcc_email,
                force_template=force_template
            )
            
            # Add status to report
            status = EmailStatus(
                hotel_code=hotel_code,
                email_address=result.data['email'] if result.data else "N/A",
                timestamp=datetime.now(),
                status="Success" if result.success else "Failed",
                message=result.message
            )
            report.add_status(status)
            
            if progress_callback:
                progress = (i + 1) / total_hotels
                progress_callback(progress)

        return True, report

    except Exception as e:
        return False, f"Error in bulk processing: {str(e)}"
    finally:
        if sender:
            sender.close()

def send_to_all_hotels(
    missing_df: pd.DataFrame,
    emails_df: pd.DataFrame,
    additional_info: str = "",
    batch_size: int = 10,
    delay_between_emails: float = 2.0,
    progress_callback: Optional[callable] = None
) -> Tuple[bool, EmailReport]:
    """Send notifications to all hotels in the missing files list"""
    all_hotels = missing_df['prop_code'].unique().tolist()
    return process_bulk_notifications(
        missing_df=missing_df,
        emails_df=emails_df,
        hotel_codes=all_hotels,
        additional_info=additional_info,
        batch_size=batch_size,
        delay_between_emails=delay_between_emails,
        progress_callback=progress_callback
    )

def get_email_statistics(missing_df: pd.DataFrame) -> Dict[str, Any]:
    """Calculate statistics about missing files"""
    try:
        return {
            'total_hotels': missing_df['prop_code'].nunique(),
            'total_missing_files': len(missing_df),
            'average_days_missing': missing_df['days_missing'].mean(),
            'files_by_type': missing_df['file_name'].value_counts().to_dict()
        }
    except Exception as e:
        raise Exception(f"Error calculating email statistics: {str(e)}")
