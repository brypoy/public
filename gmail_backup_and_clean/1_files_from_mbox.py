import mailbox
import email
import os
import re
import shutil
import json
from datetime import datetime
from email.utils import parsedate_to_datetime
from pathlib import Path

# PDF libraries
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch

def sanitize_filename(filename):
    """Remove invalid characters from filename and ensure it's valid for Windows"""
    if not filename:
        return "unnamed"
    
    # Remove invalid Windows characters
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    
    # Remove control characters
    filename = re.sub(r'[\x00-\x1f\x7f]', '', filename)
    
    # Replace multiple underscores with single
    filename = re.sub(r'_+', '_', filename)
    
    # Remove leading/trailing spaces and dots (Windows issue)
    filename = filename.strip(' .')
    
    # Don't allow empty filenames
    if not filename:
        filename = "unnamed"
    
    # Check for Windows reserved names
    reserved = ('CON', 'PRN', 'AUX', 'NUL', 'COM1', 'COM2', 'COM3', 'COM4',
                'COM5', 'COM6', 'COM7', 'COM8', 'COM9', 'LPT1', 'LPT2',
                'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9')
    name_without_ext = os.path.splitext(filename)[0].upper()
    if name_without_ext in reserved:
        filename = f"_{filename}"
    
    # Limit length but keep extension
    if len(filename) > 100:
        name, ext = os.path.splitext(filename)
        filename = name[:95] + '...' + ext
    
    return filename

def extract_tags(message):
    """Extract Gmail tags/labels from email"""
    tags = []
    x_labels = message.get('X-Gmail-Labels', '')
    if x_labels:
        raw_tags = [tag.strip() for tag in x_labels.split(',')]
        exclude_labels = ['Inbox', 'Sent', 'Draft', 'Spam', 'Trash', 'Important', 'Starred', 'Chat']
        tags = [tag for tag in raw_tags if tag not in exclude_labels and not tag.startswith('Category_')]
    if not tags:
        tags = ['Unfiled']
    return tags

def parse_email_date(date_str):
    """Safely parse email date string to datetime object"""
    if not date_str:
        return datetime.now()
    
    try:
        dt = parsedate_to_datetime(date_str)
        if dt:
            return dt
    except:
        pass
    
    try:
        for fmt in ['%a, %d %b %Y %H:%M:%S %z', '%d %b %Y %H:%M:%S %z', 
                    '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S']:
            try:
                return datetime.strptime(date_str, fmt)
            except:
                continue
    except:
        pass
    
    return datetime.now()

def extract_email_body(message):
    """Extract plain text body from email"""
    body = ""
    if message.is_multipart():
        for part in message.walk():
            if part.get_content_type() == "text/plain":
                try:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body += payload.decode('utf-8', errors='ignore')
                except:
                    pass
    else:
        try:
            payload = message.get_payload(decode=True)
            if payload:
                body += payload.decode('utf-8', errors='ignore')
        except:
            pass
    return body

def create_email_pdf(email_body, metadata, output_path):
    """Convert email body to PDF with 8pt Times New Roman"""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        leftMargin=1*inch,
        rightMargin=1*inch,
        topMargin=1*inch,
        bottomMargin=1*inch
    )
    
    style = ParagraphStyle(
        'Normal',
        fontName='Times-Roman',
        fontSize=8,
        leading=10
    )
    
    story = []
    
    # Add metadata header
    story.append(Paragraph(f"From: {metadata['from']}", style))
    story.append(Paragraph(f"To: {metadata['to']}", style))
    story.append(Paragraph(f"Date: {metadata['date']}", style))
    story.append(Paragraph(f"Subject: {metadata['subject']}", style))
    story.append(Paragraph(f"Tags: {', '.join(metadata['tags'])}", style))
    story.append(Spacer(1, 0.2*inch))
    
    # Add email body
    for line in email_body.split('\n'):
        if line.strip():
            line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(line, style))
    
    doc.build(story)

def create_safe_folder_path(base_dir, year_month, folder_name):
    """Create a safe folder path, handling long names and invalid chars"""
    # Sanitize each component
    year_month = sanitize_filename(year_month)
    folder_name = sanitize_filename(folder_name)
    
    # Create full path
    folder_path = os.path.join(base_dir, year_month, folder_name)
    
    # For Windows, ensure path isn't too long
    if os.name == 'nt':  # Windows
        # Use short path if needed
        if len(folder_path) > 240:
            # Try to use short folder name
            short_folder = folder_name[:50]
            folder_path = os.path.join(base_dir, year_month, short_folder)
            
            # If still too long, use absolute path with short names
            if len(folder_path) > 240:
                # Use drive letter and truncate more
                drive = os.path.splitdrive(base_dir)[0]
                rest = base_dir[len(drive):]
                if len(rest) > 100:
                    rest = rest[:100]
                base_short = drive + rest
                folder_path = os.path.join(base_short, year_month[:20], short_folder[:30])
    
    return folder_path

def process_mbox(mbox_path, output_dir):
    """Process mbox file and organize emails by year/month"""
    
    print(f"📂 Opening mbox: {mbox_path}")
    mbox = mailbox.mbox(mbox_path)
    total = len(mbox)
    processed = 0
    error_count = 0
    
    print(f"📊 Processing {total} messages...")
    print("=" * 60)
    
    for key, message in mbox.items():
        try:
            # Extract metadata
            from_ = str(message.get('From', 'Unknown'))
            to_ = str(message.get('To', 'Unknown'))
            subject = str(message.get('Subject', 'No Subject'))
            date_str = message.get('Date')
            
            # Parse date safely
            date = parse_email_date(date_str)
            
            tags = extract_tags(message)
            
            # Get email body (for JSON)
            email_body = extract_email_body(message)
            
            # Create folder structure
            year_month = date.strftime('%Y-%b')
            
            # Create clean folder name from date and subject
            date_prefix = date.strftime('%Y%m%d')
            # Clean subject more aggressively for folder name
            clean_subject = sanitize_filename(subject[:40])  # Even shorter for Windows
            if not clean_subject:
                clean_subject = "no_subject"
            
            folder_name = f"{date_prefix}_{clean_subject}"
            
            # Get safe folder path
            email_dir = create_safe_folder_path(output_dir, year_month, folder_name)
            
            # Create directory (with parents)
            os.makedirs(email_dir, exist_ok=True)
            
            # Save email as PDF
            pdf_filename = f"{folder_name}.pdf"
            pdf_path = os.path.join(email_dir, pdf_filename)
            
            metadata = {
                'from': from_,
                'to': to_,
                'date': date.strftime('%Y-%m-%d %H:%M:%S'),
                'subject': subject,
                'tags': tags
            }
            
            create_email_pdf(email_body, metadata, pdf_path)
            
            # Save attachments in native format
            attachment_list = []
            if message.is_multipart():
                for part in message.walk():
                    filename = part.get_filename()
                    if filename:
                        data = part.get_payload(decode=True)
                        if data:
                            safe_name = sanitize_filename(filename)
                            att_path = os.path.join(email_dir, safe_name)
                            with open(att_path, 'wb') as f:
                                f.write(data)
                            attachment_list.append(safe_name)
            
            # Save metadata as JSON with BODY field
            json_path = os.path.join(email_dir, f"{folder_name}.json")
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump({
                    'from': from_,
                    'to': to_,
                    'date': date.strftime('%Y-%m-%d %H:%M:%S'),
                    'subject': subject,
                    'tags': tags,
                    'body': email_body,  # Added body field
                    'attachments': attachment_list,
                    'attachment_count': len(attachment_list),
                    'email_pdf': pdf_filename,
                    'original_date_string': date_str,
                    'folder': folder_name
                }, f, indent=2, ensure_ascii=False)
            
            processed += 1
            if processed % 100 == 0:
                print(f"  Processed {processed}/{total}")
            
        except Exception as e:
            error_count += 1
            if error_count <= 10:
                print(f"  ❌ Error on message {key}: {e}")
                if 'subject' in locals():
                    print(f"     Subject: {subject[:100]}")
                if 'folder_name' in locals():
                    print(f"     Folder: {folder_name}")
            elif error_count == 11:
                print(f"  ... (further errors suppressed)")
    
    print("\n" + "=" * 60)
    print(f"✅ Complete!")
    print(f"   Successfully processed: {processed}")
    print(f"   Errors: {error_count}")
    print(f"📁 Output: {output_dir}/")
    print(f"   Format: YYYY-Mon/YYYYMMDD_Subject/")

if __name__ == "__main__":
    mbox_file = "C:/Users/ajipoynter/Desktop/BP/bryan backup/All mail Including Spam and Trash.mbox"  # Update this path
    output_directory = "gmail_archive"
    
    # Convert to absolute path
    output_directory = os.path.abspath(output_directory)
    
    if not os.path.exists(mbox_file):
        print(f"❌ Mbox file not found: {mbox_file}")
    else:
        process_mbox(mbox_file, output_directory)


