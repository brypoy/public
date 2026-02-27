import os
import json
import shutil
import io
import tempfile
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

# Custom exception for conversion failures
class ConversionError(Exception):
    """Raised when attachment conversion fails"""
    pass

# Optional imports with fallbacks
try:
    import magic
    MAGIC_AVAILABLE = True
except ImportError:
    magic = None
    MAGIC_AVAILABLE = False

try:
    import PyPDF2
    from PyPDF2 import PdfReader, PdfWriter
    PDF_AVAILABLE = True
except ImportError:
    PdfReader = PdfWriter = None
    PDF_AVAILABLE = False
    print("⚠️ PyPDF2 not installed - run: pip install PyPDF2")

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    print("⚠️ reportlab not installed - run: pip install reportlab")

try:
    from PIL import Image
    PIL_AVAILABLE = True
except ImportError:
    Image = None
    PIL_AVAILABLE = False
    print("⚠️ Pillow not installed - run: pip install Pillow")

try:
    import img2pdf
    IMG2PDF_AVAILABLE = True
except ImportError:
    img2pdf = None
    IMG2PDF_AVAILABLE = False

# Office document libraries (optional)
try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False

try:
    import xlrd
    XLRD_AVAILABLE = True
except ImportError:
    xlrd = None
    XLRD_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    openpyxl = None
    OPENPYXL_AVAILABLE = False

try:
    import textract
    TEXTRACT_AVAILABLE = True
except ImportError:
    textract = None
    TEXTRACT_AVAILABLE = False

# ========== CONFIGURATION ==========
MAX_EMBED_SIZE_MB = 50  # Files larger than this will be embedded only
LIBREOFFICE_TIMEOUT = 200  # seconds
LARGE_PDF_THRESHOLD = 100  # pages

# ========== FILE TYPE DETECTION ==========

OFFICE_TYPES = {
    # Word
    '.doc': 'application/msword',
    '.dot': 'application/msword',
    '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    '.docm': 'application/vnd.ms-word.document.macroEnabled.12',
    '.dotx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.template',
    '.dotm': 'application/vnd.ms-word.template.macroEnabled.12',
    '.rtf': 'application/rtf',
    '.xps': 'application/vnd.ms-xpsdocument', 
    '.writer8': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',  # Add this
    '.odt': 'application/vnd.oasis.opendocument.text',  # Add this line
    '.ods': 'application/vnd.oasis.opendocument.spreadsheet',  # Optional: for spreadsheets
    '.odp': 'application/vnd.oasis.opendocument.presentation',  # Optional: for presentations

    # visio
    '.vsd': 'application/vnd.visio',

    # Excel
    '.xls': 'application/vnd.ms-excel',
    '.xlt': 'application/vnd.ms-excel',
    '.xla': 'application/vnd.ms-excel',
    '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    '.xlsm': 'application/vnd.ms-excel.sheet.macroEnabled.12',
    '.xltx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.template',
    '.xltm': 'application/vnd.ms-excel.template.macroEnabled.12',
    '.xlsb': 'application/vnd.ms-excel.sheet.binary.macroEnabled.12',
    '.csv': 'text/csv',
    
    # PowerPoint
    '.ppt': 'application/vnd.ms-powerpoint',
    '.pot': 'application/vnd.ms-powerpoint',
    '.pps': 'application/vnd.ms-powerpoint',
    '.ppa': 'application/vnd.ms-powerpoint',
    '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    '.pptm': 'application/vnd.ms-powerpoint.presentation.macroEnabled.12',
    '.potx': 'application/vnd.openxmlformats-officedocument.presentationml.template',
    '.potm': 'application/vnd.ms-powerpoint.template.macroEnabled.12',
    '.ppam': 'application/vnd.ms-powerpoint.addin.macroEnabled.12',
    '.ppsx': 'application/vnd.openxmlformats-officedocument.presentationml.slideshow',
    '.ppsm': 'application/vnd.ms-powerpoint.slideshow.macroEnabled.12',
}

def get_file_type(file_path):
    """Detect actual file type using magic numbers or extension fallback"""
    if MAGIC_AVAILABLE and magic:
        try:
            mime = magic.from_file(file_path, mime=True)
            return mime
        except:
            pass
    ext = os.path.splitext(file_path)[1].lower()
    return OFFICE_TYPES.get(ext, 'application/octet-stream')

# ========== LIBREOFFICE PATH DETECTION ==========

def find_libreoffice():
    """Find LibreOffice executable (soffice.exe) in common locations."""
    # Common installation paths
    candidates = [
        'soffice',
        'soffice.exe',
        r'C:\Program Files\LibreOffice\program\soffice.exe',
        r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
        r'C:\Program Files\LibreOffice\program\swriter.exe',
        r'C:\Program Files (x86)\LibreOffice\program\swriter.exe',
    ]
    for path in candidates:
        if shutil.which(path):
            return shutil.which(path)
        if os.path.exists(path):
            return path
    # Try searching in common directories
    for base in [r'C:\Program Files\LibreOffice\program', r'C:\Program Files (x86)\LibreOffice\program']:
        if os.path.exists(base):
            for exe in ['soffice.exe', 'swriter.exe']:
                full = os.path.join(base, exe)
                if os.path.exists(full):
                    return full
    return None

def get_soffice_path():
    """Return path to soffice.exe (required for conversion)."""
    lo = find_libreoffice()
    if not lo:
        return None
    if 'soffice' in os.path.basename(lo).lower():
        return lo
    # If swriter.exe, look for soffice in same directory
    d = os.path.dirname(lo)
    soffice = os.path.join(d, 'soffice.exe')
    if os.path.exists(soffice):
        return soffice
    return None

def create_isolated_env():
    """Environment to avoid MS Office conflicts."""
    env = os.environ.copy()
    env.update({
        'OOO_DISABLE_RECOVERY': '1',
        'SAL_USE_VCLPLUGIN': 'gen',
        'HOME': tempfile.gettempdir(),
        'USERPROFILE': tempfile.gettempdir(),
        'APPDATA': tempfile.gettempdir(),
        'TEMP': tempfile.gettempdir(),
        'TMP': tempfile.gettempdir(),
        'NO_MSI': '1',
        'OOO_FORCE_DESKTOP': '1'
    })
    return env

# ========== CONVERSION FUNCTIONS ==========

def convert_pdf_to_pdf_pages(pdf_path, temp_dir):
    """
    Convert a PDF to a list of single-page PDFs.
    For large PDFs (>100 pages) or PDFs that can't be parsed, return the original path.
    """
    # First attempt: strict=True (default) - proper PDF parsing
    try:
        reader = PdfReader(pdf_path, strict=True)
        total = len(reader.pages)
        
        if total > LARGE_PDF_THRESHOLD:
            print(f"          📚 Large PDF: {total} pages - adding directly")
            return [pdf_path]
        
        page_paths = []
        for i, page in enumerate(reader.pages):
            writer = PdfWriter()
            writer.add_page(page)
            out = os.path.join(temp_dir, f"pdf_page_{i+1:04d}.pdf")
            with open(out, 'wb') as f:
                writer.write(f)
            page_paths.append(out)
            if total > 50 and (i+1) % 20 == 0:
                print(f"          📄 Processed {i+1}/{total} pages")
        return page_paths
        
    except Exception as e1:
        print(f"          ⚠️ Strict parsing failed: {e1}")
        print(f"          🔄 Attempting with strict=False...")
        
        # Second attempt: strict=False - more forgiving for corrupt PDFs
        try:
            reader = PdfReader(pdf_path, strict=False)
            total = len(reader.pages)
            
            if total > LARGE_PDF_THRESHOLD:
                print(f"          📚 Large PDF: {total} pages - adding directly")
                return [pdf_path]
            
            page_paths = []
            failed_pages = 0
            
            for i, page in enumerate(reader.pages):
                try:
                    writer = PdfWriter()
                    writer.add_page(page)
                    out = os.path.join(temp_dir, f"pdf_page_{i+1:04d}.pdf")
                    with open(out, 'wb') as f:
                        writer.write(f)
                    page_paths.append(out)
                except Exception as page_error:
                    print(f"          ⚠️ Could not process page {i+1}: {page_error}")
                    failed_pages += 1
                    
                if total > 50 and (i+1) % 20 == 0:
                    print(f"          📄 Processed {i+1}/{total} pages ({failed_pages} failed)")
            
            if page_paths:
                if failed_pages > 0:
                    print(f"          ⚠️ {failed_pages} out of {total} pages failed - using {len(page_paths)} pages")
                return page_paths
            else:
                print(f"          📄 No pages could be extracted with strict=False")
                
        except Exception as e2:
            print(f"          ⚠️ Forgiving parsing also failed: {e2}")
        
        # Third attempt: Try LibreOffice
        print(f"          🔄 Attempting LibreOffice PDF import...")
        try:
            soffice = get_soffice_path()
            if soffice:
                # Create a temporary output PDF
                base = os.path.splitext(os.path.basename(pdf_path))[0]
                lo_output = os.path.join(temp_dir, f"{base}_libreoffice.pdf")
                
                # Use LibreOffice to import PDF and export as PDF again
                # This can sometimes fix corrupted PDFs
                env = create_isolated_env()
                cmd = [
                    soffice, '--headless', '--safe-mode', '--nologo',
                    '--nodefault', '--nofirststartwizard', '--norestore',
                    '--infilter=writer_pdf_import',  # Import as PDF
                    '--convert-to', 'pdf',
                    '--outdir', temp_dir, pdf_path
                ]
                
                result = subprocess.run(cmd, env=env, capture_output=True,
                                      timeout=LIBREOFFICE_TIMEOUT * 2)  # Double timeout for PDFs
                
                # LibreOffice might output with different name pattern
                possible_outputs = [
                    os.path.join(temp_dir, f"{base}.pdf"),
                    os.path.join(temp_dir, os.path.basename(pdf_path))
                ]
                
                for possible in possible_outputs:
                    if os.path.exists(possible) and os.path.getsize(possible) > 0:
                        # Now try to split this repaired PDF
                        try:
                            repaired_reader = PdfReader(possible)
                            repaired_total = len(repaired_reader.pages)
                            
                            if repaired_total > LARGE_PDF_THRESHOLD:
                                print(f"          📚 Large PDF after repair: {repaired_total} pages - adding directly")
                                return [possible]
                            
                            page_paths = []
                            for j, repaired_page in enumerate(repaired_reader.pages):
                                writer = PdfWriter()
                                writer.add_page(repaired_page)
                                out = os.path.join(temp_dir, f"pdf_page_{j+1:04d}_repaired.pdf")
                                with open(out, 'wb') as f:
                                    writer.write(f)
                                page_paths.append(out)
                            
                            if page_paths:
                                print(f"          ✅ LibreOffice repaired and split PDF into {len(page_paths)} pages")
                                return page_paths
                        except:
                            # If splitting fails, return the repaired PDF as-is
                            print(f"          ✅ LibreOffice repaired PDF (returning as single file)")
                            return [possible]
                            
                print(f"          ⚠️ LibreOffice PDF import produced no output")
            else:
                print(f"          ⚠️ LibreOffice not available")
                
        except subprocess.TimeoutExpired:
            print(f"          ⚠️ LibreOffice timed out on PDF import")
        except Exception as e3:
            print(f"          ⚠️ LibreOffice PDF import failed: {e3}")
        
        # Fourth attempt: Try to extract text from binary and create a text PDF
        print(f"          🔄 Attempting binary text extraction as last resort...")
        try:
            # Read the file as binary
            with open(pdf_path, 'rb') as f:
                content = f.read()
            
            # Try to decode as text (ignoring errors)
            text = content.decode('utf-8', errors='ignore')
            
            # Look for readable text (this is a very basic approach)
            import re
            # Find sequences of printable characters (words)
            words = re.findall(r'[a-zA-Z0-9\s\.\,\;\:\-\_]+', text)
            extracted_text = ' '.join(words)
            
            # Clean up excessive whitespace
            extracted_text = re.sub(r'\s+', ' ', extracted_text)
            
            if len(extracted_text.strip()) > 100:  # Only if we got meaningful content
                out = os.path.join(temp_dir, f"extracted_text_{os.path.basename(pdf_path)}.pdf")
                
                # Create a PDF with the extracted text
                if REPORTLAB_AVAILABLE:
                    doc = SimpleDocTemplate(out, pagesize=letter,
                                          leftMargin=1*inch, rightMargin=1*inch,
                                          topMargin=1*inch, bottomMargin=1*inch)
                    style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
                    
                    story = []
                    story.append(Paragraph(f"Extracted Text from: {os.path.basename(pdf_path)}", 
                                          ParagraphStyle('Header', fontName='Times-Bold', fontSize=10)))
                    story.append(Spacer(1, 0.2*inch))
                    story.append(Paragraph("(Original PDF could not be parsed - showing extracted text)", style))
                    story.append(Spacer(1, 0.1*inch))
                    
                    # Split into lines and add
                    for line in extracted_text.split('. ')[:200]:  # Limit to 200 sentences
                        if line.strip():
                            safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                            story.append(Paragraph(safe + '.', style))
                    
                    doc.build(story)
                    
                    if os.path.exists(out):
                        print(f"          ✅ Created text-extracted PDF with {len(extracted_text)} characters")
                        return [out]
            else:
                print(f"          ⚠️ Not enough readable text found in binary")
                
        except Exception as e4:
            print(f"          ⚠️ Binary extraction failed: {e4}")
    
    # If we get here, all parsing attempts failed
    error_msg = f"All PDF parsing attempts failed for {os.path.basename(pdf_path)}"
    print(f"          ❌ {error_msg}")
    raise ConversionError(error_msg)

def convert_image_to_pdf_pages(image_path, temp_dir):
    """Convert image to PDF pages with debug info"""
    try:
        out = os.path.join(temp_dir, f"img_{os.path.basename(image_path)}.pdf")
        print(f"          🔍 Opening image: {image_path}")
        
        if not PIL_AVAILABLE:
            raise ConversionError("Pillow library not available for image conversion")
            
        img = Image.open(image_path)
        print(f"          🔍 Image mode: {img.mode}, size: {img.size}")
        
        # Handle all image types
        if img.mode in ['P', 'PA']:  # Palette modes
            print(f"          🔍 Converting palette image to RGB")
            img = img.convert('RGB')
        elif img.mode not in ['RGB', 'L', 'CMYK']:
            print(f"          🔍 Converting {img.mode} to RGB")
            img = img.convert('RGB')
        
        # CMYK to RGB if needed
        if img.mode == 'CMYK':
            print(f"          🔍 Converting CMYK to RGB")
            img = img.convert('RGB')
        
        print(f"          🔍 Saving to PDF: {out}")
        img.save(out, format='PDF', quality=85, optimize=True)
        
        # Verify file was created
        if os.path.exists(out):
            size = os.path.getsize(out) / 1024
            print(f"          ✅ PDF created: {size:.1f} KB")
            return [out]
        else:
            error_msg = f"PDF file not created for {image_path}"
            print(f"          ❌ {error_msg}")
            raise ConversionError(error_msg)
            
    except ConversionError:
        raise
    except Exception as e:
        error_msg = f"Image conversion failed for {image_path}: {e}"
        print(f"          ❌ {error_msg}")
        import traceback
        traceback.print_exc()
        raise ConversionError(error_msg) from e

def convert_office_with_libreoffice(input_path, temp_dir, file_ext):
    soffice = get_soffice_path()
    if not soffice:
        raise ConversionError("LibreOffice not found")
    
    base = os.path.splitext(os.path.basename(input_path))[0]
    pdf_path = os.path.join(temp_dir, f"{base}.pdf")
    
    # Use absolute paths
    abs_input = os.path.abspath(input_path)
    abs_outdir = os.path.abspath(temp_dir)
    
    env = os.environ.copy()
    env.update({
        'HOME': temp_dir,
        'USERPROFILE': temp_dir,
        'APPDATA': temp_dir,
        'TEMP': temp_dir,
        'TMP': temp_dir,
    })
    
    cmd = [
        soffice,
        '--headless',
        '--safe-mode',
        '--nologo',
        '--nodefault',
        '--nofirststartwizard',
        '--norestore',
        '--convert-to', 'pdf',
        '--outdir', abs_outdir,
        abs_input
    ]
    
    print(f"          🔄 Running: {' '.join(cmd)}")
    try:
        result = subprocess.run(cmd, env=env, capture_output=True, timeout=200, check=False)
        print(f"          Return code: {result.returncode}")
        if result.stdout:
            print(f"          stdout: {result.stdout.decode('utf-8', errors='ignore')[:200]}")
        if result.stderr:
            print(f"          stderr: {result.stderr.decode('utf-8', errors='ignore')[:200]}")
        
        if result.returncode == 0 and os.path.exists(pdf_path):
            with open(pdf_path, 'rb') as f:
                if f.read(4) == b'%PDF':
                    return pdf_path
    except Exception as e:
        print(f"          Exception: {e}")
    
    # Kill lingering processes
    subprocess.run(['taskkill', '/f', '/im', 'soffice.exe'], capture_output=True)
    return None

def convert_powerpoint_with_text(input_path, temp_dir):
    """Fallback for PowerPoint files: extract text and create PDF."""
    try:
        base = os.path.splitext(os.path.basename(input_path))[0]
        pdf_path = os.path.join(temp_dir, f"{base}_ppt_text.pdf")
        
        if not REPORTLAB_AVAILABLE:
            return None
            
        text = ""
        
        # Method 1: Try textract if available
        if TEXTRACT_AVAILABLE:
            try:
                text = textract.process(input_path).decode('utf-8', errors='ignore')
                if text.strip():
                    print(f"          ✅ Textract extracted {len(text)} characters from PowerPoint")
            except Exception as e:
                print(f"          ⚠️ Textract failed for PowerPoint: {e}")
        
        # Method 2: If textract failed, try catppt if available (Linux/Mac)
        if not text.strip():
            try:
                # Check if catppt is installed (part of catdoc package on some systems)
                result = subprocess.run(['catppt', input_path], 
                                      capture_output=True, timeout=30)
                if result.returncode == 0:
                    text = result.stdout.decode('utf-8', errors='ignore')
                    print(f"          ✅ Catppt extracted {len(text)} characters")
            except (subprocess.SubprocessError, FileNotFoundError):
                pass
        
        # Method 3: Last resort - try to extract any readable text from binary
        if not text.strip():
            try:
                with open(input_path, 'rb') as f:
                    content = f.read()
                    # Try to decode as latin-1 and extract printable chars
                    raw_text = content.decode('latin-1')
                    # Keep only printable characters and basic punctuation
                    import string
                    printable = set(string.printable)
                    text = ''.join(c for c in raw_text if c in printable)
                    # Look for text that might be slide content (words separated by spaces)
                    # Clean up excessive whitespace and non-text
                    text = re.sub(r'[^\x20-\x7E\n\r\t]', ' ', text)
                    text = re.sub(r'\s+', ' ', text)
                    # Split into lines at reasonable intervals
                    words = text.split()
                    if len(words) > 10:  # Only if we got meaningful content
                        lines = []
                        current_line = []
                        for word in words:
                            current_line.append(word)
                            if len(' '.join(current_line)) > 80:
                                lines.append(' '.join(current_line))
                                current_line = []
                        if current_line:
                            lines.append(' '.join(current_line))
                        text = '\n'.join(lines)
                        print(f"          ✅ Binary extraction got {len(text)} characters")
            except Exception as e:
                print(f"          ⚠️ Binary extraction failed: {e}")
        
        if not text.strip():
            return None
        
        # Create PDF from extracted text
        doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                leftMargin=1*inch, rightMargin=1*inch,
                                topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
        
        story = []
        story.append(Paragraph(f"PowerPoint File: {os.path.basename(input_path)}", 
                              ParagraphStyle('Header', fontName='Times-Bold', fontSize=10, leading=12)))
        story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph("(Extracted text content - slide formatting may be lost)", style))
        story.append(Spacer(1, 0.2*inch))
        
        # Add text content
        lines = text.split('\n')
        for line in lines[:300]:  # Limit to 300 lines
            if line.strip():
                safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(safe, style))
        
        if len(lines) > 300:
            story.append(Paragraph("... (content truncated)", style))
        
        doc.build(story)
        
        if os.path.exists(pdf_path):
            size = os.path.getsize(pdf_path) / 1024
            print(f"          ✅ PowerPoint text PDF created: {size:.1f} KB")
            return pdf_path
        else:
            return None
            
    except Exception as e:
        print(f"          ⚠️ PowerPoint conversion error: {e}")
        return None
    
def convert_excel_with_pandas(input_path, temp_dir):
    """Fallback: use pandas to read Excel and create a simple PDF."""
    if not PANDAS_AVAILABLE:
        return None
    try:
        base = os.path.splitext(os.path.basename(input_path))[0]
        pdf_path = os.path.join(temp_dir, f"{base}_pandas.pdf")
        
        # Try different engines
        engines = []
        if XLRD_AVAILABLE:
            engines.append('xlrd')
        if OPENPYXL_AVAILABLE:
            engines.append('openpyxl')
        engines.append(None)
        
        for engine in engines:
            try:
                df_dict = pd.read_excel(input_path, sheet_name=None, engine=engine)
                break
            except:
                continue
        else:
            return None
        
        if not REPORTLAB_AVAILABLE:
            return None
            
        doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                leftMargin=0.5*inch, rightMargin=0.5*inch,
                                topMargin=0.5*inch, bottomMargin=0.5*inch)
        style_normal = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=7, leading=9)
        style_bold = ParagraphStyle('Bold', fontName='Times-Bold', fontSize=8, leading=10)
        
        story = []
        for sheet_idx, (sheet_name, df) in enumerate(df_dict.items()):
            story.append(Paragraph(f"Sheet: {sheet_name}", style_bold))
            story.append(Spacer(1, 0.1*inch))
            if not df.empty:
                # Convert to text table (first 50 rows)
                text = df.head(50).to_string(index=False, max_colwidth=20)
                for line in text.split('\n'):
                    if line.strip():
                        safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        story.append(Paragraph(safe, style_normal))
                if len(df) > 50:
                    story.append(Paragraph(f"... and {len(df)-50} more rows", style_normal))
            else:
                story.append(Paragraph("(Empty sheet)", style_normal))
            if sheet_idx < len(df_dict)-1:
                story.append(PageBreak())
        doc.build(story)
        return pdf_path
    except Exception as e:
        print(f"          ⚠️ pandas conversion error: {e}")
        return None

def convert_excel_with_xlrd(input_path, temp_dir):
    """Fallback: use xlrd for old .xls files."""
    if not XLRD_AVAILABLE:
        return None
    try:
        import xlrd
        wb = xlrd.open_workbook(input_path, formatting_info=False)
        base = os.path.splitext(os.path.basename(input_path))[0]
        pdf_path = os.path.join(temp_dir, f"{base}_xlrd.pdf")
        
        if not REPORTLAB_AVAILABLE:
            return None
            
        doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                leftMargin=0.5*inch, rightMargin=0.5*inch,
                                topMargin=0.5*inch, bottomMargin=0.5*inch)
        style_normal = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=7, leading=9)
        style_bold = ParagraphStyle('Bold', fontName='Times-Bold', fontSize=8, leading=10)
        
        story = []
        for sheet_idx in range(min(wb.nsheets, 5)):
            sheet = wb.sheet_by_index(sheet_idx)
            story.append(Paragraph(f"Sheet: {sheet.name}", style_bold))
            story.append(Spacer(1, 0.1*inch))
            if sheet.nrows > 0:
                # Headers
                headers = [str(sheet.cell_value(0, c))[:15] for c in range(min(sheet.ncols, 10))]
                story.append(Paragraph(" | ".join(headers), style_bold))
                story.append(Paragraph("-"*60, style_normal))
                # Data rows
                for r in range(1, min(sheet.nrows, 51)):
                    row = [str(sheet.cell_value(r, c))[:15] for c in range(min(sheet.ncols, 10))]
                    story.append(Paragraph(" | ".join(row), style_normal))
                if sheet.nrows > 51:
                    story.append(Paragraph(f"... and {sheet.nrows-51} more rows", style_normal))
            else:
                story.append(Paragraph("(Empty sheet)", style_normal))
            if sheet_idx < min(wb.nsheets,5)-1:
                story.append(PageBreak())
        doc.build(story)
        return pdf_path
    except Exception as e:
        print(f"          ⚠️ xlrd error: {e}")
        return None

def convert_word_document(input_path, temp_dir):
    """Enhanced Word document conversion with multiple fallbacks."""
    try:
        base = os.path.splitext(os.path.basename(input_path))[0]
        pdf_path = os.path.join(temp_dir, f"{base}_word.pdf")
        
        if not REPORTLAB_AVAILABLE:
            return None
            
        text = ""
        
        # Method 1: Try textract if available
        if TEXTRACT_AVAILABLE:
            try:
                text = textract.process(input_path).decode('utf-8', errors='ignore')
                if text.strip():
                    print(f"          ✅ Textract extracted {len(text)} characters")
            except Exception as e:
                print(f"          ⚠️ Textract failed: {e}")
        
        # Method 2: If textract failed or returned empty, try antiword (if available)
        if not text.strip():
            try:
                # Check if antiword is installed (common on Linux/Mac for .doc files)
                result = subprocess.run(['antiword', input_path], 
                                      capture_output=True, timeout=30)
                if result.returncode == 0:
                    text = result.stdout.decode('utf-8', errors='ignore')
                    print(f"          ✅ Antiword extracted {len(text)} characters")
            except (subprocess.SubprocessError, FileNotFoundError):
                pass
        
        # Method 3: Try catdoc if available
        if not text.strip():
            try:
                result = subprocess.run(['catdoc', input_path], 
                                      capture_output=True, timeout=30)
                if result.returncode == 0:
                    text = result.stdout.decode('utf-8', errors='ignore')
                    print(f"          ✅ Catdoc extracted {len(text)} characters")
            except (subprocess.SubprocessError, FileNotFoundError):
                pass
        
        # Method 4: Last resort - try to read as binary and extract any readable text
        if not text.strip():
            try:
                with open(input_path, 'rb') as f:
                    content = f.read()
                    # Try to decode as latin-1 (which never fails) and extract printable chars
                    raw_text = content.decode('latin-1')
                    # Keep only printable characters and basic punctuation
                    import string
                    printable = set(string.printable)
                    text = ''.join(c for c in raw_text if c in printable)
                    # Clean up excessive whitespace
                    text = re.sub(r'\s+', ' ', text)
                    print(f"          ✅ Binary extraction got {len(text)} characters")
            except Exception as e:
                print(f"          ⚠️ Binary extraction failed: {e}")
        
        if not text.strip():
            return None
        
        # Create PDF from extracted text
        doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                leftMargin=1*inch, rightMargin=1*inch,
                                topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
        
        story = []
        story.append(Paragraph(f"File: {os.path.basename(input_path)}", 
                              ParagraphStyle('Header', fontName='Times-Bold', fontSize=10, leading=12)))
        story.append(Spacer(1, 0.1*inch))
        
        # Add text content line by line (limit to first 500 lines to avoid huge PDFs)
        lines = text.split('\n')
        for line in lines[:500]:
            if line.strip():
                safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(safe, style))
        
        if len(lines) > 500:
            story.append(Paragraph(f"... (text truncated, {len(lines)-500} more lines)", style))
        
        doc.build(story)
        
        if os.path.exists(pdf_path):
            size = os.path.getsize(pdf_path) / 1024
            print(f"          ✅ Word PDF created: {size:.1f} KB")
            return pdf_path
        else:
            return None
            
    except Exception as e:
        print(f"          ⚠️ Word conversion error: {e}")
        return None
    
def convert_office_document(input_path, filename, temp_dir):
    """Main office conversion: tries LibreOffice first, then fallbacks."""
    ext = os.path.splitext(filename)[1].lower()
    print(f"          🔄 Converting {ext} document with LibreOffice")
    
    # Try LibreOffice
    pdf = convert_office_with_libreoffice(input_path, temp_dir, ext)
    if pdf:
        return pdf
    
    # If LibreOffice fails, raise error
    error_msg = f"LibreOffice could not convert {ext} document: {filename}"
    print(f"          ❌ {error_msg}")
    raise ConversionError(error_msg)

def convert_text_to_pdf_pages(text_path, temp_dir, font_size=8):
    """Convert a text file to PDF pages with specified font size."""
    try:
        out = os.path.join(temp_dir, f"text_{os.path.basename(text_path)}.pdf")
        print(f"          🔍 Reading text file: {text_path}")
        
        if not REPORTLAB_AVAILABLE:
            raise ConversionError("reportlab not available for text conversion")
        
        # Read the text file
        with open(text_path, 'r', encoding='utf-8', errors='ignore') as f:
            text_content = f.read()
        
        # Create PDF with specified font size
        doc = SimpleDocTemplate(out, pagesize=letter,
                                leftMargin=1*inch, rightMargin=1*inch,
                                topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=font_size, leading=font_size+2)
        header_style = ParagraphStyle('Header', fontName='Times-Bold', fontSize=font_size+2, leading=font_size+4)
        
        story = []
        # Add filename as header
        story.append(Paragraph(f"File: {os.path.basename(text_path)}", header_style))
        story.append(Spacer(1, 0.1*inch))
        
        # Add text content line by line
        for line in text_content.split('\n'):
            if line.strip():  # Skip empty lines
                safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(safe, style))
        
        doc.build(story)
        
        # Verify file was created
        if os.path.exists(out):
            size = os.path.getsize(out) / 1024
            print(f"          ✅ Text PDF created: {size:.1f} KB")
            return [out]
        else:
            raise ConversionError(f"PDF file not created for {text_path}")
            
    except ConversionError:
        raise
    except Exception as e:
        error_msg = f"Text conversion failed for {text_path}: {e}"
        print(f"          ❌ {error_msg}")
        raise ConversionError(error_msg) from e

def convert_rtf_to_pdf_pages(rtf_path, temp_dir):
    """Convert RTF file to PDF by extracting text content."""
    try:
        out = os.path.join(temp_dir, f"rtf_{os.path.basename(rtf_path)}.pdf")
        print(f"          🔍 Reading RTF file: {rtf_path}")
        
        if not REPORTLAB_AVAILABLE:
            raise ConversionError("reportlab not available for RTF conversion")
        
        text = ""
        
        # Method 1: Try textract if available
        if TEXTRACT_AVAILABLE:
            try:
                text = textract.process(rtf_path).decode('utf-8', errors='ignore')
                if text.strip():
                    print(f"          ✅ Textract extracted {len(text)} characters from RTF")
            except Exception as e:
                print(f"          ⚠️ Textract failed for RTF: {e}")
        
        # Method 2: Try unrtf if available (common Linux/Mac tool)
        if not text.strip():
            try:
                result = subprocess.run(['unrtf', '--text', rtf_path], 
                                      capture_output=True, timeout=30)
                if result.returncode == 0:
                    text = result.stdout.decode('utf-8', errors='ignore')
                    print(f"          ✅ UnRTF extracted {len(text)} characters")
            except (subprocess.SubprocessError, FileNotFoundError):
                pass
        
        # Method 3: Manual RTF stripping (basic)
        if not text.strip():
            try:
                with open(rtf_path, 'r', encoding='utf-8', errors='ignore') as f:
                    rtf_content = f.read()
                
                # Very basic RTF tag stripping
                # Remove RTF control words and groups
                text = re.sub(r'\\[a-z]+[-\d]*', ' ', rtf_content)  # Remove control words
                text = re.sub(r'\{[^}]*\}', ' ', text)  # Remove groups
                text = re.sub(r'\\\'[a-f0-9]{2}', ' ', text)  # Remove hex escapes
                text = re.sub(r'[{}]', ' ', text)  # Remove braces
                # Clean up whitespace
                text = re.sub(r'\s+', ' ', text)
                text = re.sub(r'\n\s*\n', '\n\n', text)
                
                if text.strip():
                    print(f"          ✅ Manual RTF stripping extracted {len(text)} characters")
            except Exception as e:
                print(f"          ⚠️ Manual RTF stripping failed: {e}")
        
        if not text.strip():
            raise ConversionError(f"No text could be extracted from RTF file {rtf_path}")
        
        # Create PDF from extracted text
        doc = SimpleDocTemplate(out, pagesize=letter,
                                leftMargin=1*inch, rightMargin=1*inch,
                                topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
        
        story = []
        story.append(Paragraph(f"RTF File: {os.path.basename(rtf_path)}", 
                              ParagraphStyle('Header', fontName='Times-Bold', fontSize=10, leading=12)))
        story.append(Spacer(1, 0.1*inch))
        
        # Add text content
        lines = text.split('\n')
        for line in lines[:500]:
            if line.strip():
                safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                story.append(Paragraph(safe, style))
        
        if len(lines) > 500:
            story.append(Paragraph("... (content truncated)", style))
        
        doc.build(story)
        
        if os.path.exists(out):
            size = os.path.getsize(out) / 1024
            print(f"          ✅ RTF PDF created: {size:.1f} KB")
            return [out]
        else:
            raise ConversionError(f"PDF file not created for {rtf_path}")
            
    except ConversionError:
        raise
    except Exception as e:
        error_msg = f"RTF conversion failed for {rtf_path}: {e}"
        print(f"          ❌ {error_msg}")
        raise ConversionError(error_msg) from e
    
def convert_html_to_pdf_pages(html_path, temp_dir):
    """Convert HTML file to PDF pages by extracting text content."""
    try:
        out = os.path.join(temp_dir, f"html_{os.path.basename(html_path)}.pdf")
        print(f"          🔍 Reading HTML file: {html_path}")
        
        if not REPORTLAB_AVAILABLE:
            raise ConversionError("reportlab not available for HTML conversion")
        
        # Read the HTML file
        with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
            html_content = f.read()
        
        # Simple HTML tag stripping (basic approach)
        import re
        # Remove scripts and style tags and their content
        text = re.sub(r'<script.*?>.*?</script>', '', html_content, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<style.*?>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        # Remove HTML tags
        text = re.sub(r'<[^>]+>', ' ', text)
        # Decode HTML entities
        text = re.sub(r'&nbsp;', ' ', text)
        text = re.sub(r'&amp;', '&', text)
        text = re.sub(r'&lt;', '<', text)
        text = re.sub(r'&gt;', '>', text)
        text = re.sub(r'&quot;', '"', text)
        text = re.sub(r'&#39;', "'", text)
        # Clean up whitespace
        text = re.sub(r'\s+', ' ', text)
        text = re.sub(r'\n\s*\n', '\n\n', text)
        
        # Split into lines for PDF
        lines = text.split('\n')
        
        # Create PDF
        doc = SimpleDocTemplate(out, pagesize=letter,
                                leftMargin=1*inch, rightMargin=1*inch,
                                topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
        
        story = []
        story.append(Paragraph(f"HTML File: {os.path.basename(html_path)}", 
                              ParagraphStyle('Header', fontName='Times-Bold', fontSize=10, leading=12)))
        story.append(Spacer(1, 0.1*inch))
        
        # Add text content
        line_count = 0
        for line in lines:
            if line.strip():
                # Limit line length to avoid PDF issues
                line = line.strip()
                if len(line) > 200:
                    # Split long lines
                    parts = [line[i:i+200] for i in range(0, len(line), 200)]
                    for part in parts:
                        safe = part.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                        story.append(Paragraph(safe, style))
                        line_count += 1
                else:
                    safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    story.append(Paragraph(safe, style))
                    line_count += 1
                
                if line_count > 1000:  # Limit total lines
                    story.append(Paragraph("... (content truncated due to length)", style))
                    break
        
        doc.build(story)
        
        # Verify file was created
        if os.path.exists(out):
            size = os.path.getsize(out) / 1024
            print(f"          ✅ HTML PDF created: {size:.1f} KB")
            return [out]
        else:
            raise ConversionError(f"PDF file not created for {html_path}")
            
    except ConversionError:
        raise
    except Exception as e:
        error_msg = f"HTML conversion failed for {html_path}: {e}"
        print(f"          ❌ {error_msg}")
        raise ConversionError(error_msg) from e
    
def handle_unconvertible_file(file_path, filename, temp_dir, file_type):
    """Create a summary PDF for files that cannot be converted to PDF."""
    try:
        out = os.path.join(temp_dir, f"skip_{sanitize_filename(filename)}.pdf")
        doc = SimpleDocTemplate(out, pagesize=letter,
                                leftMargin=1*inch, rightMargin=1*inch,
                                topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
        bold = ParagraphStyle('Bold', fontName='Times-Bold', fontSize=9, leading=12)
        
        stat = os.stat(file_path)
        size_mb = stat.st_size / (1024 * 1024)
        
        # File type descriptions
        type_descriptions = {
            '.p7s': 'Digital Signature (PKCS#7)',
            '.p7m': 'Digital Signature (PKCS#7)',
            '.p7c': 'Digital Signature (PKCS#7)',
            '.mp3': 'Audio File (MP3)',
            '.mp4': 'Video File (MP4)',
            '.wav': 'Audio File (WAV)',
            '.avi': 'Video File (AVI)',
            '.mov': 'Video File (QuickTime)',
            '.wmv': 'Video File (Windows Media)',
            '.flv': 'Video File (Flash)',
            '.mkv': 'Video File (Matroska)',
            '.m4a': 'Audio File (AAC)',
            '.aac': 'Audio File (AAC)',
            '.ogg': 'Audio File (OGG)',
            '.flac': 'Audio File (FLAC)',
            '.wma': 'Audio File (Windows Media)',
            '.zip': 'Compressed Archive',
            '.zap': 'Compressed Archive',
            '.pub': 'public key',
            '.sig': 'signature',
            '.rar': 'Compressed Archive',
            '.7z': 'Compressed Archive',
            '.tar': 'Compressed Archive',
            '.gz': 'Compressed Archive',
            '.exe': 'Executable File',
            '.dll': 'Dynamic Link Library',
            '.msi': 'Windows Installer Package',
            '.iso': 'Disk Image',
            '.bin': 'Binary File',
            '.dat': 'Data File'
        }
        
        description = type_descriptions.get(file_type.lower(), f'File Type: {file_type.upper()}')
        
        story = [
            Paragraph(f"MEDIA/SIGNATURE FILE: {filename}", bold),
            Spacer(1, 0.1*inch),
            Paragraph(f"Type: {description}", style),
            Paragraph(f"Size: {size_mb:.2f} MB", style),
            Paragraph(f"Modified: {datetime.fromtimestamp(stat.st_mtime)}", style),
            Spacer(1, 0.2*inch),
            Paragraph("This file type cannot be converted to PDF content.", style),
            Paragraph(f"The original {file_type.upper()} file is embedded below.", style),
            Spacer(1, 0.1*inch),
            Paragraph("To access this file:", style),
            Paragraph("• Adobe Acrobat Reader: Click on the paperclip icon in the Attachments panel", style),
            Paragraph("• Other PDF viewers: Look for an attachment or paperclip icon", style),
            Paragraph("• Save the PDF and use a file extraction tool to get the original file", style)
        ]
        doc.build(story)
        return [out]
    except Exception as e:
        print(f"          ⚠️ Could not create skip summary: {e}")
        return None

def convert_calendar_to_pdf_pages(cal_path, filename, temp_dir, ext):
    """Convert calendar/contact files to PDF."""
    try:
        print(f"          📅 {ext.upper()} calendar/contact file detected - converting to text PDF")
        
        out = os.path.join(temp_dir, f"calendar_{sanitize_filename(filename)}.pdf")
        
        # Read the file as text
        with open(cal_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        if not REPORTLAB_AVAILABLE:
            raise ConversionError("reportlab not available for calendar conversion")
        
        import re
        
        doc = SimpleDocTemplate(out, pagesize=letter,
                              leftMargin=1*inch, rightMargin=1*inch,
                              topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
        bold = ParagraphStyle('Bold', fontName='Times-Bold', fontSize=9, leading=12)
        
        story = []
        story.append(Paragraph(f"Calendar/Contact File: {filename}", bold))
        story.append(Paragraph(f"File Type: {ext.upper()}", style))
        story.append(Spacer(1, 0.1*inch))
        
        # Handle different calendar formats
        if ext in ['.ics', '.ical', '.icalendar', '.ifb', '.vcs', '.vcalendar']:
            # iCalendar format
            events = re.findall(r'BEGIN:VEVENT(.*?)END:VEVENT', content, re.DOTALL | re.IGNORECASE)
            todos = re.findall(r'BEGIN:VTODO(.*?)END:VTODO', content, re.DOTALL | re.IGNORECASE)
            journals = re.findall(r'BEGIN:VJOURNAL(.*?)END:VJOURNAL', content, re.DOTALL | re.IGNORECASE)
            
            story.append(Paragraph(f"Events: {len(events)}", bold))
            story.append(Paragraph(f"Tasks: {len(todos)}", bold))
            story.append(Paragraph(f"Journals: {len(journals)}", bold))
            story.append(Spacer(1, 0.2*inch))
            
            # Process events
            if events:
                story.append(Paragraph("EVENTS:", bold))
                for i, event in enumerate(events[:30]):
                    summary = re.search(r'SUMMARY[:\s]+(.*?)[\r\n]', event, re.IGNORECASE)
                    dtstart = re.search(r'DTSTART[:\s]+(.*?)[\r\n]', event, re.IGNORECASE)
                    dtend = re.search(r'DTEND[:\s]+(.*?)[\r\n]', event, re.IGNORECASE)
                    location = re.search(r'LOCATION[:\s]+(.*?)[\r\n]', event, re.IGNORECASE)
                    
                    story.append(Paragraph(f"  Event {i+1}: {summary.group(1) if summary else 'Unnamed'}", style))
                    if dtstart:
                        story.append(Paragraph(f"    Start: {dtstart.group(1)}", style))
                    if dtend:
                        story.append(Paragraph(f"    End: {dtend.group(1)}", style))
                    if location:
                        story.append(Paragraph(f"    Location: {location.group(1)}", style))
                    story.append(Spacer(1, 0.05*inch))
                
                if len(events) > 30:
                    story.append(Paragraph(f"  ... and {len(events)-30} more events", style))
            
            # Process tasks
            if todos:
                story.append(Spacer(1, 0.1*inch))
                story.append(Paragraph("TASKS:", bold))
                for i, todo in enumerate(todos[:20]):
                    summary = re.search(r'SUMMARY[:\s]+(.*?)[\r\n]', todo, re.IGNORECASE)
                    due = re.search(r'DUE[:\s]+(.*?)[\r\n]', todo, re.IGNORECASE)
                    status = re.search(r'STATUS[:\s]+(.*?)[\r\n]', todo, re.IGNORECASE)
                    
                    story.append(Paragraph(f"  Task {i+1}: {summary.group(1) if summary else 'Unnamed'}", style))
                    if due:
                        story.append(Paragraph(f"    Due: {due.group(1)}", style))
                    if status:
                        story.append(Paragraph(f"    Status: {status.group(1)}", style))
                    story.append(Spacer(1, 0.05*inch))
        
        elif ext in ['.vcf', '.vcard']:
            # vCard contact format
            contacts = re.findall(r'BEGIN:VCARD(.*?)END:VCARD', content, re.DOTALL | re.IGNORECASE)
            
            story.append(Paragraph(f"Contacts: {len(contacts)}", bold))
            story.append(Spacer(1, 0.2*inch))
            
            for i, contact in enumerate(contacts[:50]):
                # Extract contact details
                fn = re.search(r'FN[:\s]+(.*?)[\r\n]', contact, re.IGNORECASE)
                n = re.search(r'N[:\s]+(.*?)[\r\n]', contact, re.IGNORECASE)
                email = re.search(r'EMAIL[:\s]+(.*?)[\r\n]', contact, re.IGNORECASE)
                tel = re.search(r'TEL[:\s]+(.*?)[\r\n]', contact, re.IGNORECASE)
                org = re.search(r'ORG[:\s]+(.*?)[\r\n]', contact, re.IGNORECASE)
                
                story.append(Paragraph(f"Contact {i+1}:", bold))
                if fn:
                    story.append(Paragraph(f"  Name: {fn.group(1)}", style))
                elif n:
                    story.append(Paragraph(f"  Name: {n.group(1)}", style))
                if email:
                    story.append(Paragraph(f"  Email: {email.group(1)}", style))
                if tel:
                    story.append(Paragraph(f"  Phone: {tel.group(1)}", style))
                if org:
                    story.append(Paragraph(f"  Organization: {org.group(1)}", style))
                story.append(Spacer(1, 0.05*inch))
            
            if len(contacts) > 50:
                story.append(Paragraph(f"... and {len(contacts)-50} more contacts", style))
        
        else:
            # Fallback for unknown calendar formats - show raw text
            story.append(Paragraph("Raw calendar data:", bold))
            lines = content.split('\n')[:100]
            for line in lines:
                if line.strip():
                    safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                    story.append(Paragraph(safe, style))
        
        doc.build(story)
        
        if os.path.exists(out):
            size = os.path.getsize(out) / 1024
            print(f"          ✅ Created calendar PDF: {size:.1f} KB")
            return [out]
        else:
            raise ConversionError(f"PDF file not created for {filename}")
            
    except Exception as e:
        print(f"          ⚠️ Calendar conversion failed: {e}")
        # Fall back to text conversion
        return convert_text_to_pdf_pages(cal_path, temp_dir)

def convert_latex_to_pdf_pages(tex_path, temp_dir):
    """Convert LaTeX file to PDF."""
    try:
        print(f"          📄 LaTeX file detected - converting to PDF")
        
        # First try to compile with pdflatex if available
        import subprocess
        import shutil
        
        # Check if pdflatex is installed
        pdflatex_path = shutil.which('pdflatex')
        
        if pdflatex_path:
            print(f"          🔄 Attempting to compile with pdflatex...")
            # Create a temporary directory for LaTeX compilation
            latex_temp = os.path.join(temp_dir, 'latex_compile')
            os.makedirs(latex_temp, exist_ok=True)
            
            # Copy tex file to temp dir
            tex_filename = os.path.basename(tex_path)
            tex_temp_path = os.path.join(latex_temp, tex_filename)
            shutil.copy2(tex_path, tex_temp_path)
            
            # Run pdflatex (need to run twice for references)
            for i in range(2):
                result = subprocess.run(
                    [pdflatex_path, '-interaction=nonstopmode', '-output-directory', latex_temp, tex_temp_path],
                    cwd=latex_temp,
                    capture_output=True,
                    timeout=60
                )
                if result.returncode != 0:
                    print(f"          ⚠️ pdflatex run {i+1} had issues")
            
            # Look for generated PDF
            pdf_name = os.path.splitext(tex_filename)[0] + '.pdf'
            pdf_path = os.path.join(latex_temp, pdf_name)
            
            if os.path.exists(pdf_path):
                print(f"          ✅ LaTeX compilation successful")
                # Convert the PDF to pages
                return convert_pdf_to_pdf_pages(pdf_path, temp_dir)
            else:
                print(f"          ⚠️ pdflatex did not produce a PDF")
                
        # Fallback: treat as text file
        print(f"          🔄 Falling back to text extraction")
        return convert_text_to_pdf_pages(tex_path, temp_dir)
        
    except Exception as e:
        print(f"          ⚠️ LaTeX conversion failed: {e}")
        # Fallback to text
        return convert_text_to_pdf_pages(tex_path, temp_dir)
    
def convert_attachment_to_pdf(attachment_path, filename, temp_dir):
    """Route attachment to appropriate converter."""
    ext = os.path.splitext(filename)[1].lower()
    
    try:
        # If no extension, try multiple formats in sequence
        if not ext:
            print(f"          🔍 No file extension detected - trying multiple formats")
            
            # Try these extensions in order
            try_extensions = [
                '.jpg', '.txt', '.xls', '.doc', '.docx', '.pdf', '.html', '.rtf', 
                '.png', '.csv', '.ppt', '.xml', '.log', '.ics', '.vcs', '.vcf',
                '.vsd', '.vsdx', '.vss', '.vst', '.vsw', '.vsdm', '.vssx', '.vssm', '.vstx', '.vstm',
                '.tex', '.asc', '.odt', '.ods', '.odp', '.spv', '.sav', '.emz', '.mso',
                '.py', '.js', '.java', '.c', '.cpp', '.sql', '.json', '.md'
            ]
            
            for try_ext in try_extensions:
                print(f"          🔄 Attempting as {try_ext} file...")
                try:
                    # Temporarily pretend the file has this extension
                    if try_ext == '.jpg':
                        # Try as image
                        result = convert_image_to_pdf_pages(attachment_path, temp_dir)
                        if result:
                            print(f"          ✅ Successfully processed as image")
                            return result
                    elif try_ext == '.txt':
                        # Try as text
                        result = convert_text_to_pdf_pages(attachment_path, temp_dir)
                        if result:
                            print(f"          ✅ Successfully processed as text")
                            return result
                    elif try_ext in ['.xls', '.doc', '.ppt', '.csv', '.xml']:
                        # Try as office document
                        result = convert_office_document(attachment_path, filename + try_ext, temp_dir)
                        if result:
                            pdf_pages = convert_pdf_to_pdf_pages(result, temp_dir)
                            if pdf_pages:
                                print(f"          ✅ Successfully processed as office document")
                                return pdf_pages
                    elif try_ext == '.pdf':
                        # Try as PDF
                        result = convert_pdf_to_pdf_pages(attachment_path, temp_dir)
                        if result:
                            print(f"          ✅ Successfully processed as PDF")
                            return result
                    elif try_ext in ['.html', '.htm']:
                        # Try as HTML
                        result = convert_html_to_pdf_pages(attachment_path, temp_dir)
                        if result:
                            print(f"          ✅ Successfully processed as HTML")
                            return result
                    elif try_ext == '.rtf':
                        # Try as RTF
                        result = convert_rtf_to_pdf_pages(attachment_path, temp_dir)
                        if result:
                            print(f"          ✅ Successfully processed as RTF")
                            return result
                    elif try_ext in ['.ics', '.vcs', '.vcf', '.ifb', '.ical', '.icalendar', '.cal']:
                        # Try as calendar/contact file
                        result = convert_calendar_to_pdf_pages(attachment_path, filename, temp_dir, try_ext)
                        if result:
                            print(f"          ✅ Successfully processed as calendar/contact file")
                            return result
                    elif try_ext in ['.vsd', '.vsdx', '.vss', '.vst', '.vsw', '.vsdm', '.vssx', '.vssm', '.vstx', '.vstm']:
                        # Try as Visio file
                        result = convert_office_document(attachment_path, filename + try_ext, temp_dir)
                        if result:
                            pdf_pages = convert_pdf_to_pdf_pages(result, temp_dir)
                            if pdf_pages:
                                print(f"          ✅ Successfully processed as Visio diagram")
                                return pdf_pages
                    elif try_ext == '.tex':
                        # Try as LaTeX file
                        result = convert_latex_to_pdf_pages(attachment_path, temp_dir)
                        if result:
                            print(f"          ✅ Successfully processed as LaTeX document")
                            return result
                    elif try_ext in ['.py', '.js', '.java', '.c', '.cpp', '.sql', '.json', '.md']:
                        # Try as code/text file
                        result = convert_text_to_pdf_pages(attachment_path, temp_dir)
                        if result:
                            print(f"          ✅ Successfully processed as code/text file")
                            return result
                except Exception as e:
                    print(f"          ⚠️ Failed as {try_ext}: {e}")
                    continue
            
            # If all attempts fail, treat as .mp3 to trigger the unconvertible handler
            print(f"          🔄 All format attempts failed - treating as audio file to trigger skip")
            ext = '.mp3'
        
        # Regular extension handling
        # PDF
        if ext == '.pdf':
            return convert_pdf_to_pdf_pages(attachment_path, temp_dir)
        
        # XPS files - handle with LibreOffice
        if ext == '.xps':
            print(f"          🔄 Converting XPS document with LibreOffice...")
            # Try LibreOffice first
            pdf = convert_office_with_libreoffice(attachment_path, temp_dir, ext)
            if pdf:
                return convert_pdf_to_pdf_pages(pdf, temp_dir)
            
            # If LibreOffice fails, create a summary
            print(f"          ⚠️ Could not convert XPS document")
            summary = handle_unconvertible_file(attachment_path, filename, temp_dir, ext)
            if summary:
                return summary
            else:
                raise ConversionError(f"Could not convert XPS document: {filename}")
        
        # Visio files
        visio_extensions = ['.vsd', '.vsdx', '.vss', '.vst', '.vsw', '.vsdm', '.vssx', '.vssm', '.vstx', '.vstm']
        
        if ext in visio_extensions:
            print(f"          📊 Visio diagram file detected - attempting conversion with LibreOffice...")
            # Try LibreOffice first (it can open some Visio files)
            pdf = convert_office_with_libreoffice(attachment_path, temp_dir, ext)
            if pdf:
                return convert_pdf_to_pdf_pages(pdf, temp_dir)
            
            # If LibreOffice fails, create a summary
            print(f"          ⚠️ Could not convert Visio diagram")
            summary = handle_unconvertible_file(attachment_path, filename, temp_dir, ext)
            if summary:
                return summary
            else:
                raise ConversionError(f"Could not convert Visio diagram: {filename}")
        
 
        # Images - comprehensive list of all common image extensions
        image_extensions = [
            # JPEG variants
            '.jpg', '.jpeg', '.jpe', '.jfif', '.jif', '.jfi',
            # PNG and GIF
            '.png', '.gif',
            # BMP variants
            '.bmp', '.dib', '.rle',
            # TIFF variants
            '.tiff', '.tif',
            # WebP
            '.webp',
            # HEIC/HEIF (modern iPhone formats)
            '.heic', '.heif', '.heics', '.heifs',
            # Icons
            '.ico', '.cur',
            # Vector formats (may be handled as text/images)
            '.svg', '.svgz', '.eps', '.ai', '.cdr',
            # Photoshop and GIMP
            '.psd', '.psb', '.xcf',
            # Camera RAW formats
            '.raw', '.cr2', '.cr3', '.nef', '.nrw', '.arw', '.srf', '.sr2',
            '.dng', '.orf', '.ptx', '.pef', '.rw2', '.raf', '.3fr', '.kdc',
            '.dcr', '.mrw', '.bay', '.erf', '.mef', '.mos', '.iiq',
            # Other common image formats
            '.jp2', '.j2k', '.jpf', '.jpx', '.jpm',  # JPEG 2000
            '.pgm', '.ppm', '.pbm', '.pnm',  # Netpbm formats
            '.pcx', '.tga', '.icns', '.hdp', '.jxr', '.wdp',  # Other formats
            '.dds', '.dcm', '.dicm',  # Medical/texture formats
            '.exr', '.hdr',  # HDR formats
        ]
        
        if ext in image_extensions:
            return convert_image_to_pdf_pages(attachment_path, temp_dir)
        
        # Calendar and contact files
        calendar_extensions = [
            '.ics', '.ical', '.icalendar', '.ifb', '.vcs',  # iCalendar formats
            '.vcf', '.vcard',  # vCard contact formats
            '.cal', '.calendar',  # Generic calendar
            '.event', '.todo', '.task',  # Event formats
            '.vcalendar',  # vCalendar
            '.xcal', '.xcs',  # XML calendar formats
        ]
        
        if ext in calendar_extensions:
            return convert_calendar_to_pdf_pages(attachment_path, filename, temp_dir, ext)
        
        # LaTeX files
        if ext == '.tex':
            return convert_latex_to_pdf_pages(attachment_path, temp_dir)
        
        # Python/Code files - treat as text
        code_extensions = [
            '.py', '.js', '.java', '.c', '.cpp', '.h', '.cs', '.php', '.rb', '.go',
            '.rs', '.swift', '.kt', '.scala', '.pl', '.pm', '.tcl', '.lua', '.r',
            '.m', '.sql', '.sh', '.bash', '.zsh', '.fish', '.ps1', '.bat', '.cmd',
            '.xml', '.json', '.yaml', '.yml', '.toml', '.ini', '.cfg', '.conf',
            '.css', '.scss', '.less', '.md', '.markdown', '.rst'
        ]
        
        if ext in code_extensions:
            print(f"          💻 Code file detected - converting as text")
            return convert_text_to_pdf_pages(attachment_path, temp_dir, font_size=6)
        
        # Text files
        if ext == '.txt':
            return convert_text_to_pdf_pages(attachment_path, temp_dir)
        
        # HTML files
        if ext in ['.html', '.htm']:
            return convert_html_to_pdf_pages(attachment_path, temp_dir)
        
        # Office (including RTF)
        if ext in OFFICE_TYPES:
            pdf = convert_office_document(attachment_path, filename, temp_dir)
            if pdf:
                return convert_pdf_to_pdf_pages(pdf, temp_dir)
        
        # Media files, signatures, and other unconvertible types - create skip summary
        unconvertible_types = [
            '.p7s', '.p7m', '.p7c', '.pub', '.sig', '.asc', '.spv', '.sav', '.emz', '.mso',  # Signatures & proprietary
            '.mp3', '.mp4', '.wav', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.opus', # Audio/Video
            '.m4a', '.aac', '.ogg', '.flac', '.wma', '.pages', '.HEIC', # More audio
            '.zip', '.zap', '.rar', '.7z', '.tar', '.gz', '.msg', '.rpmsg', # Archives & Installer packages
            '.exe', '.dll', '.msi',  # Executables
            '.iso', '.bin', '.dat', '.img',  # Disk images and binary files
            '.cab', '.dmg', '.vhd', '.vmdk',  # More disk images
            '.reg', '.ini', '.cfg', '.config',  # Configuration files
            '.log', '.tmp', '.bak', '.old',  # Temporary/backup files
        ]
        
        if ext in unconvertible_types:
            print(f"          🔍 {ext.upper()} file detected - creating skip summary")
            summary = handle_unconvertible_file(attachment_path, filename, temp_dir, ext)
            if summary:
                return summary
        
        # Unknown file type
        error_msg = f"Unsupported file type: {ext}"
        print(f"          ⚠️ {error_msg}")
        raise ConversionError(error_msg)
        
    except ConversionError:
        raise
    except Exception as e:
        error_msg = f"Unexpected error converting {filename}: {e}"
        print(f"          ❌ {error_msg}")
        raise ConversionError(error_msg) from e

    
def create_file_summary_page(file_path, filename, temp_dir):
    """Create a simple PDF summary when conversion fails."""
    try:
        out = os.path.join(temp_dir, f"summary_{sanitize_filename(filename)}.pdf")
        doc = SimpleDocTemplate(out, pagesize=letter,
                                leftMargin=1*inch, rightMargin=1*inch,
                                topMargin=1*inch, bottomMargin=1*inch)
        style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
        bold = ParagraphStyle('Bold', fontName='Times-Bold', fontSize=9, leading=12)
        
        stat = os.stat(file_path)
        mime = get_file_type(file_path)
        size_kb = stat.st_size / 1024
        
        story = [
            Paragraph(f"FILE: {filename}", bold),
            Spacer(1, 0.1*inch),
            Paragraph(f"Type: {mime}", style),
            Paragraph(f"Size: {size_kb:.1f} KB", style),
            Paragraph(f"Modified: {datetime.fromtimestamp(stat.st_mtime)}", style),
            Spacer(1, 0.2*inch),
            Paragraph("This file could not be automatically converted to PDF.", style),
            Paragraph("The original file is embedded below.", style)
        ]
        doc.build(story)
        return out
    except:
        return None

def sanitize_filename(filename):
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    filename = re.sub(r'_+', '_', filename)
    return filename.strip(' .') or "unnamed"

# ========== EMAIL PROCESSING ==========

def extract_tags(message):
    """Extract Gmail labels from X-Gmail-Labels header."""
    tags = []
    x_labels = message.get('X-Gmail-Labels', '')
    if x_labels:
        raw = [t.strip() for t in x_labels.split(',')]
        exclude = ['Inbox', 'Sent', 'Draft', 'Spam', 'Trash', 'Important', 'Starred', 'Chat']
        tags = [t for t in raw if t not in exclude and not t.startswith('Category_')]
    return tags or ['Unfiled']

def parse_email_date(date_str):
    """Safely parse email date string."""
    if not date_str:
        return datetime.now()
    try:
        from email.utils import parsedate_to_datetime
        dt = parsedate_to_datetime(date_str)
        if dt:
            return dt
    except:
        pass
    # Try common formats
    for fmt in ['%a, %d %b %Y %H:%M:%S %z', '%d %b %Y %H:%M:%S %z',
                '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S']:
        try:
            return datetime.strptime(date_str, fmt)
        except:
            continue
    return datetime.now()

def extract_email_body(message):
    """Extract plain text body from email."""
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
    """Create PDF for email body (8pt Times New Roman)."""
    doc = SimpleDocTemplate(output_path, pagesize=letter,
                            leftMargin=1*inch, rightMargin=1*inch,
                            topMargin=1*inch, bottomMargin=1*inch)
    style = ParagraphStyle('Normal', fontName='Times-Roman', fontSize=8, leading=10)
    story = []
    story.append(Paragraph(f"From: {metadata['from']}", style))
    story.append(Paragraph(f"To: {metadata['to']}", style))
    story.append(Paragraph(f"Date: {metadata['date']}", style))
    story.append(Paragraph(f"Subject: {metadata['subject']}", style))
    story.append(Paragraph(f"Tags: {', '.join(metadata['tags'])}", style))
    story.append(Spacer(1, 0.2*inch))
    for line in email_body.split('\n'):
        if line.strip():
            safe = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(safe, style))
    doc.build(story)

def process_month(source_month_path, dest_month_path):
    """Process one month: combine all emails into PDFs + consolidated JSON."""
    month_name = os.path.basename(source_month_path)
    print(f"\n📁 Processing month: {month_name}")
    os.makedirs(dest_month_path, exist_ok=True)
    
    # Find email subfolders
    email_folders = [os.path.join(source_month_path, d) for d in os.listdir(source_month_path)
                     if os.path.isdir(os.path.join(source_month_path, d)) and not d.startswith('_')]
    if not email_folders:
        print("  No email folders found.")
        return
    print(f"  Found {len(email_folders)} email folders")
    
    consolidated = []
    
    for folder_path in email_folders:
        folder_name = os.path.basename(folder_path)
        print(f"\n    Processing: {folder_name}")
        
        with tempfile.TemporaryDirectory() as tmp:
            # Find JSON metadata
            json_files = [f for f in os.listdir(folder_path) if f.endswith('.json')]
            if not json_files:
                print(f"      ⚠️ No JSON found, skipping")
                continue
            json_path = os.path.join(folder_path, json_files[0])
            with open(json_path, 'r', encoding='utf-8') as f:
                meta = json.load(f)
            
            # Find email PDF
            pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
            email_pdf = None
            for pf in pdf_files:
                if pf.startswith(folder_name[:8]):
                    email_pdf = pf
                    break
            if not email_pdf and pdf_files:
                email_pdf = pdf_files[0]
            if not email_pdf:
                print(f"      ⚠️ No PDF found")
                continue
            email_pdf_path = os.path.join(folder_path, email_pdf)
            
            # List attachments (everything except JSON and email PDF)
            attachments = [f for f in os.listdir(folder_path)
                           if not f.endswith('.json') and f != email_pdf]
            print(f"      Found {len(attachments)} attachment(s)")
            
            # Create new PDF for this email
            new_pdf_name = f"{folder_name}_complete.pdf"
            new_pdf_path = os.path.join(dest_month_path, new_pdf_name)
            writer = PdfWriter()
            
            # Add email PDF
            try:
                reader = PdfReader(email_pdf_path)
                for page in reader.pages:
                    writer.add_page(page)
                print(f"      ✅ Added email ({len(reader.pages)} pages)")
            except Exception as e:
                error_msg = f"Could not read email PDF {email_pdf_path}: {e}"
                print(f"      ❌ {error_msg}")
                raise ConversionError(error_msg) from e
            
            # Add attachments - ONLY if there are attachments
            if attachments:
                # Separator page
                sep = io.BytesIO()
                sep_doc = SimpleDocTemplate(sep, pagesize=letter)
                sep_doc.build([Paragraph("ATTACHMENTS", ParagraphStyle('H', fontName='Times-Bold', fontSize=10))])
                sep.seek(0)
                writer.add_page(PdfReader(sep).pages[0])
                
                # Process attachments
                converted = 0
                for att_file in attachments:
                    att_path = os.path.join(folder_path, att_file)
                    print(f"        Processing: {att_file}")
                    
                    # Size check
                    size_mb = os.path.getsize(att_path) / (1024*1024)
                    if size_mb > MAX_EMBED_SIZE_MB:
                        print(f"          ⚠️ Large file ({size_mb:.1f}MB) - embedding only")
                        with open(att_path, 'rb') as f:
                            writer.add_attachment(filename=att_file, data=f.read())
                        continue
                    
                    try:
                        # Try conversion
                        pages = convert_attachment_to_pdf(att_path, att_file, tmp)
                        if pages:
                            for p in pages:
                                try:
                                    if p == att_path:  # original PDF
                                        r = PdfReader(p)
                                        for page in r.pages:
                                            writer.add_page(page)
                                        converted += len(r.pages)
                                    else:
                                        r = PdfReader(p)
                                        for page in r.pages:
                                            writer.add_page(page)
                                        converted += 1
                                except Exception as e:
                                    error_msg = f"Error adding page from {p}: {e}"
                                    print(f"          ❌ {error_msg}")
                                    raise ConversionError(error_msg) from e
                        else:
                            # This should not happen - convert_attachment_to_pdf should raise exception
                            error_msg = f"Conversion returned None without raising exception for {att_file}"
                            print(f"          ❌ {error_msg}")
                            raise ConversionError(error_msg)
                            
                    except ConversionError:
                        # Re-raise to stop processing
                        raise
                    except Exception as e:
                        error_msg = f"Unexpected error processing {att_file}: {e}"
                        print(f"          ❌ {error_msg}")
                        raise ConversionError(error_msg) from e
                    
                    # Always embed original (even if conversion succeeded)
                    with open(att_path, 'rb') as f:
                        writer.add_attachment(filename=att_file, data=f.read())
                
                if converted:
                    print(f"      ✅ Added {converted} attachment pages")
            
            # Save final PDF
            with open(new_pdf_path, 'wb') as f:
                writer.write(f)
            final_size = os.path.getsize(new_pdf_path) / (1024*1024)
            print(f"      ✅ Created: {new_pdf_name} ({final_size:.1f} MB)")
            
            # Add to consolidated metadata
            meta['pdf_file'] = new_pdf_name
            meta['original_folder'] = folder_name
            meta['attachments'] = attachments
            consolidated.append(meta)
    
    # Save consolidated JSON
    if consolidated:
        consolidated.sort(key=lambda x: x.get('date', ''))
        json_out = os.path.join(dest_month_path, f"{month_name}_consolidated.json")
        with open(json_out, 'w', encoding='utf-8') as f:
            json.dump({
                'month': month_name,
                'generated': datetime.now().isoformat(),
                'total_emails': len(consolidated),
                'emails': consolidated
            }, f, indent=2, ensure_ascii=False)
        print(f"\n  ✅ Consolidated JSON saved")

def process_all_months(source_dir, dest_dir):
    """Walk through all month folders."""
    print("="*70)
    print("📧 EMAIL ARCHIVE CONSOLIDATION - COMPLETE EDITION")
    print("="*70)
    print(f"Source: {source_dir}")
    print(f"Destination: {dest_dir}")
    print("-"*70)
    
    # Check dependencies
    print("\n🔍 Dependency check:")
    soffice = get_soffice_path()
    if soffice:
        print(f"  ✅ LibreOffice: {soffice}")
    else:
        print("  ⚠️ LibreOffice not found - Office docs will be embedded only")
    
    print(f"  ✅ PyPDF2: {'yes' if PDF_AVAILABLE else 'no'}")
    print(f"  ✅ reportlab: {'yes' if REPORTLAB_AVAILABLE else 'no'}")
    print(f"  ✅ Pillow: {'yes' if PIL_AVAILABLE else 'no'}")
    print(f"  ✅ img2pdf: {'yes' if IMG2PDF_AVAILABLE else 'no'}")
    print(f"  ✅ pandas: {'yes' if PANDAS_AVAILABLE else 'no'}")
    print(f"  ✅ xlrd: {'yes' if XLRD_AVAILABLE else 'no'}")
    print(f"  ✅ openpyxl: {'yes' if OPENPYXL_AVAILABLE else 'no'}")
    print(f"  ✅ textract: {'yes' if TEXTRACT_AVAILABLE else 'no'}")
    print("-"*70)
    
    os.makedirs(dest_dir, exist_ok=True)
    
    # Find month folders (YYYY-Mon)
    month_folders = []
    for item in os.listdir(source_dir):
        path = os.path.join(source_dir, item)
        if os.path.isdir(path) and re.match(r'\d{4}-[A-Z][a-z]{2}', item):
            month_folders.append(path)
    month_folders.sort()
    print(f"\nFound {len(month_folders)} month folders")
    
    try:
        for mpath in month_folders:
            mname = os.path.basename(mpath)
            dest_m = os.path.join(dest_dir, mname)
            process_month(mpath, dest_m)
    except ConversionError as e:
        print("\n" + "="*70)
        print(f"❌ CONVERSION FAILED - STOPPING")
        print("="*70)
        print(f"Error: {e}")
        print("\nPlease fix the issue and run the program again.")
        print(f"Failed while processing: {mpath if 'mpath' in locals() else 'unknown folder'}")
        sys.exit(1)
    
    print("\n" + "="*70)
    print("✅ CONSOLIDATION COMPLETE")
    print("="*70)

# ========== MAIN ==========
if __name__ == "__main__":
    source = "gmail_archive"
    dest = "gmail_consolidated"
    if len(sys.argv) > 1:
        source = sys.argv[1]
    if len(sys.argv) > 2:
        dest = sys.argv[2]
    if not os.path.exists(source):
        print(f"❌ Source directory not found: {source}")
    else:
        process_all_months(source, dest)