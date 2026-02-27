import os
import time
from pathlib import Path
import sys
import subprocess
from PyPDF2 import PdfReader, PdfWriter

# ===== CONFIG =====
SOURCE_DIR = "gmail_consolidated"
DEST_DIR = "gmail_pdf_fixed"
BASE_DIR = "C:/Users/ajipoynter/Desktop/BP/projects/study/gmail_mbox/"
# =================

def install_package(package):
    """Install required packages if missing"""
    try:
        __import__(package)
        return True
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package])
        return False

def ensure_directory_exists(path):
    """Create directory if it doesn't exist"""
    Path(path).mkdir(parents=True, exist_ok=True)

def fix_pdf_with_pypdf2(input_path, output_path):
    """
    Fix PDF by re-saving it with PyPDF2.
    This preserves text and formatting.
    """
    try:
        # Read the original PDF
        reader = PdfReader(str(input_path))
        writer = PdfWriter()
        
        # Copy all pages
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            writer.add_page(page)
        
        # Save to new file
        with open(output_path, 'wb') as output_file:
            writer.write(output_file)
        
        return True
        
    except Exception as e:
        print(f"  PyPDF2 error: {e}")
        return False

def fix_pdf_with_pikepdf(input_path, output_path):
    """
    Alternative: Use pikepdf for better PDF preservation.
    Install with: pip install pikepdf
    """
    try:
        import pikepdf
        
        # Open the PDF
        pdf = pikepdf.open(input_path)
        
        # Save it with compression
        pdf.save(output_path, compress_streams=True)
        pdf.close()
        
        return True
        
    except Exception as e:
        print(f"  pikepdf error: {e}")
        return False

def fix_pdf_with_pymupdf(input_path, output_path):
    """
    Alternative: Use PyMuPDF (fitz) which often works best.
    Install with: pip install PyMuPDF
    """
    try:
        import fitz  # PyMuPDF
        
        # Open the PDF
        doc = fitz.open(input_path)
        
        # Save with optimization options that preserve text
        doc.save(output_path, 
                garbage=4,  # Remove unused objects
                deflate=True,  # Compress streams
                clean=True,  # Clean and sanitize
                linear=True)  # Linearize for web viewing
        doc.close()
        
        return True
        
    except Exception as e:
        print(f"  PyMuPDF error: {e}")
        return False

def verify_pdf_quality(pdf_path):
    """
    Quick check if PDF has selectable text
    """
    try:
        import fitz
        doc = fitz.open(pdf_path)
        
        # Check first page for text
        if len(doc) > 0:
            page = doc[0]
            text = page.get_text()
            has_text = len(text.strip()) > 0
            doc.close()
            return has_text
        doc.close()
        return False
    except:
        return False

def main():
    # Install required packages
    install_package("pypdf2")
    
    # Try to install optional better libraries
    try:
        install_package("pikepdf")
    except:
        print("Note: pikepdf not available, will use PyPDF2")
    
    try:
        install_package("pymupdf")
    except:
        print("Note: PyMuPDF not available, will use PyPDF2")
    
    # Setup paths
    src = Path(BASE_DIR) / SOURCE_DIR
    dst = Path(BASE_DIR) / DEST_DIR
    
    print(f"Source directory: {src}")
    print(f"Destination directory: {dst}")
    
    if not src.exists():
        print(f"Source directory '{src}' not found!")
        return
    
    # Create root destination directory
    ensure_directory_exists(dst)
    
    # Get all PDFs recursively
    all_pdfs = list(src.rglob("*.pdf")) + list(src.rglob("*.PDF"))
    print(f"Found {len(all_pdfs)} total PDF files")
    
    # Filter out files that have already been processed
    pdfs_to_process = []
    already_done = []
    
    for pdf in all_pdfs:
        rel_path = pdf.relative_to(src)
        output_path = dst / rel_path
        
        if output_path.exists() and output_path.stat().st_size > 1024:
            already_done.append(pdf)
            print(f"  ✅ Already exists: {rel_path}")
        else:
            pdfs_to_process.append((pdf, output_path, rel_path))
    
    print(f"\nSummary:")
    print(f"  ✅ Already fixed: {len(already_done)}")
    print(f"  🔧 Need to process: {len(pdfs_to_process)}")
    
    if not pdfs_to_process:
        print("\nAll files already fixed! Nothing to do.")
        return
    
    print("\n" + "="*60)
    print("Starting processing...")
    print("="*60)
    
    # Process files
    processed_count = 0
    failed_files = []
    low_quality_files = []
    
    # Try different libraries in order of preference
    fix_methods = [
        ("PyMuPDF", fix_pdf_with_pymupdf),
        ("pikepdf", fix_pdf_with_pikepdf),
        ("PyPDF2", fix_pdf_with_pypdf2)
    ]
    
    for i, (input_path, output_path, rel_path) in enumerate(pdfs_to_process, 1):
        print(f"\n[{i}/{len(pdfs_to_process)}] Processing: {rel_path}")
        print(f"  Input: {input_path}")
        print(f"  Output: {output_path}")
        
        # Create subdirectory in destination if needed
        ensure_directory_exists(output_path.parent)
        
        # Try each fix method until one works
        success = False
        method_used = None
        
        for method_name, fix_method in fix_methods:
            try:
                print(f"  Trying {method_name}...")
                if fix_method(input_path, output_path):
                    success = True
                    method_used = method_name
                    break
            except Exception as e:
                print(f"  {method_name} failed: {e}")
                continue
        
        if success and output_path.exists() and output_path.stat().st_size > 1024:
            processed_count += 1
            print(f"  ✅ Successfully saved using {method_used}")
            
            # Check quality
            if verify_pdf_quality(output_path):
                print(f"  ✅ Text is selectable")
            else:
                print(f"  ⚠️  Warning: Text may not be selectable")
                low_quality_files.append(rel_path)
        else:
            print(f"  ❌ All methods failed for: {rel_path}")
            failed_files.append(rel_path)
    
    # Final summary
    print("\n" + "="*60)
    print("PROCESSING COMPLETE")
    print("="*60)
    print(f"Successfully processed: {processed_count}/{len(pdfs_to_process)}")
    print(f"Failed: {len(failed_files)}")
    
    if low_quality_files:
        print(f"\n⚠️  Files with potential text issues ({len(low_quality_files)}):")
        for f in low_quality_files[:10]:
            print(f"  - {f}")
        if len(low_quality_files) > 10:
            print(f"  ... and {len(low_quality_files)-10} more")
    
    if failed_files:
        print("\n❌ Failed files:")
        for f in failed_files:
            print(f"  - {f}")
    
    print(f"\nFiles saved to: {dst}")
    print("="*60)

if __name__ == "__main__":
    main()