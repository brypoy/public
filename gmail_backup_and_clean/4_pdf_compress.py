import os
import subprocess
import shutil
import threading
import time
from pathlib import Path
import sys

# ===== CONFIG =====
SOURCE_DIR = "gmail_pdf_fixed"
OUTPUT_DIR = "gmail_pdf_compressed"
BASE_DIR = "C:/Users/ajipoynter/Desktop/BP/projects/study/gmail_mbox/"

COMPRESSION_LEVEL = "ebook"
ALWAYS_KEEP_SMALLEST = True
TIMEOUT_SECONDS = 60  # Maximum seconds to wait for Ghostscript
# =================

def ensure_directory_exists(path):
    Path(path).mkdir(parents=True, exist_ok=True)

def get_file_size_mb(file_path):
    return file_path.stat().st_size / (1024 * 1024)

def format_size(size_mb):
    if size_mb < 1:
        return f"{size_mb * 1024:.2f} KB"
    return f"{size_mb:.2f} MB"

def find_ghostscript():
    possible_paths = [
        r"C:\Program Files\gs\gs10.04.0\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs10.03.0\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs10.02.0\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs10.01.0\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs10.00.0\bin\gswin64c.exe",
        r"C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe",
        r"C:\Program Files (x86)\gs\gs10.04.0\bin\gswin32c.exe",
    ]
    
    for path in possible_paths:
        if Path(path).exists():
            return path
    
    gs_path = shutil.which("gswin64c") or shutil.which("gswin32c")
    if gs_path:
        return gs_path
    
    return None

def run_with_timeout(cmd, timeout_seconds):
    """Run a command with timeout, returns (success, stdout, stderr)"""
    
    process = None
    try:
        # Start the process
        process = subprocess.Popen(
            cmd, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE,
            text=True
        )
        
        # Wait for process to complete or timeout
        stdout, stderr = process.communicate(timeout=timeout_seconds)
        
        return process.returncode == 0, stdout, stderr
        
    except subprocess.TimeoutExpired:
        print(f"  ⏰ Ghostscript timed out after {timeout_seconds} seconds")
        if process:
            try:
                process.terminate()
                time.sleep(1)
                process.kill()  # Force kill if terminate doesn't work
            except:
                pass
        return False, "", "Timeout expired"
        
    except Exception as e:
        print(f"  Error running Ghostscript: {e}")
        return False, "", str(e)

def compress_with_ghostscript(input_path, output_path, level="ebook", timeout=60):
    """
    Compress PDF using Ghostscript with timeout
    """
    try:
        gs_exe = find_ghostscript()
        if not gs_exe:
            print("  ❌ Ghostscript not found")
            return False
        
        settings_map = {
            "screen": "screen",
            "ebook": "ebook",
            "printer": "printer",
            "prepress": "prepress"
        }
        
        gs_settings = settings_map.get(level, "ebook")
        
        cmd = [
            gs_exe,
            "-sDEVICE=pdfwrite",
            f"-dPDFSETTINGS=/{gs_settings}",
            "-dCompatibilityLevel=1.4",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            "-dDetectDuplicateImages",
            "-dCompressFonts=true",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dAutoRotatePages=/None",
            "-sOutputFile=" + str(output_path),
            str(input_path)
        ]
        
        # Run with timeout
        success, stdout, stderr = run_with_timeout(cmd, timeout)
        
        if not success:
            if "Timeout" not in stderr:  # Don't show timeout as error
                print(f"  Ghostscript error: {stderr}")
            return False
        
        return True
        
    except Exception as e:
        print(f"  Error: {e}")
        return False

def main():
    src = Path(BASE_DIR) / SOURCE_DIR
    dst = Path(BASE_DIR) / OUTPUT_DIR
    
    print("="*70)
    print("PDF COMPRESSION WITH GHOSTSCRIPT")
    print("="*70)
    print(f"Parent directory: {BASE_DIR}")
    print(f"Source: {src}")
    print(f"Destination: {dst}")
    print(f"Compression level: {COMPRESSION_LEVEL}")
    print(f"Always keep smallest: {ALWAYS_KEEP_SMALLEST}")
    print(f"Timeout: {TIMEOUT_SECONDS} seconds")
    print("-"*70)
    
    if not src.exists():
        print(f"❌ Source directory '{src}' not found!")
        return
    
    gs_exe = find_ghostscript()
    if not gs_exe:
        print("\n❌ Ghostscript not found!")
        print("\nPlease install Ghostscript:")
        print("1. Download from: https://ghostscript.com/releases/gsdnld.html")
        print("2. Run the installer")
        print("3. Restart this script")
        return
    
    print(f"✅ Found Ghostscript: {gs_exe}")
    
    ensure_directory_exists(dst)
    print(f"✅ Destination directory ready: {dst}")
    
    all_pdfs = list(src.rglob("*.pdf")) + list(src.rglob("*.PDF"))
    print(f"\nFound {len(all_pdfs)} PDF files to process")
    
    if not all_pdfs:
        print("No PDF files found!")
        return
    
    total_original_size = 0
    total_final_size = 0
    successful = 0
    failed = 0
    kept_original = 0
    compressed_smaller = 0
    timed_out = 0
    
    for i, pdf_path in enumerate(all_pdfs, 1):
        rel_path = pdf_path.relative_to(src)
        output_path = dst / rel_path
        
        ensure_directory_exists(output_path.parent)
        
        original_size = get_file_size_mb(pdf_path)
        total_original_size += original_size
        
        print(f"\n[{i}/{len(all_pdfs)}] Processing: {rel_path}")
        print(f"  Original: {format_size(original_size)}")
        
        if output_path.exists():
            final_size = get_file_size_mb(output_path)
            print(f"  ⏭️  Output already exists: {format_size(final_size)}")
            total_final_size += final_size
            successful += 1
            
            if final_size < original_size:
                compressed_smaller += 1
            else:
                kept_original += 1
            continue
        
        temp_path = output_path.with_suffix('.temp.pdf')
        
        print(f"  ⏳ Compressing (timeout: {TIMEOUT_SECONDS}s)...")
        start_time = time.time()
        success = compress_with_ghostscript(pdf_path, temp_path, COMPRESSION_LEVEL, TIMEOUT_SECONDS)
        elapsed = time.time() - start_time
        
        if success and temp_path.exists():
            compressed_size = get_file_size_mb(temp_path)
            print(f"  ⏱️  Compression took {elapsed:.1f} seconds")
            
            if compressed_size < original_size or not ALWAYS_KEEP_SMALLEST:
                shutil.move(temp_path, output_path)
                final_size = compressed_size
                total_final_size += final_size
                successful += 1
                
                if compressed_size < original_size:
                    compressed_smaller += 1
                    print(f"  ✅ Compressed: {format_size(compressed_size)} (saved {((original_size - compressed_size)/original_size*100):.1f}%)")
                else:
                    kept_original += 1
                    print(f"  ⚠️  Compressed: {format_size(compressed_size)} (larger than original, but saved anyway)")
            else:
                # Compressed version is larger, keep original
                shutil.copy2(pdf_path, output_path)
                final_size = original_size
                total_final_size += final_size
                successful += 1
                kept_original += 1
                
                print(f"  ℹ️  Compressed was larger ({format_size(compressed_size)}), kept original")
                if temp_path.exists():
                    temp_path.unlink()
            
            print(f"  📁 Saved to: {output_path.relative_to(Path(BASE_DIR))}")
            
        else:
            # Compression failed or timed out, copy original
            if elapsed >= TIMEOUT_SECONDS:
                timed_out += 1
                print(f"  ⏰ Timed out after {TIMEOUT_SECONDS}s, copying original...")
            else:
                failed += 1
                print(f"  ❌ Compression failed, copying original...")
            
            shutil.copy2(pdf_path, output_path)
            final_size = original_size
            total_final_size += final_size
            print(f"  📁 Copied original to: {output_path.relative_to(Path(BASE_DIR))}")
            
            if temp_path.exists():
                temp_path.unlink()
    
    # Summary
    print("\n" + "="*70)
    print("COMPRESSION SUMMARY")
    print("="*70)
    
    print(f"\nFiles processed: {len(all_pdfs)}")
    print(f"✅ Successful: {successful}")
    print(f"  - Files where compressed was smaller: {compressed_smaller}")
    print(f"  - Files where original was kept (compressed larger): {kept_original}")
    if timed_out > 0:
        print(f"⏰  Files that timed out (original copied): {timed_out}")
    if failed > 0:
        print(f"❌ Failed (original copied): {failed}")
    
    if successful > 0 and total_original_size > 0:
        total_savings = ((total_original_size - total_final_size) / total_original_size) * 100
        
        print(f"\n📊 Size Statistics:")
        print(f"  Original total: {format_size(total_original_size)}")
        print(f"  Final total: {format_size(total_final_size)}")
        print(f"  Space saved: {format_size(total_original_size - total_final_size)} ({total_savings:.1f}%)")
    
    print(f"\n📁 Files saved to: {dst}")
    print("="*70)

if __name__ == "__main__":
    main()