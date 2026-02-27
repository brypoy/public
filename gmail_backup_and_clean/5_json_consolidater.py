import json
import shutil
from pathlib import Path
import sys
from datetime import datetime

# ===== CONFIG =====
SOURCE_DIR = "gmail_consolidated"  # Original directory with PDFs and JSONs
DEST_DIR = "gmail_pdf_fixed"  # Compressed directory (where JSONs should go)
BASE_DIR = "C:/Users/ajipoynter/Desktop/BP/projects/study/gmail_mbox/"  # Parent directory

MASTER_JSON_NAME = "master_metadata.json"  # Name of the master JSON file
# =================

def ensure_directory_exists(path):
    """Create directory if it doesn't exist"""
    Path(path).mkdir(parents=True, exist_ok=True)

def find_json_files(directory):
    """Find all JSON files in directory structure"""
    return list(Path(directory).rglob("*.json"))

def load_json_file(json_path):
    """Load and parse a JSON file"""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"  ⚠️  Error loading {json_path}: {e}")
        return None

def get_relative_path(file_path, base_path):
    """Get path relative to base directory"""
    return file_path.relative_to(base_path)

def create_master_json(all_metadata, output_path):
    """Create a master JSON file with all metadata"""
    
    # Organize metadata by relative path
    master_data = {
        "generated": datetime.now().isoformat(),
        "source_directory": SOURCE_DIR,
        "destination_directory": DEST_DIR,
        "total_json_files": len(all_metadata),
        "files": {}
    }
    
    # Add each file's metadata
    for rel_path, metadata in all_metadata.items():
        master_data["files"][str(rel_path)] = metadata
    
    # Save master JSON
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(master_data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"  ❌ Error creating master JSON: {e}")
        return False

def main():
    # Setup paths
    src = Path(BASE_DIR) / SOURCE_DIR
    dst = Path(BASE_DIR) / DEST_DIR
    
    print("="*70)
    print("JSON METADATA COPY AND MASTER CREATION")
    print("="*70)
    print(f"Source directory: {src}")
    print(f"Destination directory: {dst}")
    print("-"*70)
    
    # Check if source exists
    if not src.exists():
        print(f"❌ Source directory '{src}' not found!")
        return
    
    # Check if destination exists
    if not dst.exists():
        print(f"❌ Destination directory '{dst}' not found!")
        print(f"Please run PDF compression first to create the destination directory.")
        return
    
    # Find all JSON files in source
    json_files = find_json_files(src)
    print(f"\nFound {len(json_files)} JSON files in source directory")
    
    if not json_files:
        print("No JSON files found!")
        return
    
    # Statistics
    copied_count = 0
    skipped_count = 0
    error_count = 0
    all_metadata = {}
    
    # Process each JSON file
    for i, json_path in enumerate(json_files, 1):
        # Get relative path
        rel_path = json_path.relative_to(src)
        
        # Construct destination path
        dest_path = dst / rel_path
        
        print(f"\n[{i}/{len(json_files)}] Processing: {rel_path}")
        
        # Load metadata for master JSON
        metadata = load_json_file(json_path)
        if metadata:
            all_metadata[str(rel_path)] = metadata
        
        # Check if destination already exists
        if dest_path.exists():
            print(f"  ⏭️  Already exists in destination, skipping...")
            skipped_count += 1
            continue
        
        # Create subdirectory in destination if needed
        ensure_directory_exists(dest_path.parent)
        
        try:
            # Copy JSON file
            shutil.copy2(json_path, dest_path)
            copied_count += 1
            print(f"  ✅ Copied to: {dest_path.relative_to(Path(BASE_DIR))}")
        except Exception as e:
            error_count += 1
            print(f"  ❌ Error copying: {e}")
    
    # Create master JSON
    print("\n" + "-"*70)
    print("Creating master metadata file...")
    
    master_json_path = dst / MASTER_JSON_NAME
    
    if create_master_json(all_metadata, master_json_path):
        print(f"✅ Master JSON created: {master_json_path.relative_to(Path(BASE_DIR))}")
        print(f"   Contains metadata for {len(all_metadata)} files")
    else:
        print("❌ Failed to create master JSON")
    
    # Summary
    print("\n" + "="*70)
    print("SUMMARY")
    print("="*70)
    print(f"Total JSON files found: {len(json_files)}")
    print(f"✅ Successfully copied: {copied_count}")
    print(f"⏭️  Skipped (already exist): {skipped_count}")
    print(f"❌ Errors: {error_count}")
    print(f"\n📁 Master JSON location: {master_json_path.relative_to(Path(BASE_DIR))}")
    print("="*70)

if __name__ == "__main__":
    main()