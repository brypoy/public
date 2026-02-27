import json
import re

# ===== CONFIGURATION =====
json_file = "C:/Users/ajipoynter/Desktop/BP/projects/study/gmail_mbox/master_metadata.json"  # Your input file
search_terms = ["dcod", "ngc"]  # Change these to your keywords
# =========================

def clean_text(text):
    """Remove extra whitespace and clean up text."""
    if not text:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip()

# Load the data
try:
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    print(f"Loaded file with {data.get('total_json_files', 0)} files")
except Exception as e:
    print(f"Error: {e}")
    exit()

# Create search string for filename
search_string = "_".join(search_terms).replace(" ", "_")
output_file = f"search_return_{search_string}.txt"

matches = []
search_terms_lower = [term.lower() for term in search_terms]

# Navigate through the nested structure
for filename, file_data in data.get('files', {}).items():
    emails = file_data.get('emails', [])
    
    # Search through emails in this file
    for email in emails:
        # Combine all text fields for searching
        searchable = ' '.join([
            email.get('subject', ''),
            email.get('body', ''),
            email.get('from', ''),
            email.get('to', ''),
            ' '.join(email.get('tags', []))
        ]).lower()
        
        # Check if ALL terms are present (AND logic)
        if all(term in searchable for term in search_terms_lower):
            matches.append(email)

# Write results to file
with open(output_file, 'w', encoding='utf-8') as f:
    f.write(f"Search Results for: {', '.join(search_terms)}\n")
    f.write(f"Found: {len(matches)} matching emails\n")
    f.write("=" * 80 + "\n\n")
    
    for i, email in enumerate(matches, 1):
        f.write(f"EMAIL {i}\n")
        f.write("-" * 40 + "\n")
        f.write(f"From:    {clean_text(email.get('from', 'N/A'))}\n")
        f.write(f"To:      {clean_text(email.get('to', 'N/A'))}\n")
        f.write(f"Date:    {email.get('date', 'N/A')}\n")
        f.write(f"Subject: {clean_text(email.get('subject', 'N/A'))}\n")
        
        tags = email.get('tags', [])
        if tags:
            f.write(f"Tags:    {', '.join(tags)}\n")
        
        # Show attachment info if present
        attachments = email.get('attachments', [])
        if attachments:
            f.write(f"Attachments: {len(attachments)} file(s)\n")
        
        f.write("\n" + "=" * 40 + " BODY " + "=" * 40 + "\n")
        f.write(email.get('body', 'No body content'))
        f.write("\n" + "=" * 80 + "\n\n")

print(f"\n✓ Found {len(matches)} matching emails")
print(f"✓ Results saved to: {output_file}")