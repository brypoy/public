import os, time, pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import json
from datetime import datetime
from glob import glob

def layer_0():
    os.makedirs("output/tables", exist_ok=True)

    driver = webdriver.Chrome()
    driver.get("https://allowances.state.gov/Web920/lqa_all.asp")

    # Get all date values
    options = driver.find_elements(By.XPATH, "//select[@name='EffectiveDate']/option")
    date_values = [opt.get_attribute('value') for opt in options]

    for date_val in date_values:
        # Check if CSV already exists
        csv_path = f"output/tables/{date_val}.csv"
        if os.path.exists(csv_path):
            print(f"Skipping {date_val} - CSV already exists")
            continue
        
        # Only navigate and scrape if file doesn't exist
        driver.get("https://allowances.state.gov/Web920/lqa_all.asp")
        Select(driver.find_element(By.NAME, 'EffectiveDate')).select_by_value(date_val)
        driver.find_element(By.XPATH, "//input[@type='submit']").click()
        time.sleep(1)
        
        try:
            dfs = pd.read_html(driver.page_source)
            if dfs:
                df = dfs[0]
                # Flatten multi-index columns if needed
                if isinstance(df.columns, pd.MultiIndex):
                    df.columns = [f"{col[0]}_{col[1]}" if pd.notna(col[1]) else col[0] 
                                for col in df.columns.values]
                df.to_csv(csv_path, index=False)
                print(f"Saved: {date_val} ({len(df)} rows)")
        except Exception as e:
            print(f"Failed {date_val}: {e}")

    driver.quit()


def create_germany_csv():
    """Create a CSV with only Germany data across all dates"""
    
    os.makedirs("output/layer_2", exist_ok=True)
    
    # Get all CSV files
    csv_files = glob("output/tables/*.csv")
    
    if not csv_files:
        print("No CSV files found to process Germany data")
        return
    
    # List to hold all Germany data rows
    germany_data = []
    
    for csv_file in csv_files:
        try:
            # Extract date from filename
            date_val = os.path.basename(csv_file).replace('.csv', '')
            
            # Read the CSV file
            df = pd.read_csv(csv_file)
            
            # Find Germany rows (case-insensitive)
            # Germany might appear as "GERMANY", "Germany", etc.
            germany_mask = df.iloc[:, 0].astype(str).str.upper() == "GERMANY"
            germany_rows = df[germany_mask]
            
            if not germany_rows.empty:
                # Process each Germany row (could be multiple posts in Germany)
                for _, row in germany_rows.iterrows():
                    # Create a row for the Germany CSV
                    germany_row = {"Date": date_val}
                    
                    # Add all column values
                    for col in df.columns:
                        if col != df.columns[0]:  # Skip country name column
                            value = row[col]
                            # Clean numeric values
                            if pd.notna(value):
                                try:
                                    clean_val = str(value).replace(',', '').replace('&nbsp;', '').strip()
                                    if clean_val.replace('.', '', 1).isdigit():
                                        germany_row[col] = float(clean_val) if '.' in clean_val else int(clean_val)
                                    else:
                                        germany_row[col] = clean_val
                                except:
                                    germany_row[col] = str(value).strip()
                            else:
                                germany_row[col] = None
                    
                    germany_data.append(germany_row)
                
                print(f"Found {len(germany_rows)} Germany rows in {date_val}")
            else:
                print(f"No Germany data found in {date_val}")
                
        except Exception as e:
            print(f"Error processing Germany data from {csv_file}: {e}")
    
    if germany_data:
        # Create DataFrame
        germany_df = pd.DataFrame(germany_data)
        
        # Sort by date (newest first)
        if 'Date' in germany_df.columns:
            germany_df = germany_df.sort_values('Date', ascending=False)
        
        # Save to CSV
        output_file = "output/layer_2/germany_data.csv"
        germany_df.to_csv(output_file, index=False)
        
        print(f"\nSaved Germany data to: {output_file}")
        print(f"Total Germany rows: {len(germany_df)}")
        print(f"Date range: {germany_df['Date'].min()} to {germany_df['Date'].max()}")
        
        # Show column structure
        print(f"\nCSV structure:")
        print(f"  Date | {' | '.join([col for col in germany_df.columns if col != 'Date'])}")
        
        # Show sample data
        print(f"\nSample rows:")
        for i in range(min(3, len(germany_df))):
            row = germany_df.iloc[i]
            print(f"  {row['Date']}: {len([col for col in germany_df.columns if col != 'Date' and pd.notna(row[col])])} data points")
    else:
        print("No Germany data found in any CSV files")

def layer_1():
    """Alternative: Include post names in the structure"""
    
    os.makedirs("output/layer_2", exist_ok=True)
    today = datetime.now().strftime("%Y%m%d")
    json_filename = f"output/layer_2/state_dep_tables_with_posts_to_{today}.json"
    
    data_structure = {}
    csv_files = glob("output/tables/*.csv")
    
    for csv_file in csv_files:
        try:
            date_key = os.path.basename(csv_file).replace('.csv', '')
            df = pd.read_csv(csv_file)
            data_structure[date_key] = {}
            
            # Find the Post Name column (usually second column)
            post_column = None
            for col in df.columns:
                if 'post' in col.lower() or 'Post' in col:
                    post_column = col
                    break
            
            if post_column is None and len(df.columns) > 1:
                post_column = df.columns[1]  # Assume second column is post name
            
            for _, row in df.iterrows():
                country = row.iloc[0] if len(row) > 0 else "Unknown"
                
                if pd.isna(country) or str(country).strip() == "":
                    continue
                
                # Get post name if available
                post = row[post_column] if post_column else "Unknown"
                
                # Create unique key: Country + Post
                country_post_key = f"{country} - {post}" if post != "Unknown" else country
                
                data_structure[date_key][country_post_key] = {}
                
                # Add all columns except country and post
                columns_to_include = [col for col in df.columns if col not in [df.columns[0], post_column]]
                
                for column in columns_to_include:
                    value = row[column]
                    if pd.isna(value):
                        data_structure[date_key][country_post_key][column] = None
                    else:
                        # Clean numeric values
                        try:
                            clean_val = str(value).replace(',', '').strip()
                            if clean_val.replace('.', '', 1).replace('-', '', 1).isdigit():
                                data_structure[date_key][country_post_key][column] = float(clean_val) if '.' in clean_val else int(clean_val)
                            else:
                                data_structure[date_key][country_post_key][column] = clean_val
                        except:
                            data_structure[date_key][country_post_key][column] = str(value).strip()
            
            print(f"Processed with posts: {csv_file}")
            
        except Exception as e:
            print(f"Error processing {csv_file}: {e}")
    
    # Save JSON
    with open(json_filename, 'w', encoding='utf-8') as f:
        json.dump(data_structure, f, indent=2, ensure_ascii=False)
    
    print(f"\nAlternative JSON saved: {json_filename}")

def create_consolidated_csv():
    """Bonus: Create a consolidated CSV with all dates"""
    
    all_data = []
    csv_files = glob("output/tables/*.csv")
    
    for csv_file in csv_files:
        try:
            date_val = os.path.basename(csv_file).replace('.csv', '')
            df = pd.read_csv(csv_file)
            df['Date'] = date_val  # Add date column
            all_data.append(df)
        except Exception as e:
            print(f"Error reading {csv_file}: {e}")
    
    if all_data:
        consolidated_df = pd.concat(all_data, ignore_index=True)
        consolidated_df.to_csv("output/layer_2/consolidated_all_dates.csv", index=False)
        print(f"Consolidated CSV saved: output/layer_2/consolidated_all_dates.csv")
        print(f"Total rows: {len(consolidated_df)}")

if __name__ == "__main__":
    layer_0()
    # Create the main JSON structure
    layer_1()
    # Create Germany-specific CSV
    create_germany_csv()
    # Create consolidated CSV (optional)
    create_consolidated_csv()