import requests
import pandas as pd
import os
from datetime import datetime

API_KEY = '9268eec02f5f23c1f38ccdcd6d1a09ea'

def search_fred_and_display():
    """
    Searches FRED for a term, saves ALL results to CSV, and returns them.
    CSV is saved to: output/search_results/search_term_timestamp.csv
    """
    search_term = input("Enter search term (e.g., 'gdp', 'unemployment', 'population'): ").strip()
    
    if not search_term:
        print("Error: Search term cannot be empty")
        return None

    print(f"\nSearching FRED for '{search_term}'...")
    
    url = "https://api.stlouisfed.org/fred/series/search"
    params = {
        'search_text': search_term,
        'api_key': API_KEY,
        'file_type': 'json',
        'limit': 100,
        'order_by': 'popularity',
        'sort_order': 'desc'
    }
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        
        if 'seriess' not in data or not data['seriess']:
            print("No series found for that term.")
            return None
        
        results = data['seriess']
        
        # SAVE SEARCH RESULTS TO CSV
        os.makedirs('output/search_results', exist_ok=True)
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_term = ''.join(c for c in search_term if c.isalnum() or c in (' ', '-', '_')).rstrip()
        safe_term = safe_term[:50]
        filename = f"{safe_term}_{timestamp}.csv" if safe_term else f"search_{timestamp}.csv"
        filepath = os.path.join('output/search_results', filename)
        
        df = pd.DataFrame(results)
        # Ensure popularity is numeric for sorting
        if 'popularity' in df.columns:
            df['popularity'] = pd.to_numeric(df['popularity'], errors='coerce')
            df = df.sort_values('popularity', ascending=False)
        
        df.to_csv(filepath, index=False)
        print(f"✓ Search results saved to: {filepath}")
        
        # Display results summary
        print(f"\nFound {len(results)} results. Displaying top 20:")
        print("=" * 100)
        
        for i, series in enumerate(results[:20], 1):
            series_id = series['id']
            title = series.get('title', 'N/A')
            if len(title) > 60:
                title = title[:57] + "..."
            
            print(f"{i:2}. {series_id:15} | {title}")
            print(f"     Freq: {series.get('frequency', 'N/A'):10} | "
                  f"Units: {series.get('units', 'N/A')[:20]:20} | "
                  f"Pop: {series.get('popularity', 'N/A'):3}")
        
        if len(results) > 20:
            print(f"... and {len(results) - 20} more results")
        
        print("=" * 100)
        return results
        
    except requests.exceptions.RequestException as e:
        print(f"Search error: {e}")
        return None

def download_series_by_id(series_id):
    """
    Downloads a specific series by ID (full date range).
    """
    print(f"\n" + "="*60)
    print(f"DOWNLOADING SERIES: {series_id}")
    print("="*60)
    
    # Step 1: Get series details
    series_details = fetch_series_details(series_id)
    if not series_details:
        print(f"Error: Could not find series with ID '{series_id}'")
        return False
    
    print(f"Title: {series_details.get('title', 'N/A')}")
    print(f"Units: {series_details.get('units', 'N/A')}")
    print(f"Frequency: {series_details.get('frequency', 'N/A')}")
    print(f"Date Range: {series_details.get('observation_start', 'N/A')} to {series_details.get('observation_end', 'N/A')}")
    
    # Step 2: Get ALL observations (full range)
    print(f"\nDownloading ALL observation data...")
    observations = fetch_series_data(series_id)
    
    if not observations:
        print("Error: No observation data found.")
        return False
    
    obs_df = pd.DataFrame(observations)
    obs_df['value'] = pd.to_numeric(obs_df['value'], errors='coerce')
    obs_df['date'] = pd.to_datetime(obs_df['date'])
    
    # Remove rows with null values
    obs_df = obs_df.dropna(subset=['value'])
    
    print(f"✓ Downloaded {len(obs_df)} data points")
    print(f"\nDate range in data: {obs_df['date'].min().date()} to {obs_df['date'].max().date()}")
    print("\nPreview (first 5 and last 5):")
    print(obs_df[['date', 'value']].head())
    print("...")
    print(obs_df[['date', 'value']].tail())
    
    # Step 3: Save to CSV
    filepath = save_series_to_csv(series_id, series_details, obs_df)
    if filepath:
        print(f"\n" + "="*60)
        print(f"✓ SUCCESSFULLY SAVED TO: {filepath}")
        print("="*60)
        
        if len(obs_df) > 0:
            print(f"\nSUMMARY STATISTICS:")
            print(f"  Total observations: {len(obs_df)}")
            print(f"  Date range: {obs_df['date'].min().date()} to {obs_df['date'].max().date()}")
            print(f"  Mean value: {obs_df['value'].mean():.2f}")
            print(f"  Min value: {obs_df['value'].min():.2f}")
            print(f"  Max value: {obs_df['value'].max():.2f}")
            print(f"  Std deviation: {obs_df['value'].std():.2f}")
        
        return True
    return False

def fetch_series_details(series_id):
    """Fetches detailed metadata for a FRED series ID."""
    url = "https://api.stlouisfed.org/fred/series"
    params = {'api_key': API_KEY, 'file_type': 'json', 'series_id': series_id}
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        if 'seriess' in data and data['seriess']:
            return data['seriess'][0]
        return None
    except Exception as e:
        print(f"Error fetching series details: {e}")
        return None

def fetch_series_data(series_id, observation_start=None, observation_end=None):
    """Fetches the observation data for a FRED series ID."""
    url = "https://api.stlouisfed.org/fred/series/observations"
    params = {'api_key': API_KEY, 'file_type': 'json', 'series_id': series_id}
    
    # Always get full range - no date restrictions
    if observation_start:
        params['observation_start'] = observation_start
    if observation_end:
        params['observation_end'] = observation_end
    
    try:
        response = requests.get(url, params=params)
        response.raise_for_status()
        data = response.json()
        if 'observations' in data:
            return data['observations']
        return None
    except Exception as e:
        print(f"Error fetching series data: {e}")
        return None

def save_series_to_csv(series_id, series_details, observations_df):
    """Saves the series metadata and observations to a CSV file."""
    try:
        os.makedirs('output/series_data', exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        safe_id = series_id.replace('/', '_').replace('\\', '_')
        filename = f"{safe_id}_{timestamp}.csv"
        filepath = f"output/series_data/{filename}"
        
        metadata = {
            'Series ID': series_id,
            'Title': series_details.get('title', 'N/A'),
            'Frequency': series_details.get('frequency', 'N/A'),
            'Units': series_details.get('units', 'N/A'),
            'Seasonal Adjustment': series_details.get('seasonal_adjustment', 'N/A'),
            'Observation Start': series_details.get('observation_start', 'N/A'),
            'Observation End': series_details.get('observation_end', 'N/A'),
            'Last Updated': series_details.get('last_updated', 'N/A'),
            'Notes': series_details.get('notes', 'N/A'),
            'Popularity': series_details.get('popularity', 'N/A')
        }
        
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write("# Series Metadata\n")
            for key, value in metadata.items():
                f.write(f"# {key}: {value}\n")
            f.write("\n# Observations\n")
            observations_df.to_csv(f, index=False)
        
        return filepath
    except Exception as e:
        print(f"Error saving to CSV: {e}")
        return None

def main():
    """
    Main workflow: Search -> Display -> Choose exact series ID -> Download full range.
    """
    print("=" * 80)
    print("FRED DATA DOWNLOADER v2.0")
    print("=" * 80)
    print("Instructions:")
    print("1. Enter a search term (e.g., 'gdp', 'population')")
    print("2. Review all matching series IDs")
    print("3. Enter the EXACT series ID you want to download")
    print("4. Full date range will be downloaded automatically")
    print("=" * 80)
    
    while True:
        # Step 1: Search and display
        search_results = search_fred_and_display()
        if not search_results:
            print("\nNo results found. Try another search term.")
            continue
        
        # Step 2: Ask for EXACT series ID
        print("\n" + "="*80)
        print("ENTER THE EXACT SERIES ID TO DOWNLOAD")
        print("="*80)
        print("Examples from above: 'GDP', 'GDPC1', 'POP', 'CLF16OV', etc.")
        print("Enter 'NEW' to search for a different term")
        print("Enter 'EXIT' to quit")
        print("="*80)
        
        series_input = input("\nEnter exact series ID: ").strip().upper()
        
        if series_input == 'EXIT':
            print("Goodbye!")
            break
        elif series_input == 'NEW':
            continue
        
        # Step 3: Validate and download
        valid_ids = [series['id'] for series in search_results]
        if series_input in valid_ids:
            download_series_by_id(series_input)
        else:
            print(f"\n⚠️  Error: '{series_input}' not found in search results.")
            print(f"Valid IDs from search: {', '.join(valid_ids[:20])}{'...' if len(valid_ids) > 20 else ''}")
            
            # Offer to download anyway (user might know better)
            force_download = input(f"Download '{series_input}' anyway? (y/n): ").strip().lower()
            if force_download == 'y':
                download_series_by_id(series_input)
        
        # Ask if user wants to continue
        print("\n" + "="*80)
        cont = input("Download another series? (y/n): ").strip().lower()
        if cont != 'y':
            print("Goodbye!")
            break

if __name__ == "__main__":
    main()