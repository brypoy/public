#!/usr/bin/env python3
"""
FBI Vault Scraper - Complete Single File Solution
Run directly: python fbi_vault_scraper.py "search term"
"""

import subprocess
import sys
import os
import json
import time
import tempfile
import shutil
import platform
import urllib.request
import zipfile
from datetime import datetime
from pathlib import Path
import importlib.util
import logging
import ctypes

# ==================== DEPENDENCY MANAGEMENT ====================

def is_admin():
    """Check if script is running with admin privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def check_and_install_dependencies():
    """Check for and install required packages automatically"""
    
    required_packages = {
        'selenium': 'selenium==4.15.0',
        'webdriver_manager': 'webdriver-manager==4.0.1',
        'requests': 'requests==2.31.0',
        'PyPDF2': 'PyPDF2==3.0.1',
        'pytesseract': 'pytesseract==0.3.10',
        'pdf2image': 'pdf2image==1.16.3',
        'PIL': 'Pillow==10.1.0',
        'urllib3': 'urllib3==1.26.18'
    }
    
    missing_packages = []
    
    print("\n" + "="*60)
    print("Checking Python dependencies...")
    print("="*60)
    
    # Get Python architecture
    python_arch = platform.architecture()[0]
    print(f"Python architecture: {python_arch}")
    
    for package_name, install_name in required_packages.items():
        spec = importlib.util.find_spec(package_name)
        if spec is None:
            missing_packages.append(install_name)
            print(f"✗ {package_name} - NOT FOUND")
        else:
            print(f"✓ {package_name} - FOUND")
    
    if missing_packages:
        print("\nInstalling missing packages...")
        try:
            # Upgrade pip first
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])
            
            # Install missing packages
            for package in missing_packages:
                print(f"Installing {package}...")
                # Use --only-binary to avoid compilation issues
                subprocess.check_call([
                    sys.executable, '-m', 'pip', 'install',
                    '--only-binary', ':all:',
                    package
                ])
            
            print("\n✓ All Python packages installed successfully!")
        except Exception as e:
            print(f"\n✗ Failed to install packages: {e}")
            print("\nTrying alternative installation method...")
            try:
                for package in missing_packages:
                    subprocess.check_call([
                        sys.executable, '-m', 'pip', 'install',
                        package
                    ])
            except Exception as e2:
                print(f"✗ Installation failed: {e2}")
                return False
    
    return True

def install_tesseract_via_pip():
    """Try to install Tesseract via pip packages"""
    print("\n" + "="*60)
    print("Attempting to install Tesseract OCR tools via pip...")
    print("="*60)
    
    try:
        # Install tesseract wrapper packages that don't need system Tesseract
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install',
            'tesseract',
            'pytesseract'
        ])
        print("✓ Tesseract Python packages installed")
        return True
    except Exception as e:
        print(f"✗ Failed to install Tesseract packages: {e}")
        return False

def install_poppler_via_pip():
    """Try to install Poppler via pip packages"""
    print("\n" + "="*60)
    print("Attempting to install Poppler tools via pip...")
    print("="*60)
    
    try:
        # Install poppler-utils via pip (some Windows wheels available)
        subprocess.check_call([
            sys.executable, '-m', 'pip', 'install',
            'poppler-utils',
            'pdf2image'
        ])
        print("✓ Poppler Python packages installed")
        return True
    except Exception as e:
        print(f"✗ Failed to install Poppler packages: {e}")
        return False

def check_chrome_and_driver():
    """Check Chrome installation and get compatible ChromeDriver"""
    print("\n" + "="*60)
    print("Checking Chrome installation...")
    print("="*60)
    
    # Common Chrome paths on Windows
    chrome_paths = [
        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        os.path.expanduser(r'~\AppData\Local\Google\Chrome\Application\chrome.exe')
    ]
    
    chrome_found = False
    chrome_version = None
    
    for path in chrome_paths:
        if os.path.exists(path):
            chrome_found = True
            # Try to get version
            try:
                import win32file
                info = win32file.GetFileVersionInfo(path, "\\")
                ms = info['FileVersionMS']
                ls = info['FileVersionLS']
                chrome_version = f"{ms >> 16}.{ms & 0xFFFF}.{ls >> 16}.{ls & 0xFFFF}"
                print(f"✓ Chrome found: {path}")
                print(f"  Version: {chrome_version}")
            except:
                print(f"✓ Chrome found: {path}")
            break
    
    if not chrome_found:
        print("✗ Chrome not found in standard locations")
        print("Please install Google Chrome from: https://www.google.com/chrome/")
        return False
    
    # Check Python architecture vs Chrome architecture
    python_arch = platform.architecture()[0]
    if python_arch == '32bit':
        print("\n⚠ Warning: You're using 32-bit Python with 64-bit Chrome")
        print("  This can cause driver issues. Consider using 64-bit Python.")
    
    return True

def download_chromedriver_manual():
    """Manual ChromeDriver download as fallback"""
    print("\n" + "="*60)
    print("Attempting manual ChromeDriver download...")
    print("="*60)
    
    try:
        # Get latest ChromeDriver version
        version_url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
        with urllib.request.urlopen(version_url) as response:
            latest_version = response.read().decode('utf-8').strip()
        
        print(f"Latest ChromeDriver version: {latest_version}")
        
        # Determine architecture
        python_arch = platform.architecture()[0]
        arch = 'win32' if python_arch == '32bit' else 'win64'
        
        # Download ChromeDriver
        download_url = f"https://chromedriver.storage.googleapis.com/{latest_version}/chromedriver_{arch}.zip"
        zip_path = os.path.join(tempfile.gettempdir(), 'chromedriver.zip')
        
        print(f"Downloading from: {download_url}")
        urllib.request.urlretrieve(download_url, zip_path)
        
        # Extract to a known location
        extract_path = os.path.join(os.path.dirname(sys.executable), 'chromedriver')
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
        
        # Add to PATH
        os.environ['PATH'] += os.pathsep + extract_path
        
        print(f"✓ ChromeDriver downloaded to: {extract_path}")
        return True
        
    except Exception as e:
        print(f"✗ Failed to download ChromeDriver: {e}")
        return False

# ==================== MAIN SCRAPER CLASS ====================

class FBIVaultScraper:
    def __init__(self, download_dir=None, pdf_storage_dir=None, keep_pdfs=True):
        """Initialize the FBI Vault scraper"""
        self.setup_logging()
        self.keep_pdfs = keep_pdfs
        
        # Set temporary download directory
        if download_dir:
            self.download_dir = download_dir
        else:
            self.download_dir = os.path.join(tempfile.gettempdir(), 'fbi_vault_temp')
        
        # Set permanent PDF storage directory
        if pdf_storage_dir:
            self.pdf_storage_dir = pdf_storage_dir
        else:
            self.pdf_storage_dir = os.path.join(os.getcwd(), 'fbi_vault_pdfs')
        
        self.driver = None
        self.results = []
        self.ocr_available = False
        self.poppler_available = False
        
        # Create directories
        for directory in [self.download_dir, self.pdf_storage_dir]:
            try:
                Path(directory).mkdir(parents=True, exist_ok=True)
                self.logger.info(f"Created directory: {directory}")
            except Exception as e:
                self.logger.error(f"Could not create directory {directory}: {e}")
        
        # Import modules
        self.import_modules()
    
    def setup_logging(self):
        """Set up logging configuration"""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('fbi_vault_scraper.log', encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def import_modules(self):
        """Import required modules"""
        global webdriver, By, WebDriverWait, EC, TimeoutException
        global NoSuchElementException, WebDriverException, Service, ChromeDriverManager
        
        from selenium import webdriver
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
        from selenium.webdriver.chrome.service import Service
        
        try:
            from webdriver_manager.chrome import ChromeDriverManager
            self.wdm_available = True
        except:
            self.wdm_available = False
        
        global requests, PyPDF2
        import requests
        import PyPDF2
        
        # Try to import OCR modules
        try:
            global pytesseract, convert_from_path, Image
            import pytesseract
            from pdf2image import convert_from_path
            from PIL import Image
            self.ocr_available = True
            self.logger.info("OCR modules loaded successfully")
        except ImportError as e:
            self.logger.warning(f"OCR modules not available: {e}")
            self.ocr_available = False
    
    def setup_driver(self):
        """Set up Chrome driver with multiple fallback methods"""
        options = webdriver.ChromeOptions()
        
        # Download settings
        prefs = {
            "download.default_directory": str(self.download_dir),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True,
            "profile.default_content_setting_values.automatic_downloads": 1
        }
        options.add_experimental_option("prefs", prefs)
        
        # Browser options
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--log-level=3')
        options.add_argument('--remote-debugging-port=9222')
        
        # Try different driver setup methods
        setup_methods = [
            self.setup_driver_with_webdriver_manager,
            self.setup_driver_direct,
            self.setup_driver_manual_path
        ]
        
        for method in setup_methods:
            try:
                self.driver = method(options)
                if self.driver:
                    self.logger.info(f"Driver setup successful using: {method.__name__}")
                    self.wait = WebDriverWait(self.driver, 10)
                    return True
            except Exception as e:
                self.logger.warning(f"Driver setup failed with {method.__name__}: {e}")
                continue
        
        raise Exception("All driver setup methods failed")
    
    def setup_driver_with_webdriver_manager(self, options):
        """Setup using webdriver-manager"""
        if not self.wdm_available:
            raise Exception("webdriver-manager not available")
        
        service = Service(ChromeDriverManager().install())
        return webdriver.Chrome(service=service, options=options)
    
    def setup_driver_direct(self, options):
        """Setup using direct ChromeDriver"""
        return webdriver.Chrome(options=options)
    
    def setup_driver_manual_path(self, options):
        """Setup using manually downloaded ChromeDriver"""
        # Check common ChromeDriver locations
        common_paths = [
            os.path.join(os.path.dirname(sys.executable), 'chromedriver', 'chromedriver.exe'),
            os.path.join(os.environ.get('USERPROFILE', ''), 'chromedriver.exe'),
            r'C:\chromedriver\chromedriver.exe'
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                service = Service(path)
                return webdriver.Chrome(service=service, options=options)
        
        raise Exception("No ChromeDriver found in common locations")
    
    def search(self, term):
        """Navigate to search page"""
        search_url = f"https://vault.fbi.gov/search?SearchableText={term}"
        self.logger.info(f"Searching for: {term}")
        
        try:
            self.driver.get(search_url)
            time.sleep(3)
        except Exception as e:
            self.logger.error(f"Failed to load search page: {e}")
            raise
    
    def get_result_links(self):
        """Get all result links from current page"""
        try:
            self.wait.until(EC.presence_of_element_located((By.XPATH, "//dt/a")))
            result_links = self.driver.find_elements(By.XPATH, "//dt/a")
            self.logger.info(f"Found {len(result_links)} results on current page")
            return result_links
        except TimeoutException:
            self.logger.warning("No results found or page timed out")
            return []
    
    def navigate_to_next_page(self):
        """Navigate to next page of results"""
        try:
            next_selectors = [
                "//a[contains(text(), 'next')]",
                "//a[contains(text(), 'Next')]",
                "//a[@title='Go to next page']",
                "//li[@class='next']/a"
            ]
            
            for selector in next_selectors:
                try:
                    next_button = self.driver.find_element(By.XPATH, selector)
                    next_button.click()
                    time.sleep(3)
                    return True
                except NoSuchElementException:
                    continue
            
            return False
        except Exception as e:
            self.logger.error(f"Error navigating to next page: {e}")
            return False
    
    def check_driver_health(self):
        """Check if driver session is still valid"""
        try:
            self.driver.current_url
            return True
        except:
            return False

    def recover_driver_session(self, search_term):
        """Recover from invalid session by recreating driver"""
        self.logger.warning("Driver session invalid - attempting to recover...")
        try:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
        except:
            pass
        
        time.sleep(5)
        
        try:
            self.setup_driver()
            self.search(search_term)
            time.sleep(3)
            self.logger.info("✓ Driver session recovered successfully")
            return True
        except Exception as e:
            self.logger.error(f"Failed to recover driver session: {e}")
            return False


    def trigger_download_with_selenium(self, pdf_url):
        """Trigger download using Selenium - SINGLE TRIGGER ONLY"""
        try:
            # Check driver health before proceeding
            if not self.check_driver_health():
                self.logger.warning("Driver unhealthy before download, recovering...")
                if not self.recover_driver_session(self.current_search_term):
                    return False
            
            # Find the download link - but only click it once
            download_links = self.driver.find_elements(By.XPATH, "//a[contains(@href, 'at_download') or contains(@href, '.pdf')]")
            
            for link in download_links:
                href = link.get_attribute('href')
                if href and ('at_download' in href or '.pdf' in href):
                    self.logger.info(f"Found download link: {href}")
                    
                    # Store current window handle
                    original_window = self.driver.current_window_handle
                    original_handles = len(self.driver.window_handles)
                    
                    # Execute the download - ONLY ONCE
                    self.driver.execute_script("window.open(arguments[0]);", href)
                    time.sleep(3)
                    
                    # Check if a new tab opened
                    if len(self.driver.window_handles) > original_handles:
                        # Find and close the new tab
                        for handle in self.driver.window_handles:
                            if handle != original_window:
                                self.driver.switch_to.window(handle)
                                time.sleep(1)
                                self.driver.close()
                                break
                        
                        # Switch back to original
                        self.driver.switch_to.window(original_window)
                    
                    # SUCCESS - return immediately, do NOT continue loop
                    return True
            
            self.logger.error("No download link found")
            return False
            
        except Exception as e:
            if "invalid session id" in str(e):
                self.logger.error(f"Session died during download: {e}")
            else:
                self.logger.error(f"Failed to trigger download: {e}")
            return False

    def recover_driver_session(self, search_term):
        """Recover from invalid session by recreating driver"""
        self.logger.warning("Driver session invalid - attempting to recover...")
        try:
            if self.driver:
                try:
                    self.driver.quit()
                except:
                    pass
        except:
            pass
        
        time.sleep(5)
        
        try:
            self.setup_driver()
            self.search(search_term)
            time.sleep(3)
            self.logger.info("✓ Driver session recovered successfully")
            return True
        except Exception as e:
            self.logger.error(f"Failed to recover driver session: {e}")
            return False

    def wait_for_download_with_monitoring(self, expected_filename=None, timeout=60):
        """Wait for download with driver health monitoring"""
        self.logger.info("Waiting for download to complete...")
        start_time = time.time()
        last_file_count = 0
        stable_count_checks = 0
        health_check_interval = 10
        last_health_check = time.time()
        seen_files = set()  # Track files we've already seen
        
        while time.time() - start_time < timeout:
            # Periodic driver health check
            if time.time() - last_health_check > health_check_interval:
                if not self.check_driver_health():
                    self.logger.warning("Driver became unhealthy during download wait")
                last_health_check = time.time()
            
            time.sleep(2)
            try:
                files = os.listdir(self.download_dir)
            except:
                continue
            
            # Filter out temporary/downloading files
            current_files = [f for f in files if not any(f.endswith(ext) for ext in 
                        ['.crdownload', '.tmp', '.part', '.download'])]
            
            # Find new files (not seen before)
            new_files = [f for f in current_files if f not in seen_files]
            
            if new_files:
                self.logger.info(f"New file detected: {new_files[0]}")
                seen_files.update(new_files)
                
                # Wait a bit to ensure download is complete
                time.sleep(3)
                
                # Check if file is still there and not growing
                file_path = os.path.join(self.download_dir, new_files[0])
                if os.path.exists(file_path):
                    size1 = os.path.getsize(file_path)
                    time.sleep(2)
                    size2 = os.path.getsize(file_path)
                    
                    if size1 == size2 and size1 > 0:
                        self.logger.info(f"Download completed: {new_files[0]}")
                        return file_path
            
            # Also check for expected filename pattern
            if expected_filename:
                base_name = os.path.splitext(expected_filename)[0]
                matching = [f for f in current_files if base_name in f or 
                        f.lower().endswith('.pdf')]
                if matching:
                    file_path = os.path.join(self.download_dir, matching[0])
                    if os.path.getsize(file_path) > 0:
                        return file_path
            
            # Check if file count is stable (download complete)
            current_file_count = len(current_files)
            if current_file_count > 0:
                if current_file_count == last_file_count:
                    stable_count_checks += 1
                else:
                    stable_count_checks = 0
                    last_file_count = current_file_count
                
                if stable_count_checks >= 3:
                    try:
                        latest_file = max([os.path.join(self.download_dir, f) for f in current_files], 
                                    key=os.path.getctime)
                        if os.path.getsize(latest_file) > 0:
                            return latest_file
                    except:
                        pass
            
            # Log periodically
            if int(time.time() - start_time) % 10 == 0 and current_files:
                self.logger.info(f"Files in temp dir: {current_files}")
        
        return None

    def download_pdf(self, pdf_url, filename):
        """Main download function with duplicate prevention"""
        # Sanitize filename
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        if len(filename) > 200:
            name, ext = os.path.splitext(filename)
            filename = name[:195] + ext
        
        pdf_path = os.path.join(self.pdf_storage_dir, filename)
        
        # Check if already exists
        if os.path.exists(pdf_path):
            self.logger.info(f"PDF already exists: {filename}")
            return pdf_path
        
        # Check if any file with same base name already exists (from previous runs)
        base_name = os.path.splitext(filename)[0]
        existing = [f for f in os.listdir(self.pdf_storage_dir) 
                    if f.startswith(base_name) and f.endswith('.pdf')]
        if existing:
            self.logger.info(f"PDF with same base name exists: {existing[0]}")
            return os.path.join(self.pdf_storage_dir, existing[0])
        
        # Try Selenium method with retry on session death
        max_download_attempts = 2
        
        for attempt in range(max_download_attempts):
            try:
                # Check driver health
                if not self.check_driver_health():
                    self.logger.warning(f"Driver unhealthy before download attempt {attempt + 1}")
                    if not self.recover_driver_session(self.current_search_term):
                        continue
                
                if self.trigger_download_with_selenium(pdf_url):
                    temp_file = self.wait_for_download_with_monitoring(
                        expected_filename=filename, 
                        timeout=60
                    )
                    
                    if temp_file:
                        # Get the actual filename from temp file
                        temp_filename = os.path.basename(temp_file)
                        
                        # Check if we already have this file (by size for now)
                        temp_size = os.path.getsize(temp_file)
                        
                        # Look for duplicate in permanent storage
                        duplicate = False
                        for existing_file in os.listdir(self.pdf_storage_dir):
                            existing_path = os.path.join(self.pdf_storage_dir, existing_file)
                            if os.path.exists(existing_path) and os.path.getsize(existing_path) == temp_size:
                                self.logger.info(f"Potential duplicate found: {existing_file}")
                                duplicate = True
                                break
                        
                        if duplicate:
                            self.logger.info("Skipping duplicate file")
                            # Clean up temp file
                            try:
                                os.remove(temp_file)
                            except:
                                pass
                            return existing_path
                        
                        # Determine final filename
                        temp_name, temp_ext = os.path.splitext(temp_filename)
                        
                        # If temp file has different extension, preserve it
                        if temp_ext.lower() != '.pdf':
                            new_filename = f"{base_name}{temp_ext}"
                            pdf_path = os.path.join(self.pdf_storage_dir, new_filename)
                        
                        if self.move_downloaded_file(temp_file, pdf_path):
                            return pdf_path
                    else:
                        self.logger.warning(f"Download attempt {attempt + 1} - no file detected")
                        
                        # Clean up any orphaned temp files
                        self.cleanup_temp_directory()
                else:
                    self.logger.warning(f"Failed to trigger download on attempt {attempt + 1}")
                    
            except Exception as e:
                self.logger.error(f"Download attempt {attempt + 1} failed: {e}")
                if "invalid session id" in str(e) and attempt < max_download_attempts - 1:
                    self.logger.info("Attempting session recovery...")
                    self.recover_driver_session(self.current_search_term)
                    time.sleep(5)
                    continue
        
        # Fallback to requests
        self.logger.info("Trying requests fallback...")
        if self.download_with_requests(pdf_url, pdf_path):
            return pdf_path
        
        self.logger.error("All download methods failed")
        return None


    def download_with_requests(self, pdf_url, pdf_path):
        """Download using requests as fallback"""
        self.logger.info("Attempting download with requests...")
        
        session = requests.Session()
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
            'Referer': 'https://vault.fbi.gov/'
        }
        
        # Get cookies first
        try:
            session.get('https://vault.fbi.gov', headers=headers, timeout=10)
        except:
            pass
        
        # Try URL variations
        url_variations = [
            pdf_url,
            pdf_url.replace('/at_download/file', ''),
            pdf_url.replace('/at_download/file', '.pdf'),
            pdf_url.replace('https://vault.fbi.gov/', 'https://vault.fbi.gov/download/')
        ]
        
        for url in url_variations:
            try:
                self.logger.info(f"Trying URL: {url}")
                response = session.get(url, headers=headers, stream=True, timeout=30, allow_redirects=True)
                
                if response.status_code == 200:
                    content_type = response.headers.get('content-type', '').lower()
                    content_length = response.headers.get('content-length', '0')
                    
                    if 'pdf' in content_type or 'octet-stream' in content_type or 'application' in content_type:
                        with open(pdf_path, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=8192):
                                f.write(chunk)
                        
                        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                            return True
            except Exception as e:
                self.logger.error(f"Request failed for {url}: {e}")
                continue
        
        return False

    def move_downloaded_file(self, temp_file, pdf_path):
        """Move downloaded file to permanent storage"""
        try:
            shutil.move(temp_file, pdf_path)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                self.logger.info(f"✓ Moved to: {pdf_path}")
                return True
        except Exception as e:
            self.logger.error(f"Failed to move file: {e}")
        return False

    
    def extract_text_from_pdf(self, pdf_path):
        """Extract text from PDF"""
        text = ""
        
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                self.logger.info(f"PDF has {len(pdf_reader.pages)} pages")
                
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    page_text = page.extract_text()
                    
                    # Use OCR if available and text is minimal
                    if self.ocr_available and len(page_text.strip()) < 100:
                        self.logger.info(f"Page {page_num + 1}: Using OCR")
                        page_text = self.ocr_pdf_page(pdf_path, page_num)
                    
                    text += page_text + "\n"
        except Exception as e:
            self.logger.error(f"Error extracting text: {e}")
        
        return text.strip()
    
    def ocr_pdf_page(self, pdf_path, page_num):
        """Perform OCR on a PDF page with Windows virtual environment support"""
        if not self.ocr_available:
            return ""
        
        try:
            from pdf2image import convert_from_path
            import pytesseract
            import subprocess
            import shutil
            import winreg
            
            def find_tesseract_windows():
                """Find Tesseract installation in Windows registry"""
                try:
                    # Check common installation paths first
                    common_paths = [
                        r'C:\Program Files\Tesseract-OCR\tesseract.exe',
                        r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
                        os.path.expanduser(r'~\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'),
                        os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Programs', 'Tesseract-OCR', 'tesseract.exe'),
                        os.path.join(os.environ.get('PROGRAMFILES', ''), 'Tesseract-OCR', 'tesseract.exe'),
                        os.path.join(os.environ.get('PROGRAMFILES(X86)', ''), 'Tesseract-OCR', 'tesseract.exe')
                    ]
                    
                    for path in common_paths:
                        if os.path.exists(path):
                            return path
                    
                    # Try Windows Registry
                    try:
                        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r'SOFTWARE\Tesseract-OCR') as key:
                            install_dir, _ = winreg.QueryValueEx(key, 'InstallDir')
                            tesseract_path = os.path.join(install_dir, 'tesseract.exe')
                            if os.path.exists(tesseract_path):
                                return tesseract_path
                    except:
                        pass
                    
                    try:
                        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'SOFTWARE\Tesseract-OCR') as key:
                            install_dir, _ = winreg.QueryValueEx(key, 'InstallDir')
                            tesseract_path = os.path.join(install_dir, 'tesseract.exe')
                            if os.path.exists(tesseract_path):
                                return tesseract_path
                    except:
                        pass
                    
                    # Check PATH
                    tesseract_in_path = shutil.which('tesseract')
                    if tesseract_in_path:
                        return tesseract_in_path
                    
                except Exception as e:
                    self.logger.debug(f"Error finding Tesseract: {e}")
                
                return None
            
            # Find Tesseract
            tesseract_path = find_tesseract_windows()
            if tesseract_path:
                pytesseract.pytesseract.tesseract_cmd = tesseract_path
                self.logger.info(f"Found Tesseract at: {tesseract_path}")
            else:
                self.logger.warning("Tesseract not found - OCR unavailable")
                return ""
            
            def find_poppler_windows():
                """Find Poppler installation in Windows"""
                common_poppler_paths = [
                    r'C:\poppler\bin',
                    r'C:\poppler-utils\bin',
                    r'C:\Program Files\poppler\bin',
                    os.path.join(os.environ.get('LOCALAPPDATA', ''), 'poppler', 'bin'),
                    os.path.join(os.environ.get('PROGRAMFILES', ''), 'poppler', 'bin')
                ]
                
                for path in common_poppler_paths:
                    if os.path.exists(path):
                        return path
                
                # Check PATH for pdfinfo
                pdfinfo_path = shutil.which('pdfinfo')
                if pdfinfo_path:
                    return os.path.dirname(pdfinfo_path)
                
                return None
            
            # Convert PDF to image with poppler path if needed
            poppler_path = find_poppler_windows()
            
            try:
                if poppler_path:
                    images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, poppler_path=poppler_path)
                else:
                    images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1)
            except Exception as e:
                self.logger.warning(f"PDF to image conversion failed: {e}")
                return ""
            
            if images:
                try:
                    text = pytesseract.image_to_string(images[0])
                    return text
                except Exception as e:
                    self.logger.error(f"OCR processing failed: {e}")
                    return ""
                
        except Exception as e:
            self.logger.error(f"OCR failed: {e}")
            return ""
            
            # Try to convert PDF to image
            try:
                images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1)
            except Exception as e:
                self.logger.warning(f"PDF to image conversion failed: {e}")
                # Try with different poppler path if available
                common_poppler_paths = [
                    r'C:\poppler\bin',
                    r'C:\poppler-utils\bin',
                    r'C:\Program Files\poppler\bin'
                ]
                
                for path in common_poppler_paths:
                    if os.path.exists(path):
                        try:
                            images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, poppler_path=path)
                            break
                        except:
                            continue
                else:
                    return ""
            
            if images:
                # Perform OCR on the image
                try:
                    text = pytesseract.image_to_string(images[0])
                    return text
                except Exception as e:
                    self.logger.error(f"OCR processing failed: {e}")
                    return ""
                
        except Exception as e:
            self.logger.error(f"OCR failed: {e}")
            return ""
    
    def extract_document_details(self):
        """Extract document details from current page"""
        details = {
            'title': '',
            'date': datetime.now().strftime("%Y-%m-%d"),
            'url': self.driver.current_url
        }
        
        try:
            title_elem = self.driver.find_element(By.XPATH, "//h1")
            details['title'] = title_elem.text.strip()
        except:
            pass
        
        return details
    
    def process_document_page(self):
        """Process current document page"""
        try:
            time.sleep(2)
            doc_details = self.extract_document_details()
            
            # Find download link
            download_selectors = [
                "/html/body/div/div[1]/div/table/tbody/tr/td[1]/div/div/div/div[2]/div[2]/div/p/span/a",
                "//a[contains(@href, '.pdf')]",
                "//a[contains(@href, 'at_download')]",
                "//a[contains(text(), 'Download')]"
            ]
            
            download_link = None
            for selector in download_selectors:
                try:
                    download_link = self.driver.find_element(By.XPATH, selector)
                    if download_link:
                        href = download_link.get_attribute('href')
                        if href and ('.pdf' in href or 'at_download' in href):
                            self.logger.info(f"Found download link with selector: {selector}")
                            break
                except:
                    continue
            
            if not download_link:
                self.logger.error("Download link not found")
                return None
            
            pdf_url = download_link.get_attribute('href')
            safe_title = "".join(c for c in doc_details['title'] if c.isalnum() or c in (' ', '-', '_')).rstrip()
            if not safe_title:
                safe_title = f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            pdf_filename = f"{safe_title}_{timestamp}.pdf"
            
            self.logger.info(f"Downloading from: {pdf_url}")
            pdf_path = self.download_pdf(pdf_url, pdf_filename)
            
            if pdf_path and os.path.exists(pdf_path):
                self.logger.info(f"PDF saved to: {pdf_path}")
                body_text = self.extract_text_from_pdf(pdf_path)
                
                document = {
                    'date': doc_details['date'],
                    'title': doc_details['title'],
                    'body': body_text,
                    'url': doc_details['url'],
                    'pdf_url': pdf_url
                }
                
                # Clean up if not keeping PDFs
                if not self.keep_pdfs:
                    try:
                        os.remove(pdf_path)
                        self.logger.info(f"Deleted temporary PDF: {pdf_filename}")
                    except Exception as e:
                        self.logger.warning(f"Could not delete PDF: {e}")
                
                return document
            
            return None
            
        except Exception as e:
            self.logger.error(f"Error processing document page: {e}")
            return None
    
    def download_only(self, pdf_url, title):
        """Download PDF only, no OCR"""
        try:
            safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
            if not safe_title:
                safe_title = f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            pdf_filename = f"{safe_title}_{timestamp}.pdf"
            
            self.logger.info(f"Downloading from: {pdf_url}")
            pdf_path = self.download_pdf(pdf_url, pdf_filename)
            
            return pdf_path if pdf_path and os.path.exists(pdf_path) else None
            
        except Exception as e:
            self.logger.error(f"Error in download_only: {e}")
            return None

    def initialize_download_phase(self, search_term):
        """Initialize download phase and load existing downloads"""
        tracking_file = f"downloaded_{search_term}.json"
        downloaded_pdfs = []
        
        if os.path.exists(tracking_file):
            try:
                with open(tracking_file, 'r', encoding='utf-8') as f:
                    downloaded_pdfs = json.load(f)
                self.logger.info(f"Loaded {len(downloaded_pdfs)} previously downloaded PDFs from {tracking_file}")
            except Exception as e:
                self.logger.warning(f"Could not load tracking file: {e}")
        
        return downloaded_pdfs, tracking_file

    def get_expected_document_count(self):
        """Get expected total document count from page"""
        try:
            count_text = self.driver.find_element(By.XPATH, "//span[@class='results-count']").text
            import re
            numbers = re.findall(r'\d+', count_text)
            if numbers:
                return int(numbers[-1])
        except:
            pass
        return None

    def process_single_document_download(self, link, search_page_url, downloaded_pdfs, 
                                        tracking_file, documents_processed, max_retries=3):
        """Process download for a single document with retry logic"""
        link_url = link.get_attribute('href')
        link_text = link.text.strip() or "Unknown Document"
        
        # Check if already downloaded
        if any(pdf['url'] == link_url for pdf in downloaded_pdfs):
            self.logger.info(f"⏭ Already downloaded: {link_text[:50]}...")
            return documents_processed, True
        
        retry_count = 0
        while retry_count < max_retries:
            try:
                # Check driver health
                if not self.check_driver_health():
                    self.logger.warning("Driver session invalid, recovering...")
                    self.recover_driver_session(self.current_search_term)
                    search_page_url = self.driver.current_url
                
                self.logger.info(f"\nDownloading: {link_text[:100]}...")
                
                # Navigate to document page
                self.driver.get(link_url)
                time.sleep(3)
                
                # Get PDF URL
                pdf_url = self.extract_pdf_url_from_page()
                if not pdf_url:
                    self.logger.error("No PDF download link found")
                    self.driver.get(search_page_url)
                    time.sleep(3)
                    retry_count += 1
                    continue
                
                # Create filename
                pdf_filename = self.create_pdf_filename(link_text)
                
                # Download PDF
                self.logger.info(f"Downloading from: {pdf_url}")
                pdf_path = self.download_pdf(pdf_url, pdf_filename)
                
                if pdf_path and os.path.exists(pdf_path):
                    downloaded_pdfs.append({
                        'path': pdf_path,
                        'url': link_url,
                        'title': link_text,
                        'ocr': 'pending'
                    })
                    documents_processed += 1
                    
                    # Save tracking file
                    with open(tracking_file, 'w', encoding='utf-8') as f:
                        json.dump(downloaded_pdfs, f, indent=2)
                    
                    self.logger.info(f"✓ Downloaded ({documents_processed}): {os.path.basename(pdf_path)}")
                    
                    # Return to search results
                    self.driver.get(search_page_url)
                    time.sleep(2)
                    return documents_processed, True
                else:
                    self.logger.error("Download failed - no file created")
                    retry_count += 1
                    
            except Exception as e:
                retry_count += 1
                self.logger.error(f"Error (attempt {retry_count}/{max_retries}): {e}")
                
                if "invalid session id" in str(e):
                    self.recover_driver_session(self.current_search_term)
                    search_page_url = self.driver.current_url
                
                if retry_count < max_retries:
                    self.logger.info("Retrying in 5 seconds...")
                    time.sleep(5)
                    try:
                        self.driver.get(search_page_url)
                        time.sleep(3)
                    except:
                        pass
        
        self.logger.error(f"Failed after {max_retries} attempts")
        return documents_processed, False

    def extract_pdf_url_from_page(self):
        """Extract PDF download URL from current document page"""
        download_selectors = [
            "/html/body/div/div[1]/div/table/tbody/tr/td[1]/div/div/div/div[2]/div[2]/div/p/span/a",
            "//a[contains(@href, '.pdf')]",
            "//a[contains(@href, 'at_download')]",
            "//a[contains(text(), 'Download')]"
        ]
        
        for selector in download_selectors:
            try:
                potential_links = self.driver.find_elements(By.XPATH, selector)
                for link in potential_links:
                    href = link.get_attribute('href')
                    if href and ('.pdf' in href or 'at_download' in href):
                        return href
            except:
                continue
        return None

    def create_pdf_filename(self, title):
        """Create safe filename from document title"""
        safe_title = "".join(c for c in title if c.isalnum() or c in (' ', '-', '_')).rstrip()
        if not safe_title:
            safe_title = f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"{safe_title}_{timestamp}.pdf"

    def process_page_downloads(self, page_num, search_page_url, downloaded_pdfs, 
                            tracking_file, documents_processed, max_documents):
        """Process all documents on current page"""
        result_links = self.get_result_links()
        self.logger.info(f"Found {len(result_links)} documents on page {page_num}")
        
        for i in range(len(result_links)):
            if max_documents and documents_processed >= max_documents:
                self.logger.info(f"Reached maximum documents ({max_documents})")
                return documents_processed, True
            
            # Refresh links to avoid staleness
            current_links = self.get_result_links()
            if i >= len(current_links):
                self.logger.warning(f"Link index {i} out of range")
                break
            
            documents_processed, _ = self.process_single_document_download(
                current_links[i], search_page_url, downloaded_pdfs, 
                tracking_file, documents_processed
            )
        
        return documents_processed, False

    def verify_downloads_complete(self, search_term, downloaded_pdfs, tracking_file, 
                                max_verification_attempts=3):
        """Verify all documents are downloaded, download any missing"""
        self.logger.info("\n" + "="*50)
        self.logger.info("PHASE 1.5: Verifying all PDFs downloaded")
        self.logger.info("="*50)
        
        for attempt in range(1, max_verification_attempts + 1):
            self.logger.info(f"\nVerification attempt {attempt}/{max_verification_attempts}")
            
            # Count total available
            self.driver.get(f"https://vault.fbi.gov/search?SearchableText={search_term}")
            time.sleep(3)
            
            total_available = 0
            page_num = 1
            
            while True:
                result_links = self.get_result_links()
                total_available += len(result_links)
                
                if not self.navigate_to_next_page():
                    break
                page_num += 1
                time.sleep(2)
            
            self.logger.info(f"Available: {total_available}, Downloaded: {len(downloaded_pdfs)}")
            
            if len(downloaded_pdfs) >= total_available:
                self.logger.info("✓ All documents downloaded!")
                return True
            
            if attempt < max_verification_attempts:
                self.logger.warning(f"Missing {total_available - len(downloaded_pdfs)} documents")
                self.download_missing_documents(search_term, downloaded_pdfs, tracking_file)
                time.sleep(5)
        
        return False

    def download_missing_documents(self, search_term, downloaded_pdfs, tracking_file):
        """Download any missing documents"""
        downloaded_urls = {pdf['url'] for pdf in downloaded_pdfs}
        
        self.driver.get(f"https://vault.fbi.gov/search?SearchableText={search_term}")
        time.sleep(3)
        
        while True:
            result_links = self.get_result_links()
            
            for link in result_links:
                link_url = link.get_attribute('href')
                
                if link_url not in downloaded_urls:
                    link_text = link.text.strip() or "Unknown"
                    self.logger.info(f"Found missing: {link_text[:50]}...")
                    
                    try:
                        self.driver.get(link_url)
                        time.sleep(3)
                        
                        pdf_url = self.extract_pdf_url_from_page()
                        if pdf_url:
                            pdf_filename = self.create_pdf_filename(link_text)
                            pdf_path = self.download_pdf(pdf_url, pdf_filename)
                            
                            if pdf_path and os.path.exists(pdf_path):
                                downloaded_pdfs.append({
                                    'path': pdf_path,
                                    'url': link_url,
                                    'title': link_text,
                                    'ocr': 'pending'
                                })
                                downloaded_urls.add(link_url)
                                
                                with open(tracking_file, 'w', encoding='utf-8') as f:
                                    json.dump(downloaded_pdfs, f, indent=2)
                        
                        self.driver.back()
                        time.sleep(2)
                        
                    except Exception as e:
                        self.logger.error(f"Error downloading missing: {e}")
            
            if not self.navigate_to_next_page():
                break

    def run(self, search_term, max_documents=None):
        """Main execution method - download all PDFs first, then OCR"""
        self.current_search_term = search_term  # Store for recovery
        
        try:
            self.setup_driver()
            self.search(search_term)
            
            # PHASE 1: Download all PDFs
            downloaded_pdfs, tracking_file = self.initialize_download_phase(search_term)
            documents_processed = len(downloaded_pdfs)
            page_num = 1
            expected_count = self.get_expected_document_count()
            
            self.logger.info("\n" + "="*50)
            self.logger.info("PHASE 1: Downloading all PDFs")
            self.logger.info("="*50)
            
            # Main download loop
            while True:
                self.logger.info(f"\n{'='*50}")
                self.logger.info(f"Page {page_num} - Download Phase")
                self.logger.info(f"{'='*50}")
                
                result_links = self.get_result_links()
                if not result_links:
                    self.logger.info("No results found")
                    break
                
                search_page_url = self.driver.current_url
                
                # Process current page
                documents_processed, reached_max = self.process_page_downloads(
                    page_num, search_page_url, downloaded_pdfs, tracking_file,
                    documents_processed, max_documents
                )
                
                if reached_max:
                    break
                
                # Next page
                if not self.navigate_to_next_page():
                    break
                
                page_num += 1
                time.sleep(3)
            
            # Verify all downloads complete
            all_downloaded = self.verify_downloads_complete(
                search_term, downloaded_pdfs, tracking_file
            )
            
            self.logger.info(f"\n{'='*50}")
            self.logger.info(f"Download phase complete! Downloaded {len(downloaded_pdfs)} PDFs")
            self.logger.info(f"{'='*50}")
            
            # PHASE 2: OCR Processing
            if downloaded_pdfs:
                self.process_ocr_phase(search_term, downloaded_pdfs, tracking_file)
                
                # Final verification
                if not self.verify_all_ocr_complete(search_term):
                    self.logger.info("Re-entering OCR phase...")
                    self.process_ocr_phase(search_term, downloaded_pdfs, tracking_file)
            
            # Cleanup
            if os.path.exists(tracking_file) and self.verify_all_ocr_complete(search_term):
                os.remove(tracking_file)
            
            self.logger.info(f"\n{'='*50}")
            self.logger.info(f"Scraping completed! Processed {len(self.results)} documents")
            self.logger.info(f"PDFs saved to: {self.pdf_storage_dir}")
            self.logger.info(f"{'='*50}")
            
        except KeyboardInterrupt:
            self.logger.info("\nScraping interrupted by user")
            self.save_results(search_term)
            
        except Exception as e:
            self.logger.error(f"Fatal error during scraping: {e}")
            self.save_results(search_term)
            
        finally:
            if self.driver:
                self.driver.quit()

    def process_ocr_phase(self, search_term, downloaded_pdfs, tracking_file):
        """Process OCR for all PDFs - will retry until all are complete or user intervenes"""
        
        self.logger.info("\n" + "="*50)
        self.logger.info("PHASE 2: OCR Processing (Required for Completion)")
        self.logger.info("="*50)
        
        # Load existing tracking data
        if os.path.exists(tracking_file):
            try:
                with open(tracking_file, 'r', encoding='utf-8') as f:
                    downloaded_pdfs = json.load(f)
                self.logger.info(f"Loaded {len(downloaded_pdfs)} PDFs from tracking file")
            except Exception as e:
                self.logger.warning(f"Could not load tracking file: {e}")
        
        # Load existing JSON results
        json_filename = f"fbi_vault_{search_term}.json"
        if os.path.exists(json_filename):
            try:
                with open(json_filename, 'r', encoding='utf-8') as f:
                    self.results = json.load(f)
                self.logger.info(f"Loaded {len(self.results)} existing OCR results")
            except:
                self.results = []
        
        # OCR PROCESSING LOOP - WILL CONTINUE UNTIL ALL ARE COMPLETE
        ocr_attempts = 0
        max_ocr_attempts = 3  # Maximum retry attempts for failed OCR
        
        while True:
            # Count current OCR status
            ocr_complete = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'complete')
            ocr_failed = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'failed')
            ocr_pending = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'pending')
            total_pdfs = len(downloaded_pdfs)
            
            self.logger.info(f"\nOCR Status: {ocr_complete}/{total_pdfs} complete, {ocr_pending} pending, {ocr_failed} failed")
            
            # Check if all PDFs are successfully OCR'd
            if ocr_complete == total_pdfs:
                self.logger.info("\n" + "="*50)
                self.logger.info("✓ ALL PDFs SUCCESSFULLY OCR'd - PROCESS COMPLETE")
                self.logger.info("="*50)
                break
            
            # If we've tried too many times and still have failures, warn but continue
            if ocr_attempts >= max_ocr_attempts and ocr_failed > 0:
                self.logger.warning(f"\n⚠ Reached maximum OCR attempts ({max_ocr_attempts})")
                self.logger.warning(f"Failed to OCR {ocr_failed} documents after multiple attempts")
                response = input("Continue anyway? (y/n): ").lower()
                if response != 'y':
                    self.logger.info("OCR phase incomplete - script will exit")
                    break
            
            # Process pending and failed PDFs
            for idx, pdf_info in enumerate(downloaded_pdfs, 1):
                # Skip if already complete
                if pdf_info.get('ocr') == 'complete':
                    continue
                
                # Skip failed ones after max attempts
                if pdf_info.get('ocr') == 'failed' and ocr_attempts >= max_ocr_attempts:
                    continue
                
                self.logger.info(f"\nProcessing PDF {idx}/{total_pdfs}: {os.path.basename(pdf_info['path'])}")
                self.logger.info(f"Current status: {pdf_info.get('ocr', 'pending')}")
                
                try:
                    # Extract text with OCR
                    body_text = self.extract_text_from_pdf(pdf_info['path'])
                    
                    # Get document details
                    doc_title = pdf_info['title']
                    doc_url = pdf_info['url']
                    
                    try:
                        self.driver.get(pdf_info['url'])
                        time.sleep(2)
                        doc_details = self.extract_document_details()
                        if doc_details['title']:
                            doc_title = doc_details['title']
                    except:
                        self.logger.warning("Could not fetch document details from URL, using stored title")
                    
                    # Check if OCR actually extracted text
                    if body_text and len(body_text.strip()) > 50:  # Arbitrary minimum text length
                        document = {
                            'date': datetime.now().strftime("%Y-%m-%d"),
                            'title': doc_title,
                            'body': body_text,
                            'url': doc_url,
                            'pdf_path': pdf_info['path']
                        }
                        
                        self.results.append(document)
                        
                        # Mark as complete
                        pdf_info['ocr'] = 'complete'
                        self.logger.info(f"✓ OCR successful - Text length: {len(body_text)} characters")
                    else:
                        # Mark as failed if no text extracted
                        pdf_info['ocr'] = 'failed'
                        self.logger.warning(f"⚠ OCR produced insufficient text - Marked as failed")
                    
                    # Save tracking file after each document
                    with open(tracking_file, 'w', encoding='utf-8') as f:
                        json.dump(downloaded_pdfs, f, indent=2)
                    
                    # Save JSON results after each document
                    self.save_results(search_term)
                    
                except Exception as e:
                    self.logger.error(f"OCR failed: {e}")
                    pdf_info['ocr'] = 'failed'
                    
                    # Save tracking file even on failure
                    with open(tracking_file, 'w', encoding='utf-8') as f:
                        json.dump(downloaded_pdfs, f, indent=2)
            
            # Increment attempt counter
            ocr_attempts += 1
            
            # If we still have pending/failed after this pass, ask user
            ocr_complete = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'complete')
            ocr_failed = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'failed')
            ocr_pending = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'pending')
            
            if ocr_complete < total_pdfs:
                self.logger.info(f"\nAfter pass {ocr_attempts}:")
                self.logger.info(f"Complete: {ocr_complete}, Pending: {ocr_pending}, Failed: {ocr_failed}")
                
                if ocr_attempts < max_ocr_attempts:
                    self.logger.info(f"Waiting 10 seconds before retry {ocr_attempts + 1}/{max_ocr_attempts}...")
                    time.sleep(10)
                else:
                    self.logger.warning("Max retry attempts reached for remaining documents")
                    response = input("Continue with incomplete OCR? (y/n): ").lower()
                    if response != 'y':
                        self.logger.info("OCR phase incomplete - script will exit")
                        break
        
        # Final status report
        ocr_complete = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'complete')
        ocr_failed = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'failed')
        ocr_pending = sum(1 for pdf in downloaded_pdfs if pdf.get('ocr') == 'pending')
        
        self.logger.info(f"\n{'='*50}")
        self.logger.info(f"OCR Phase Final Status:")
        self.logger.info(f"✓ Successfully OCR'd: {ocr_complete} documents")
        self.logger.info(f"✗ Failed OCR: {ocr_failed} documents")
        self.logger.info(f"⏳ Pending: {ocr_pending} documents")
        self.logger.info(f"Tracking file: {tracking_file}")
        self.logger.info(f"{'='*50}")

    def verify_all_ocr_complete(self, search_term):
        """Verify that all PDFs have been OCR'd before allowing script to exit"""
        tracking_file = f"downloaded_{search_term}.json"
        
        if os.path.exists(tracking_file):
            with open(tracking_file, 'r', encoding='utf-8') as f:
                downloaded_pdfs = json.load(f)
            
            incomplete = [pdf for pdf in downloaded_pdfs if pdf.get('ocr') != 'complete']
            
            if incomplete:
                self.logger.warning(f"\n⚠ WARNING: {len(incomplete)} PDFs have not been successfully OCR'd:")
                for pdf in incomplete:
                    self.logger.warning(f"  - {os.path.basename(pdf['path'])}: {pdf.get('ocr', 'pending')}")
                
                response = input("\nOCR phase incomplete. Exit anyway? (y/n): ").lower()
                if response != 'y':
                    self.logger.info("Restarting OCR phase...")
                    return False
        
        return True
    
    def save_results(self, search_term):
        """Save results to JSON file"""
        safe_term = "".join(c for c in search_term if c.isalnum() or c in (' ', '-', '_')).rstrip()
        filename = f"fbi_vault_{safe_term}.json"
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(self.results, f, indent=2, ensure_ascii=False)
            self.logger.info(f"Saved {len(self.results)} results to {filename}")
        except Exception as e:
            self.logger.error(f"Failed to save results: {e}")
    
    def cleanup_temp_directory(self):
        """Clean up orphaned files in temp directory"""
        try:
            files = os.listdir(self.download_dir)
            for f in files:
                file_path = os.path.join(self.download_dir, f)
                try:
                    # Only remove files older than 10 minutes
                    if time.time() - os.path.getctime(file_path) > 600:
                        os.remove(file_path)
                        self.logger.info(f"Cleaned up old temp file: {f}")
                except:
                    pass
        except:
            pass


# ==================== MAIN ENTRY POINT ====================

def main():
    """Main function"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Scrape FBI Vault for documents')
    parser.add_argument('term', help='Search term')
    parser.add_argument('--max', type=int, help='Maximum number of documents to process')
    parser.add_argument('--download-dir', help='Directory for temporary PDF downloads')
    parser.add_argument('--skip-deps', action='store_true', help='Skip dependency checking')
    
    args = parser.parse_args()
    
    print("\n" + "="*60)
    print("FBI VAULT SCRAPER")
    print("="*60)
    print(f"Python: {sys.version}")
    print(f"Architecture: {platform.architecture()[0]}")
    print(f"System: {platform.system()} {platform.release()}")
    
    if not args.skip_deps:
        # Check Python dependencies
        if not check_and_install_dependencies():
            print("\n⚠ Some Python dependencies failed to install")
            response = input("Continue anyway? (y/n): ").lower()
            if response != 'y':
                sys.exit(1)
        
        # Try to install OCR tools via pip
        install_tesseract_via_pip()
        install_poppler_via_pip()
        
        # Check Chrome
        check_chrome_and_driver()
        
        # Try manual ChromeDriver download as fallback
        try:
            from webdriver_manager.chrome import ChromeDriverManager
        except:
            print("\nAttempting manual ChromeDriver download...")
            download_chromedriver_manual()
    
    # Run the scraper
    scraper = FBIVaultScraper(download_dir=args.download_dir)
    
    try:
        scraper.run(args.term, max_documents=args.max)
    except Exception as e:
        print(f"\nError: {e}")
        print("\nTroubleshooting tips:")
        print("1. Make sure Chrome is installed")
        print("2. Try running as administrator")
        print("3. Check if Chrome is 64-bit and Python is 64-bit")
        print("4. Run with --skip-deps to skip dependency checking")
        sys.exit(1)

if __name__ == "__main__":
    main()