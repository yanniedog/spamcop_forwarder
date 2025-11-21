# Standard library imports
import imaplib
import smtplib
import email
import os
import sys
import datetime
import time
import re
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import parsedate_to_datetime
from email.header import decode_header

# Third-party imports
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeRemainingColumn, TimeElapsedColumn  # type: ignore
from rich.console import Console  # type: ignore

# Local imports
import messages as msg  # type: ignore

# Initialize rich console
console = Console()

# ==========================================
#              CONFIGURATION LOADING
# ==========================================

def load_config():
    """Loads configuration from config.py, creating it with placeholders if missing"""
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.py')
    
    # Check if config.py exists, create it if not
    if not os.path.exists(config_path):
        print("=" * 70)
        print("CONFIG FILE NOT FOUND")
        print("=" * 70)
        print(f"Creating {config_path} with placeholder values...")
        print("Please edit this file and replace the placeholders with your actual values.")
        print("=" * 70)
        
        # Create config.py with placeholders
        config_content = '''# ==========================================
#              CONFIGURATION
# ==========================================
# This file contains all configuration settings for the SpamCop Forwarder.
# IMPORTANT: This file is in .gitignore - do NOT commit your actual credentials!
# ==========================================

# 1. SOURCE ACCOUNT (The one receiving spam and sending to SpamCop)
#    Permissions: READ-ONLY for IMAP. Will use SMTP to forward spam.
GMAIL_ACCOUNT = 'YOUR_GMAIL_ACCOUNT_HERE'  # e.g., 'yourname@gmail.com'
APP_PASS = 'YOUR_APP_PASSWORD_HERE'  # Gmail App Password (16 characters, spaces will be auto-removed)

# 2. SMTP CONFIGURATION (Using same account as download)
#    Gmail SMTP settings - these are standard and should not be changed
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587  # Standard Gmail SMTP port with STARTTLS

# 3. DESTINATION
SPAMCOP_ADDRESS = 'YOUR_SPAMCOP_QUICK_SEND_ADDRESS_HERE'  # e.g., 'quick.xxxxxxxxxxxxx@spam.spamcop.net'

# 4. LOCAL STORAGE
#    Base directory for downloaded spam emails
#    Default: 'downloads' folder in the current working directory
BASE_DIRECTORY = 'downloads'  # Relative to script location, or use absolute path

# 5. LOOP CONFIGURATION
#    Frequency in hours: Must be > 0 and <= 48
#    This determines how often the script runs
LOOP_FREQUENCY_HOURS = 5  # Run every 5 hours
# For testing, you can use a smaller value like 0.1 (6 minutes) or 0.05 (3 minutes)

# 6. SPAM SEARCH WINDOW
#    How far back (in hours) to search for spam emails
#    Must be > 0 and <= 168 (7 days)
#    This should typically be >= LOOP_FREQUENCY_HOURS to avoid missing spam
SPAM_SEARCH_WINDOW_HOURS = 5  # Search for spam from the past 5 hours

# 7. SIMULATION MODE
#    When True: Performs all steps (connect, search, download, save files) but does NOT forward to SpamCop
#    When False: Performs all steps including forwarding to SpamCop
#    Default: True (simulation mode ON) - set to False when ready to actually forward spam
SIMULATION_MODE = True  # Set to False to enable actual forwarding to SpamCop
'''
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(config_content)
        print(f"Created {config_path}")
        print("\nPlease edit the config file and replace the placeholders, then run the script again.")
        sys.exit(1)
    
    # Import config
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("config", config_path)
        if spec is None or spec.loader is None:
            raise ImportError(f"Could not load config from {config_path}")
        config = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(config)
        
        # Check for placeholders
        placeholders_found = []
        if hasattr(config, 'GMAIL_ACCOUNT') and 'YOUR_GMAIL_ACCOUNT_HERE' in str(config.GMAIL_ACCOUNT):
            placeholders_found.append('GMAIL_ACCOUNT')
        if hasattr(config, 'APP_PASS') and 'YOUR_APP_PASSWORD_HERE' in str(config.APP_PASS):
            placeholders_found.append('APP_PASS')
        if hasattr(config, 'SPAMCOP_ADDRESS') and 'YOUR_SPAMCOP_QUICK_SEND_ADDRESS_HERE' in str(config.SPAMCOP_ADDRESS):
            placeholders_found.append('SPAMCOP_ADDRESS')
        
        if placeholders_found:
            print_config_instructions(placeholders_found)
            sys.exit(1)
        
        # Return config values
        # Strip spaces from APP_PASS (Google provides it with spaces, but IMAP/SMTP need it without)
        app_pass = str(config.APP_PASS).replace(' ', '')
        
        return {
            'GMAIL_ACCOUNT': config.GMAIL_ACCOUNT,
            'APP_PASS': app_pass,
            'SMTP_SERVER': config.SMTP_SERVER,
            'SMTP_PORT': config.SMTP_PORT,
            'SPAMCOP_ADDRESS': config.SPAMCOP_ADDRESS,
            'BASE_DIRECTORY': config.BASE_DIRECTORY,
            'LOOP_FREQUENCY_HOURS': config.LOOP_FREQUENCY_HOURS,
            'SPAM_SEARCH_WINDOW_HOURS': getattr(config, 'SPAM_SEARCH_WINDOW_HOURS', config.LOOP_FREQUENCY_HOURS),
            'SIMULATION_MODE': getattr(config, 'SIMULATION_MODE', True)
        }
    except Exception as e:
        print(f"Error loading configuration: {e}")
        print(f"Please check that {config_path} exists and is valid Python code.")
        sys.exit(1)

def print_config_instructions(missing_fields):
    """Prints detailed instructions for obtaining configuration values"""
    print("\n" + "=" * 70)
    print("CONFIGURATION INCOMPLETE")
    print("=" * 70)
    print("The following configuration values still contain placeholders:")
    for field in missing_fields:
        print(f"  - {field}")
    print("\n" + "=" * 70)
    print("STEP-BY-STEP INSTRUCTIONS TO OBTAIN ALL CONFIGURATION VALUES")
    print("=" * 70)
    
    # Get instructions from messages module
    instructions = msg.get_config_instructions(missing_fields)  # type: ignore
    if instructions:
        print(instructions)
    
    print("\n" + "=" * 70)
    print("SMTP CONFIGURATION (Already Set - No Action Needed)")
    print("-" * 70)
    print(msg.SMTP_CONFIG_INFO)  # type: ignore
    print()
    print("=" * 70)
    print("NEXT STEPS:")
    print("=" * 70)
    print("1. Open config.py in a text editor")
    print("2. Replace all placeholder values with your actual values")
    print("3. Save the file")
    print("4. Run this script again")
    print("=" * 70 + "\n")

# Load configuration
config = load_config()
GMAIL_ACCOUNT = config['GMAIL_ACCOUNT']
APP_PASS = config['APP_PASS']
SMTP_SERVER = config['SMTP_SERVER']
SMTP_PORT = config['SMTP_PORT']
SPAMCOP_ADDRESS = config['SPAMCOP_ADDRESS']
BASE_DIRECTORY = config['BASE_DIRECTORY']
LOOP_FREQUENCY_HOURS = config['LOOP_FREQUENCY_HOURS']
SPAM_SEARCH_WINDOW_HOURS = config['SPAM_SEARCH_WINDOW_HOURS']
SIMULATION_MODE = config['SIMULATION_MODE']

# Convert BASE_DIRECTORY to absolute path if it's relative
if not os.path.isabs(BASE_DIRECTORY):
    BASE_DIRECTORY = os.path.join(os.path.dirname(os.path.abspath(__file__)), BASE_DIRECTORY)

# ==========================================

def decode_email_header(header_value):
    """Decodes email header values that may be encoded (UTF-8, base64, etc.)"""
    if not header_value:
        return ""
    
    decoded_parts = decode_header(header_value)
    decoded_str = ""
    for part, encoding in decoded_parts:
        if isinstance(part, bytes):
            if encoding:
                try:
                    decoded_str += part.decode(encoding)
                except (UnicodeDecodeError, LookupError):
                    decoded_str += part.decode('utf-8', errors='ignore')
            else:
                decoded_str += part.decode('utf-8', errors='ignore')
        else:
            decoded_str += str(part)
    
    return decoded_str

def safe_print_subject(subject, max_len=50):
    """Safely prints subject line, handling unicode/emoji for Windows console"""
    if not subject:
        return "(No Subject)"
    
    # Try to encode/decode to handle unicode characters that Windows console can't display
    try:
        # First try to print directly
        display = subject[:max_len]
        # Test if it can be encoded to console encoding
        display.encode(sys.stdout.encoding or 'utf-8', errors='strict')
        return display
    except (UnicodeEncodeError, AttributeError):
        # If encoding fails, replace problematic characters
        try:
            return subject.encode('ascii', errors='replace').decode('ascii')[:max_len]
        except:
            return subject[:max_len].encode('utf-8', errors='replace').decode('utf-8', errors='replace')

def sanitize_filename(filename):
    """Cleans filenames to ensure they are valid for Windows/Linux"""
    if not filename:
        return "NoSubject"
    
    # Remove newlines, carriage returns, and other control characters
    filename = re.sub(r'[\r\n\t]', '', str(filename))
    
    # Windows invalid characters: < > : " / \ | ? *
    # Also remove other problematic characters
    filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', filename)
    
    # Remove any remaining non-printable characters
    filename = re.sub(r'[^\w\s\-\.]', '', filename)
    
    # Replace multiple spaces/underscores with single underscore
    filename = re.sub(r'[\s_]+', '_', filename)
    
    # Remove leading/trailing dots and spaces
    filename = filename.strip('. ')
    
    # If empty after sanitization, use default
    if not filename:
        filename = "NoSubject"
    
    return filename

def get_size_str(bytes_val):
    """Returns formatted size string"""
    if bytes_val < 1024:
        return f"{bytes_val}b"
    elif bytes_val < 1024 * 1024:
        return f"{int(bytes_val / 1024)}kb"
    else:
        return f"{int(bytes_val / (1024 * 1024))}mb"

def validate_loop_frequency(hours):
    """Validates that loop frequency is > 0 and <= 48 hours"""
    if hours is None or hours == 0:
        raise ValueError("LOOP_FREQUENCY_HOURS must be greater than 0. Cannot run in rapid loop mode.")
    if hours > 48:
        raise ValueError(f"LOOP_FREQUENCY_HOURS ({hours}) exceeds maximum of 48 hours.")
    if hours < 0:
        raise ValueError(f"LOOP_FREQUENCY_HOURS ({hours}) cannot be negative.")
    return True

def validate_search_window(hours):
    """Validates that search window is > 0 and <= 168 hours (7 days)"""
    if hours is None or hours == 0:
        raise ValueError("SPAM_SEARCH_WINDOW_HOURS must be greater than 0.")
    if hours > 168:
        raise ValueError(f"SPAM_SEARCH_WINDOW_HOURS ({hours}) exceeds maximum of 168 hours (7 days).")
    if hours < 0:
        raise ValueError(f"SPAM_SEARCH_WINDOW_HOURS ({hours}) cannot be negative.")
    return True

def process_spam_iteration():
    """Processes one iteration of spam download and forwarding"""
    print("--- STARTING SPAM PROCESSOR ITERATION ---")
    
    mail = None
    spam_candidates = []
    total_size_bytes = 0
    
    # ---------------------------------------------------------
    # PHASE 1: CONNECT AND IDENTIFY (READ-ONLY)
    # ---------------------------------------------------------
    try:
        print(f"Connecting to {GMAIL_ACCOUNT}...")
        mail = imaplib.IMAP4_SSL('imap.gmail.com')
        mail.login(GMAIL_ACCOUNT, APP_PASS)
        
        # First, list all folders to get exact folder names
        print("Listing all available mailboxes to find spam folder...")
        all_folder_names = []
        spam_candidate_names = []
        try:
            status, folders = mail.list()
            if status == 'OK':
                for folder in folders:
                    folder_str = ""
                    if isinstance(folder, bytes):
                        folder_str = folder.decode('utf-8', errors='ignore')
                    elif isinstance(folder, tuple):
                        decoded_parts = []
                        for part in folder:
                            if isinstance(part, bytes):
                                decoded_parts.append(part.decode('utf-8', errors='ignore'))
                            else:
                                decoded_parts.append(str(part))
                        folder_str = ' '.join(decoded_parts)
                    else:
                        folder_str = str(folder)
                    
                    # Parse IMAP LIST response format: (\\HasNoChildren) "/" "INBOX"
                    # or: (\\HasChildren \\Noselect) "/" "[Gmail]"
                    # Extract the folder name (last quoted string or unquoted part)
                    folder_name = None
                    # Try to extract quoted folder name
                    quoted_match = re.search(r'"([^"]+)"', folder_str)
                    if quoted_match:
                        folder_name = quoted_match.group(1)
                    else:
                        # If no quotes, try to get the last part after spaces
                        parts = folder_str.split()
                        if len(parts) > 0:
                            folder_name = parts[-1]
                    
                    if folder_name:
                        all_folder_names.append(folder_name)
                        folder_upper = folder_name.upper()
                        # Identify spam/junk folders (but exclude INBOX)
                        if ('SPAM' in folder_upper or 'JUNK' in folder_upper) and 'INBOX' not in folder_upper:
                            spam_candidate_names.append(folder_name)
        except Exception as list_err:
            print(f"Warning: Could not list mailboxes: {list_err}")
        
        # Print found folders for debugging
        if all_folder_names:
            print(f"Found {len(all_folder_names)} mailboxes")
            if spam_candidate_names:
                print(f"Spam/Junk folder candidates: {spam_candidate_names}")
        
        # On initial run, show message counts for all folders
        initial_listing_flag_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.spamcop_initial_listing_complete')
        is_initial_run = not os.path.exists(initial_listing_flag_file)
        
        if is_initial_run and all_folder_names:
            print("\n" + "="*70)
            print("INITIAL RUN - FOLDER MESSAGE COUNTS")
            print("="*70)
            print("Getting message counts for all folders...")
            print()
            
            folder_counts = []
            for folder_name in all_folder_names:
                # Skip INBOX for security
                if folder_name.upper() == 'INBOX':
                    continue
                
                msg_count = None
                try:
                    # Get message count using STATUS command
                    # Note: Some folders (like parent folders) may not support STATUS
                    status, data = mail.status(folder_name, "(MESSAGES)")
                    if status == 'OK' and data:
                        # Parse message count from response like: (MESSAGES 123)
                        # Response format can vary: b'(MESSAGES 123)' or ('MESSAGES', 123)
                        count_str = str(data[0])
                        count_match = re.search(r'MESSAGES\s+(\d+)', count_str)
                        if count_match:
                            msg_count = int(count_match.group(1))
                        else:
                            # Try alternative format
                            if isinstance(data[0], (bytes, str)):
                                # Try to find number directly
                                num_match = re.search(r'(\d+)', count_str)
                                if num_match:
                                    msg_count = int(num_match.group(1))
                except imaplib.IMAP4.error:
                    # Folder might not support STATUS (e.g., parent folders like [Gmail])
                    msg_count = None
                except Exception as status_err:
                    # Other errors - mark as unknown
                    msg_count = None
                
                # Add to list (None will be displayed as "N/A")
                folder_counts.append((folder_name, msg_count))
            
            # Sort by message count (descending) for better visibility
            # Put None values at the end
            folder_counts.sort(key=lambda x: (x[1] is None, x[1] if x[1] is not None else 0), reverse=True)
            
            # Display the counts
            print(f"{'Folder Name':<40} {'Messages':>10}")
            print("-" * 70)
            total_messages = 0
            folders_with_counts = 0
            for folder_name, count in folder_counts:
                if count is not None:
                    print(f"{folder_name:<40} {count:>10,}")
                    total_messages += count
                    folders_with_counts += 1
                else:
                    print(f"{folder_name:<40} {'N/A':>10}")
            print("-" * 70)
            if folders_with_counts > 0:
                print(f"{'TOTAL (countable folders)':<40} {total_messages:>10,}")
            print(f"{'Folders with message counts':<40} {folders_with_counts:>10}")
            print("="*70)
            print()
            
            # Mark initial listing as complete
            try:
                with open(initial_listing_flag_file, 'w') as f:
                    f.write(f"Initial folder listing completed: {datetime.datetime.now().isoformat()}\n")
            except Exception as e:
                print(f"Warning: Could not create initial listing flag file: {e}")
        
        # CRITICAL SAFETY: readonly=True ensures NO deletion/modification possible
        # ABSOLUTE REQUIREMENT: ONLY access spam folder, NEVER INBOX
        # Build list of folders to try: start with candidates from LIST, then fallback to common names
        spam_folders = []
        
        # First, try the exact folder names we found from LIST
        for candidate in spam_candidate_names:
            if candidate not in spam_folders:
                spam_folders.append(candidate)
        
        # Then add common Gmail spam folder names (if not already added)
        common_names = [
            '[Gmail]/Spam',      # Most common Gmail format
            '"[Gmail]/Spam"',    # Quoted version
            '[Google Mail]/Spam', # Alternative Gmail folder name
            'Spam',              # Simple name (some IMAP clients)
            '[Gmail]/Junk',      # Some accounts use Junk
            'Junk'               # Simple Junk folder name
        ]
        for name in common_names:
            if name not in spam_folders:
                spam_folders.append(name)
        
        mailbox_selected = False
        selected_folder = None
        
        # FORBIDDEN: Never try INBOX - this is a hard requirement
        FORBIDDEN_FOLDERS = ['INBOX', 'inbox', 'Inbox', '"INBOX"']
        
        for folder_name in spam_folders:
            # Safety check: never allow INBOX
            if folder_name.upper().replace('"', '').replace("'", '') == 'INBOX':
                print(f"SECURITY ERROR: Attempted to access forbidden folder: {folder_name}")
                raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Cannot access {folder_name}. Only spam folders are allowed.")
            
            try:
                status, data = mail.select(folder_name, readonly=True)
                if status == 'OK':
                    # CRITICAL SAFETY CHECK 1: Verify we did NOT select INBOX
                    normalized_name = folder_name.upper().replace('"', '').replace("'", '').replace('\\', '/').replace('[', '').replace(']', '')
                    if normalized_name == 'INBOX' or normalized_name.endswith('/INBOX') or normalized_name.startswith('INBOX/'):
                        print(f"SECURITY ERROR: Selected forbidden folder: {folder_name}")
                        mail.close()
                        raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Cannot access {folder_name}. Only spam/junk folders are allowed.")
                    
                    # CRITICAL SAFETY CHECK 2: Verify folder name contains SPAM or JUNK
                    if 'SPAM' not in normalized_name and 'JUNK' not in normalized_name:
                        print(f"SECURITY ERROR: Selected folder '{folder_name}' does not contain 'SPAM' or 'JUNK'!")
                        mail.close()
                        raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Folder '{folder_name}' is not a spam/junk folder. Access denied.")
                    
                    selected_folder = folder_name
                    mailbox_selected = True
                    
                    # Get message count to verify this is the right folder
                    msg_count = 0
                    try:
                        status, message_count = mail.status(folder_name, "(MESSAGES)")
                        if status == 'OK' and message_count:
                            # Parse message count from response like: (MESSAGES 123)
                            count_match = re.search(r'MESSAGES\s+(\d+)', str(message_count[0]))
                            if count_match:
                                msg_count = int(count_match.group(1))
                                print(f"Successfully selected SPAM/JUNK mailbox: {folder_name} ({msg_count} messages)")
                            else:
                                print(f"Successfully selected SPAM/JUNK mailbox: {folder_name}")
                        else:
                            print(f"Successfully selected SPAM/JUNK mailbox: {folder_name}")
                    except Exception as status_err:
                        print(f"Successfully selected SPAM/JUNK mailbox: {folder_name} (could not get message count: {status_err})")
                    
                    # Warn if folder is empty - might be the wrong folder
                    if msg_count == 0:
                        print(f"WARNING: Selected folder '{folder_name}' has 0 messages. This might not be the correct spam folder.")
                        print("If you know there is spam, please check the folder list above and verify the correct folder name.")
                    
                    break
                else:
                    print(f"Failed to select '{folder_name}': status={status}")
            except imaplib.IMAP4.error as select_err:
                # Check if error is because we tried to access INBOX
                if 'INBOX' in str(select_err).upper() or folder_name.upper().replace('"', '').replace("'", '') == 'INBOX':
                    print(f"SECURITY ERROR: Attempted to access forbidden folder: {folder_name}")
                    raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Cannot access {folder_name}. Only spam folders are allowed.")
                print(f"Error selecting '{folder_name}': {select_err}")
                continue
        
        if not mailbox_selected:
            print("=" * 70)
            print("CRITICAL ERROR: Could not select SPAM mailbox!")
            print("=" * 70)
            print("The script will NOT proceed without a valid spam folder selection.")
            print("Listing ALL available mailboxes to help identify the correct spam folder name...")
            print()
            try:
                status, folders = mail.list()
                if status == 'OK':
                    print("ALL available mailboxes:")
                    all_folders = []
                    spam_candidates_found = []
                    for folder in folders:
                        folder_str = ""
                        if isinstance(folder, bytes):
                            folder_str = folder.decode('utf-8', errors='ignore')
                        elif isinstance(folder, tuple):
                            decoded_parts = []
                            for part in folder:
                                if isinstance(part, bytes):
                                    decoded_parts.append(part.decode('utf-8', errors='ignore'))
                                else:
                                    decoded_parts.append(str(part))
                            folder_str = ' '.join(decoded_parts)
                        else:
                            folder_str = str(folder)
                        
                        all_folders.append(folder_str)
                        folder_upper = folder_str.upper()
                        # Identify spam/junk folders
                        if ('SPAM' in folder_upper or 'JUNK' in folder_upper) and 'INBOX' not in folder_upper:
                            spam_candidates_found.append(folder_str)
                    
                    # Print all folders
                    for folder in all_folders:
                        print(f"  {folder}")
                    
                    if spam_candidates_found:
                        print()
                        print("Spam/Junk folder candidates found:")
                        for candidate in spam_candidates_found:
                            print(f"  - {candidate}")
                        print()
                        print("Please check the exact folder name and update the code if needed.")
                    else:
                        print()
                        print("WARNING: No folders containing 'SPAM' or 'JUNK' were found in the list above.")
                        print("Please manually identify the spam folder from the list above.")
            except Exception as list_err:
                print(f"Could not list mailboxes: {list_err}")
            
            print("=" * 70)
            if mail:
                try: 
                    mail.close()
                    mail.logout()
                except: pass
            raise imaplib.IMAP4.error("SECURITY: Failed to select spam mailbox. Script aborted to prevent INBOX access.")
        
        
        # FINAL SAFETY CHECK: Verify selected folder is SPAM or JUNK ONLY
        if selected_folder:
            normalized = selected_folder.upper().replace('"', '').replace("'", '').replace('\\', '/').replace('[', '').replace(']', '')
            
            # Check 1: Must NOT be INBOX
            if normalized == 'INBOX' or normalized.endswith('/INBOX') or normalized.startswith('INBOX/'):
                print(f"SECURITY ERROR: Selected folder '{selected_folder}' appears to be INBOX!")
                mail.close()
                mail.logout()
                raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Selected folder '{selected_folder}' is forbidden. Only spam/junk folders allowed.")
            
            # Check 2: MUST contain SPAM or JUNK
            if 'SPAM' not in normalized and 'JUNK' not in normalized:
                print(f"SECURITY ERROR: Selected folder '{selected_folder}' does not contain 'SPAM' or 'JUNK'!")
                mail.close()
                mail.logout()
                raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Selected folder '{selected_folder}' is not a spam/junk folder. Access denied.")
        
        # Time window: Based on SPAM_SEARCH_WINDOW_HOURS
        # IMAP SINCE uses date format, so we use the date from hours_ago
        hours_ago = datetime.datetime.now() - datetime.timedelta(hours=SPAM_SEARCH_WINDOW_HOURS)
        date_since = hours_ago.strftime("%d-%b-%Y")
        
        # Format time window for display
        if SPAM_SEARCH_WINDOW_HOURS < 1:
            time_window_str = f"{int(SPAM_SEARCH_WINDOW_HOURS * 60)} minutes"
        elif SPAM_SEARCH_WINDOW_HOURS == 1:
            time_window_str = "1 hour"
        else:
            time_window_str = f"{SPAM_SEARCH_WINDOW_HOURS} hours"
        
        print(f"Searching for spam received since {date_since} (past {time_window_str})...")
        print(f"Current time: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Search window start: {hours_ago.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Fetch list of IDs - use SINCE for received date (not SENTSINCE which is sent date)
        # Note: IMAP SINCE is date-based, not time-based, so for very small windows (< 1 day),
        # we may get emails from earlier in the same day, but we'll filter by actual received time later if needed
        try:
            status, messages = mail.search(None, f'(SINCE "{date_since}")')
            if status != 'OK':
                print(f"Search failed with status: {status}")
                if messages and len(messages) > 0:
                    print(f"Search response: {messages}")
                if mail:
                    try: mail.logout()
                    except: pass
                return
        except Exception as search_err:
            print(f"Error executing IMAP search: {search_err}")
            if mail:
                try: mail.logout()
                except: pass
            return
            
        email_ids = messages[0].split() if messages and len(messages) > 0 and messages[0] else []
        
        # Log search results
        if email_ids:
            print(f"IMAP search returned {len(email_ids)} message ID(s)")
        else:
            print("IMAP search returned 0 messages")
        
        if not email_ids:
            if SPAM_SEARCH_WINDOW_HOURS < 1:
                hours_str = f"{int(SPAM_SEARCH_WINDOW_HOURS * 60)} minutes"
            elif SPAM_SEARCH_WINDOW_HOURS == 1:
                hours_str = "1 hour"
            else:
                hours_str = f"{SPAM_SEARCH_WINDOW_HOURS} hours"
            print(f"No spam found in the past {hours_str}.")
            if mail:
                try: mail.logout()
                except: pass
            return

        print(f"\nFound {len(email_ids)} candidate messages.")
        
        # Pre-fetch Headers to calculate stats BEFORE downloading bodies
        print("Analyzing headers...")
        for e_id in email_ids:
            try:
                # Fetch size and header separately for more reliable parsing
                # First get the size
                res_size, data_size = mail.fetch(e_id, '(RFC822.SIZE)')
                size = 0
                if res_size == 'OK' and data_size:
                    for item in data_size:
                        if isinstance(item, tuple) and len(item) >= 2:
                            size_str = item[1].decode() if isinstance(item[1], bytes) else str(item[1])
                            size_match = re.search(r'RFC822\.SIZE\s+(\d+)', size_str)
                            if size_match:
                                size = int(size_match.group(1))
                                break
                            # Try alternative format
                            size_match = re.search(r'\(RFC822\.SIZE\s+(\d+)\)', size_str)
                            if size_match:
                                size = int(size_match.group(1))
                                break
                
                # If size still not found, use default
                if size == 0:
                    size = 1024  # Default 1KB if we can't determine
                
                # Now get the header
                res_header, data_header = mail.fetch(e_id, '(BODY.PEEK[HEADER])')
                raw_header = None
                if res_header == 'OK' and data_header:
                    for item in data_header:
                        if isinstance(item, tuple) and len(item) >= 2:
                            raw_header = item[1]
                            if isinstance(raw_header, bytes):
                                break
                            elif isinstance(raw_header, str):
                                raw_header = raw_header.encode()
                                break
                
                if raw_header is None:
                    print(f"Warning: Could not fetch header for message {e_id.decode()}")
                    continue
                
                total_size_bytes += size
                
                msg_header = email.message_from_bytes(raw_header)
                subject_raw = msg_header['Subject']
                # Decode the subject if it's encoded
                subject = decode_email_header(subject_raw) if subject_raw else "(No Subject)"
                date_str = msg_header['Date']
                
                # Store for later
                spam_candidates.append({
                    'id': e_id,
                    'subject': subject,
                    'date': date_str,
                    'size': size
                })
                # Display subject safely (handle unicode/emoji for Windows console)
                display_subject = safe_print_subject(subject, 50)
                print(f" - Identified: {display_subject}... ({int(size/1024)} KB)")
            except Exception as e:
                print(f"Warning: Error processing message {e_id.decode()}: {e}")
                import traceback
                traceback.print_exc()
                continue

    except imaplib.IMAP4.error as e:
        print(f"IMAP Error: {e}")
        print("\nAUTHENTICATION TROUBLESHOOTING:")
        print("=" * 50)
        print("Option 1: Try your regular Gmail password")
        print("  - Some accounts can still use regular passwords")
        print("  - If this doesn't work, proceed to Option 2")
        print("\nOption 2: Use Gmail App Password (requires 2FA)")
        print("  - Go to: https://myaccount.google.com/security")
        print("  - Enable '2-Step Verification' first")
        print("  - Then go to: https://myaccount.google.com/apppasswords")
        print("  - Generate an App Password for 'Mail'")
        print("  - Use that 16-character password (no spaces) in APP_PASS")
        print("\nOption 3: Enable 'Less Secure App Access' (if available)")
        print("  - Go to: https://myaccount.google.com/lesssecureapps")
        print("  - Note: This option is deprecated and may not be available")
        print("=" * 50)
        if mail:
            try: mail.logout()
            except: pass
        return
    except Exception as e:
        print(f"Error during connection/search: {e}")
        if mail:
            try: mail.logout()
            except: pass
        return

    # ---------------------------------------------------------
    # PHASE 2: DOWNLOAD (Non-interactive - always proceed)
    # ---------------------------------------------------------
    if not spam_candidates:
        print("No spam candidates to download.")
        if mail:
            try: mail.logout()
            except: pass
        return
        
    print("\n" + "="*40)
    print(f"Identified {len(spam_candidates)} emails. Total Size: {get_size_str(total_size_bytes)}.")
    print("Proceeding with download (non-interactive mode)...")

    # ---------------------------------------------------------
    # PHASE 3: CREATE FOLDER AND DOWNLOAD
    # ---------------------------------------------------------
    # Construct Folder Name: [Account]__YYYY-MM-DD__HHMMSS__Count__Size
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d__%H%M%S")
    account_tag = f"[{GMAIL_ACCOUNT.split('@')[0]}]"
    folder_name = f"{account_tag}__{timestamp}__{len(spam_candidates)}__{get_size_str(total_size_bytes)}"
    
    # Ensure base downloads directory exists
    if not os.path.exists(BASE_DIRECTORY):
        os.makedirs(BASE_DIRECTORY)
    
    download_path = os.path.join(BASE_DIRECTORY, folder_name)
    
    if not os.path.exists(download_path):
        os.makedirs(download_path)
        
    print(f"\nCreated Directory: {download_path}")
    print("Downloading messages...")
    
    downloaded_files = []
    timestamps = []

    for item in spam_candidates:
        try:
            # Fetch Full Body (Safe PEEK)
            res, data = mail.fetch(item['id'], '(BODY.PEEK[])')
            if res != 'OK' or not data:
                print(f"Warning: Could not fetch body for message {item['id'].decode()}")
                continue
            
            # Extract raw email from fetch response
            raw_email = None
            for response_item in data:
                if isinstance(response_item, tuple) and len(response_item) >= 2:
                    raw_email = response_item[1]
                    if isinstance(raw_email, bytes):
                        break
                    elif isinstance(raw_email, str):
                        raw_email = raw_email.encode('utf-8')
                        break
            
            if raw_email is None:
                print(f"Warning: Failed to fetch raw email for message ID {item['id'].decode()}. Skipping.")
                continue
            
            # Parse timestamp if available
            if item['date']:
                try:
                    dt = parsedate_to_datetime(item['date'])
                    if dt: 
                        timestamps.append(dt)
                except: 
                    pass

            # Save file
            clean_sub = sanitize_filename(item['subject'])
            # Shorten filename if too long
            clean_sub = (clean_sub[:50] + '..') if len(clean_sub) > 50 else clean_sub
            filename = f"{clean_sub}_{item['id'].decode()}.eml"
            filepath = os.path.join(download_path, filename)
            
            try:
                with open(filepath, 'wb') as f:
                    f.write(raw_email)
                downloaded_files.append(filepath)
                print(f"Saved: {filename}")
            except Exception as file_err:
                print(f"Error saving message {item['id'].decode()} to file: {file_err}")
                
        except Exception as e:
            print(f"Error downloading message {item['id'].decode()}: {e}")
            import traceback
            traceback.print_exc()
            continue

    # Close IMAP connection (Mission Accomplished)
    mail.close()
    mail.logout()

    # ---------------------------------------------------------
    # PHASE 4: REPORT STATISTICS
    # ---------------------------------------------------------
    print("\n" + "="*40)
    print("DOWNLOAD STATISTICS")
    print("="*40)
    print(f"Source Folder:     [Gmail]/Spam (READ-ONLY ACCESS)")
    print(f"Total Emails:      {len(downloaded_files)}")
    print(f"Total Size:        {get_size_str(total_size_bytes)}")
    
    if timestamps:
        timestamps.sort()
        earliest = timestamps[0]
        latest = timestamps[-1]
        span = latest - earliest
        print(f"Earliest Email:    {earliest}")
        print(f"Latest Email:      {latest}")
        print(f"Time Span:         {span}")
    else:
        print("Time Span:         Could not calculate (invalid date headers)")

    # ---------------------------------------------------------
    # PHASE 5: FORWARD TO SPAMCOP (With first-run confirmation)
    # ---------------------------------------------------------
    if len(downloaded_files) > 0:
        # Check if this is the first run (first-time confirmation required)
        # Skip first-run confirmation in simulation mode since we're not actually sending
        first_run_flag_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.spamcop_first_run_complete')
        is_first_run = not os.path.exists(first_run_flag_file) and not SIMULATION_MODE
        
        if is_first_run:
            print("\n" + "="*70)
            print("FIRST RUN - VERIFICATION REQUIRED")
            print("="*70)
            print(msg.FIRST_RUN_HEADER)  # type: ignore
            print("\n" + "-"*70)
            print("CRITICAL WARNING:")
            print("-"*70)
            print(msg.FIRST_RUN_WARNING)  # type: ignore
            print("\n" + "-"*70)
            print("DOWNLOADED EMAILS LIST:")
            print("-"*70)
            print(f"Total emails downloaded: {len(downloaded_files)}")
            print(f"Download location: {download_path}")
            print("\nEmail subjects (please verify these match your Gmail Spam folder):")
            print()
            
            for idx, item in enumerate(spam_candidates, 1):
                display_subject = safe_print_subject(item['subject'], 60)
                print(f"  {idx}. {display_subject}")
                if item['date']:
                    print(f"     Date: {item['date']}")
            
            print("\n" + "="*70)
            print("VERIFICATION STEPS:")
            print("="*70)
            print(msg.FIRST_RUN_VERIFICATION_STEPS)  # type: ignore
            print("="*70)
            print()
            
            while True:
                response = input("Have you verified that ALL downloaded emails are spam? (yes/no): ").strip().lower()
                if response in ['yes', 'y']:
                    print("\n✓ Confirmation received. Proceeding with forwarding to SpamCop...")
                    # Create flag file to mark first run as complete
                    try:
                        with open(first_run_flag_file, 'w') as f:
                            f.write(f"First run completed: {datetime.datetime.now().isoformat()}\n")
                    except Exception as e:
                        print(f"Warning: Could not create first-run flag file: {e}")
                    break
                elif response in ['no', 'n']:
                    print("\n" + "="*70)
                    print("FORWARDING CANCELLED")
                    print("="*70)
                    print(msg.FIRST_RUN_CANCELLED.format(download_path=download_path))  # type: ignore
                    print("="*70)
                    print("\n--- ITERATION COMPLETE (NO EMAILS FORWARDED) ---")
                    return
                else:
                    print("Please enter 'yes' or 'no'.")
        
        # Normal operation (after first run)
        if not is_first_run:
            if SIMULATION_MODE:
                print("\n" + "-"*40)
                print(f"SIMULATION MODE: Would bundle {len(downloaded_files)} files and send to SpamCop...")
            else:
                print("\n" + "-"*40)
                print(f"Bundling {len(downloaded_files)} files and sending to SpamCop...")
        
        # Check simulation mode
        if SIMULATION_MODE:
            print("\n" + "="*70)
            print("SIMULATION MODE ENABLED - Email will NOT be sent to SpamCop")
            print("="*70)
            print(f"Would send email:")
            print(f"  From: {GMAIL_ACCOUNT}")
            print(f"  To: {SPAMCOP_ADDRESS}")
            print(f"  Subject: Spam Report: {len(downloaded_files)} messages")
            print(f"  Attachments: {len(downloaded_files)} EML files")
            print()
            for filepath in downloaded_files:
                filename = os.path.basename(filepath)
                file_size = os.path.getsize(filepath)
                print(f"    - {filename} ({get_size_str(file_size)})")
            print()
            print("All steps completed successfully (connect, search, download, save).")
            print("Set SIMULATION_MODE = False in config.py to enable actual forwarding.")
            print("="*70)
        else:
            # Actual forwarding mode
            print(f"\nConnecting to Gmail SMTP ({SMTP_SERVER})...")
            try:
                # Construct the Email
                msg = MIMEMultipart()
                msg['From'] = GMAIL_ACCOUNT
                msg['To'] = SPAMCOP_ADDRESS
                msg['Subject'] = f"Spam Report: {len(downloaded_files)} messages"
                
                # Attach files
                print("Attaching files...")
                for filepath in downloaded_files:
                    filename = os.path.basename(filepath)
                    with open(filepath, 'rb') as f:
                        part = MIMEBase('message', 'rfc822')
                        part.set_payload(f.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                        msg.attach(part)
                
                # Send using Gmail SMTP
                print("Sending data...")
                server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
                server.starttls()
                server.login(GMAIL_ACCOUNT, APP_PASS)
                server.send_message(msg)
                server.quit()
                
                print("\nSUCCESS: Report sent to SpamCop.")
                
            except Exception as e:
                print(f"\nFAILED to send email: {e}")
                print("The files are still safe in your local folder.")
                import traceback
                traceback.print_exc()
    else:
        print("\nSkipping send. Files remain in your local folder.")

    print("\n--- ITERATION COMPLETE ---")

def run_spam_processor():
    """Main loop that runs spam processing continuously"""
    # Validate configuration
    try:
        validate_loop_frequency(LOOP_FREQUENCY_HOURS)
        validate_search_window(SPAM_SEARCH_WINDOW_HOURS)
    except ValueError as e:
        print(f"CONFIGURATION ERROR: {e}")
        if "LOOP_FREQUENCY_HOURS" in str(e):
            print("Please fix LOOP_FREQUENCY_HOURS in the configuration section.")
        elif "SPAM_SEARCH_WINDOW_HOURS" in str(e):
            print("Please fix SPAM_SEARCH_WINDOW_HOURS in the configuration section.")
        return
    
    # Format display strings
    if LOOP_FREQUENCY_HOURS < 1:
        freq_str = f"{int(LOOP_FREQUENCY_HOURS * 60)} minutes"
    elif LOOP_FREQUENCY_HOURS == 1:
        freq_str = "1 hour"
    else:
        freq_str = f"{LOOP_FREQUENCY_HOURS} hours"
    
    if SPAM_SEARCH_WINDOW_HOURS < 1:
        window_str = f"{int(SPAM_SEARCH_WINDOW_HOURS * 60)} minutes"
    elif SPAM_SEARCH_WINDOW_HOURS == 1:
        window_str = "1 hour"
    else:
        window_str = f"{SPAM_SEARCH_WINDOW_HOURS} hours"
    
    print("="*60)
    print("SPAM PROCESSOR - AUTONOMOUS MODE")
    print("="*60)
    print(f"Frequency: Every {freq_str}")
    print(f"Search window: Past {window_str}")
    print("Mode: Non-interactive (automatic download and forward)")
    if SIMULATION_MODE:
        print("SIMULATION MODE: ENABLED - Emails will NOT be forwarded to SpamCop")
    else:
        print("SIMULATION MODE: DISABLED - Emails WILL be forwarded to SpamCop")
    print("="*60)
    print("Starting first iteration immediately...\n")
    
    iteration_count = 0
    last_run_time = None
    
    # Start processing immediately - no delay
    while True:
        iteration_count += 1
        current_time = datetime.datetime.now()
        
        print(f"\n{'='*60}")
        print(f"ITERATION #{iteration_count} - {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
        if last_run_time:
            time_since_last = current_time - last_run_time
            hours = int(time_since_last.total_seconds() // 3600)
            minutes = int((time_since_last.total_seconds() % 3600) // 60)
            print(f"Last run: {last_run_time.strftime('%Y-%m-%d %H:%M:%S')} ({hours}h {minutes}m ago)")
        else:
            print("Last run: (First iteration)")
        print(f"{'='*60}\n")
        
        try:
            process_spam_iteration()
        except KeyboardInterrupt:
            print("\n\nReceived interrupt signal. Shutting down gracefully...")
            break
        except Exception as e:
            print(f"\nERROR in iteration #{iteration_count}: {e}")
            print("Continuing to next iteration...")
            import traceback
            traceback.print_exc()
        
        # Record completion time
        last_run_time = datetime.datetime.now()
        
        # Calculate sleep time
        sleep_seconds = LOOP_FREQUENCY_HOURS * 3600
        sleep_hours = LOOP_FREQUENCY_HOURS
        next_run = last_run_time + datetime.timedelta(seconds=sleep_seconds)
        
        console.print(f"\n{'='*60}")
        console.print(f"Waiting {sleep_hours} hour(s) until next iteration...")
        console.print(f"Next run scheduled for: {next_run.strftime('%Y-%m-%d %H:%M:%S')}")
        console.print(f"{'='*60}\n")
        
        # Countdown timer with progress bar using rich library
        try:
            with Progress(
                SpinnerColumn(),
                TextColumn("[progress.description]{task.description}"),
                BarColumn(),
                TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
                TimeRemainingColumn(),
                TimeElapsedColumn(),
                console=console,
                transient=False
            ) as progress:
                task = progress.add_task(
                    f"[cyan]Waiting until {next_run.strftime('%Y-%m-%d %H:%M:%S')}",
                    total=sleep_seconds
                )
                
                remaining_seconds = sleep_seconds
                update_interval = 1  # Update every second
                
                while remaining_seconds > 0:
                    # Update progress
                    elapsed = sleep_seconds - remaining_seconds
                    progress.update(task, completed=elapsed)
                    
                    # Sleep in smaller increments for responsive countdown
                    sleep_chunk = min(update_interval, remaining_seconds)
                    time.sleep(sleep_chunk)
                    remaining_seconds -= sleep_chunk
                
                # Complete the progress bar
                progress.update(task, completed=sleep_seconds)
            
            console.print("[green]Countdown complete! Starting next iteration...\n")
        except KeyboardInterrupt:
            console.print("\n[yellow]Received interrupt signal during sleep. Shutting down gracefully...")
            break

if __name__ == "__main__":
    run_spam_processor()