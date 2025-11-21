# Standard library imports
import imaplib
import smtplib
import email
import os
import sys
import datetime
import time
import re
import atexit
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.utils import parsedate_to_datetime, parseaddr
from email.header import decode_header

# Third-party imports
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeRemainingColumn, TimeElapsedColumn  # type: ignore
from rich.console import Console  # type: ignore

# Local imports
import messages as msg  # type: ignore

# Initialize rich console
console = Console()

# ==========================================
#              CONSTANTS
# ==========================================

IMAP_SERVER = 'imap.gmail.com'
FORBIDDEN_FOLDERS = [
    'INBOX', 'inbox', 'Inbox', '"INBOX"',
    '[Google Mail]/All Mail', '"[Google Mail]/All Mail"',
    '[Google Mail]/Important', '"[Google Mail]/Important"',
    '[Google Mail]/Sent Mail', '"[Google Mail]/Sent Mail"',
    '[Google Mail]/Starred', '"[Google Mail]/Starred"',
    '[Google Mail]/Drafts', '"[Google Mail]/Drafts"',
    '[Gmail]/All Mail', '"[Gmail]/All Mail"',
    '[Gmail]/Important', '"[Gmail]/Important"',
    '[Gmail]/Sent Mail', '"[Gmail]/Sent Mail"',
    '[Gmail]/Starred', '"[Gmail]/Starred"',
    '[Gmail]/Drafts', '"[Gmail]/Drafts"'
]
COMMON_SPAM_FOLDER_NAMES = [
    '[Google Mail]/Spam', # Default suggestion - Alternative Gmail folder name
    '[Gmail]/Spam',      # Most common Gmail format
    '"[Gmail]/Spam"',    # Quoted version
    'Spam',              # Simple name (some IMAP clients)
    '[Gmail]/Junk',      # Some accounts use Junk
    'Junk'               # Simple Junk folder name
]
FIRST_RUN_FLAG_FILE = '.spamcop_first_run_complete'
DEFAULT_SIZE_KB = 1024  # Default 1KB if size can't be determined
LOG_FILE = 'spamcop_forwarder.log'  # Logfile that appends all output
SENT_UIDS_FILE = '.spamcop_sent_uids.txt'  # File to track UIDs of emails already sent to SpamCop
SPAM_FOLDER_CACHE_FILE = '.spamcop_folder_cache.txt'  # File to cache the selected spam folder name

# ==========================================
#              LOGGING SETUP
# ==========================================

class TeeOutput:
    """Class that writes to both a terminal stream and a logfile simultaneously"""
    def __init__(self, terminal_stream, logfile_handle):
        self.terminal = terminal_stream
        self.logfile = logfile_handle
    
    def write(self, message):
        """Writes to both terminal and logfile"""
        # Write to terminal
        try:
            self.terminal.write(message)
            self.terminal.flush()
        except Exception:
            # Fallback for terminal encoding issues
            try:
                encoding = getattr(self.terminal, 'encoding', 'utf-8') or 'utf-8'
                self.terminal.write(message.encode(encoding, errors='replace').decode(encoding))
                self.terminal.flush()
            except Exception:
                pass

        # Write to logfile
        if self.logfile:
            try:
                # Strip ANSI codes before writing to log file
                clean_message = self._strip_ansi(message)
                self.logfile.write(clean_message)
                self.logfile.flush()
            except Exception:
                pass

    def flush(self):
        """Flushes both terminal and logfile"""
        try:
            self.terminal.flush()
        except Exception:
            pass
        if self.logfile:
            try:
                self.logfile.flush()
            except Exception:
                pass
    
    def _strip_ansi(self, text):
        """Strips ANSI escape codes from text"""
        ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
        return ansi_escape.sub('', text)
    
    def isatty(self):
        """Delegate isatty to the terminal stream"""
        return getattr(self.terminal, 'isatty', lambda: False)()

def is_initial_run_internal():
    """Determines if this is the initial run by checking logfile and download history"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(script_dir, LOG_FILE)
    
    # Check if logfile exists
    if not os.path.exists(log_path):
        return True
    
    # Check if logfile is empty or very small
    try:
        if os.path.getsize(log_path) < 100:
            return True
    except Exception:
        pass
    
    # Check logfile content for STRONG evidence of successful previous runs
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            content = f.read()
            run_ended_count = content.count('RUN ENDED:')
            
            has_actual_processing = (
                ('DOWNLOAD STATISTICS' in content and 'Total Emails:' in content and 'Total Emails:      0' not in content) or
                ('SUCCESS: Report sent to SpamCop' in content) or
                ('Would send email:' in content and 'SIMULATION MODE' in content and 'Total Emails:' in content)
            )
            
            if run_ended_count >= 10 and has_actual_processing:
                return False
    except Exception:
        return True
    
    # Check download directory if possible
    try:
        import sys
        current_module = sys.modules[__name__]
        if hasattr(current_module, 'BASE_DIRECTORY'):
            base_dir = getattr(current_module, 'BASE_DIRECTORY')
            if base_dir and os.path.exists(base_dir):
                download_items = os.listdir(base_dir)
                if download_items:
                    return False
    except Exception:
        pass
    
    return True

def setup_logging():
    """Sets up logging to both console and logfile, capturing stdout and stderr"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    log_path = os.path.join(script_dir, LOG_FILE)
    
    # Determine if this is the initial run using internal checks
    is_initial_run = is_initial_run_internal()
    
    # Open logfile
    mode = 'w' if is_initial_run else 'a'
    try:
        log_file = open(log_path, mode, encoding='utf-8', buffering=1)
    except Exception as e:
        print(f"Warning: Could not open logfile '{log_path}': {e}", file=sys.stderr)
        log_file = None
    
    # Replace stdout and stderr with TeeOutput
    stdout_tee = TeeOutput(sys.stdout, log_file)
    stderr_tee = TeeOutput(sys.stderr, log_file)
    
    sys.stdout = stdout_tee
    sys.stderr = stderr_tee
    
    # Write header
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if not is_initial_run:
        print(f"\n{'='*80}")
    print(f"NEW RUN STARTED: {timestamp}")
    print(f"{'='*80}\n")
    sys.stdout.flush()
    
    return log_file, is_initial_run

# Set up logging before any other output
_log_file, _is_initial_run_global = setup_logging()

# Register cleanup function to close logfile on exit
_log_cleanup_done = False

def cleanup_logging():
    """Closes the logfile when script exits"""
    global _log_cleanup_done
    if _log_cleanup_done:
        return
    _log_cleanup_done = True
    
    # Restore original streams if possible (wrapped in TeeOutput)
    if hasattr(sys.stdout, 'terminal'):
        sys.stdout = sys.stdout.terminal
    if hasattr(sys.stderr, 'terminal'):
        sys.stderr = sys.stderr.terminal
    
    if _log_file:
        try:
            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            _log_file.write(f"\n{'='*80}\n")
            _log_file.write(f"RUN ENDED: {timestamp}\n")
            _log_file.write(f"{'='*80}\n\n")
            _log_file.flush()
            _log_file.close()
        except Exception:
            pass

atexit.register(cleanup_logging)

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

# 8. FOLDER PREVIEW
#    When True: Displays all folders with message counts before selecting spam folder
#    When False: Skips folder preview and only looks for the specified spam folder
#    Default: False (preview disabled) - set to True to see all folders
PREVIEW_ALL_FOLDERS = False  # Set to True to enable folder preview

# 9. SPAM FOLDER NAME
#    The name of the spam folder to use when PREVIEW_ALL_FOLDERS is False
#    Default: '[Google Mail]/Spam' (standard Gmail spam folder)
SPAM_FOLDER_NAME = '[Google Mail]/Spam'  # Spam folder name (use exact IMAP folder name)

# 10. EXCLUSION LISTS
#    Emails matching these criteria will NOT be sent to SpamCop
#    EXCLUDED_SENDERS: List of email addresses or domains to exclude (case-insensitive)
#    EXCLUDED_SUBJECT_KEYWORDS: List of keywords/phrases in subject lines and body to exclude (case-insensitive)
#      - Supports both single keywords and multi-word phrases
#      - Keywords/phrases are searched in BOTH subject line AND email body
#      - Cannot conflict with FORCE_INCLUDE_KEYWORDS (script will error if conflicts detected)
#    Examples:
#    EXCLUDED_SENDERS = ['noreply@example.com', '@newsletters.com']  # Exclude specific senders or domains
#    EXCLUDED_SUBJECT_KEYWORDS = ['newsletter', 'unsubscribe', 'marketing update', 'policy change']  # Exclude emails with these keywords/phrases
EXCLUDED_SENDERS = []  # List of email addresses or domains (e.g., ['sender@example.com', '@domain.com'])
EXCLUDED_SUBJECT_KEYWORDS = []  # List of keywords/phrases (e.g., ['newsletter', 'unsubscribe', 'marketing update'])

# 11. FORCE INCLUDE KEYWORDS
#    Emails matching these keywords in subject or body will ALWAYS be sent to SpamCop
#    This overrides exclusion rules - if an email matches force-include keywords, it will be sent
#    even if it would normally be excluded
#    FORCE_INCLUDE_KEYWORDS: List of keywords/phrases to force include (case-insensitive)
#      - Supports both single keywords and multi-word phrases
#      - Keywords/phrases are searched ONLY in subject line and email body (NOT in sender email address)
#      - This prevents false matches when keywords appear in your own email address
#      - Any email with subject or body containing these keywords/phrases will be force-included
#      - This takes priority over exclusion rules
#      - Cannot conflict with EXCLUDED_SUBJECT_KEYWORDS (script will error if conflicts detected)
#    Examples:
#    FORCE_INCLUDE_KEYWORDS = ['phishing', 'scam', 'bitcoin', 'cryptocurrency', 'nigerian prince']
FORCE_INCLUDE_KEYWORDS = []  # List of keywords/phrases (e.g., ['phishing', 'scam', 'bitcoin'])
'''
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write(config_content)
        print(f"Created {config_path}")
        print("\nPlease edit the config file and replace the placeholders, then run the script again.")
        cleanup_logging()
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
            cleanup_logging()
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
            'SIMULATION_MODE': getattr(config, 'SIMULATION_MODE', True),
            'PREVIEW_ALL_FOLDERS': getattr(config, 'PREVIEW_ALL_FOLDERS', False),
            'SPAM_FOLDER_NAME': getattr(config, 'SPAM_FOLDER_NAME', '[Google Mail]/Spam'),
            'EXCLUDED_SENDERS': getattr(config, 'EXCLUDED_SENDERS', []),
            'EXCLUDED_SUBJECT_KEYWORDS': getattr(config, 'EXCLUDED_SUBJECT_KEYWORDS', []),
            'FORCE_INCLUDE_KEYWORDS': getattr(config, 'FORCE_INCLUDE_KEYWORDS', [])
        }
    except Exception as e:
        print(f"Error loading configuration: {e}")
        print(f"Please check that {config_path} exists and is valid Python code.")
        cleanup_logging()
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

def validate_keyword_conflicts():
    """Validates that there are no conflicts between exclusion and force-include keywords.
    Raises ValueError if conflicts are found.
    """
    conflicts = []
    
    # Normalize all keywords for comparison (lowercase, strip)
    excluded_keywords = [kw.lower().strip() for kw in EXCLUDED_SUBJECT_KEYWORDS if kw and kw.strip()]
    force_include_keywords = [kw.lower().strip() for kw in FORCE_INCLUDE_KEYWORDS if kw and kw.strip()]
    
    # Check for exact matches
    for excluded in excluded_keywords:
        if excluded in force_include_keywords:
            conflicts.append(f"Exact match: '{excluded}' appears in both EXCLUDED_SUBJECT_KEYWORDS and FORCE_INCLUDE_KEYWORDS")
    
    # Check for substring matches (one keyword contains another)
    for excluded in excluded_keywords:
        for force_include in force_include_keywords:
            # Check if excluded keyword is contained in force-include keyword
            if excluded in force_include and excluded != force_include:
                conflicts.append(f"Substring match: EXCLUDED keyword '{excluded}' is contained in FORCE_INCLUDE keyword '{force_include}'")
            # Check if force-include keyword is contained in excluded keyword
            elif force_include in excluded and excluded != force_include:
                conflicts.append(f"Substring match: FORCE_INCLUDE keyword '{force_include}' is contained in EXCLUDED keyword '{excluded}'")
    
    if conflicts:
        error_msg = "\n" + "=" * 70 + "\n"
        error_msg += "CONFIGURATION CONFLICT ERROR\n"
        error_msg += "=" * 70 + "\n"
        error_msg += "Conflicts detected between EXCLUDED_SUBJECT_KEYWORDS and FORCE_INCLUDE_KEYWORDS:\n\n"
        for i, conflict in enumerate(conflicts, 1):
            error_msg += f"  {i}. {conflict}\n"
        error_msg += "\n"
        error_msg += "Please resolve these conflicts by removing or modifying the conflicting keywords.\n"
        error_msg += "A keyword cannot be both excluded and force-included.\n"
        error_msg += "=" * 70 + "\n"
        raise ValueError(error_msg)
    
    return True

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
PREVIEW_ALL_FOLDERS = config['PREVIEW_ALL_FOLDERS']
SPAM_FOLDER_NAME = config['SPAM_FOLDER_NAME']
EXCLUDED_SENDERS = config['EXCLUDED_SENDERS']
EXCLUDED_SUBJECT_KEYWORDS = config['EXCLUDED_SUBJECT_KEYWORDS']
FORCE_INCLUDE_KEYWORDS = config['FORCE_INCLUDE_KEYWORDS']

# Validate keyword conflicts
try:
    validate_keyword_conflicts()
except ValueError as e:
    print(str(e))
    sys.stdout.flush()
    cleanup_logging()
    sys.exit(1)

# Convert BASE_DIRECTORY to absolute path if it's relative
if not os.path.isabs(BASE_DIRECTORY):
    BASE_DIRECTORY = os.path.join(os.path.dirname(os.path.abspath(__file__)), BASE_DIRECTORY)

# Ensure base directory exists
if not os.path.exists(BASE_DIRECTORY):
    try:
        os.makedirs(BASE_DIRECTORY)
        print(f"Created base directory: {BASE_DIRECTORY}")
    except Exception as e:
        print(f"Warning: Could not create base directory '{BASE_DIRECTORY}': {e}")
        print("The script may fail when trying to save downloaded emails.")
elif not os.path.isdir(BASE_DIRECTORY):
    print(f"Error: '{BASE_DIRECTORY}' exists but is not a directory!")
    cleanup_logging()
    sys.exit(1)

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

def format_hours_as_string(hours):
    """Converts hours to a human-readable string"""
    if hours < 1:
        return f"{int(hours * 60)} minutes"
    elif hours == 1:
        return "1 hour"
    else:
        return f"{int(hours)} hours"

def normalize_folder_name(folder_name):
    """Normalizes folder name for comparison (removes quotes, brackets, etc.)"""
    return folder_name.upper().replace('"', '').replace("'", '').replace('\\', '/').replace('[', '').replace(']', '')

def load_sent_uids():
    """Loads the set of UIDs that have already been sent to SpamCop"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    uids_file = os.path.join(script_dir, SENT_UIDS_FILE)
    sent_uids = set()
    
    if os.path.exists(uids_file):
        try:
            with open(uids_file, 'r', encoding='utf-8') as f:
                for line in f:
                    uid = line.strip()
                    if uid:
                        sent_uids.add(uid)
        except Exception as e:
            print(f"Warning: Could not load sent UIDs file: {e}")
            sys.stdout.flush()
    
    return sent_uids

def save_sent_uids(sent_uids):
    """Saves the set of UIDs that have been sent to SpamCop"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    uids_file = os.path.join(script_dir, SENT_UIDS_FILE)
    
    try:
        with open(uids_file, 'w', encoding='utf-8') as f:
            for uid in sorted(sent_uids):
                f.write(f"{uid}\n")
    except Exception as e:
        print(f"Warning: Could not save sent UIDs file: {e}")
        sys.stdout.flush()

def add_sent_uids(new_uids):
    """Adds new UIDs to the sent UIDs set and saves to file"""
    sent_uids = load_sent_uids()
    sent_uids.update(new_uids)
    save_sent_uids(sent_uids)

def load_spam_folder_cache():
    """Loads the cached spam folder name from file"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cache_file = os.path.join(script_dir, SPAM_FOLDER_CACHE_FILE)
    
    if os.path.exists(cache_file):
        try:
            with open(cache_file, 'r', encoding='utf-8') as f:
                folder_name = f.read().strip()
                if folder_name:
                    return folder_name
        except Exception as e:
            print(f"Warning: Could not load spam folder cache: {e}")
            sys.stdout.flush()
    
    return None

def save_spam_folder_cache(folder_name):
    """Saves the spam folder name to cache file"""
    script_dir = os.path.dirname(os.path.abspath(__file__))
    cache_file = os.path.join(script_dir, SPAM_FOLDER_CACHE_FILE)
    
    try:
        with open(cache_file, 'w', encoding='utf-8') as f:
            f.write(folder_name)
    except Exception as e:
        print(f"Warning: Could not save spam folder cache: {e}")
        sys.stdout.flush()

def quote_folder_name_for_imap(folder_name):
    """Quotes folder name for IMAP commands if it contains special characters.
    Returns the folder name properly quoted for IMAP commands.
    """
    if not folder_name:
        return folder_name
    
    # If already quoted, return as-is
    if folder_name.startswith('"') and folder_name.endswith('"'):
        return folder_name
    
    # Check if folder name contains special characters that require quoting
    # IMAP requires quoting for: spaces, brackets, and other special chars
    needs_quoting = (
        ' ' in folder_name or
        '[' in folder_name or
        ']' in folder_name or
        '(' in folder_name or
        ')' in folder_name or
        '{' in folder_name or
        '}' in folder_name or
        '"' in folder_name or
        '\\' in folder_name
    )
    
    if needs_quoting:
        # Escape any existing quotes and wrap in quotes
        escaped = folder_name.replace('\\', '\\\\').replace('"', '\\"')
        return f'"{escaped}"'
    
    return folder_name

def is_forbidden_folder(folder_name):
    """Checks if a folder is in the forbidden list (e.g., INBOX, All Mail, Important, etc.)"""
    # Check exact matches first
    if folder_name in FORBIDDEN_FOLDERS:
        return True
    
    # Normalize for comparison
    normalized = normalize_folder_name(folder_name)
    
    # Check INBOX variations
    if normalized == 'INBOX' or normalized.endswith('/INBOX') or normalized.startswith('INBOX/'):
        return True
    
    # Check forbidden Gmail/Google Mail folders
    forbidden_patterns = [
        'ALL MAIL',
        'IMPORTANT',
        'SENT MAIL',
        'STARRED',
        'DRAFTS'
    ]
    
    for pattern in forbidden_patterns:
        if pattern in normalized:
            return True
    
    return False

def is_spam_folder(folder_name):
    """Checks if a folder name indicates it's a spam/junk folder"""
    normalized = normalize_folder_name(folder_name)
    return 'SPAM' in normalized or 'JUNK' in normalized

def safe_logout(mail):
    """Safely logout from IMAP connection"""
    if mail:
        try:
            mail.logout()
        except Exception:
            pass

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

def parse_folder_from_list_response(folder):
    """Parses folder name from IMAP LIST response"""
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
    # Or: (\\HasChildren \\Noselect) "/" "[Gmail]"
    # Or: (\\HasNoChildren) "/" "[Gmail]/Spam"
    # Gmail format: (attributes) "delimiter" "folder_name" or (attributes) delimiter folder_name
    
    folder_name = None
    
    # Method 1: Look for quoted strings (most common format)
    # Find all quoted strings
    quoted_matches = re.findall(r'"([^"]+)"', folder_str)
    if quoted_matches:
        # The last quoted string is usually the folder name
        # But sometimes there are multiple (delimiter and folder)
        # For Gmail: (\\HasNoChildren) "/" "[Gmail]/Spam" -> folder is "[Gmail]/Spam"
        # For simple: (\\HasNoChildren) "/" "INBOX" -> folder is "INBOX"
        if len(quoted_matches) >= 2:
            # If multiple quoted strings, the last one is the folder name
            folder_name = quoted_matches[-1]
        else:
            folder_name = quoted_matches[0]
    
    # Method 2: If no quoted strings, try to extract from the end
    if not folder_name:
        # Split by spaces and take the last non-empty part
        parts = folder_str.split()
        if len(parts) > 0:
            # Skip attributes (parts in parentheses) and delimiter
            for part in reversed(parts):
                if part and part not in ['/', '\\'] and not part.startswith('(') and not part.endswith(')'):
                    folder_name = part
                    break
    
    # Method 3: If still nothing, try to extract anything that looks like a folder name
    if not folder_name:
        # Look for patterns like [Gmail]/Spam or INBOX
        pattern_match = re.search(r'([A-Za-z0-9\[\]/_-]+)', folder_str)
        if pattern_match:
            potential_name = pattern_match.group(1)
            # Skip common delimiters and attributes
            if potential_name not in ['/', '\\', 'HasNoChildren', 'HasChildren', 'Noselect']:
                folder_name = potential_name
    
    # Clean up the folder name
    if folder_name:
        # Remove any remaining quotes
        folder_name = folder_name.strip('"\'')
        # If it's just "/" or empty, return None
        if folder_name in ['/', '']:
            return None
    
    return folder_name

def get_message_count(mail, folder_name):
    """Gets message count for a folder using multiple methods"""
    # Quote folder name for IMAP commands
    quoted_folder = quote_folder_name_for_imap(folder_name)
    
    # Method 1: STATUS command (most efficient, doesn't require selecting folder)
    try:
        status, data = mail.status(quoted_folder, "(MESSAGES)")
        if status == 'OK' and data:
            count_str = str(data[0])
            count_match = re.search(r'MESSAGES\s+(\d+)', count_str)
            if count_match:
                return int(count_match.group(1))
            # Try alternative format
            if isinstance(data[0], (bytes, str)):
                num_match = re.search(r'(\d+)', count_str)
                if num_match:
                    return int(num_match.group(1))
    except (imaplib.IMAP4.error, AttributeError):
        pass
    
    # Method 2: EXAMINE (read-only SELECT)
    try:
        status, data = mail.examine(quoted_folder)
        if status == 'OK':
            for item in (data if data else []):
                item_str = item.decode() if isinstance(item, bytes) else str(item)
                exists_match = re.search(r'(\d+)\s*EXISTS', item_str, re.IGNORECASE)
                if exists_match:
                    msg_count = int(exists_match.group(1))
                    try:
                        mail.close()
                    except Exception:
                        pass
                    return msg_count
            
            # Search the full response string
            response_text = str(data) if data else ""
            exists_match = re.search(r'(\d+)\s*EXISTS', response_text, re.IGNORECASE)
            if exists_match:
                msg_count = int(exists_match.group(1))
                try:
                    mail.close()
                except Exception:
                    pass
                return msg_count
            
            # Try to find number in response
            if data:
                for item in data:
                    if isinstance(item, bytes):
                        item_str = item.decode('utf-8', errors='ignore')
                    else:
                        item_str = str(item)
                    num_match = re.search(r'\b(\d+)\b', item_str)
                    if num_match:
                        potential_count = int(num_match.group(1))
                        if potential_count >= 0:
                            try:
                                mail.close()
                            except Exception:
                                pass
                            return potential_count
            
            try:
                mail.close()
            except Exception:
                pass
    except (imaplib.IMAP4.error, AttributeError):
        pass
    except Exception:
        pass
    
    return None

def get_most_recent_email_info(mail, folder_name):
    """Gets the date/time and subject of the most recent email in a folder
    Returns: (date_datetime, subject_string) or (None, None) if not found
    Uses fetch('*') which is significantly faster than UID SEARCH
    """
    try:
        # Select folder (readonly) - quote folder name for IMAP
        quoted_folder = quote_folder_name_for_imap(folder_name)
        status, data = mail.select(quoted_folder, readonly=True)
        
        if status != 'OK':
             return None, None

        # Fetch the last message (*)
        # We want INTERNALDATE and SUBJECT
        status, data = mail.fetch('*', '(INTERNALDATE BODY.PEEK[HEADER.FIELDS (SUBJECT)])')
        
        if status != 'OK' or not data or data == [None]:
            try:
                mail.close()
            except:
                pass
            return None, None
            
        # Parse response using existing helpers
        date_str = parse_internal_date(data)
        
        # Extract subject
        subject_str = "(No Subject)"
        for item in data:
             if isinstance(item, tuple) and len(item) >= 2:
                 header_chunk = item[1]
                 if isinstance(header_chunk, bytes):
                     try:
                         msg_header = email.message_from_bytes(header_chunk)
                         raw_sub = msg_header.get('Subject', '')
                         if raw_sub:
                             subject_str = decode_email_header(raw_sub)
                     except:
                         pass
        
        msg_date = None
        if date_str:
            try:
                msg_date = parsedate_to_datetime(date_str)
            except:
                pass
                
        try:
            mail.close()
        except:
            pass
            
        return msg_date, subject_str

    except Exception:
        # Fail silently for folder access issues (common with Gmail special folders)
        try:
            mail.close()
        except:
            pass
        return None, None

def parse_internal_date(data_date):
    """Parses INTERNALDATE from IMAP FETCH response
    Handles both regular FETCH and UID FETCH response formats
    """
    date_str = None
    if not data_date:
        return None
    
    # Build full response string from all items
    full_response = ""
    for item in data_date:
        if isinstance(item, tuple):
            # Tuple format: (b'1 (INTERNALDATE "21-Nov-2025 13:10:45 +0000")', b'...')
            # or: (b'1 (INTERNALDATE "21-Nov-2025 13:10:45 +0000")')
            for part in item:
                if isinstance(part, bytes):
                    full_response += part.decode('utf-8', errors='ignore') + " "
                else:
                    full_response += str(part) + " "
        elif isinstance(item, bytes):
            full_response += item.decode('utf-8', errors='ignore') + " "
        else:
            full_response += str(item) + " "
    
    # Try various patterns for INTERNALDATE
    # Pattern 1: INTERNALDATE "21-Nov-2025 13:10:45 +0000" (with quotes)
    date_match = re.search(r'INTERNALDATE\s+"([^"]+)"', full_response)
    if date_match:
        return date_match.group(1)
    
    # Pattern 2: (INTERNALDATE "21-Nov-2025 13:10:45 +0000") (in parentheses with quotes)
    date_match = re.search(r'\(INTERNALDATE\s+"([^"]+)"\)', full_response)
    if date_match:
        return date_match.group(1)
    
    # Pattern 3: INTERNALDATE 21-Nov-2025 13:10:45 +0000 (without quotes, with timezone)
    date_match = re.search(r'INTERNALDATE\s+(\d{1,2}-[A-Za-z]{3}-\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[+-]\d{4})', full_response)
    if date_match:
        return date_match.group(1)
    
    # Pattern 4: INTERNALDATE 21-Nov-2025 13:10:45 +0000 (without quotes, try to get full)
    date_match = re.search(r'INTERNALDATE\s+([^\s)]+\s+[^\s)]+\s+[^\s)]+)', full_response)
    if date_match:
        return date_match.group(1)
    
    # Pattern 5: Look for date-like patterns in the response (IMAP date format)
    # IMAP date format: DD-MMM-YYYY HH:MM:SS +HHMM
    date_match = re.search(r'(\d{1,2}-[A-Za-z]{3}-\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[+-]\d{4})', full_response)
    if date_match:
        return date_match.group(1)
    
    # Pattern 6: Date without timezone
    date_match = re.search(r'(\d{1,2}-[A-Za-z]{3}-\d{4}\s+\d{1,2}:\d{2}:\d{2})', full_response)
    if date_match:
        return date_match.group(1)
    
    return None

def parse_rfc822_size(data_size):
    """Parses RFC822.SIZE from IMAP FETCH response"""
    size = 0
    if data_size:
        for item in data_size:
            if isinstance(item, tuple) and len(item) >= 2:
                size_str = item[1].decode() if isinstance(item[1], bytes) else str(item[1])
                size_match = re.search(r'RFC822\.SIZE\s+(\d+)', size_str)
                if size_match:
                    return int(size_match.group(1))
                size_match = re.search(r'\(RFC822\.SIZE\s+(\d+)\)', size_str)
                if size_match:
                    return int(size_match.group(1))
    return size

def extract_raw_email(data):
    """Extracts raw email bytes from IMAP FETCH response"""
    for response_item in data:
        if isinstance(response_item, tuple) and len(response_item) >= 2:
            raw_email = response_item[1]
            if isinstance(raw_email, bytes):
                return raw_email
            elif isinstance(raw_email, str):
                return raw_email.encode('utf-8')
    return None

def calculate_cutoff_times(search_window_hours):
    """Calculates cutoff time and IMAP search date for time window"""
    now = datetime.datetime.now()
    cutoff_time = now - datetime.timedelta(hours=search_window_hours)
    
    # IMAP SINCE uses date format, so we use the date from cutoff_time
    if search_window_hours < 24:
        # For windows less than 24 hours, search from yesterday to catch all possible matches
        date_since = (now - datetime.timedelta(days=1)).strftime("%d-%b-%Y")
    else:
        # For windows >= 24 hours, use the date from cutoff_time
        date_since = cutoff_time.strftime("%d-%b-%Y")
    
    # Convert cutoff_time to UTC for comparison with INTERNALDATE
    tz_offset_seconds = time.timezone if (time.daylight == 0) else time.altzone
    tz_offset = datetime.timedelta(seconds=-tz_offset_seconds)
    cutoff_time_utc = cutoff_time - tz_offset
    
    return now, cutoff_time, cutoff_time_utc, date_since

def search_messages_by_date(mail, date_since):
    """Searches for messages using IMAP UID SEARCH to get persistent UIDs"""
    try:
        status, messages = mail.uid('SEARCH', None, f'(SINCE "{date_since}")')
        if status != 'OK':
            print(f"Search failed with status: {status}")
            if messages and len(messages) > 0:
                print(f"Search response: {messages}")
            sys.stdout.flush()
            return []
    except Exception as search_err:
        print(f"Error executing IMAP UID search: {search_err}")
        sys.stdout.flush()
        return []
    
    email_uids = messages[0].split() if messages and len(messages) > 0 and messages[0] else []
    
    if email_uids:
        print(f"IMAP UID search returned {len(email_uids)} message UID(s)")
    else:
        print("IMAP UID search returned 0 messages")
    sys.stdout.flush()
    
    return email_uids

def filter_messages_by_time(mail, email_uids, cutoff_time_utc, time_window_str):
    """Filters messages by INTERNALDATE to match time window using UID FETCH"""
    print(f"\nFound {len(email_uids)} candidate messages from date search.")
    print("Filtering by actual received time...")
    sys.stdout.flush()
    
    filtered_email_uids = []
    
    for uid in email_uids:
        try:
            # Use UID FETCH to get INTERNALDATE
            res_date, data_date = mail.uid('FETCH', uid, '(INTERNALDATE)')
            if res_date == 'OK' and data_date:
                date_str = parse_internal_date(data_date)
                
                if date_str:
                    try:
                        msg_date = parsedate_to_datetime(date_str)
                        if msg_date:
                            # Convert to UTC naive datetime for comparison
                            if msg_date.tzinfo:
                                msg_date_utc = msg_date.astimezone(datetime.timezone.utc).replace(tzinfo=None)
                            else:
                                msg_date_utc = msg_date
                            
                            # Compare both in UTC
                            if msg_date_utc >= cutoff_time_utc:
                                filtered_email_uids.append(uid)
                        else:
                            # Can't parse, include it to be safe
                            filtered_email_uids.append(uid)
                    except Exception as date_parse_err:
                        uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
                        print(f"Warning: Could not parse INTERNALDATE for message UID {uid_str}: {date_parse_err}")
                        filtered_email_uids.append(uid)
                else:
                    # If we can't get INTERNALDATE, include the message to be safe
                    filtered_email_uids.append(uid)
            else:
                # If we can't fetch INTERNALDATE, include the message to be safe
                filtered_email_uids.append(uid)
        except Exception as fetch_err:
            uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
            print(f"Warning: Could not fetch INTERNALDATE for message UID {uid_str}: {fetch_err}")
            filtered_email_uids.append(uid)
    
    print(f"After time filtering: {len(filtered_email_uids)} messages within the {time_window_str} window.")
    sys.stdout.flush()
    return filtered_email_uids

def extract_sender_from_header(msg_header):
    """Extracts sender email address from email header"""
    from_header = msg_header.get('From', '')
    if not from_header:
        return ''
    
    # Decode the From header
    from_decoded = decode_email_header(from_header)
    
    # Try to extract email address from "Name <email@domain.com>" or just "email@domain.com"
    # Use email.utils.parseaddr to handle various formats
    name, email_addr = parseaddr(from_decoded)
    
    # If parseaddr didn't find an email, try regex
    if not email_addr:
        email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', from_decoded)
        if email_match:
            email_addr = email_match.group(0)
    
    return email_addr.lower() if email_addr else ''

def is_email_excluded(sender, subject, body_text=""):
    """Checks if an email should be excluded based on sender, subject, or body keywords/phrases.
    
    Supports both single keywords and multi-word phrases.
    """
    sender_lower = sender.lower()
    subject_lower = subject.lower()
    body_lower = body_text.lower() if body_text else ""
    
    # Check excluded senders
    for excluded in EXCLUDED_SENDERS:
        excluded_lower = excluded.lower().strip()
        if not excluded_lower:
            continue
        
        # Extract domain from sender email (everything after @)
        sender_domain = ''
        if '@' in sender_lower:
            sender_domain = sender_lower.split('@', 1)[1]
        
        # Check if it's a domain exclusion (starts with @)
        if excluded_lower.startswith('@'):
            # Remove the @ for comparison
            excluded_domain = excluded_lower[1:]
            # Check if sender domain ends with the excluded domain (for partial domain matching)
            # e.g., @.gov.au matches ths.tas.gov.au
            if sender_domain and (sender_domain == excluded_domain or sender_domain.endswith('.' + excluded_domain)):
                return True, f"sender domain matches '{excluded}'"
            # Also check if excluded string is in full sender (for backwards compatibility)
            if excluded_lower in sender_lower:
                return True, f"sender domain matches '{excluded}'"
        # Check if it's a partial domain (starts with .) - matches domains ending with it
        elif excluded_lower.startswith('.'):
            if sender_domain and sender_domain.endswith(excluded_lower):
                return True, f"sender domain ends with '{excluded}'"
        # Check if it's an exact email match
        elif excluded_lower == sender_lower:
            return True, f"sender matches '{excluded}'"
        # Check if sender contains the excluded string (for partial matches)
        elif excluded_lower in sender_lower:
            return True, f"sender contains '{excluded}'"
    
    # Check excluded subject/body keywords/phrases (supports multi-word phrases)
    for keyword in EXCLUDED_SUBJECT_KEYWORDS:
        keyword_lower = keyword.lower().strip()
        if not keyword_lower:
            continue
        
        # Check subject
        if keyword_lower in subject_lower:
            return True, f"subject contains keyword/phrase '{keyword}'"
        
        # Check body
        if body_lower and keyword_lower in body_lower:
            return True, f"body contains keyword/phrase '{keyword}'"
    
    return False, None

def extract_body_text(raw_email_bytes):
    """Extracts plain text from email body for keyword searching"""
    try:
        msg = email.message_from_bytes(raw_email_bytes)
        body_text = ""
        
        # Try to get plain text body
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/plain':
                    payload = part.get_payload(decode=True)
                    if payload:
                        try:
                            # Try to decode with charset from part
                            charset = part.get_content_charset() or 'utf-8'
                            body_text += payload.decode(charset, errors='ignore')
                        except:
                            # Fallback to utf-8
                            body_text += payload.decode('utf-8', errors='ignore')
        else:
            # Single part message
            payload = msg.get_payload(decode=True)
            if payload:
                try:
                    charset = msg.get_content_charset() or 'utf-8'
                    body_text = payload.decode(charset, errors='ignore')
                except:
                    body_text = payload.decode('utf-8', errors='ignore')
        
        return body_text.lower()
    except Exception:
        return ""

def is_email_force_included(subject, body_text):
    """Checks if an email should be force-included based on keywords in subject or body.
    
    NOTE: This function intentionally does NOT check the sender email address.
    Keywords are only matched against the subject line and email body text.
    This prevents false matches when keywords appear in the sender's email address
    (e.g., keyword 'jkokavec' should not match sender 'jkokavec@gmail.com').
    """
    if not FORCE_INCLUDE_KEYWORDS:
        return False, None
    
    subject_lower = subject.lower()
    body_lower = body_text.lower() if body_text else ""
    
    for keyword in FORCE_INCLUDE_KEYWORDS:
        keyword_lower = keyword.lower().strip()
        if not keyword_lower:
            continue
        
        # Check subject only (NOT sender address) - supports phrases
        if keyword_lower in subject_lower:
            return True, f"subject contains force-include keyword/phrase '{keyword}'"
        
        # Check body only (NOT sender address) - supports phrases
        if body_lower and keyword_lower in body_lower:
            return True, f"body contains force-include keyword/phrase '{keyword}'"
    
    return False, None

def analyze_message_headers(mail, email_uids):
    """Analyzes message headers and returns spam candidates with metadata using UID FETCH"""
    print(f"\nFound {len(email_uids)} candidate messages.")
    print("Analyzing headers...")
    sys.stdout.flush()
    
    spam_candidates = []
    total_size_bytes = 0
    excluded_count = 0
    
    for uid in email_uids:
        try:
            # Fetch size using UID FETCH
            res_size, data_size = mail.uid('FETCH', uid, '(RFC822.SIZE)')
            size = parse_rfc822_size(data_size)
            if size == 0:
                size = DEFAULT_SIZE_KB
            
            # Fetch header using UID FETCH
            res_header, data_header = mail.uid('FETCH', uid, '(BODY.PEEK[HEADER])')
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
                uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
                print(f"Warning: Could not fetch header for message UID {uid_str}")
                continue
            
            msg_header = email.message_from_bytes(raw_header)
            subject_raw = msg_header['Subject']
            subject = decode_email_header(subject_raw) if subject_raw else "(No Subject)"
            date_str = msg_header['Date']
            
            # Extract sender
            sender = extract_sender_from_header(msg_header)
            
            # Fetch body to check for keywords (needed for both exclusion and force-include)
            body_text = ""
            if FORCE_INCLUDE_KEYWORDS or EXCLUDED_SUBJECT_KEYWORDS:
                try:
                    res_body, data_body = mail.uid('FETCH', uid, '(BODY.PEEK[])')
                    if res_body == 'OK' and data_body:
                        raw_email = extract_raw_email(data_body)
                        if raw_email:
                            body_text = extract_body_text(raw_email)
                except Exception:
                    pass  # If body fetch fails, just use empty body
            
            # Check force-include keywords first (these override exclusions)
            is_force_included, force_include_reason = is_email_force_included(subject, body_text)
            
            if is_force_included:
                # Force include - always process this email regardless of exclusion rules
                print(f" - FORCE INCLUDED: {safe_print_subject(subject, 50)}... ({force_include_reason})")
                sys.stdout.flush()
                # Continue processing - skip exclusion check
            else:
                # Check if email should be excluded (now checks both subject and body)
                is_excluded, exclusion_reason = is_email_excluded(sender, subject, body_text)
                if is_excluded:
                    excluded_count += 1
                    display_subject = safe_print_subject(subject, 50)
                    print(f" - EXCLUDED: {display_subject}... ({exclusion_reason})")
                    sys.stdout.flush()
                    continue
            
            total_size_bytes += size
            
            # Store UID as string for easier handling
            uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
            
            spam_candidates.append({
                'uid': uid_str,
                'uid_bytes': uid,  # Keep bytes version for IMAP operations
                'subject': subject,
                'sender': sender,
                'date': date_str,
                'size': size
            })
            
            display_subject = safe_print_subject(subject, 50)
            print(f" - Identified: {display_subject}... ({int(size/1024)} KB)")
            sys.stdout.flush()
        except Exception as e:
            uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
            print(f"Warning: Error processing message UID {uid_str}: {e}")
            sys.stdout.flush()
            import traceback
            traceback.print_exc()
            continue
    
    if excluded_count > 0:
        print(f"\nExcluded {excluded_count} email(s) based on exclusion rules.")
        sys.stdout.flush()
    
    return spam_candidates, total_size_bytes

def download_messages(mail, spam_candidates, total_size_bytes):
    """Downloads messages to disk and returns list of file paths"""
    print("\n" + "="*40)
    print(f"Identified {len(spam_candidates)} emails. Total Size: {get_size_str(total_size_bytes)}.")
    print("Proceeding with download (non-interactive mode)...")
    sys.stdout.flush()

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
    sys.stdout.flush()
    
    downloaded_files = []
    timestamps = []

    for item in spam_candidates:
        try:
            # Fetch Full Body (Safe PEEK) using UID FETCH
            uid_bytes = item['uid_bytes']
            res, data = mail.uid('FETCH', uid_bytes, '(BODY.PEEK[])')
            if res != 'OK' or not data:
                print(f"Warning: Could not fetch body for message UID {item['uid']}")
                continue
            
            # Extract raw email
            raw_email = extract_raw_email(data)
            if raw_email is None:
                print(f"Warning: Failed to fetch raw email for message UID {item['uid']}. Skipping.")
                continue
            
            # Parse timestamp if available
            if item['date']:
                try:
                    dt = parsedate_to_datetime(item['date'])
                    if dt: 
                        timestamps.append(dt)
                except Exception:
                    pass

            # Save file
            clean_sub = sanitize_filename(item['subject'])
            clean_sub = (clean_sub[:50] + '..') if len(clean_sub) > 50 else clean_sub
            filename = f"{clean_sub}_{item['uid']}.eml"
            filepath = os.path.join(download_path, filename)
            
            try:
                with open(filepath, 'wb') as f:
                    f.write(raw_email)
                downloaded_files.append(filepath)
                print(f"Saved: {filename}")
                sys.stdout.flush()
            except Exception as file_err:
                print(f"Error saving message UID {item['uid']} to file: {file_err}")
                sys.stdout.flush()
                
        except Exception as e:
            print(f"Error downloading message UID {item['uid']}: {e}")
            sys.stdout.flush()
            import traceback
            traceback.print_exc()
            continue
    
    return downloaded_files, timestamps

def _handle_imap_error(e, mail):
    """Handles IMAP errors with authentication troubleshooting"""
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
    safe_logout(mail)

def connect_imap():
    """Connects to IMAP server and logs in"""
    print(f"Connecting to {GMAIL_ACCOUNT}...")
    sys.stdout.flush()
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(GMAIL_ACCOUNT, APP_PASS)
    print("Connected successfully.")
    sys.stdout.flush()
    return mail

def list_all_folders(mail):
    """Lists all folders and identifies spam folder candidates"""
    print("Listing all available mailboxes to find spam folder...")
    sys.stdout.flush()
    all_folder_names = []
    spam_candidate_names = []
    
    try:
        status, folders = mail.list()
        if status == 'OK':
            for folder in folders:
                folder_name = parse_folder_from_list_response(folder)
                if folder_name and folder_name != '/' and folder_name.strip():
                    all_folder_names.append(folder_name)
                    folder_upper = folder_name.upper()
                    if ('SPAM' in folder_upper or 'JUNK' in folder_upper) and 'INBOX' not in folder_upper:
                        spam_candidate_names.append(folder_name)
    except Exception as list_err:
        print(f"Warning: Could not list mailboxes: {list_err}")
        sys.stdout.flush()
    
    if all_folder_names:
        print(f"Found {len(all_folder_names)} mailboxes")
        if spam_candidate_names:
            print(f"Spam/Junk folder candidates: {spam_candidate_names}")
        sys.stdout.flush()
    
    return all_folder_names, spam_candidate_names

def display_folder_counts(mail, all_folder_names, is_first_run=False):
    """Displays message counts and most recent email date for all folders
    Reports in real-time for each folder before moving to the next.
    Aborts if folder names, messages, or most recent email cannot be found (where messages > 0).
    """
    if not all_folder_names:
        if is_first_run:
            print("\n" + "="*70)
            print("FIRST RUN FAILED: Could not retrieve folder names")
            print("="*70)
            print("The script is aborting to commandline as this is the first run.")
            print("Please verify your Gmail account settings and IMAP access.")
            print("="*70)
            sys.stdout.flush()
            cleanup_logging()
            sys.exit(1)
        return
    
    print("\n" + "="*70)
    print("FOLDER MESSAGE COUNTS")
    print("="*70)
    print("Getting message counts and most recent email dates for all folders...")
    print()
    sys.stdout.flush()
    
    folder_counts = []
    for folder_name in all_folder_names:
        # Skip INBOX and other forbidden folders for security
        if not folder_name or folder_name == '/':
            continue
        if is_forbidden_folder(folder_name):
            continue
        
        print(f"Checking {folder_name}...")
        sys.stdout.flush()
        
        msg_count = get_message_count(mail, folder_name)
        
        # Log warning if we can't get message count, but don't abort
        # The actual spam folder selection will handle critical failures
        if msg_count is None:
            print(f"  Warning: Could not retrieve message count for folder '{folder_name}'")
            print(f"  Continuing...")
            sys.stdout.flush()
        
        most_recent_date = None
        most_recent_subject = None
        
        # Always try to get the most recent date and subject if we have messages
        if msg_count is not None and msg_count > 0:
            most_recent_date, most_recent_subject = get_most_recent_email_info(mail, folder_name)
            
            # Log warning if we can't get most recent email info, but don't abort
            # The actual spam folder selection will handle critical failures
            if most_recent_date is None:
                print(f"  Warning: Could not retrieve most recent email info for folder '{folder_name}' ({msg_count} messages)")
                print(f"  Continuing...")
                sys.stdout.flush()
        
        folder_counts.append((folder_name, msg_count, most_recent_date, most_recent_subject))
        
        # Report immediately for this folder
        count_str = f"{msg_count:>10,}" if msg_count is not None else "N/A".rjust(10)
        date_str = "N/A"
        subject_str = "N/A"
        if most_recent_date:
            try:
                date_str = most_recent_date.strftime('%Y-%m-%d %H:%M:%S')
            except Exception:
                date_str = "N/A"
        if most_recent_subject:
            subject_str = safe_print_subject(most_recent_subject, 50)
        
        print(f"  Messages: {count_str}")
        print(f"  Most Recent: {date_str}")
        print(f"  Subject: {subject_str}")
        print()
        sys.stdout.flush()
    
    # Sort by message count (descending)
    folder_counts.sort(key=lambda x: (x[1] is None, x[1] if x[1] is not None else 0), reverse=True)
    
    # Display summary table
    print(f"{'Folder Name':<40} {'Messages':>10} {'Most Recent Email':>20} {'Subject':>30}")
    print("-" * 100)
    sys.stdout.flush()
    total_messages = 0
    folders_with_counts = 0
    for folder_name, count, recent_date, recent_subject in folder_counts:
        # Format message count
        count_str = f"{count:>10,}" if count is not None else "N/A".rjust(10)
        
        # Format date - always try to show it if available
        date_str = "N/A"
        if recent_date:
            try:
                date_str = recent_date.strftime('%Y-%m-%d %H:%M')
            except Exception:
                date_str = "N/A"
        
        # Format subject
        subject_str = "N/A"
        if recent_subject:
            subject_str = safe_print_subject(recent_subject, 30)
        
        print(f"{folder_name:<40} {count_str} {date_str:>20} {subject_str:>30}")
        
        if count is not None:
            total_messages += count
            folders_with_counts += 1
        sys.stdout.flush()
    print("-" * 100)
    if folders_with_counts > 0:
        print(f"{'TOTAL (countable folders)':<40} {total_messages:>10,}")
    print(f"{'Folders with message counts':<40} {folders_with_counts:>10}")
    print("="*70)
    print()
    sys.stdout.flush()

def find_and_select_spam_folder(mail, spam_candidate_names):
    """Finds and selects the spam folder, with security checks. Prompts user to select by number."""
    print("\n" + "="*70)
    print("SPAM FOLDER SELECTION")
    print("="*70)
    
    # If preview is disabled, only check the specified spam folder
    if not PREVIEW_ALL_FOLDERS:
        print(f"Preview disabled. Checking specified spam folder: {SPAM_FOLDER_NAME}")
        sys.stdout.flush()
        
        # CRITICAL: Safety check - never allow forbidden folders
        if is_forbidden_folder(SPAM_FOLDER_NAME):
            print(f"\nSECURITY ERROR: Configured spam folder '{SPAM_FOLDER_NAME}' is FORBIDDEN!")
            print("This folder cannot be used for security reasons.")
            print("Please update SPAM_FOLDER_NAME in config.py to a valid spam/junk folder.")
            sys.stdout.flush()
            raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Configured spam folder '{SPAM_FOLDER_NAME}' is forbidden.")
        
        # Verify it's a spam/junk folder
        if not is_spam_folder(SPAM_FOLDER_NAME):
            print(f"\nSECURITY ERROR: Configured spam folder '{SPAM_FOLDER_NAME}' is not a spam/junk folder!")
            print("The folder name must contain 'SPAM' or 'JUNK'.")
            print("Please update SPAM_FOLDER_NAME in config.py to a valid spam/junk folder.")
            sys.stdout.flush()
            raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Configured spam folder '{SPAM_FOLDER_NAME}' is not a spam/junk folder.")
        
        # Try to select the specified folder
        try:
            quoted_spam_folder = quote_folder_name_for_imap(SPAM_FOLDER_NAME)
            status, data = mail.select(quoted_spam_folder, readonly=True)
            if status != 'OK':
                print(f"ERROR: Could not select folder '{SPAM_FOLDER_NAME}'")
                print(f"Status: {status}")
                sys.stdout.flush()
                raise imaplib.IMAP4.error(f"Failed to select spam mailbox '{SPAM_FOLDER_NAME}'")
            
            # Final security checks
            normalized_name = normalize_folder_name(SPAM_FOLDER_NAME)
            if normalized_name == 'INBOX' or normalized_name.endswith('/INBOX') or normalized_name.startswith('INBOX/'):
                mail.close()
                print(f"SECURITY ERROR: Folder '{SPAM_FOLDER_NAME}' appears to be INBOX!")
                mail.logout()
                raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Folder '{SPAM_FOLDER_NAME}' is forbidden.")
            
            if is_forbidden_folder(SPAM_FOLDER_NAME):
                mail.close()
                print(f"SECURITY ERROR: Folder '{SPAM_FOLDER_NAME}' is FORBIDDEN!")
                mail.logout()
                raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Folder '{SPAM_FOLDER_NAME}' is forbidden.")
            
            # Get message count
            msg_count = 0
            try:
                status, message_count = mail.status(quoted_spam_folder, "(MESSAGES)")
                if status == 'OK' and message_count:
                    count_match = re.search(r'MESSAGES\s+(\d+)', str(message_count[0]))
                    if count_match:
                        msg_count = int(count_match.group(1))
            except Exception:
                pass
            
            # Still require user confirmation
            print(f"\nFound spam folder: {SPAM_FOLDER_NAME} ({msg_count} messages)")
            print("="*70)
            sys.stdout.flush()
            
            while True:
                try:
                    response = input(f"\nUse this spam folder? (yes/no) [yes]: ").strip().lower()
                    if not response:
                        response = 'yes'  # Default to yes if empty
                    
                    if response in ['yes', 'y']:
                        print(f"\nSelected spam folder: {SPAM_FOLDER_NAME} ({msg_count} messages)")
                        print("="*70)
                        sys.stdout.flush()
                        return SPAM_FOLDER_NAME
                    elif response in ['no', 'n']:
                        print("\nFolder selection cancelled by user.")
                        sys.stdout.flush()
                        cleanup_logging()
                        sys.exit(1)
                    else:
                        print("Please enter 'yes' or 'no'.")
                        sys.stdout.flush()
                except KeyboardInterrupt:
                    print("\n\nSelection cancelled by user.")
                    sys.stdout.flush()
                    cleanup_logging()
                    sys.exit(1)
        
        except Exception as e:
            if isinstance(e, imaplib.IMAP4.error):
                raise
            print(f"ERROR: Could not select folder '{SPAM_FOLDER_NAME}': {e}")
            sys.stdout.flush()
            raise imaplib.IMAP4.error(f"Failed to select spam mailbox '{SPAM_FOLDER_NAME}'")
    
    # Preview enabled - show all spam folders
    print("Checking available spam/junk folders...")
    sys.stdout.flush()
    
    spam_folders = []
    
    # Add candidates from LIST
    for candidate in spam_candidate_names:
        if candidate not in spam_folders:
            spam_folders.append(candidate)
    
    # Add common names
    for name in COMMON_SPAM_FOLDER_NAMES:
        if name not in spam_folders:
            spam_folders.append(name)
    
    # First pass: Check all folders and get message counts
    folder_candidates = []  # List of (folder_name, msg_count, can_select)
    
    for folder_name in spam_folders:
        # CRITICAL: Safety check - never allow forbidden folders
        if is_forbidden_folder(folder_name):
            continue
        
        try:
            quoted_folder = quote_folder_name_for_imap(folder_name)
            status, data = mail.select(quoted_folder, readonly=True)
            if status == 'OK':
                # CRITICAL: Double-check forbidden folders
                if is_forbidden_folder(folder_name):
                    mail.close()
                    continue
                
                # Verify we did NOT select INBOX
                normalized_name = normalize_folder_name(folder_name)
                if normalized_name == 'INBOX' or normalized_name.endswith('/INBOX') or normalized_name.startswith('INBOX/'):
                    mail.close()
                    continue
                
                # Verify folder name contains SPAM or JUNK
                if not is_spam_folder(folder_name):
                    mail.close()
                    continue
                
                # Get message count
                msg_count = 0
                try:
                    status, message_count = mail.status(quoted_folder, "(MESSAGES)")
                    if status == 'OK' and message_count:
                        count_match = re.search(r'MESSAGES\s+(\d+)', str(message_count[0]))
                        if count_match:
                            msg_count = int(count_match.group(1))
                except Exception:
                    pass
                
                folder_candidates.append((folder_name, msg_count, True))
                mail.close()
            else:
                folder_candidates.append((folder_name, 0, False))
        except imaplib.IMAP4.error as select_err:
            if 'INBOX' in str(select_err).upper() or is_forbidden_folder(folder_name):
                continue
            folder_candidates.append((folder_name, 0, False))
            continue
        except Exception:
            folder_candidates.append((folder_name, 0, False))
            continue
    
    # Filter to only folders that can be selected
    selectable_folders = [(name, count) for name, count, can_select in folder_candidates if can_select]
    
    if not selectable_folders:
        _print_folder_selection_error(mail)
        raise imaplib.IMAP4.error("SECURITY: Failed to find any selectable spam mailboxes. Script aborted to prevent INBOX access.")
    
    # Find default suggestion: "[Google Mail]/Spam"
    default_folder = "[Google Mail]/Spam"
    default_index = None
    
    # Try to find default folder in selectable folders
    for idx, (folder_name, _) in enumerate(selectable_folders):
        if folder_name == default_folder:
            default_index = idx
            break
    
    # If default not found, try to find it in candidates and add it if it's selectable
    if default_index is None:
        for folder_name, count, can_select in folder_candidates:
            if folder_name == default_folder and can_select:
                selectable_folders.append((folder_name, count))
                default_index = len(selectable_folders) - 1
                break
    
    # Sort by message count (descending) for display, but prioritize default
    selectable_folders.sort(key=lambda x: (x[0] != default_folder, -x[1] if x[1] is not None else 0))
    
    # Recalculate default_index after sorting
    if default_index is not None:
        for idx, (folder_name, _) in enumerate(selectable_folders):
            if folder_name == default_folder:
                default_index = idx
                break
    
    # Display numbered list
    print("\nAvailable spam/junk folders:")
    print("-" * 70)
    for idx, (folder_name, msg_count) in enumerate(selectable_folders, 1):
        count_str = f"{msg_count:>10,}" if msg_count is not None else "N/A".rjust(10)
        default_marker = " (suggested)" if idx == default_index + 1 else ""
        print(f"  {idx}. {folder_name:<50} ({count_str} messages){default_marker}")
    print("-" * 70)
    sys.stdout.flush()
    
    # Prompt user for selection (MUST get user input, no auto-selection)
    default_prompt = f" [{default_index + 1}]" if default_index is not None else ""
    while True:
        try:
            response = input(f"\nPlease select the spam folder by number (1-{len(selectable_folders)}){default_prompt}: ").strip()
            if not response:
                print("Please enter a number. User input is required.")
                sys.stdout.flush()
                continue
            
            selection = int(response)
            if 1 <= selection <= len(selectable_folders):
                selected_folder, msg_count = selectable_folders[selection - 1]
                
                # CRITICAL: Double-check that selected folder is not forbidden
                if is_forbidden_folder(selected_folder):
                    print(f"\nSECURITY ERROR: Selected folder '{selected_folder}' is FORBIDDEN!")
                    print("This folder cannot be selected for security reasons.")
                    print("Please select a different folder.")
                    sys.stdout.flush()
                    continue
                
                # Verify it's a spam/junk folder
                if not is_spam_folder(selected_folder):
                    print(f"\nSECURITY ERROR: Selected folder '{selected_folder}' is not a spam/junk folder!")
                    print("Please select a folder that contains 'SPAM' or 'JUNK' in its name.")
                    sys.stdout.flush()
                    continue
                
                break
            else:
                print(f"Please enter a number between 1 and {len(selectable_folders)}.")
                sys.stdout.flush()
        except ValueError:
            print("Please enter a valid number.")
            sys.stdout.flush()
        except KeyboardInterrupt:
            print("\n\nSelection cancelled by user.")
            sys.stdout.flush()
            cleanup_logging()
            sys.exit(1)
    
    # Verify and select the chosen folder
    try:
        # FINAL SECURITY CHECK: Prevent forbidden folders from EVER being selected
        if is_forbidden_folder(selected_folder):
            print(f"\nSECURITY VIOLATION: Selected folder '{selected_folder}' is FORBIDDEN!")
            print("The following folders are NEVER allowed:")
            print("  - INBOX")
            print("  - [Google Mail]/All Mail or [Gmail]/All Mail")
            print("  - [Google Mail]/Important or [Gmail]/Important")
            print("  - [Google Mail]/Sent Mail or [Gmail]/Sent Mail")
            print("  - [Google Mail]/Starred or [Gmail]/Starred")
            print("  - [Google Mail]/Drafts or [Gmail]/Drafts")
            sys.stdout.flush()
            raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Selected folder '{selected_folder}' is forbidden. Script aborted.")
        
        quoted_selected_folder = quote_folder_name_for_imap(selected_folder)
        status, data = mail.select(quoted_selected_folder, readonly=True)
        if status != 'OK':
            print(f"ERROR: Could not select folder '{selected_folder}'")
            sys.stdout.flush()
            raise imaplib.IMAP4.error(f"Failed to select spam mailbox '{selected_folder}'")
        
        # Final security checks after selection
        normalized_name = normalize_folder_name(selected_folder)
        if normalized_name == 'INBOX' or normalized_name.endswith('/INBOX') or normalized_name.startswith('INBOX/'):
            mail.close()
            print(f"SECURITY ERROR: Selected folder '{selected_folder}' appears to be INBOX!")
            mail.logout()
            raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Selected folder '{selected_folder}' is forbidden. Only spam/junk folders allowed.")
        
        # Check for forbidden Gmail folders
        forbidden_check = is_forbidden_folder(selected_folder)
        if forbidden_check:
            mail.close()
            print(f"SECURITY ERROR: Selected folder '{selected_folder}' is FORBIDDEN!")
            mail.logout()
            raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Selected folder '{selected_folder}' is forbidden. Access denied.")
        
        if not is_spam_folder(selected_folder):
            mail.close()
            print(f"SECURITY ERROR: Selected folder '{selected_folder}' does not contain 'SPAM' or 'JUNK'!")
            mail.logout()
            raise imaplib.IMAP4.error(f"SECURITY VIOLATION: Selected folder '{selected_folder}' is not a spam/junk folder. Access denied.")
        
        print(f"\nSuccessfully selected SPAM/JUNK mailbox: {selected_folder} ({msg_count} messages)")
        sys.stdout.flush()
        
        if msg_count == 0:
            print(f"WARNING: Selected folder '{selected_folder}' has 0 messages. This might not be the correct spam folder.")
            sys.stdout.flush()
        
        print("="*70)
        sys.stdout.flush()
        
    except Exception as e:
        if isinstance(e, imaplib.IMAP4.error):
            raise
        print(f"ERROR: Could not select folder '{selected_folder}': {e}")
        sys.stdout.flush()
        raise imaplib.IMAP4.error(f"Failed to select spam mailbox '{selected_folder}'")
    
    return selected_folder

def _print_folder_selection_error(mail):
    """Prints detailed error when spam folder cannot be selected"""
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
                if ('SPAM' in folder_upper or 'JUNK' in folder_upper) and 'INBOX' not in folder_upper:
                    spam_candidates_found.append(folder_str)
            
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

def process_spam_iteration(is_first_run=False):
    """Processes one iteration of spam download and forwarding"""
    print("--- STARTING SPAM PROCESSOR ITERATION ---")
    sys.stdout.flush()
    
    mail = None
    spam_candidates = []
    total_size_bytes = 0
    
    # ---------------------------------------------------------
    # PHASE 1: CONNECT AND IDENTIFY (READ-ONLY)
    # ---------------------------------------------------------
    try:
        mail = connect_imap()
        sys.stdout.flush()
        
        all_folder_names, spam_candidate_names = list_all_folders(mail)
        sys.stdout.flush()
        
        # On first run, abort if folder listing failed or returned empty
        if is_first_run:
            if not all_folder_names:
                print("\n" + "="*70)
                print("FIRST RUN FAILED: Could not retrieve folder names")
                print("="*70)
                print("The script is aborting to commandline as this is the first run.")
                print("Please verify your Gmail account settings and IMAP access.")
                print("="*70)
                safe_logout(mail)
                cleanup_logging()
                sys.exit(1)
        
        if all_folder_names:
            # Only display folder counts if preview is enabled
            if PREVIEW_ALL_FOLDERS:
                display_folder_counts(mail, all_folder_names, is_first_run=is_first_run)
                sys.stdout.flush()
        elif is_first_run:
            # On first run, abort if we couldn't display folder counts
            print("\n" + "="*70)
            print("FIRST RUN FAILED: Could not retrieve folder counts")
            print("="*70)
            print("The script is aborting to commandline as this is the first run.")
            print("Please verify your Gmail account settings and IMAP access.")
            print("="*70)
            safe_logout(mail)
            cleanup_logging()
            sys.exit(1)
        
        # Load cached spam folder or select it on first run
        selected_folder = load_spam_folder_cache()
        if selected_folder is None or is_first_run:
            # Need to select spam folder (first run or cache missing)
            selected_folder = find_and_select_spam_folder(mail, spam_candidate_names)
            save_spam_folder_cache(selected_folder)
            sys.stdout.flush()
        else:
            # Use cached folder name - just select it without prompting
            print(f"Using cached spam folder: {selected_folder}")
            sys.stdout.flush()
            quoted_folder = quote_folder_name_for_imap(selected_folder)
            status, data = mail.select(quoted_folder, readonly=True)
            if status != 'OK':
                print(f"ERROR: Could not select cached folder '{selected_folder}'. Re-selecting...")
                sys.stdout.flush()
                selected_folder = find_and_select_spam_folder(mail, spam_candidate_names)
                save_spam_folder_cache(selected_folder)
            sys.stdout.flush()
        
        # Calculate time window and search dates
        now, cutoff_time, cutoff_time_utc, date_since = calculate_cutoff_times(SPAM_SEARCH_WINDOW_HOURS)
        time_window_str = format_hours_as_string(SPAM_SEARCH_WINDOW_HOURS)
        
        print(f"Searching for spam received since {date_since} (past {time_window_str})...")
        print(f"Current time: {now.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Search window start: {cutoff_time.strftime('%Y-%m-%d %H:%M:%S')}")
        sys.stdout.flush()
        
        # Load already-sent UIDs to filter them out (stored as strings)
        sent_uids = load_sent_uids()
        if sent_uids:
            print(f"Filtering out {len(sent_uids)} already-sent email(s)...")
            sys.stdout.flush()
        
        # Search for messages (returns UIDs as bytes)
        email_uids = search_messages_by_date(mail, date_since)
        sys.stdout.flush()
        if not email_uids:
            print(f"No spam found in the past {time_window_str}.")
            sys.stdout.flush()
            safe_logout(mail)
            return
        
        # Filter out already-sent UIDs (convert bytes to strings for comparison)
        new_email_uids = []
        skipped_count = 0
        for uid in email_uids:
            uid_str = uid.decode() if isinstance(uid, bytes) else str(uid)
            if uid_str not in sent_uids:
                new_email_uids.append(uid)  # Keep as bytes for IMAP operations
            else:
                skipped_count += 1
        
        if skipped_count > 0:
            print(f"Filtered out {skipped_count} already-sent email(s).")
            sys.stdout.flush()
        
        if not new_email_uids:
            print(f"No new spam found in the past {time_window_str} (all messages already sent).")
            sys.stdout.flush()
            safe_logout(mail)
            return
        
        # Filter by INTERNALDATE
        filtered_uids = filter_messages_by_time(mail, new_email_uids, cutoff_time_utc, time_window_str)
        sys.stdout.flush()
        if not filtered_uids:
            print(f"No new spam found in the past {time_window_str}.")
            sys.stdout.flush()
            safe_logout(mail)
            return

        # Analyze headers
        spam_candidates, total_size_bytes = analyze_message_headers(mail, filtered_uids)
        sys.stdout.flush()
    
    except imaplib.IMAP4.error as e:
        # If we failed to find spam folders, re-raise so run_spam_processor can handle first-run abort
        if "Failed to select spam mailbox" in str(e) or "spam mailbox" in str(e).lower():
            safe_logout(mail)
            sys.stdout.flush()
            raise  # Re-raise so run_spam_processor can check if it's first run
        _handle_imap_error(e, mail)
        sys.stdout.flush()
        return
    except Exception as e:
        error_str = str(e).lower()
        # If error is about spam folder detection, re-raise for first-run handling
        if "spam" in error_str and ("folder" in error_str or "mailbox" in error_str):
            safe_logout(mail)
            sys.stdout.flush()
            raise  # Re-raise so run_spam_processor can check if it's first run
        print(f"Error during connection/search: {e}")
        sys.stdout.flush()
        safe_logout(mail)
        return

    # ---------------------------------------------------------
    # PHASE 2: DOWNLOAD (Non-interactive - always proceed)
    # ---------------------------------------------------------
    if not spam_candidates:
        print("No spam candidates to download.")
        sys.stdout.flush()
        safe_logout(mail)
        return
    
    downloaded_files, timestamps = download_messages(mail, spam_candidates, total_size_bytes)
    sys.stdout.flush()
    
    # Close IMAP connection
    mail.close()
    mail.logout()

    # ---------------------------------------------------------
    # PHASE 4: REPORT STATISTICS
    # ---------------------------------------------------------
    print_statistics(downloaded_files, total_size_bytes, timestamps)
    sys.stdout.flush()
    
    # ---------------------------------------------------------
    # PHASE 5: FORWARD TO SPAMCOP (With first-run confirmation)
    # ---------------------------------------------------------
    if len(downloaded_files) > 0:
        forward_to_spamcop(downloaded_files, spam_candidates, total_size_bytes)
        sys.stdout.flush()
    else:
        print("\nSkipping send. Files remain in your local folder.")
        sys.stdout.flush()

    print("\n--- ITERATION COMPLETE ---")
    sys.stdout.flush()

def print_statistics(downloaded_files, total_size_bytes, timestamps):
    """Prints download statistics"""
    print("\n" + "="*40)
    print("DOWNLOAD STATISTICS")
    print("="*40)
    print(f"Source Folder:     [Gmail]/Spam (READ-ONLY ACCESS)")
    print(f"Total Emails:      {len(downloaded_files)}")
    print(f"Total Size:        {get_size_str(total_size_bytes)}")
    sys.stdout.flush()
    
    if timestamps:
        # Normalize timestamps to all be timezone-aware (UTC) before sorting
        # This fixes the error when comparing offset-naive and offset-aware datetimes
        normalized_timestamps = []
        for dt in timestamps:
            if dt.tzinfo is None:
                # If naive, assume UTC
                normalized_timestamps.append(dt.replace(tzinfo=datetime.timezone.utc))
            else:
                # If already aware, convert to UTC for consistency
                normalized_timestamps.append(dt.astimezone(datetime.timezone.utc))
        
        normalized_timestamps.sort()
        earliest = normalized_timestamps[0]
        latest = normalized_timestamps[-1]
        span = latest - earliest
        print(f"Earliest Email:    {earliest}")
        print(f"Latest Email:      {latest}")
        print(f"Time Span:         {span}")
    else:
        print("Time Span:         Could not calculate (invalid date headers)")
    sys.stdout.flush()

def forward_to_spamcop(downloaded_files, spam_candidates, total_size_bytes):
    """Handles forwarding emails to SpamCop with first-run confirmation"""
    # Check if this is the first run
    first_run_flag_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), FIRST_RUN_FLAG_FILE)
    is_first_run = not os.path.exists(first_run_flag_file) and not SIMULATION_MODE
    
    if is_first_run:
        if not _handle_first_run_confirmation(first_run_flag_file, downloaded_files, spam_candidates):
            return
    
    # Normal operation (after first run)
    if not is_first_run:
        if SIMULATION_MODE:
            print("\n" + "-"*40)
            print(f"SIMULATION MODE: Would bundle {len(downloaded_files)} files and send to SpamCop...")
            sys.stdout.flush()
        else:
            print("\n" + "-"*40)
            print(f"Bundling {len(downloaded_files)} files and sending to SpamCop...")
            sys.stdout.flush()
    
    # Check simulation mode
    if SIMULATION_MODE:
        _print_simulation_mode_info(downloaded_files)
    else:
        success = _send_to_spamcop(downloaded_files)
        # If sending was successful, save UIDs to prevent re-sending
        if success and spam_candidates:
            sent_uids = [candidate['uid'] for candidate in spam_candidates if 'uid' in candidate]
            if sent_uids:
                add_sent_uids(sent_uids)
                print(f"Marked {len(sent_uids)} email(s) as sent to prevent duplicate forwarding.")
                sys.stdout.flush()

def _handle_first_run_confirmation(first_run_flag_file, downloaded_files, spam_candidates):
    """Handles first-run user confirmation and returns True if confirmed"""
    download_path = os.path.dirname(downloaded_files[0]) if downloaded_files else ""
    
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
    sys.stdout.flush()
    
    for idx, item in enumerate(spam_candidates, 1):
        display_subject = safe_print_subject(item['subject'], 60)
        print(f"  {idx}. {display_subject}")
        if item['date']:
            print(f"     Date: {item['date']}")
        sys.stdout.flush()
    
    print("\n" + "="*70)
    print("VERIFICATION STEPS:")
    print("="*70)
    print(msg.FIRST_RUN_VERIFICATION_STEPS)  # type: ignore
    print("="*70)
    print()
    sys.stdout.flush()
    
    while True:
        response = input("Have you verified that ALL downloaded emails are spam? (yes/no): ").strip().lower()
        if response in ['yes', 'y']:
            print("\n[OK] Confirmation received. Proceeding with forwarding to SpamCop...")
            sys.stdout.flush()
            try:
                with open(first_run_flag_file, 'w') as f:
                    f.write(f"First run completed: {datetime.datetime.now().isoformat()}\n")
            except Exception as e:
                print(f"Warning: Could not create first-run flag file: {e}")
                sys.stdout.flush()
            return True
        elif response in ['no', 'n']:
            print("\n" + "="*70)
            print("FORWARDING CANCELLED")
            print("="*70)
            print(msg.FIRST_RUN_CANCELLED.format(download_path=download_path))  # type: ignore
            print("="*70)
            print("\n--- ITERATION COMPLETE (NO EMAILS FORWARDED) ---")
            sys.stdout.flush()
            return False
        else:
            print("Please enter 'yes' or 'no'.")
            sys.stdout.flush()

def _print_simulation_mode_info(downloaded_files):
    """Prints simulation mode information"""
    print("\n" + "="*70)
    print("SIMULATION MODE ENABLED - Email will NOT be sent to SpamCop")
    print("="*70)
    print(f"Would send email:")
    print(f"  From: {GMAIL_ACCOUNT}")
    print(f"  To: {SPAMCOP_ADDRESS}")
    print(f"  Subject: Spam Report: {len(downloaded_files)} messages")
    print(f"  Attachments: {len(downloaded_files)} EML files")
    print()
    sys.stdout.flush()
    for filepath in downloaded_files:
        filename = os.path.basename(filepath)
        file_size = os.path.getsize(filepath)
        print(f"    - {filename} ({get_size_str(file_size)})")
        sys.stdout.flush()
    print()
    print("All steps completed successfully (connect, search, download, save).")
    print("Set SIMULATION_MODE = False in config.py to enable actual forwarding.")
    print("="*70)
    sys.stdout.flush()

def _send_to_spamcop(downloaded_files):
    """Sends emails to SpamCop via SMTP. Returns True if successful, False otherwise."""
    print(f"\nConnecting to Gmail SMTP ({SMTP_SERVER})...")
    sys.stdout.flush()
    try:
        # Construct the Email
        msg = MIMEMultipart()
        msg['From'] = GMAIL_ACCOUNT
        msg['To'] = SPAMCOP_ADDRESS
        msg['Subject'] = f"Spam Report: {len(downloaded_files)} messages"
        
        # Attach files
        print("Attaching files...")
        sys.stdout.flush()
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
        sys.stdout.flush()
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(GMAIL_ACCOUNT, APP_PASS)
        server.send_message(msg)
        server.quit()
        
        print("\nSUCCESS: Report sent to SpamCop.")
        sys.stdout.flush()
        return True
        
    except Exception as e:
        print(f"\nFAILED to send email: {e}")
        print("The files are still safe in your local folder.")
        sys.stdout.flush()
        import traceback
        traceback.print_exc()
        return False

def run_spam_processor():
    """Main loop that runs spam processing continuously"""
    # Determine if this is the first run using internal checks (no flag files)
    is_first_run = is_initial_run_internal()
    
    # Track if initial setup has been completed in this session
    initial_setup_completed = False
    
    # Validate configuration
    try:
        validate_loop_frequency(LOOP_FREQUENCY_HOURS)
        validate_search_window(SPAM_SEARCH_WINDOW_HOURS)
        validate_keyword_conflicts()  # Re-validate in case config was changed
    except ValueError as e:
        print(f"CONFIGURATION ERROR: {e}")
        if "LOOP_FREQUENCY_HOURS" in str(e):
            print("Please fix LOOP_FREQUENCY_HOURS in the configuration section.")
        elif "SPAM_SEARCH_WINDOW_HOURS" in str(e):
            print("Please fix SPAM_SEARCH_WINDOW_HOURS in the configuration section.")
        sys.stdout.flush()
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
    sys.stdout.flush()
    
    if is_first_run:
        print("FIRST RUN DETECTED - Will abort if spam folders are not found")
        print("="*60)
        sys.stdout.flush()
    
    print("Starting first iteration immediately...\n")
    sys.stdout.flush()
    
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
        sys.stdout.flush()
        
        try:
            process_spam_iteration(is_first_run=is_first_run)
            
            # After successful first run iteration, mark initial setup as completed
            if is_first_run and not initial_setup_completed:
                # Check if the iteration was successful by verifying we got past folder checking
                # If we're here, it means we successfully:
                # - Connected to IMAP
                # - Listed folders
                # - Displayed folder counts (or would have if not aborted)
                # - Selected spam folder (or would have if not aborted)
                print(f"\n[OK] Initial setup completed successfully.")
                print("The script has successfully connected and verified folder access.")
                sys.stdout.flush()
                initial_setup_completed = True
                is_first_run = False  # No longer first run (for this session)
                    
        except KeyboardInterrupt:
            print("\n\nReceived interrupt signal. Shutting down gracefully...")
            sys.stdout.flush()
            cleanup_logging()
            break
        except imaplib.IMAP4.error as e:
            # If it's the first run and we couldn't find spam folders, abort to commandline
            if is_first_run and ("Failed to select spam mailbox" in str(e) or "SPAM mailbox" in str(e) or "spam mailbox" in str(e).lower()):
                print("\n" + "="*70)
                print("FIRST RUN FAILED: Could not find spam folders")
                print("="*70)
                print("The script is aborting to commandline as this is the first run.")
                print("Please verify your Gmail account settings and spam folder configuration.")
                print("="*70)
                sys.stdout.flush()
                cleanup_logging()
                sys.exit(1)
            else:
                print(f"\nERROR in iteration #{iteration_count}: {e}")
                print("Continuing to next iteration...")
                sys.stdout.flush()
                import traceback
                traceback.print_exc()
        except Exception as e:
            # If it's the first run and we had an error finding spam folders, abort
            error_str = str(e).lower()
            if is_first_run and ("spam" in error_str and ("folder" in error_str or "mailbox" in error_str)):
                print("\n" + "="*70)
                print("FIRST RUN FAILED: Error during spam folder detection")
                print("="*70)
                print(f"Error: {e}")
                print("The script is aborting to commandline as this is the first run.")
                print("Please verify your Gmail account settings and spam folder configuration.")
                print("="*70)
                sys.stdout.flush()
                import traceback
                traceback.print_exc()
                cleanup_logging()
                sys.exit(1)
            else:
                print(f"\nERROR in iteration #{iteration_count}: {e}")
                print("Continuing to next iteration...")
                sys.stdout.flush()
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
            cleanup_logging()
            break

if __name__ == "__main__":
    run_spam_processor()