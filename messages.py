"""
User instruction messages for SpamCop Forwarder
This file contains all the large text blocks for user instructions
"""

# Configuration setup instructions
GMAIL_ACCOUNT_INSTRUCTIONS = """
   A. Gmail Account (GMAIL_ACCOUNT):
      - This is simply your Gmail email address
      - Example: 'yourname@gmail.com'
      - Use the Gmail account that receives spam emails
"""

APP_PASSWORD_INSTRUCTIONS = """
   B. Gmail App Password (APP_PASS):
      Step 1: Enable 2-Step Verification
        1. Go to: https://myaccount.google.com/security
        2. Sign in with your Google account
        3. Under 'Signing in to Google', find '2-Step Verification'
        4. Click '2-Step Verification' and follow the prompts to enable it
        5. You'll need a phone number for verification
      
      Step 2: Generate App Password
        1. After enabling 2-Step Verification, go to:
           https://myaccount.google.com/apppasswords
        2. You may be asked to sign in again
        3. Under 'Select app', choose 'Mail'
        4. Under 'Select device', choose 'Other (Custom name)'
        5. Type a name like 'SpamCop Forwarder' and click 'Generate'
        6. Google will show you a 16-character password
        7. Copy this password EXACTLY as Google displays it
        8. The password will look like: 'abcd efgh ijkl mnop' (with spaces)
        9. In config.py, you can enter it with or without spaces
        10. The script will automatically remove spaces when using it
      
      Important Notes:
        - App passwords are 16 characters long
        - You can paste the password with spaces (as Google shows it)
        - Spaces will be automatically removed by the script
        - Each app password can only be viewed once
        - If you lose it, generate a new one
        - App passwords are different from your regular Gmail password
"""

SPAMCOP_ADDRESS_INSTRUCTIONS = """
   Step 1: Create a SpamCop Account
      1. Go to: https://www.spamcop.net/anonsignup.shtml
      2. Fill out the registration form
      3. Verify your email address when prompted
      4. Complete the account setup process
   
   Step 2: Log in to SpamCop
      1. Go to: https://www.spamcop.net/
      2. Click 'Login' and sign in with your credentials
   
   Step 3: Obtain Your Quick Send Address
      1. After logging in, go to: https://www.spamcop.net/mcgi?action=setup
      2. Look for the 'Quick Submit' section
      3. You'll see your Quick Send address, which looks like:
         'quick.xxxxxxxxxxxxx@spam.spamcop.net'
      4. Copy this entire address
      5. In config.py, set SPAMCOP_ADDRESS to this value
   
   Alternative Method:
      1. Log in to SpamCop
      2. Go to: https://www.spamcop.net/mcgi?action=account
      3. Look for 'Quick Submit Email Address' or similar
      4. Copy the address shown
   
   Important Notes:
     - The Quick Send address is unique to your account
     - It allows you to forward spam emails directly
     - Keep this address private
     - Format: quick.XXXXXXXXXXXXX@spam.spamcop.net
"""

SMTP_CONFIG_INFO = """
   SMTP_SERVER = 'smtp.gmail.com' (standard Gmail SMTP server)
   SMTP_PORT = 587 (standard Gmail SMTP port with STARTTLS)
   These values are correct for all Gmail accounts and should not be changed.
"""

# First-run verification messages
FIRST_RUN_HEADER = """
⚠️  IMPORTANT: This is the first time the script has downloaded emails.
   You MUST manually verify that ALL downloaded emails are spam before
   forwarding them to SpamCop.
"""

FIRST_RUN_WARNING = """
   Forwarding legitimate emails to SpamCop will:
   • Flag legitimate senders as fraudulent/spammers
   • May cause legitimate senders to be blacklisted
   • May prevent you from receiving legitimate emails in the future
   • Can damage your relationship with legitimate businesses/services
"""

FIRST_RUN_VERIFICATION_STEPS = """
1. Open your Gmail account in a web browser
2. Navigate to the SPAM folder
3. Compare the emails listed above with your Gmail Spam folder
4. Verify that ALL emails shown are indeed spam
5. Double-check that NO legitimate emails are in the list
6. If you find ANY legitimate emails, answer 'no' below
"""

FIRST_RUN_CANCELLED = """
The downloaded emails remain in your local folder:
  {download_path}

Please review the emails and remove any legitimate ones.
You can manually delete legitimate emails from the download folder.
After cleaning, you can run the script again.
"""

def get_config_instructions(missing_fields):
    """Returns formatted configuration instructions based on missing fields"""
    instructions = []
    
    if 'GMAIL_ACCOUNT' in missing_fields or 'APP_PASS' in missing_fields:
        instructions.append("\n1. GMAIL ACCOUNT AND APP PASSWORD SETUP")
        instructions.append("-" * 70)
        if 'GMAIL_ACCOUNT' in missing_fields:
            instructions.append(GMAIL_ACCOUNT_INSTRUCTIONS)
        if 'APP_PASS' in missing_fields:
            instructions.append(APP_PASSWORD_INSTRUCTIONS)
    
    if 'SPAMCOP_ADDRESS' in missing_fields:
        instructions.append("\n2. SPAMCOP QUICK SEND ADDRESS SETUP")
        instructions.append("-" * 70)
        instructions.append(SPAMCOP_ADDRESS_INSTRUCTIONS)
    
    return "\n".join(instructions)

