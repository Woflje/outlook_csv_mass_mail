import os
from dotenv import load_dotenv

load_dotenv()

### Config

action = "send" # "send" or "draft"

waitForInput = True

replacers = {
    # All mail
    'consistent': {
    },
    # Contact specific
    'csv': {
        '{name}': os.getenv('CSV_FILE_NAME_COLUMN_LABEL'),
        '{nickname}': os.getenv('CSV_FILE_NICKNAME_COLUMN_LABEL'),
    }
}

encodings = ['utf-8', 'utf-16', 'iso-8859-1', 'windows-1252']