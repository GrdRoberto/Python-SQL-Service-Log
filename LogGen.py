import os
import pandas as pd
import pyodbc
import zipfile
from datetime import datetime, timedelta
import logging
import time

# Get yesterday's date to use it in the log
yesterday = datetime.now() - timedelta(days=1)

# Configure logging system with yesterday's date
log_directory = "location logs"

logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'  # Log message format
)

def adjust_time(base_time, hours_offset):
    """Adjusts `base_time` by the specified number of hours."""
    return base_time + timedelta(hours=hours_offset)

def main():
    server = r'SQLSERVER'
    database = 'NAME'
    username = 'User'
    password = 'Password'

    today = datetime.now()
    adjusted_time = adjust_time(today, 0)  # Adjust for the desired timezone
    yesterday = adjusted_time - timedelta(days=1)

    txt_file_name = os.path.join(log_directory, f"{yesterday.strftime('%Y_%m_%d')}.txt")
    zip_file_name = os.path.join(log_directory, f"{yesterday.strftime('%Y_%m_%d')}.zip")

    # Check if the zip file already exists
    if os.path.exists(zip_file_name):
        logging.error(f"Log file for {yesterday.strftime('%Y-%m-%d')} already exists. Skipping creation.")
        return  # Exit the function early if the log file exists

    # Create log directory if it doesn't exist
    if not os.path.exists(log_directory):
        os.makedirs(log_directory)

    user_query = f"""
    insert query
WHERE DataOra >= '{yesterday.strftime('%Y%m%d 00:00')}' AND DataOra <= '{yesterday.strftime('%Y%m%d 23:59')}'
ORDER BY [ID] ASC
"""

    while True:
        try:
            conn = pyodbc.connect(
                f'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}')
            df = pd.read_sql(user_query, conn)

            if not df.empty:
                df.to_csv(txt_file_name, index=False, sep='\t')
                zip_txt_file(txt_file_name, zip_file_name)
                os.remove(txt_file_name)

            break  # Exit the loop when the connection is successful

        except Exception as e:
            logging.error(f"Connection failed: {e}")
            time.sleep(5)  # Wait for 5 seconds in case of failure

        finally:
            if 'conn' in locals():
                conn.close()

def zip_txt_file(txt_file, zip_file):
    """Create a ZIP file from the TXT file."""
    with zipfile.ZipFile(zip_file, 'w', compression=zipfile.ZIP_BZIP2) as zipf:
        zipf.write(txt_file, os.path.basename(txt_file))

if __name__ == '__main__':
    while True:
        now = datetime.now()
        if now.hour == 1 and now.minute == 0:
            main()
            time.sleep(60)  # Wait for 60 seconds
        elif now.hour > 1 or (now.hour == 1 and now.minute > 0):
            main()
            time.sleep(60)  # Wait for 60 seconds
        else:
            time.sleep(30)  # Check every 30 seconds