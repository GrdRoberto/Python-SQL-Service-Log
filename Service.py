#########################       SETUP       #####################
#   RUN THE FOLLOWING COMMAND IN CMD WITH ADMINISTRATOR:
#       pip install pywin32 pandas pyodbc
#================================================================
#       HOW TO CREATE THE SERVICE:
#       python myservice.py install
#       python myservice.py start
#       python myservice.py remove

import os
import pandas as pd
import pyodbc
import zipfile
import win32serviceutil
import win32service
import win32event
from datetime import datetime, timedelta
import logging
import time

# Set up logging
log_directory = "C:\\Logs" # Define default log folder location
logging.basicConfig(
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Define the service class
class LogService(win32serviceutil.ServiceFramework):
    _svc_name_ = "LogPythonSQL"
    _svc_display_name_ = "LogPythonSQL"
    _svc_description_ = "A service that extracts log data from SQL Server daily."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        win32event.SetEvent(self.stop_event)
        self.running = False

    def SvcRun(self):
        while self.running:
            now = datetime.now()
            if now.hour == 11 and now.minute == 41:  # Adjust time as needed
                self.main()
                time.sleep(60)  # Wait for a minute after execution
            else:
                time.sleep(30)  # Check every 30 seconds

    def adjust_time(self, base_time, hours_offset):
        return base_time + timedelta(hours=hours_offset)

    def main(self):
        server = r'DBAdress' # SQL DB adress
        database = 'DBname' #DB name
        username = 'user'
        password = 'pass'

        today = datetime.now()
        adjusted_time = self.adjust_time(today, 0)
        yesterday = adjusted_time - timedelta(days=1)

        txt_file_name = os.path.join(log_directory, f"{yesterday.strftime('%Y_%m_%d')}.txt")
        zip_file_name = os.path.join(log_directory, f"{yesterday.strftime('%Y_%m_%d')}.zip")

        # Create log directory if it doesn't exist
        if not os.path.exists(log_directory):
            os.makedirs(log_directory)

        user_query = f"""	
            #Insert Query
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
                    self.zip_txt_file(txt_file_name, zip_file_name)
                    os.remove(txt_file_name)

                break  # Exit the loop when the connection is successful

            except Exception as e:
                logging.error(f"Connection failed: {e}")
                time.sleep(5)

            finally:
                if 'conn' in locals():
                    conn.close()

    def zip_txt_file(self, txt_file, zip_file):
        with zipfile.ZipFile(zip_file, 'w') as zipf:
            zipf.write(txt_file, os.path.basename(txt_file))

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(LogService)