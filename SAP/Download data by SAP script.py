import subprocess
import time
import win32com.client
import logging
import os
from datetime import datetime

# Setup logging to both file and console
log_file = "sap_automation.log"

# Create logging formatter
formatter = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")

# File Handler (Save logs to a file)
file_handler = logging.FileHandler(log_file)
file_handler.setFormatter(formatter)

# Console Handler (Print logs in CMD)
console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)

# Setup Logging
logging.basicConfig(level=logging.DEBUG, handlers=[file_handler, console_handler])

logging.info("SAP automation script started.")

# Step 1: Determine the Fiscal Year and Current Quarter
today = datetime.today()
year = today.year

# Fiscal year starts in October
if today.month >= 10:
    fiscal_year = year  # If today is October or later, we're in the new fiscal year
else:
    fiscal_year = year - 1  # Otherwise, we're still in the previous fiscal year

# Define fiscal quarters with proper fiscal year allocation
quarters = {
    1: (f"01.10.{fiscal_year}", f"31.12.{fiscal_year}"),  # Q1: Oct 1 - Dec 31
    2: (f"01.01.{fiscal_year + 1}", f"31.03.{fiscal_year + 1}"),  # Q2: Jan 1 - Mar 31
    3: (f"01.04.{fiscal_year + 1}", f"30.06.{fiscal_year + 1}"),  # Q3: Apr 1 - Jun 30
    4: (f"01.07.{fiscal_year + 1}", f"30.09.{fiscal_year + 1}"),  # Q4: Jul 1 - Sep 30
}

# Determine the current quarter based on today's month
if today.month >= 10:  # October - December
    current_quarter = 1
elif today.month >= 7:  # July - September
    current_quarter = 4
elif today.month >= 4:  # April - June
    current_quarter = 3
else:  # January - March
    current_quarter = 2

logging.info(f"Fiscal Year: {fiscal_year}, Downloading Data from Q1 to Q{current_quarter}")

try:
    # Step 2: Open SAP Logon
    sap_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    logging.info("Starting SAP Logon...")
    subprocess.Popen(sap_path)
    time.sleep(5)  # Wait for SAP Logon to open

    # Step 3: Connect to SAP GUI
    logging.info("Connecting to SAP GUI...")
    sap = win32com.client.GetObject("SAPGUI")
    app = sap.GetScriptingEngine
    connection = app.OpenConnection("LIGER Production [EP1]", True)
    session = connection.Children(0)

    # Step 4: Enter SAP Login Credentials
    logging.info("Entering SAP credentials...")
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "Your ID"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Your Password"

    # Step 5: Click Logon
    logging.info("Logging into SAP...")
    session.findById("wnd[0]").sendVKey(0)
    logging.info("Successfully logged into SAP LIGER Production [EP1]")

except Exception as e:
    logging.error("SAP Logon error: %s", str(e))
    print(f"An error occurred. Check the log file: {log_file}")
    exit()  # Stop execution if SAP login fails

try:
    for quarter in range(1, current_quarter+1):  # Loop from Q1 to the current quarter //for quarter in range(1, current_quarter + 1):
        start_date, end_date = quarters[quarter]
        logging.info(f"Processing Fiscal Year {fiscal_year}, Quarter {quarter} ({start_date} - {end_date})")

        # Step 6: Open Transaction Code ZMMR02
        logging.info("Entering transaction: ZMMR02")
        session.findById("wnd[0]/tbar[0]/okcd").text = "zmmr02"
        session.findById("wnd[0]").sendVKey(0)

        # Step 7: Enter First Set of Parameters
        logging.info("Entering first parameter set (15*, 21*)...")
        session.findById("wnd[0]/usr/btn%_S_BANFN_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "15*"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "21*"
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # Step 8: Enter Date and Material Number Selection
        logging.info("Entering selection criteria...")
        session.findById("wnd[0]/usr/ctxtS_BADAT-LOW").text = start_date
        session.findById("wnd[0]/usr/ctxtS_BADAT-HIGH").text = end_date
        session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = "200000000"
        session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").text = "499999999"
        session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "7200"

        # Step 9: Enter Second Parameter Set (2001, 2002)
        logging.info("Selecting additional parameters (2001, 2002)...")
        session.findById("wnd[0]/usr/btn%_S_LGORT2_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "2001"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "2002"
        session.findById("wnd[1]/tbar[0]/btn[8]").press()

        # Step 10: Execute the Report
        logging.info("Executing report...")
        session.findById("wnd[0]/tbar[1]/btn[8]").press()

        # Step 11: Select Layout Before Exporting the Data
        logging.info("Selecting a predefined layout...")
        session.findById("wnd[0]/tbar[1]/btn[33]").press()
        shell = session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
        shell.currentCellRow = 2
        shell.firstVisibleRow = 0
        shell.selectedRows = "2"
        shell.clickCurrentCell()
        logging.info("Layout selection completed successfully.")
        time.sleep(5)

        # Step 12: Save the Exported Data
        logging.info("Saving exported data...")
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(2)
        session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = "D:\\My\\Project\\KPI_mspha_automate\\"
        time.sleep(1)

        export_filename = f"KPI_mspha_{fiscal_year}_Q{quarter}.XLSX"
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").setFocus()
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = export_filename
        time.sleep(1)
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(10)

        logging.info(f"Completed processing for Q{quarter}. Moving to the next quarter...\n")

        # Step 14: Restart the loop (Re-enter transaction)
        logging.info("Restarting transaction ZMMR02 for the next quarter...")
        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # Press "Back"
        session.findById("wnd[0]/tbar[0]/btn[15]").press()  # Press "Back"

except Exception as e:
    logging.error("Data extraction error: %s", str(e))
    print(f"An error occurred. Check the log file: {log_file}")

logging.info("Closing all running Excel instances...")
os.system("taskkill /f /im excel.exe")  # Force close all Excel instances
time.sleep(3)  # Ensure Excel is completely closed
logging.info("All Excel instances closed successfully.")

logging.info("Ensuring SAP is completely closed...")
os.system("taskkill /f /im saplogon.exe")  # Force close SAP GUI
os.system("taskkill /f /im sapgui.exe")    # Ensure SAP GUI is fully terminated
time.sleep(3)
