import subprocess
import time
import win32com.client
import logging

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

try:
    # Step 1: Open SAP Logon
    sap_path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"
    logging.info("Starting SAP Logon...")
    subprocess.Popen(sap_path)
    time.sleep(5)  # Wait for SAP Logon to open

    # Step 2: Connect to SAP GUI
    logging.info("Connecting to SAP GUI...")
    sap = win32com.client.GetObject("SAPGUI")
    app = sap.GetScriptingEngine
    connection = app.OpenConnection("LIGER Production [EP1]", True)
    session = connection.Children(0)

    # Step 3: Enter SAP Login Credentials
    logging.info("Entering SAP credentials...")
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "Your ID"
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "Your Password"

    # Step 4: Click Logon
    logging.info("Logging into SAP...")
    session.findById("wnd[0]").sendVKey(0)
    logging.info("Successfully logged into SAP LIGER Production [EP1]")

except Exception as e:
    logging.error("SAP Logon error: %s", str(e))
    print(f"An error occurred. Check the log file: {log_file}")
    exit()  # Stop execution if SAP login fails

try:
    # Step 5: Open Transaction Code ZMMR02
    logging.info("Entering transaction: ZMMR02")
    session.findById("wnd[0]/tbar[0]/okcd").text = "zmmr02"
    session.findById("wnd[0]").sendVKey(0)

    # Step 6: Enter First Set of Parameters
    logging.info("Entering first parameter set (15*, 21*)...")
    session.findById("wnd[0]/usr/btn%_S_BANFN_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "15*"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "21*"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Step 7: Enter Date and Material Number Selection
    logging.info("Entering selection criteria...")
    session.findById("wnd[0]/usr/ctxtS_BADAT-LOW").text = "01.10.2024"
    session.findById("wnd[0]/usr/ctxtS_BADAT-HIGH").text = "31.12.2024"
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = "200000000"
    session.findById("wnd[0]/usr/ctxtS_MATNR-HIGH").text = "499999999"
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = "7200"

    # Step 8: Enter Second Parameter Set (2001, 2002)
    logging.info("Selecting additional parameters (2001, 2002)...")
    session.findById("wnd[0]/usr/btn%_S_LGORT2_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "2001"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "2002"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    # Step 9: Select Layout Before Exporting the Data
    logging.info("Selecting a predefined layout...")
    session.findById("wnd[0]/tbar[1]/btn[33]").press()  # Open "Select Layout" (Ctrl + F8)
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 2
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 0
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()

    # Step 10: Save the Exported Data
    logging.info("Exporting data to file...")
    session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 0
    session.findById("wnd[1]").sendVKey(4)
    session.findById("wnd[2]/usr/ctxtDY_PATH").setFocus()
    session.findById("wnd[2]/usr/ctxtDY_PATH").caretPosition = 42
    session.findById("wnd[2]").sendVKey(4)
    session.findById("wnd[3]/usr/ctxtDY_PATH").text = "D:\\My\\Project\\KPI ยา เวช automate"
    session.findById("wnd[3]/usr/ctxtDY_FILENAME").text = "KPI_mspha_data.xlsx"
    session.findById("wnd[3]/usr/ctxtDY_FILENAME").caretPosition = 19
    session.findById("wnd[3]/tbar[0]/btn[0]").press()

    logging.info("Data extraction completed successfully and saved.")

except Exception as e:
    logging.error("Data extraction error: %s", str(e))
    print(f"An error occurred. Check the log file: {log_file}")
