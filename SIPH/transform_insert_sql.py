import pandas as pd
from datetime import datetime, time, timedelta
import openpyxl
import re
import pyodbc
import logging
from logging.handlers import RotatingFileHandler
import urllib
import win32com.client
import os
from pathlib import Path
from sqlalchemy import create_engine, event, text


excel_file = r"\\siphvmdata01\opd_stat_time\data\data.xlsx"

# --------------------------- Logging Setup ---------------------------
log_path = Path(r"\\siphvmdata01\opd_stat_time\logs\opd_stat_time.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(str(log_path), mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ],
    force=True
)

def refresh_excel(excel_file):
    logging.info("Starting Process Refresh Excel File")
    abs_path = os.path.abspath(excel_file)

    excel = None
    workbook = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.UserControl = False

        logging.info(f"Processing Refresh: {abs_path}")

        workbook = excel.Workbooks.Open(abs_path)
        workbook.RefreshAll()

        excel.CalculateUntilAsyncQueriesDone()

        workbook.Save()
        logging.info("Refresh Complete!")

    except Exception as e:
        logging.error(f"Error: {e}")

    finally:
        if workbook is not None:
            workbook.Close(False)
        if excel is not None:
            excel.Quit()

refresh_excel(excel_file)

# --------------------------- Core Processing ---------------------------
try:
    logging.info("Starting Core Processing")

    # --- 1. Load Data ---
    try:
        df = pd.read_excel(excel_file, sheet_name='data', converters={'CaseNo':int,'Med_Number':int}, engine='openpyxl')
        df1 = pd.read_excel(excel_file, sheet_name='Department', engine='openpyxl')
        df2 = pd.read_excel(excel_file, sheet_name='Clinic', engine='openpyxl')
        df3 = pd.read_excel(excel_file, sheet_name='Form Responses 1', engine='openpyxl')
    except FileNotFoundError:
        logging.error(f"Error: Not found file {excel_file}")
        raise
    except ValueError as e:
        logging.error(f"Error:{e}")
        raise

    # --- 2. Data Transformation (Date/Time) ---
    df3['วันที่'] = pd.to_datetime(df3['วันที่'], errors='coerce')
    
    first_day_this_month = pd.Timestamp.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    first_day_last_month = first_day_this_month - pd.offsets.MonthBegin(1)
    last_day_last_month = first_day_this_month - pd.Timedelta(seconds=1)

    df3 = df3[df3['วันที่'].between(first_day_last_month, last_day_last_month)].copy()

    # --- 3. Create Keys & Merging ---
    logging.info("Creating Keys and Merging DataFrames")
   
    df['time'] = pd.to_datetime(df['New']).dt.time
    df['New'] = pd.to_datetime(df['New']).dt.date
    df['key_department'] = df['Med_Number'].astype(str) + '_' + df['Department'].astype(str)
    df['key_clinic'] = df['Med_Number'].astype(str) + '_' + df['Clinic-Ward'].astype(str)

    # --- 4. Lookup Tables ---
    df1_lookup = (df1[['key', 'Material Description','type']]
                  .drop_duplicates(subset=['key'])
                  .rename(columns={'key': 'key_department', 'Material Description': 'Mat_Department','type':'Type_Department'}))

    df2_lookup = (df2[['key', 'Material Description','type']]
                  .drop_duplicates(subset=['key'])
                  .rename(columns={'key': 'key_clinic', 'Material Description': 'Mat_Clinic','type':'Type_Clinic'}))

    merged_df = df.merge(df1_lookup, on='key_department', how='left') \
                  .merge(df2_lookup, on='key_clinic', how='left')

    # --- 5. Filtering & Splitting ---
    logging.info("Filtering DataFrames")
    floor_list = df['Department'].dropna().unique().tolist()
    clinic_list = df['Clinic-Ward'].dropna().unique().tolist()

    is_floor = merged_df['Department'].isin(floor_list) & merged_df['Mat_Department'].notna()
    is_clinic = merged_df['Clinic-Ward'].isin(clinic_list) & merged_df['Mat_Clinic'].notna()

    df_remaining = merged_df[~(is_floor | is_clinic)].copy()
    df_excluded = merged_df[is_floor | is_clinic].copy()

    # --- 6. Calculate Summary Times ---
    logging.info("Calculating Summary Times")
    df3_lookup = (df3.dropna(subset=['VN'])
                  .drop_duplicates(subset='VN', keep='first')
                  [['VN', 'เวลาปลายทางได้รับ', 'ส่งทาง']])

    df_calc = (df_remaining.merge(df3_lookup, how='left', left_on='CaseNo', right_on='VN')
               .query("ส่งทาง.notna() and VN.notna()"))
    
    try:
        df_calc = df_calc.assign(
            time_td = lambda x: pd.to_timedelta(x['time'].astype(str)),
            Summary = lambda x: pd.to_timedelta(x['เวลาปลายทางได้รับ'].astype(str).str.split().str[-1]) - 
                               pd.to_timedelta(x['time'].astype(str))
        )
    except Exception as e:
        logging.warning(f"Warning: การคำนวณ Summary ผิดพลาด (ตรวจสอบ Format เวลา) - {e}")
        df_calc['Summary'] = pd.NaT

    # --- 7. Final Clean up ---
    df_final = (df_calc[df_calc['Summary'] >= pd.Timedelta(0)]
                .sort_values(by=['CaseNo', 'New', 'Summary'])
                .drop_duplicates(subset=['CaseNo', 'New'], keep='first')
                .drop(columns=['Mat_Department', 'Mat_Clinic', 'VN', 'time_td'], errors='ignore'))

    df_15min = df_final[df_final['Summary'] <= pd.Timedelta(minutes=15)]

    new_columns = ["MK", "HN", "CaseNo", "Med_Number", "Med_Description", "OrderID", "Priority", 
                   "Type", "Department", "Clinic-Ward", "User", "New", "time", "Active", 
                   "Final", "Sum of New_to_Active_minutes", "Sum of Active_to_Final_minutes", 
                   "Sum of New_to_Final_minutes", "เวลาปลายทางได้รับ", "Summary", "ส่งทาง"]

    def reorder_columns(df, cols):
        return df.reindex(columns=[c for c in cols if c in df.columns])

# Rename for SQL compatibility
    column_mapping = {
            'Priority': 'Med_Priority',
            'Type': 'Med_Type',
            'User': 'User_Staff',
            'New': 'New_Date',
            'time': 'New_Time',
            'Clinic-Ward': 'Clinic_Ward',
            'เวลาปลายทางได้รับ': 'Received_Time',
            'Summary': 'Summary_Interval',
            'ส่งทาง': 'Transport_Method'
        }
    
    df_final = reorder_columns(df_final, new_columns).rename(columns=column_mapping)
    df_excluded = reorder_columns(df_excluded, new_columns).rename(columns=column_mapping)

    df_final['is_excluded'] = 0
    df_excluded['is_excluded'] = 1
    
    df_final['Summary_Interval'] = df_final['Summary_Interval'].dt.total_seconds().fillna(0).astype(int)

    df_total = pd.concat([df_final, df_excluded], ignore_index=True)
    # --- 8. Output ---
    target = len(df_15min) + len(df_excluded)
    overall = len(df_final) + len(df_excluded)

    # print(f"Date >> {first_day_last_month.strftime('%d/%m/%Y')}")
    # print(f"Target >> {target}")
    # print(f"Overall >> {overall}")

    data = {
        'Report_Date': [first_day_last_month]
        ,'Target_Count': [target]
        ,'Overall_Count': [overall]
    }
    
    df_summary = pd.DataFrame(data)
    logging.info("Transformation Complete!")

except KeyError as e:
    logging.error(f"Error: {e}")
except Exception as e:
    logging.error(f"Unexpected Error:{e}")

# ---------------- SQL Insertion ---------------------------

SERVER = "xxx"
DATABASE = "xxx"
USERNAME = "xxx"
PASSWORD = "xxx"

quoted_password = urllib.parse.quote_plus(PASSWORD)

# --- Checking ---
date_str = first_day_last_month.strftime('%Y-%m-%d')
check_query = text("SELECT TOP 1 report_date FROM med_stat_summary WHERE report_date = :d")

# --- Setup Connection String ---
connection_string = (
    f"mssql+pyodbc://{USERNAME}:{quoted_password}@{SERVER}/{DATABASE}"
    "?driver=ODBC+Driver+17+for+SQL+Server"
)
engine = create_engine(connection_string)

# --- fast_executemany ---
@event.listens_for(engine, "before_cursor_execute")
def receive_before_cursor_execute(conn, cursor, statement, parameters, context, executemany):
    if executemany:
        cursor.fast_executemany = True

# --- Execute Truncate and Insert ---
try:
    with engine.begin() as conn:
        logging.info("Truncating table: med_stat")
        conn.execute(text("TRUNCATE TABLE dbo.med_stat"))
        
        logging.info("Inserting new data...med_stat")
        df_total.to_sql('med_stat', con=conn, if_exists='append', index=False)
    with engine.connect() as conn_read:
        existing_data = pd.read_sql(check_query, conn_read, params={'d': date_str})

    if existing_data.empty:
        logging.info(f"No existing data for {date_str}. Inserting new data...med_stat_summary")
        
        with engine.begin() as conn_summary:
            df_summary.to_sql('med_stat_summary', con=conn_summary, if_exists='append', index=False)
            logging.info(f"Success: Data for {date_str} inserted into summary.")
    else:
        logging.info(f"Skip: Data for {date_str} already exists. Skipping insertion.")
        
    logging.info("Process Complete successfully.")

except Exception as e:
    logging.error(f"SQL Error: {e}")
