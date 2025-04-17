import pandas as pd
import dbf
import datetime
from dbfread import DBF
import pyodbc

# SQL Server connection details
server = "xxxxx"
database = "xxxxx"
username = "xxxxx"
password = "xxxxx"
view_name = "xxxxx"  # Use the name of your view here

# Connect to SQL Server
try:
    conn = pyodbc.connect(
        f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
    )
    print("Connection successful!")

    query = f"SELECT * FROM {view_name}"
    df = pd.read_sql_query(query, conn)

    # Close connection
    conn.close()
except Exception as e:
    print(f"Process failed: {e}")

# Optional: Replace "NaN" strings with empty values
df.replace({pd.NA: '', 'nan': '', 'NaN': ''}, inplace=True)

# Define the DBF schema
table_structure = (
    'DOB D; Sex C(1); DateAdm D; TimeAdm C(4); DateDsc D; TimeDsc C(4); Discht C(1); AdmWt N(7,3); Age C(3); AgeDay C(3); '
    'PDx C(6); SDx1 C(6); SDx2 C(6); SDx3 C(6); SDx4 C(6); SDx5 C(6); SDx6 C(6); SDx7 C(6); SDx8 C(6); '
    'SDx9 C(6); SDx10 C(6); SDx11 C(6); SDx12 C(6); Proc1 C(7); Proc2 C(7); Proc3 C(7); Proc4 C(7); '
    'Proc5 C(7); Proc6 C(7); Proc7 C(7); Proc8 C(7); Proc9 C(7); Proc10 C(7); Proc11 C(7); Proc12 C(7); '
    'Proc13 C(7); Proc14 C(7); Proc15 C(7); Proc16 C(7); Proc17 C(7); Proc18 C(7); Proc19 C(7); Proc20 C(7); '
    'LeaveDay N(4,0); ActLOS N(3,0); Warn N(4,0); Err N(2,0); DRG C(5); MDC C(2); '
    'RW N(7,4); WTLOS N(6,2); OT N(4,0); ADJRW N(8,4)'
)

# Create and open the DBF file
table = dbf.Table('SQL_TO_DBF.dbf', table_structure)
table.open(mode=dbf.READ_WRITE)

def to_str(val, length=1, default=''):
    if pd.isna(val) or str(val).strip() == '':
        return default.ljust(length)[:length]
    
    val = str(val).strip()

    # Remove trailing ".0" if it's a float string like "2211.0"
    if val.endswith('.0'):
        val = val[:-2]

    return val.ljust(length)[:length]


def to_num(val, default=0, is_int=False):
    try:
        val = float(val)
        if pd.isna(val) or pd.isnull(val) or val != val:  # val != val catches float('nan')
            return default
        return int(val) if is_int else val
    except (ValueError, TypeError):
        return default


# Convert date fields
def to_date(val):
    if pd.isnull(val):
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    return None

def to_hhmm(val):
    if pd.isna(val) or str(val).strip() == '':
        return '0000'
    try:
        if isinstance(val, datetime.time):
            return val.strftime('%H%M')
        elif isinstance(val, datetime.datetime):
            return val.strftime('%H%M')
        elif isinstance(val, str):
            # Try parsing from string
            parsed = pd.to_datetime(val).time()
            return parsed.strftime('%H%M')
        else:
            return '0000'
    except:
        return '0000'


# Write data to DBF
for i, row in df.iterrows():
    try:
        table.append((
            to_date(row['DOB']),
            to_str(row['Sex'], 1, 'U'),
            to_date(row['DateAdm']),
            to_str(row['TimeAdm'], 4, '0000'),
            to_date(row['DateDsc']),
            to_str(row['TimeDsc'], 4, '0000'),
            to_str(row['Discht'], 1, 'U'),
            to_num(row['AdmWt'], 0.0),
            to_str(row['Age'], 3, '000'),
            to_str(row['AgeDay'], 3, '000'),
            to_str(row['PDx'], 6, 'UNK'),
            to_str(row['SDx1'], 6), to_str(row['SDx2'], 6), to_str(row['SDx3'], 6),
            to_str(row['SDx4'], 6), to_str(row['SDx5'], 6), to_str(row['SDx6'], 6),
            to_str(row['SDx7'], 6), to_str(row['SDx8'], 6), to_str(row['SDx9'], 6),
            to_str(row['SDx10'], 6), to_str(row['SDx11'], 6), to_str(row['SDx12'], 6),
            to_str(row['Proc1'], 7), to_str(row['Proc2'], 7), to_str(row['Proc3'], 7),
            to_str(row['Proc4'], 7), to_str(row['Proc5'], 7), to_str(row['Proc6'], 7),
            to_str(row['Proc7'], 7), to_str(row['Proc8'], 7), to_str(row['Proc9'], 7),
            to_str(row['Proc10'], 7), to_str(row['Proc11'], 7), to_str(row['Proc12'], 7),
            to_str(row['Proc13'], 7), to_str(row['Proc14'], 7), to_str(row['Proc15'], 7),
            to_str(row['Proc16'], 7), to_str(row['Proc17'], 7), to_str(row['Proc18'], 7),
            to_str(row['Proc19'], 7), to_str(row['Proc20'], 7),
            to_num(row['LeaveDay'], 0, is_int=True),
            to_num(row['ActLOS'], 0, is_int=True),
            to_num(row['Warn'], 0, is_int=True),
            to_num(row['Err'], 0, is_int=True),
            to_str(row['DRG'], 5, '000'),
            to_str(row['MDC'], 2, '00'),
            to_num(row['RW'], 0.0),
            to_num(row['WTLOS'], 0.0),
            to_num(row['OT'], 0, is_int=True),
            to_num(row['ADJRW'], 0.0)

        ))
    except Exception as e:
        print(f"❌ Error on row {i}: {e}")

# Save and close
table.close()
print("✅ DBF file saved successfully as 'SQL_TO_DBF.dbf'")
