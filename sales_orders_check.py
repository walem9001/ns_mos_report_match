from config import *
from queries import *
from utils import *
from openpyxl import Workbook
from dotenv import load_dotenv, find_dotenv
import pyodbc

if len(find_dotenv()) == 0:
    raise RuntimeError("Can't find your .env file.")
load_dotenv(find_dotenv())

USER = os.getenv("user")
PASSWORD = os.getenv("password")
pg_conn = pyodbc.connect(f"DSN=NETSUITE;uid={USER};pwd={PASSWORD}")
ns_df = pd.read_sql(ns_odbc_so, con=pg_conn)
pg_conn.close()

# ns_df['FECHA_SO'] = ns_df['FECHA_SO'].dt.tz_localize("UTC").dt.tz_convert("America/Mexico_City")
ns_df["FECHA_SO"] = ns_df["FECHA_SO"].dt.strftime("%Y-%m-%d")
ns_df = ns_df.loc[ns_df["FECHA_SO"] >= "2022-05-01"]
mos_df = pd.read_sql(mos_so, con=engine)
mos_df["created_at"] = (
    mos_df["created_at"].dt.tz_localize("UTC").dt.tz_convert("America/Mexico_City")
)
mos_df["created_at"] = mos_df["created_at"].dt.strftime("%Y-%m-%d")
mos_df = mos_df.loc[mos_df["created_at"] >= "2022-05-01"]

wb = Workbook()
ws = wb.active
ws.title = "Total"

merged_full = pd.merge(
    ns_df, mos_df, how="outer", left_on=["SO_ID"], right_on=["folio"]
)
merged_full["SYNC_STATUS"] = merged_full.apply(get_status, axis=1)
format_sheet(ws, merged_full)

today_day = datetime.today().strftime("%d")
this_month = datetime.today().strftime("%m")

sync_statuses = merged_full["SYNC_STATUS"].unique()

for i in sync_statuses:

    merged_df = merged_full.loc[merged_full["SYNC_STATUS"] == i]
    ws = wb.create_sheet(title=i)
    format_sheet(ws, merged_df)
wb.save("sale_order_sync_status.xlsx")
