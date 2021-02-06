# =================================================================
# ### IMPORTS ###
# =================================================================
# -------------------------------------
from datetime import datetime
import os
from os.path import abspath, dirname, exists, join
import signal
import sys
from time import perf_counter
# -------------------------------------
import pandas as pd
# -------------------------------------
# =================================================================


# =================================================================
# ### CONSTANTS ###
# =================================================================
# -------------------------------------
START_DT = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
BASE_DIR = dirname(abspath(__file__))
OUT_DIR = join(BASE_DIR, 'Auswertung')
DEF_IN_FILE_PATH = join(BASE_DIR, 'performance 15mins - Münster.xlsx')
DEF_OUT_FILE_PATH = join(OUT_DIR, f'Münster_KPI_overview_{START_DT}.xlsx')
IDX_COL = [
  'Country', 'Scooberjobregion', 'Day of Orderday'
  , 'Minute of Startof15mininterval'
]
KPI = 'KPI Type'
LARGE_COL = 'Idle Hrs'
NL = '\n'
VAL = 'Wert'
# -------------------------------------
# =================================================================


# =================================================================
# ### HANDLER FOR KEYBOARD INTERRUPT ###
# =================================================================
# -------------------------------------
def keyboard_interrupt_handler(signal, frame):
  print(f'{NL}KeyboardInterrupt has been caught. Close and clean up ...')
  sys.exit(signal)
# -------------------------------------
signal.signal(signal.SIGINT, keyboard_interrupt_handler)
# -------------------------------------
# =================================================================


# =================================================================
# ### FUNCTIONS ###
# =================================================================
# -------------------------------------
def extract_sheet_kpi_data(df_kpi, df_sheet):
  string_KPIs = []
  df_sheet[IDX_COL[2]] = format_day(df_sheet)
  df_sheet[IDX_COL[3]] = format_time(df_sheet)
  cols = df_sheet.columns.to_list()
  time_data_idx = cols.index('0')
  *str_kpi_cols, unnamed_col = cols[:time_data_idx]
  df_sheet.rename(columns={unnamed_col: KPI}, inplace=True)
  kpi_cnt = df_sheet[KPI].nunique()
  for str_kpi_col in str_kpi_cols:
    if str_kpi_col in IDX_COL:
      df_sheet[str_kpi_col].fillna(method='ffill', inplace=True)
    else:
      df_sheet[str_kpi_col].fillna(method='ffill', inplace=True, limit=kpi_cnt)
      df_sheet[str_kpi_col].fillna('', inplace=True)
      string_KPIs.append(str_kpi_col)
  df_sheet[VAL] = df_sheet.sum(axis=1, numeric_only=True)
  df_pivot = df_sheet.pivot(index=IDX_COL, columns=KPI, values=VAL)
  if df_kpi is None:
    df_kpi = df_pivot
  else:
    for col in df_pivot.columns:
      if col not in df_kpi.columns:
        df_kpi[col] = df_pivot[col]
    df_sheet.set_index(IDX_COL, drop=False, inplace=True)
    for string_KPI in string_KPIs:
      df_kpi[string_KPI] = df_sheet[string_KPI][~df_sheet.index.duplicated()]
  return df_kpi
# -------------------------------------

# -------------------------------------
def format_day(df):
  return pd.to_datetime(df[IDX_COL[2]], infer_datetime_format=True).dt.date
# -------------------------------------

# -------------------------------------
def format_time(df):
  return pd.to_datetime(df[IDX_COL[3]], infer_datetime_format=True).dt.time
# -------------------------------------

# -------------------------------------
def kpi_data_to_formated_xlsx(df_kpi, out_path):
  cols = df_kpi.columns.tolist()
  writer = pd.ExcelWriter(out_path)
  df_kpi.to_excel(writer, 'overview', merge_cells=False)
  workbook = writer.book
  perc_fmt = workbook.add_format()
  perc_fmt.set_num_format(10)
  float_fmt = workbook.add_format()
  float_fmt.set_num_format('0.0000')
  worksheet = writer.sheets['overview']
  worksheet.autofilter(0, 0, 0, len(cols) + 3)
  worksheet.set_zoom(65)
  worksheet.freeze_panes(1, 4)
  index_widths = ((x, int(len(x) * 1.25)) for x in df_kpi.index.names)
  for n, (col, width) in enumerate(index_widths):
    worksheet.set_column(n, n, width)
  widths = ((x, int(len(x) * (1.5 if x == LARGE_COL else 1.25))) for x in cols)
  for n, (col, width) in enumerate(widths, 4):
    if col.endswith('%'):
      col_fmt = perc_fmt
    elif df_kpi[col].dtype == 'float64':
      col_fmt = float_fmt
    else:
      col_fmt = {}
    worksheet.set_column(n, n, width, col_fmt)
  writer.save()
# -------------------------------------

# -------------------------------------
def print_loaded_sheets_msg(start):
  print(f'loaded all sheets in {perf_counter() - start:.2f} seconds')
# -------------------------------------

# -------------------------------------
def print_runtime_msg(start):
  print(f' RUNTIME: {perf_counter() - start:.2f} '.center(80, '='))
# -------------------------------------
# =================================================================


# =================================================================
# ### MAIN FUNCTION ###
# =================================================================
# -------------------------------------
def kpi_deserializer(in_path, out_path):
  if not exists(in_path):
    print(f'Couldn`t find {in_path}')
    return
  start = perf_counter()
  df_kpi = None
  dfs = pd.read_excel(in_path, sheet_name=None, header=1, parse_dates=True)
  print_loaded_sheets_msg(start)
  for df_sheet in dfs.values():
    df_kpi = extract_sheet_kpi_data(df_kpi, df_sheet)
  if not exists(dirname(out_path)):
    os.makedirs(dirname(out_path))
  kpi_data_to_formated_xlsx(df_kpi, out_path)
  print_runtime_msg(start)
# -------------------------------------

# -------------------------------------
def main():
  from argparse import ArgumentParser
  parser = ArgumentParser()
  parser.add_argument('-i', '--input_filepath', default=DEF_IN_FILE_PATH)
  parser.add_argument('-o', '--output_filepath', default=DEF_OUT_FILE_PATH)
  kpi_deserializer(*parser.parse_args().__dict__.values())
# -------------------------------------
# =================================================================


# =================================================================
# ### START SCRIPT ###
# =================================================================
# -------------------------------------
if __name__ == '__main__':
  sys.exit(main())
# -------------------------------------
# =================================================================
