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
# ### DATE AND TIME RELATED ###
# -------------------------------------
START_DT = datetime.now().strftime('%Y_%m_%d_%H_%M_%S')
# -------------------------------------

# -------------------------------------
# ### GLOBAL FILENAMES AND PATHS ###
# -------------------------------------
BASE_DIR = dirname(abspath(__file__))
OUT_DIR = join(BASE_DIR, 'Auswertung')
DEF_IN_FILE_PATH = join(BASE_DIR, 'performance 15mins - Münster.xlsx')
DEF_OUT_FILE_PATH = join(OUT_DIR, f'Münster_KPI_overview_{START_DT}.xlsx')
# -------------------------------------

# -------------------------------------
# ### STRINGS ###
# -------------------------------------
COUNTRY = 'Country'
FLOAT = 'float64'
KPI = 'KPI Type'
LARGE_COL = 'Idle Hrs'
MIN_OF_START = 'Minute of Startof15mininterval'
NL = '\n'
ORDERDAY = 'Day of Orderday'
REGION = 'Scooberjobregion'
VAL = 'Wert'
# -------------------------------------

# -------------------------------------
# ### LISTS ###
# -------------------------------------
IDX_COL = [COUNTRY, REGION, ORDERDAY, MIN_OF_START]
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
  df_sheet[ORDERDAY] = format_day(df_sheet)
  df_sheet[MIN_OF_START] = format_time(df_sheet)
  df_sheet, string_KPIs = ffill_string_KPIs(df_sheet)
  df_sheet[VAL] = df_sheet.sum(axis=1, numeric_only=True)
  df_pivot = df_sheet.pivot(index=IDX_COL, columns=KPI, values=VAL)
  if df_kpi is None:
    df_kpi = df_pivot
    extracted_KPIs = df_pivot.columns.tolist()
  else:
    df_sheet.set_index(IDX_COL, drop=False, inplace=True)
    extracted_KPIs = []
    for col in df_pivot.columns:
      if col not in df_kpi.columns:
        extracted_KPIs.append(col)
        df_kpi[col] = df_pivot[col]
    for string_KPI in string_KPIs:
      extracted_KPIs.append(string_KPI)
      df_kpi[string_KPI] = df_sheet[string_KPI][~df_sheet.index.duplicated()]
  return df_kpi, extracted_KPIs
# -------------------------------------

# -------------------------------------
def ffill_string_KPIs(df_sheet):
  cols = df_sheet.columns.to_list()
  time_data_idx = cols.index('0')
  *str_kpi_cols, unnamed_col = cols[:time_data_idx]
  df_sheet.rename(columns={unnamed_col: KPI}, inplace=True)
  string_KPIs = []
  for str_kpi_col in str_kpi_cols:
    if str_kpi_col in IDX_COL:
      df_sheet[str_kpi_col].fillna(method='ffill', inplace=True)
      continue
    if 'color' in str_kpi_col:
      df_sheet[str_kpi_col].fillna(method='ffill', inplace=True)
    string_KPIs.append(str_kpi_col)
  return df_sheet, string_KPIs
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
def kpi_data_to_formated_xlsx(df, out_path):
  writer = pd.ExcelWriter(out_path)
  df.to_excel(writer, 'overview', merge_cells=False)
  workbook = writer.book
  f_fmt = workbook.add_format({'num_format': '0.0000'})
  p_fmt = workbook.add_format({'num_format': '0.00%'})
  worksheet = writer.sheets['overview']
  worksheet.autofilter(0, 0, 0, len(df.columns) + 3)
  worksheet.set_zoom(65)
  worksheet.freeze_panes(1, 4)
  for n, index_col in enumerate(df.index.names):
    worksheet.set_column(n, n, int(len(index_col) * 1.25))
  for n, col in enumerate(df.columns, 4):
    width = int(len(col) * (1.5 if col == LARGE_COL else 1.25))
    fmt = p_fmt if col[-1] == '%' else f_fmt if df[col].dtype == FLOAT else {}
    worksheet.set_column(n, n, width, fmt)
  writer.save()
# -------------------------------------

# -------------------------------------
def print_extracted_kpis(sheet_name, extracted_cols):
  print(f'{sheet_name=}{NL}unique KPIs={extracted_cols or None}{NL}-----')
# -------------------------------------

# -------------------------------------
def print_header(text, fill_char='='):
  print(f' {text} '.center(80, fill_char))
# -------------------------------------

# -------------------------------------
def print_saved_xlsx(df_kpi, out_path):
  row_cnt, column_cnt = df_kpi.shape
  KPIs = df_kpi.columns.tolist()
  print(f'{row_cnt=}{NL}{column_cnt=}{NL}{KPIs=}{NL}{out_path=}{NL}-----')
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
  start = perf_counter()
  print_header('READ IN XLSX')
  if not exists(in_path):
    print(f'Couldn`t find {in_path}')
    return
  print(f'loading all sheets {in_path=} ...')
  dfs = pd.read_excel(in_path, sheet_name=None, header=1, parse_dates=True)
  print(f'loaded all sheets in {perf_counter() - start:.2f} seconds{NL}-----')
  print_header('EXTRACT SHEET DATA')
  df_kpi = None
  for sheet_name, df_sheet in dfs.items():
    df_kpi, extracted_KPIs = extract_sheet_kpi_data(df_kpi, df_sheet)
    print_extracted_kpis(sheet_name, extracted_KPIs)
  if not exists(dirname(out_path)):
    os.makedirs(dirname(out_path))
  print_header('SAVE FORMATED XLSX')
  kpi_data_to_formated_xlsx(df_kpi, out_path)
  print_saved_xlsx(df_kpi, out_path)
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
