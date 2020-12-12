# =================================================================
# ### IMPORTS ###
# =================================================================
# -------------------------------------
from collections import defaultdict
from copy import deepcopy
from datetime import date, datetime
import json
from operator import ne, eq
from os import get_terminal_size, listdir, makedirs
from os.path import abspath, basename, dirname, exists, join
from shutil import copyfileobj
import signal
import sys
from time import perf_counter
from zipfile import ZipFile
# -------------------------------------
import cv2 as cv
from fuzzywuzzy import fuzz
import numpy as np
from pandas import DataFrame, ExcelWriter, read_excel, to_datetime
from PIL import Image
from pytesseract import pytesseract
# -------------------------------------
# =================================================================

# =================================================================
# ### COMMENTS, CONSTANTS, FORMATS, GLOBAL PATHS, LOGS, PRINTS ###
# =================================================================
# -------------------------------------
# ### DATE AND TIME FORMATS ###
# -------------------------------------
DMY = '%d.%m.%Y'
HM = '%H:%M'
HMS = HM + ':%S'
MD = '%m-%d'
YMD = '%Y-%m-%d'
Y_S = '%Y_%m_%d_%H_%M_%S'
# -------------------------------------

# -------------------------------------
# ### NUMERIC ###
# -------------------------------------
MARGIN = 4
NAME_SHARE = .29
SHARP_KERNEL = np.array(([0, -1, 0], [-1, 5, -1], [0, -1, 0]), dtype="int")
START_DT = datetime.now().strftime(Y_S)
# -------------------------------------

# -------------------------------------
# ### STRINGS ###
# -------------------------------------
ALI = 'align'
AVA = 'avail'
AVAILS = 'availablities'
BG_COL = 'bg_color'
BOLD = 'bold'
BOR = 'border'
BR = '\n-----'
CAL = 'call'
CEN = 'center'
CHE = 'check'
CITY_LOG_PRE = ' LOG FOR CITY: '
CIT = 'city'
CMT = 'comment'
CNT = 'counter'
CONFIG_MISSING_MSG = '##### missing config file, using default values ...\n'
CON_TYP = 'contract type'
CO_SC = '3_color_scale'
CO_TY = 'Contract Type'
CREATE_XLSX_MSG = ' CREATE WEEKLY XLSX REPORT '
CRI = 'criteria'
DNA = 'det_name'
DON = 'done'
DRI = 'Driver'
DR_ID = 'Driver ID'
DUPL = 'duplicates'
EE = 'Rider_Ersterfassung'
FIR_ENT = 'first_entry'
FO_SI = 'font_size'
FR_HO = 'From Hour'
GIV = 'given'
GIV_AVA = 'given/avail'
GIV_MAX = 'given/max'
GIV_SHI = 'given shifts'
HRS = 'hours'
ID = 'ID'
INITIAL_MSG = f'SHIFTPLAN CHECK {START_DT}'
LAS_ENT = 'last_entry'
LEF = 'left'
LINK = 'linked'
MAX = 'max'
MAX_HOURS = 'mehr als max. Std.'
MAX_T = 'max_type'
MAX_V = 'max_value'
MERGE_FILES_MSG = ' MERGING DAILY PNGS TO DAILY FILE '
MH_AVAIL_MSG = ' "Monatsstunden" file available: '
MID_T = 'mid_type'
MID_V = 'mid_value'
MINI_LIMIT = 'Minijobber Monatsmax Std. prüfen'
MIN = 'min'
MIN_HOURS = 'weniger als Min.Std.'
MIN_T = 'min_type'
MIN_V = 'min_value'
MISSING_FILE_MSG = ' MISSING MEDATORY FILE '
MON = 'month'
MORE_HOURS = 'mehr Stunden'
MORE_THAN_AVAIL = 'mehr Std. als Verfügbarkeiten'
MUL_MAT = 'MULTIPLE MATCHES '
NF = 'NOT FOUND '
NL = '\n'
NOAV = 'no_avail'
NOCR = 'no_ocr'
NODA = 'no_data'
NOT_IN_MON = ' not in "Monatsstunden": '
NO_AVAILS = 'keine Verfügbarkeiten'
NU = 'num'
NU_FO = 'num_format'
PAI = 'paid'
PAI_MAX = 'paid/max'
PROCESS_PNG_MSG = ' SCAN AVAILABILITES FROM PNGS '
PROCESS_XLSX_MSG = ' PROCESS RAW XLSX DATA '
REDUCE_HOURS = ' -> auf Min.Std. reduzieren'
RID_NAM = 'rider name'
SHI = 'shift'
SCAN = 'scanned'
SIC = 'sick'
SH_DA = 'Shift Date'
STD_REP = 'Stundenreports'
SYNCH_MIN_H_MSG = ' SYNCHRONIZE NAMES IN MINDESTSTUNDEN LIST '
TAB = '\t'
TOP = 'top'
TO_HO = 'To Hour'
TYP = 'type'
UNK = 'unknown'
UNP = 'unpaid'
UNZIP_MSG = ' UNZIP CITY PNG FILES '
USE_NAM = 'User Name'
USE_TYP = 'User Type'
U_ID = 'User ID'
VAC = 'vacation'
VAL = 'value'
V_AL = 'valign'
V_CE = 'vcenter'
WOR = 'worked'
WO_HO = 'Worked hours'
WO_RA = 'Working Ratio'
W_AV_WO_SHIFT = '[000] availabilities, but no shifts:'
W_SHIFT_WO_AV = '-----\n[XXX] shifts without availabilities:'
XTR = 'extra'
ZIP_PNG_NAME_CHECK_MSG = 'Check file names in zip files: '
ZP = '0%'
# -------------------------------------

# -------------------------------------
# ### GLOBAL FILENAMES AND PATHS ###
# -------------------------------------
FILE_DIR = dirname(abspath(__file__))
CONFIG_FP = join(FILE_DIR, 'config_report.json')
LOG_FN = f'report_{START_DT}.log'
OUTPUT_DIR = join(FILE_DIR, 'Schichtplan_bearbeitet')
if not exists(OUTPUT_DIR):
  makedirs(OUTPUT_DIR)
OUT_FILE_PRE = ''
SPD_DIR = join(FILE_DIR, 'Schichtplan_Daten')
UNASSIGNED_AVAILS_FN = f'unassigned_avails_{START_DT}.json'
# -------------------------------------

# -------------------------------------
# ### TUPLES ###
# -------------------------------------
CONVERT_COLS_MONTH = (
  (PAI_MAX, None)
  , (WOR, 'Worked hours')
  , (VAC, 'Paid leaves (hours)')
  , (SIC, 'Sick leaves (hours)')
  , (PAI, 'Total paid hours')
  , (UNP, 'Unpaid leaves (hours)')
)
DF_DET_COLUMNS = ('kw', CIT, 'day', 'index', 'row', 'avail', 'name', 'ocr')
REPORT_HEADER = (
  ID, RID_NAM, CON_TYP, MAX, MIN, AVA, GIV, GIV_AVA, GIV_MAX, GIV_SHI, AVAILS
  , PAI_MAX, WOR, VAC, SIC, PAI, UNP, CMT, CHE, CAL, 'cmt scoober coordinator'
)
RIDER_MIN_HEADER = (RID_NAM, CON_TYP, MIN, CIT, FIR_ENT, LAS_ENT)
TIMEBLOCK_STRINGS = (
  '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30'
  , '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30'
  , '19:00', '19:30', '20:00', '20:30', '21:00', '21:30', '22:00', '22:30'
  , '23:00', '23:30'
)
WEEKDAYS = (
  'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag'
  , 'Samstag', 'Sonntag'
)
WEEKDAY_ABREVATIONS = ('Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So')
# -------------------------------------

# -------------------------------------
# ### DICTS ###
# -------------------------------------
COLOR = {
  'black': 0
  , 'filled': 126
  , 'name_box': 64
  , 'NP': {222, 228, 238}
  , 'NP 8d': 222
  , 'NP 50d': 238
  , 'scroll bar': 241
  , 'thin line': {215, 217, 221}
  , 'time_field_bg': 245
  , 'white': 255
}
CONTRACT_H = defaultdict(
  lambda: ('NO DATA', 'NO DATA'), {
    'Foodora_Minijob': (5, 11)
    , 'Foodora_Working Student': (12, 20)
    , 'Midijob': (12, 28)
    , 'Minijob': (5, 11)
    , 'Minijobber': (5, 11)
    , 'Mini-Jobber': (5, 11)
    , 'TE Midijob': (12, 28)
    , 'TE Minijob': (5, 11)
    , 'TE Teilzeit': (30, 48)
    , 'TE Werkstudent': (12, 20)
    , 'TE WS': (12, 20)
    , 'Teilzeit': (30, 48)
    , 'Vollzeit': (30, 48)
    , 'Werk Student': (12, 20)
  }
)
EXTRA_HOURS = {22: 1, 23: 1, 24: .5, 25: 0, 26: 0}
PNG_PROCESSING_DICT = {
  AVA: defaultdict(list)
  , DON: defaultdict(set)
  , HRS: defaultdict(int)
  , XTR: defaultdict(int)
  , CNT: defaultdict(int)
  , DNA: list()
}
TIMES = {
  22: TIMEBLOCK_STRINGS[1:-2]
  , 23: TIMEBLOCK_STRINGS[:-2]
  , 24: TIMEBLOCK_STRINGS[:-1]
  , 25: TIMEBLOCK_STRINGS
  , 26: TIMEBLOCK_STRINGS
}
# -------------------------------------

# -------------------------------------
# ### XLSX FORMATS ###
# -------------------------------------
COND_FMT = {
  AVA: {TYP: 'text', CRI: 'ends with', VAL: ' '}
  , 'ee only': {TYP: 'text', CRI: 'ends with', VAL: 'NOT IN AVAILS OR MONTH'}
  , 'not avail': {TYP: 'text', CRI: 'ends with', VAL: 'NOT IN AVAILS'}
  , 'date': {TYP: 'formula', CRI: '$F2 + 60 < today()'}
  , MIN: {TYP: 'cell', CRI: '==', VAL: '"NO DATA"'}
  , 'scale': {
    TYP: CO_SC, MIN_T: NU, MID_T: NU, MAX_T: NU, MIN_V: 0, MID_V: .5, MAX_V: 1
  }
}
FMT_DICT = {
  'border': {TOP: 1}
  , 'comment': {FO_SI: 10, V_AL: V_CE, BOLD: True}
  , 'error': {BOR: 1, BG_COL: '#ffff0f'}
  , 'ee only': {BOR: 1, BG_COL: '#00ffff'}
  , 'not avail': {BOR: 1, BG_COL: '#ff8f0f'}
  , 'int': {FO_SI: 10, V_AL: V_CE, ALI: CEN}
  , 'ratio': {FO_SI: 10, V_AL: V_CE, ALI: CEN, NU_FO: ZP}
  , 'old': {'italic': True, BG_COL: '#CE6161'}
  , 'red': {BOLD: True, BG_COL: '#CE2121'}
  , 'text': {FO_SI: 10, V_AL: V_CE, ALI: LEF}
}
RIDER_EE_COL_FMT = (
  ('A:A', 35, 'text')
  , ('B:B', 20, 'text')
  , ('C:C', 15, 'int')
  , ('D:D', 15, 'text')
  , ('E:E', 15, 'text')
  , ('F:F', 15, 'text')
)
RIDER_EE_CONDS = (('C2:C', MIN, 'red'), ('F2:F', 'date', 'old'))
RIDER_EE_FMTS = ('int', 'old', 'red', 'text')
XLS_REPORT_COL_FMT = (
  ('A:A', 5, 'int')
  , ('B:B', 28, 'text')
  , ('C:C', 11, 'text')
  , ('D:D', 4, 'int')
  , ('E:E', 9, 'int')
  , ('F:F', 4, 'int')
  , ('G:G', 5, 'int')
  , ('H:I', 9, 'ratio')
  , ('J:J', 31, 'text')
  , ('K:K', 23, 'text')
  , ('L:L', 8, 'ratio')
  , ('M:M', 6, 'int')
  , ('N:P', 7, 'int')
  , ('Q:Q', 6, 'int')
  , ('R:R', 27, 'comment')
  , ('S:S', 5, 'int')
  , ('T:T', 4, 'int')
  , ('U:U', 12, 'comment')
)
XLS_REPORT_COND_FMT = (
  ('H2:I', 'scale', None)
  , ('L2:L', 'scale', None)
  , ('E2:E', MIN, 'red')
  , ('K2:K', AVA, 'error')
  , ('K2:K', 'ee only', 'ee only')
  , ('K2:K', 'not avail', 'not avail')
)
# -------------------------------------
# =================================================================

# =================================================================
# ### INITIAL SETUP ###
# =================================================================
# -------------------------------------
# ### LOAD CONFIGURATION FILE ###
# -------------------------------------
if not exists(CONFIG_FP):
  print(CONFIG_MISSING_MSG)
  DEFAULT_CITIES = ['Frankfurt', 'Offenbach']
  ALIAS = {
    'Frankfurt': ('frankfurt', 'ffm', 'frankfurt am main')
    , 'Offenbach': ('offenbach', 'of', 'offenbach am main')
    , AVA: ('Verfügbarkeit', 'Verfügbarkeiten', 'Availabilities')
    , MON: ('Monatsstunden', 'Stunden', 'Working Hours')
    , SHI: ('Schichtplan', 'Schichtplanung', 'Working Shifts')
  }
else:
  with open(CONFIG_FP) as file_path:
    config = json.load(file_path)
  DEFAULT_CITIES = config['cities']
  ALIAS = {k: tuple(a for a in als) for k, als in config['aliases'].items()}
  if 'win' in sys.platform:
    pytesseract.tesseract_cmd = config['tesseract']['cmd_path']
try:
  pytesseract.get_tesseract_version()
  TESSERACT_AVAILABLE = True
except pytesseract.TesseractNotFoundError:
  TESSERACT_AVAILABLE = False
# -------------------------------------

# -------------------------------------
# ### HANDLER FOR KEYBOARD INTERRUPT ###
# -------------------------------------
def keyboard_interrupt_handler(signal, frame):
  print(f'{NL}KeyboardInterrupt (ID: {signal}) has been caught. Cleaning up..')
  sys.exit(0)
# -------------------------------------
signal.signal(signal.SIGINT, keyboard_interrupt_handler)
# -------------------------------------
# =================================================================

# =================================================================
# ### FUNCTIONS ###
# =================================================================
# -------------------------------------
def check_data_and_make_comment(rider_data):
  comment = []
  check = ''
  call = ''
  min_available = not isinstance(rider_data[MIN], str)
  if rider_data[GIV_MAX] > 1 and rider_data[MAX] != 40:
    comment.append(MAX_HOURS)
    call = 'X'
  if rider_data[GIV_AVA] == 10:
    comment.append(NO_AVAILS)
    if min_available and rider_data[GIV] > rider_data[MIN]:
      comment[-1] += REDUCE_HOURS
      call = 'X'
  else:
    if min_available and rider_data[GIV] < rider_data[MIN]:
      comment.append(MIN_HOURS)
      check = 'X'
      call = 'X'
    if rider_data[GIV_AVA] > 1:
      comment.append(MORE_THAN_AVAIL)
      check = 'X'
  if 'mini' not in rider_data[CON_TYP].casefold():
    threshold = .75
  else:
    if not isinstance(rider_data[PAI_MAX], str) and rider_data[PAI_MAX] > .9:
      threshold = 0
      comment.append(MINI_LIMIT)
      check = 'X'
      call = ''
    else:
      threshold = .55
  if rider_data[GIV_MAX] < threshold and rider_data[GIV_AVA] < threshold:
    comment.append(MORE_HOURS)
    check = 'X'
  return {CMT: NL.join(comment), CHE: check, CAL: call, AVAILS: ''}
# -------------------------------------

# -------------------------------------
def check_make_dir(*args):
  dir_path = join(*args)
  if not exists(dir_path):
    makedirs(dir_path)
  return dir_path
# -------------------------------------

# -------------------------------------
def city_in_xlsx_filename(filename, city):
  return (
    filename.endswith('xlsx')
    and filename[0].isalpha()
    and any(fuzz.partial_ratio(alias, filename) > 86 for alias in ALIAS[city])
  )
# -------------------------------------

# -------------------------------------
def get_availability_data(df_row, full=False):
  return {
  AVA: df_row['Hours Available']
  , MAX: df_row['Total Availability']
  , **(
    {ID: df_row[U_ID], RID_NAM: df_row[USE_NAM], CON_TYP: df_row[USE_TYP]}
    if full is True 
    else {}
  )
}
# -------------------------------------

# -------------------------------------
def get_avail_and_month_data(df_row, source, dfs, log):
  data = new_report_data_entry()
  if source == AVA:
    data.update(get_availability_data(df_row, full=True))
    m_data, log = get_month_data(dfs[MON], data[ID], data[RID_NAM], log)
    data.update(m_data)
  else:
    m_data, log = get_month_data(df_row, df_row[DR_ID], df_row[DRI], log, True)
    data.update(m_data)
    df_avail_row = dfs[AVA].loc[dfs[AVA][U_ID] == data[ID]]
    if df_avail_row.empty:
      data[AVA] = ''
    else:
      data.update(get_availability_data(df_avail_row.squeeze()))
  return data, log
# -------------------------------------

# -------------------------------------
def get_given_hour_ratios(avail, given, max_h):
  if isinstance(avail, str):
    return {GIV_MAX: 0, GIV_AVA: 0}
  else:
    return {
      GIV_MAX: round(given / max_h, 2)
      , GIV_AVA: round(given / avail, 2) if avail else 10
    }
# -------------------------------------

# -------------------------------------
def get_min_hours(df_ee, riders_ee, data):
  return {
    MIN: (
      CONTRACT_H[data[CON_TYP]][0]
      if 'TE' in data[CON_TYP] or data[RID_NAM] not in riders_ee
      else df_ee[df_ee[RID_NAM] == data[RID_NAM]][MIN].item()
    )
  }
# -------------------------------------

# -------------------------------------
def get_month_data(df_month, rider_id, rider_name, log, full=False):
  month_data = {}
  if df_month is None:
    for output_col, _ in CONVERT_COLS_MONTH:
      month_data[output_col] = 'N/A'
    return month_data, log
  if isinstance(df_month, DataFrame):
    df_month = df_month[df_month[DR_ID] == rider_id].copy()
  if df_month.empty:
    log += print_log(NOT_IN_MON + rider_name, '|MIS|', BR)
    return month_data, log
  for output_col, input_col in CONVERT_COLS_MONTH[1:]:
    try:
      month_data[output_col] = df_month[input_col]
    except ValueError:
      month_data[output_col] = 0
    except KeyError as ex:
      print(ex)
  try:
    work_ratio = float(str(df_month[WO_RA]).strip('%'))
    if work_ratio > 5: 
      work_ratio /= 100
  except ValueError:
    work_ratio = 0
  month_data[PAI_MAX] = round(work_ratio, 2)
  if full:
    month_data.update({
      ID: df_month[DR_ID]
      , RID_NAM: df_month[DRI]
      , CON_TYP: df_month[CO_TY]
      , MIN: CONTRACT_H[df_month[CO_TY]][0]
      , MAX: CONTRACT_H[df_month[CO_TY]][1]
    })
  return month_data, log
# -------------------------------------

# -------------------------------------
def get_shifts(df_shifts, dates, rider_id):
  given = 0
  shifts = ''
  for _, d in df_shifts[df_shifts[DR_ID] == rider_id].iterrows():
    if len(dates) < 7:
      dates.add(d[SH_DA].isoformat())
    given += d[WO_HO]
    shifts += f'{d[SH_DA]} | {d[FR_HO]} - {d[TO_HO]} | {d[WO_HO]}h{NL}'
  return {GIV: given, GIV_SHI: shifts}, dates
# -------------------------------------

# -------------------------------------
def load_avail_xlsx_into_df(df):
  return df.sort_values(USE_NAM, ignore_index=True)
# -------------------------------------

# -------------------------------------
def load_ersterkennung_xlsx_into_df(city, ree_dir):
  try:
    df_re = read_excel(join(ree_dir, f'{EE}_{city}.xlsx'))
  except FileNotFoundError:
    try:
      df_re = read_excel(f'{EE}.xlsx', city)
    except FileNotFoundError:
      df_re = DataFrame(columns=RIDER_MIN_HEADER)
  df_re[FIR_ENT] = to_datetime(df_re[FIR_ENT], format=YMD).dt.date
  df_re[LAS_ENT] = to_datetime(df_re[LAS_ENT], format=YMD).dt.date
  return df_re
# -------------------------------------

# -------------------------------------
def load_shift_xlsx_into_df(df):
  df[SH_DA] = to_datetime(df[SH_DA], format=DMY).dt.date
  try:
    df[FR_HO] = to_datetime(df[FR_HO], format=HM).dt.time
    df[TO_HO] = to_datetime(df[TO_HO], format=HM).dt.time
  except (TypeError, ValueError):
    df[FR_HO] = to_datetime(df[FR_HO], format=HMS).dt.time
    df[TO_HO] = to_datetime(df[TO_HO], format=HMS).dt.time
  return df.sort_values([DRI, SH_DA, FR_HO], ignore_index=True)
# -------------------------------------

# -------------------------------------
def load_xlsx_data_into_dfs(dirs, city, log):
  missing_files = [ALIAS[AVA][0], ALIAS[SHI][0]]
  dfs = {MON: None}
  for filename in listdir(dirs[0]):
    if not city_in_xlsx_filename(filename, city):
      continue
    if fuzz.WRatio(STD_REP, filename) > 86:
      log += print_log(f'|O.O| {STD_REP} file available, {filename = }')
      continue
    df = read_excel(join(dirs[0], filename))
    df.rename(columns=lambda x: str(x).strip(), inplace=True)
    if any(fuzz.partial_ratio(alias, filename) > 86 for alias in ALIAS[AVA]):
      dfs[AVA] = load_avail_xlsx_into_df(df)
      missing_files.remove(ALIAS[AVA][0])
    elif any(fuzz.partial_ratio(alias, filename) > 86 for alias in ALIAS[SHI]):
      dfs[SHI] = load_shift_xlsx_into_df(df)
      missing_files.remove(ALIAS[SHI][0])
    elif any(fuzz.partial_ratio(alias, filename) > 86 for alias in ALIAS[MON]):
      dfs[MON] = df
  if missing_files:
    dfs = None
    log += print_log_header(f'{MISSING_FILE_MSG}{missing_files}', '#', '', '')
  else:
    dfs[EE] = load_ersterkennung_xlsx_into_df(city, dirs[4])
  return dfs, log
# -------------------------------------

# -------------------------------------
def log_multi_match(name_data, ocr_read, row_n, png, log):
  log += f'[] [] {MUL_MAT}{png = }, {row_n = }, {ocr_read = }{NL}'
  for similarity, name, ocr, source in name_data:
    log += f'{TAB}{ocr = }, {name = }, {source = }, {similarity = }{NL}'
  det_name = max(name_data)[1]
  return det_name, log + f'{TAB} stored at: {det_name}{BR}'
# -------------------------------------

# -------------------------------------
def new_ee_only_entry(rider, daily_avail, cache, names):
  ee_only = new_report_data_entry()
  ee_only[RID_NAM] = rider
  ee_only[AVA] = cache[HRS][rider]
  ee_only[CON_TYP] = names[2][names[1].index(rider)]
  ee_only[MIN], ee_only[MAX] = CONTRACT_H[ee_only[CON_TYP]]
  ee_only[AVAILS] = png_parse_availabilities_string(
    daily_avail=daily_avail
    , week_h=ee_only[AVA]
    , max_h=ee_only[AVA] + cache[XTR][rider]
    , stored_avails=ee_only[AVA]
  ) + 'NOT IN AVAILS OR MONTH'
  for key in (GIV, GIV_AVA, GIV_MAX, PAI_MAX, WOR, VAC, SIC, PAI, UNP):
    ee_only[key] = 0
  return ee_only
# -------------------------------------

# -------------------------------------
def new_report_data_entry():
  return {report_column: '' for report_column in REPORT_HEADER}
# -------------------------------------

# -------------------------------------
def parse_break_line(fil='=', text=''):
  return text.center(80, fil)
# -------------------------------------

# -------------------------------------
def parse_progress_bar(bar_len, prog, pre, suf):
  done = int(bar_len * prog)
  return f'{pre} [{"#" * done + "-" * (bar_len - done)}] {prog:.2%} {suf}'
  # return f'{pre} [{"█" * done + "-" * (bar_len - done)}] {prog:.2%} {suf}'
# -------------------------------------

# -------------------------------------
def parse_stats_msg(counter:defaultdict):
  return (
    f'| S | total scanned        : {counter[SCAN]:3d}{NL}'
    f'| T | not readable         : {counter[NOCR]:3d}{NL}'
    f'| A | rows without data    : {counter[NODA]:3d}{NL}'
    f'| T | no availabilites     : {counter[NOAV]:3d}{NL}'
    f'| S | duplicate data       : {counter[DUPL]:3d}{NL}'
    f'<===> linked availabilites : {counter[LINK]:3d}'
  )
# -------------------------------------

# -------------------------------------
def png_avail_cell_is_filled(img, top, bot, x_test):
  return any(img[y_test, x_test] == COLOR['filled'] for y_test in (top, bot))
# -------------------------------------

# -------------------------------------
def png_capture_grid(image):
  img_height, img_width = image.shape
  rows, first_x = png_capture_grid_rows(image, img_height)
  rows, columns = png_capture_grid_cols(image, img_width, rows, first_x)
  return rows, len(rows) - 1, (columns, len(columns), first_x, columns[0])
# -------------------------------------

# -------------------------------------
def png_capture_grid_cols(image, img_width, rows, first_x):
  columns = None
  for row in rows[(1 if len(rows) == 2 else 2):]:
    columns = [first_x]
    for x_value in range(first_x, img_width - MARGIN):
      pixel_color = image[row - MARGIN, x_value]
      if pixel_color in COLOR['thin line']:
        if x_value <= columns[-1] + 5:
          columns.pop()
        columns.append(x_value)
        if len(columns) == 26:
          break
      elif pixel_color == COLOR['scroll bar']:
        if x_value < img_width // 2:
          rows.pop()
        elif x_value <= columns[-1] + 1:
          columns.pop()
        break
      elif pixel_color == COLOR['name_box']:
        break
    if len(columns) >= 22:
      break
  columns.pop(0)
  if image[rows[1] - 15, columns[0] + 1] == COLOR['time_field_bg']:
    rows.pop(0)
  return rows, columns
# -------------------------------------

# -------------------------------------
def png_capture_grid_rows(image, img_height, first_x=1, x_search_range=50):
  rows = None
  small_img = img_height < 185
  for x_value in range(first_x, x_search_range):
    line_cnt = 0
    rows = [0]
    y_value = 0
    for y_value in range(MARGIN, img_height - MARGIN):
      pixel_color = image[y_value, x_value]
      if pixel_color != COLOR['white']:
        if pixel_color in COLOR['NP']:
          continue
        if line_cnt == 10:
          break
        if y_value <= rows[-1] + 10:
          line_cnt += 1
          rows.pop()
        else:
          line_cnt = 0
        rows.append(y_value)
      if png_grid_invalid_row_values(small_img, rows, len(rows), y_value):
        break
      if y_value == 185 and len(rows) < 3:
        break
    if y_value == img_height - 5:
      first_x = x_value
      break
  rows.append(img_height)
  return rows, first_x
# -------------------------------------

# -------------------------------------
def png_grid_invalid_row_values(small_img, rows, row_cnt, y_value):
  return (
    (small_img and row_cnt == 1 and y_value > 70)
    or (row_cnt == 2 and y_value > rows[1] + 70)
    or (y_value == 185 and row_cnt < 3)
  )
# -------------------------------------

# -------------------------------------
def png_name_determination_algo(det_name, ocr_name, name_list):
  char_cnt = len(ocr_name)
  if char_cnt >= 5:
    for name in name_list:
      query = name[:char_cnt]
      slice_similarity = fuzz.WRatio(ocr_name, query)
      if slice_similarity >= 89:
        det_name = name
        break
      partial_similarity = fuzz.partial_ratio(ocr_name, query)
      if partial_similarity >= 89:
        det_name = name
        break
      if slice_similarity >= 75:
        det_name.append((slice_similarity, name, ocr_name, 'WRatio'))
      if partial_similarity >= 75:
        det_name.append((partial_similarity, name, ocr_name, 'partial'))
  return det_name
# -------------------------------------

# -------------------------------------
def png_ocr_yield_name_frames(frames, images):
  for image, size in images:
    top, bot, left, right = frames[size]
    yield (
      pytesseract.image_to_string(image[top:bot,left:right], config='--psm 7')
      .strip().split('..')[0].split('__')[0]
      .split('NP')[0].split('Np')[0].split('DA')[0]
    )
# -------------------------------------

# -------------------------------------
def png_one_row_determine_name(name_lists, row_n, frames, imgs, png, log):
  ocr_read = []
  det_name = []
  for ocr_name in png_ocr_yield_name_frames(frames, imgs):
    if not ocr_name:
      continue
    ocr_read.append(ocr_name)
    det_name = png_name_determination_algo(det_name, ocr_name, name_lists[0])
    if isinstance(det_name, str) or det_name:
      break
  if not det_name:
    for ocr_name in ocr_read:
      det_name = png_name_determination_algo(det_name, ocr_name, name_lists[1])
      if det_name:
        break
    if not det_name:
      log += print_no_name_determined(det_name, ocr_read, row_n, png)
    elif isinstance(det_name, list):
      det_name = max(det_name)[1]
  elif isinstance(det_name, list):
    if len(det_name) > len(ocr_read):
      det_name, log = log_multi_match(det_name, ocr_read, row_n, png, log)
    else:
      det_name = max(det_name)[1]
  return det_name, ocr_read, log
# -------------------------------------

# -------------------------------------
def png_one_row_get_availabities(date_str, cols, col_cnt, top, bot, img):
  daily_avail = ''
  daily_hours = 0
  extra_hours = 0
  hours_block = 0
  in_availablity_block = False
  for col_idx, column in enumerate(cols):
    if png_avail_cell_is_filled(img, top, bot, column + 1):
      hours_block += .5
      if col_idx == 0 and col_cnt == 22:
        extra_hours += .5
      if not in_availablity_block:
        daily_avail += f'{date_str} | {TIMES[col_cnt][col_idx]}'
        in_availablity_block = True
      elif col_idx == col_cnt - 1:
        extra_hours += EXTRA_HOURS[col_cnt]
        daily_avail += f' - {TIMES[col_cnt][-1]} | {hours_block:.1f}h{NL}'
        daily_hours += hours_block
        break
    elif in_availablity_block:
      daily_avail += f' - {TIMES[col_cnt][col_idx]} | {hours_block:.1f}h{NL}'
      daily_hours += hours_block
      hours_block = 0
      in_availablity_block = False
  return daily_avail, daily_hours, extra_hours
# -------------------------------------

# -------------------------------------
def png_one_row_no_data(x_name, x_avail, top, bot, img):
  return (
    all(
      img[y_val, x_name] == img[y_val, x_name + 4] == img[y_val, x_name + 8]
      for y_val in range(top, bot)
    )
    or all(
      img[y_val, x_avail] == COLOR['white'] for y_val in range(top, bot)
    )
  )
# -------------------------------------

# -------------------------------------
def png_parse_availabilities_string(daily_avail, week_h, max_h, stored_avails):
  return (
    ''.join(sorted(daily_avail))
    + f'total: {week_h}h | avail <= {max_h}{NL}'
    + (
      'NOT IN AVAILS'
      if isinstance(stored_avails, str)
      else '' if week_h <= stored_avails <= max_h else ' '
    )
  )
# -------------------------------------

# -------------------------------------
def png_parse_date(day_str, kw_dates:list):
  return date.fromisoformat(kw_dates[WEEKDAYS.index(day_str)]).strftime(MD)
# -------------------------------------

# -------------------------------------
def png_read_row(top, bot, row_n, frames, imgs, x_vals, png, cache, log, args):
  cols, col_cnt, first_x, first_col = x_vals
  x_name = (39 * first_x + first_col) // 40
  x_avail = first_col * 24 // 25
  if png_one_row_no_data(x_name, x_avail, top, bot, imgs[2][0]):
    # print(f'{top = }, {bot = }, {x_name = }, {x_avail = }')
    cache[CNT][NODA] += 1
    return cache, log
  name_lists, date_str, day, png_idx, city, kw = args
  avail_str, daily_h, extra_h = png_one_row_get_availabities(
    date_str, cols, col_cnt, top, bot, imgs[2][0]
  )
  rider, ocr_read, log = png_one_row_determine_name(
    name_lists, row_n, frames, imgs, png, log
  )
  cache[DNA].append(
    (kw, city, day, png_idx, row_n, avail_str[:-1], rider, ocr_read)
  )
  if not avail_str:
    cache[CNT][NOAV] += 1
  elif not rider:
    cache[CNT][NOCR] += 1
  elif date_str in cache[DON][rider]:
    cache[CNT][DUPL] += 1
  else:
    cache[CNT][LINK] += 1
    cache[AVA][rider].append(avail_str)
    cache[HRS][rider] += daily_h
    cache[XTR][rider] += extra_h
    cache[DON][rider].add(date_str)
  # cache = cache, avail_str, ocr_read
  return cache, log
# -------------------------------------

# -------------------------------------
def png_read_screenshot(png_n, png, png_cnt, cache, png_dir, log, *args):
  img = cv.imread(join(png_dir, png), cv.IMREAD_GRAYSCALE)
  rows, row_cnt, x_vals = png_capture_grid(img)
  cache[CNT][SCAN] += row_cnt
  _, _, left, first_col = x_vals
  # print(
  #   f'{png = }, {left = }, {first_col = }, {row_cnt = }, {rows = }, cols = '
  #   + grid_values[0]
  # )
  right = int(first_col * NAME_SHARE)
  res_f = 109 / (rows[1] - rows[0])
  res_rows = [int(res_f * row) for row in rows]
  res_l = int(res_f * left)
  res_r = int(res_f * right)
  res_width = res_r - res_l
  res_height = int(res_f * img.shape[0])
  resize_img = cv.resize(img[:, left:right].copy(), (res_width, res_height))
  imgs = (
    (resize_img, 'resize')
    , (cv.filter2D(resize_img, -1, SHARP_KERNEL), 'resize')
    , (img, 'orig')
    , (cv.threshold(img, 220, 255, cv.THRESH_BINARY)[1], 'orig')
  )
  frames = {'orig': [0, 0, left, right], 'resize': [0, 0, res_l, res_r]}
  for row_n, row in enumerate(rows[:-1], 1):
    top = frames['orig'][0] = row + 4
    bot = frames['orig'][1] = rows[row_n] - 4
    frames['resize'][0] = res_rows[row_n - 1]
    frames['resize'][1] = res_rows[row_n]
    cache, log = png_read_row(
      top, bot, row_n, frames, imgs, x_vals, png, cache, log, args
    )
    print_progress_bar(png_cnt, png_n, row_cnt, row_n, png)
    # cache, avails, ocr_read = cache
    # print(f'{row_n = }, {avails = }, {ocr_read = }')
  return cache, log
# -------------------------------------

# -------------------------------------
def png_update_report_dataframe(cache, df, names):
  only_in_ee = []
  for rider, daily_avail in cache[AVA].items():
    try:
      rider_df_idx = df[df[RID_NAM] == rider].index[0]
    except IndexError:
      only_in_ee.append(new_ee_only_entry(rider, daily_avail, cache, names))
    else:
      df.at[rider_df_idx, AVAILS] = png_parse_availabilities_string(
        daily_avail=daily_avail
        , week_h=cache[HRS][rider]
        , max_h=cache[HRS][rider] + cache[XTR][rider]
        , stored_avails=df.at[rider_df_idx, AVA]
      )
      if df.at[rider_df_idx, AVA] == '':
        df.at[rider_df_idx, AVA] = cache[HRS][rider]
  if only_in_ee:
    df = df.append(only_in_ee, ignore_index=True)
    df.sort_values(RID_NAM, inplace=True, ignore_index=True)
  df.loc[df[AVA] == df[AVAILS], AVA] = 0
  df.loc[(df[AVA] != 0) & (df[AVAILS] == ''), AVAILS] = ' '
  return df[(df[AVA] != 0) | (df[AVAILS] != '') | (df[GIV_SHI] != '')]
# -------------------------------------

# -------------------------------------
def print_log(text='', pre='', end=''):
  print(pre + text + end)
  return pre + text + end + NL
# -------------------------------------

# -------------------------------------
def print_log_header(text='', fil='=', pre='-', suf='-', brk=BR):
  if len(pre) == 1:
    pre = parse_break_line(pre) + NL
  if len(suf) == 1:
    suf = NL + parse_break_line(suf)
  msg = pre + (f' {text} ' if text else '').center(80, fil) + suf + brk
  print(msg)
  return msg + NL
# -------------------------------------

# -------------------------------------
def print_log_w_avails_wo_shift_and_vice_versa(df):
  log = ''
  for msg, op_1, op_2 in ((W_AV_WO_SHIFT, ne, eq), (W_SHIFT_WO_AV, eq, ne)):
    log += print_log(msg)
    for rider in df[op_1(df[AVA], 0) & op_2(df[GIV_SHI], '')][RID_NAM]:
      log += print_log(rider, '\t- ')
  return log + print_log('-----')
# -------------------------------------

# -------------------------------------
def print_no_name_determined(ee_name, ocr_read, row_n, png):
  print('\r', end='')
  return print_log(
    f'##### {NF}{png = }, {row_n = }{" " * 30}{NL}'
    + (f'|OCR| {ocr_read = }' if ocr_read else '')
    + (f', {ee_name = }' if ee_name else '')
    + BR
  )
# -------------------------------------

# -------------------------------------
def print_progress_bar(
  main_total
  , main_progress
  , sub_total=None
  , sub_progress=None
  , suf=''
  , bar_len=30
  , print_end='\r'
  , pre=' ==> '
  , min_pad=12
):
  if sub_total is None:
    prog = main_progress / main_total
  else:
    prog = (main_progress + sub_progress / sub_total) / main_total
    suf += f' ({sub_progress}/{sub_total})'
  if prog >= 1.:
    prog = 1
    pre = '|FIN|'
    suf = f'... DONE'
    print_end = '\r\n-----\n'
  bar_str = parse_progress_bar(bar_len, prog, pre, suf)
  pad = get_terminal_size().columns - len(bar_str) - min_pad -1
  if pad < 0:
    bar_str = '\r' + parse_progress_bar(bar_len + pad, prog, pre, suf)
  print(bar_str, ' ' * min_pad, sep='', end=print_end, flush=True)
# -------------------------------------

# -------------------------------------
def process_raw_xlsx_data_store_in_df(dirs, city, kw, log):
  log += print_log_header(PROCESS_XLSX_MSG)
  # ----- read weekly xlsx data, check availability of mendatory raw files ----
  dfs, log = load_xlsx_data_into_dfs(dirs, city, log)
  if dfs is None:
    return None, None, None, log
  dfs, log = rider_ersterfassung_update_names(kw, city, dirs[4], dfs, log)
  # ----- store data from mendatory xlsx files in dataframes -----
  dates = set()
  data_list = []
  source, col = (AVA, USE_NAM) if dfs[MON] is None else (MON, DRI)
  name_lists = (
    dfs[source][col].to_numpy()
    , dfs[EE][RID_NAM].to_list()
    , dfs[EE][CON_TYP].to_list()
  )
  for _, df_row in dfs[source].iterrows():
    data, log = get_avail_and_month_data(df_row, source, dfs, log)
    rider_shifts, dates = get_shifts(dfs[SHI], dates, data[ID])
    data.update(rider_shifts)
    data.update(get_given_hour_ratios(data[AVA], data[GIV], data[MAX]))
    data.update(get_min_hours(dfs[EE], name_lists[1], data))
    data.update(check_data_and_make_comment(data))
    data_list.append(data)
  # ---- create dataframe, extract additional information, filter unneeded rows
  df = DataFrame(data_list)
  log += print_log_w_avails_wo_shift_and_vice_versa(df)
  # ----------
  return df, name_lists, sorted(dates), log
# -------------------------------------

# -------------------------------------
def process_screenshots_store_avails(df, name_lists, dates, dirs, log, *args):
  log += print_log_header(PROCESS_PNG_MSG)
  cache = deepcopy(PNG_PROCESSING_DICT)
  pngs = sorted(listdir(dirs[3]))
  png_cnt = len(pngs)
  for png_n, png in enumerate(pngs):
    # if png != 'Samstag_4.png':
    #   continue
    day, file_suf = png.split('_')
    file_idx = file_suf.split('.')[0]
    date_str = png_parse_date(day, dates)
    cache, log = png_read_screenshot(
      png_n, png, png_cnt, cache, dirs[3], log
      , name_lists, date_str, day, file_idx, *args
    )
  log += print_log(parse_stats_msg(cache[CNT]), end=BR)
  df_determined = DataFrame(cache[DNA], columns=DF_DET_COLUMNS)
  df_determined.to_excel(
    join(dirs[1], f'det_names_{args[0]}_{START_DT}.xlsx')
    , args[0]
    , columns=DF_DET_COLUMNS
    , index=False
  )
  return png_update_report_dataframe(cache, df, name_lists), log
# -------------------------------------

# -------------------------------------
def rider_ee_new_entry(city, new_name, contract, kw_monday_date):
  return {
    RID_NAM: new_name
    , CON_TYP: contract
    , MIN: CONTRACT_H[contract][0]
    , CIT: city
    , FIR_ENT: kw_monday_date
    , LAS_ENT: kw_monday_date
  }
# -------------------------------------

# -------------------------------------
def rider_ersterfassung_format_and_save_xlsx(city, ree_dir, df_min):
  row_cnt = df_min.shape[0] + 1
  writer = ExcelWriter(join(ree_dir, f'{EE}_{city}.xlsx'), engine='xlsxwriter')
  df_min.to_excel(writer, city, index=False, freeze_panes=(1, 0))
  workbook = writer.book
  worksheet = writer.sheets[city]
  worksheet.autofilter('A1:F1')
  worksheet.freeze_panes(1, 0)
  fmt = {k: workbook.add_format(FMT_DICT[k]) for k in RIDER_EE_FMTS}
  for column, width, fmt_key in RIDER_EE_COL_FMT:
    worksheet.set_column(column, width, fmt[fmt_key])
  for columns, cond_key, fmt_key in RIDER_EE_CONDS:
    worksheet.conditional_format(
      f'{columns}{row_cnt}', {**COND_FMT[cond_key], 'format': fmt[fmt_key]}
    )
  writer.save()
# -------------------------------------

# -------------------------------------
def rider_ersterfassung_update_names(kw, city, ree_dir, dfs, log):
  log += print_log_header(SYNCH_MIN_H_MSG)
  data_list = []
  kw_date = date.fromisocalendar(date.today().year, kw, 1)
  names_mon = set() if dfs[MON] is None else {*dfs[MON][DRI]}
  log += print_log(f'CHECK {MH_AVAIL_MSG}{bool(names_mon)} {BR}')
  ree_names = dfs[EE][RID_NAM]
  known = {*ree_names}
  names_av = {*dfs[AVA][USE_NAM]}
  new_in_mon = names_mon - known
  params = [(dfs[MON], CO_TY, DRI, new_in_mon)] if new_in_mon else []
  params.append((dfs[AVA], USE_TYP, USE_NAM, names_av - (known | new_in_mon)))
  for df_q, contract_key, name_key, new_names in params:
    for _, d in df_q.iterrows():
      name = d[name_key]
      if name in new_names:
        contract = df_q[df_q[name_key] == name][contract_key].item()
        data_list.append(rider_ee_new_entry(city, name, contract, kw_date))
        log += print_log(f'{TAB}- {name = }, {contract = }')
      elif name in known:
        rider_ee_idx = dfs[EE][ree_names == name].index[0]
        if kw_date > dfs[EE].at[rider_ee_idx, LAS_ENT]:
          dfs[EE].at[rider_ee_idx, LAS_ENT] = kw_date
    if new_names:
      log += print_log('-----')
  dfs[EE] = dfs[EE].append(data_list, ignore_index=True)
  dfs[EE].sort_values([LAS_ENT, FIR_ENT, RID_NAM], inplace=True)
  rider_ersterfassung_format_and_save_xlsx(city, ree_dir, dfs[EE])
  return dfs, log
# -------------------------------------

# -------------------------------------
def save_df_in_formated_xlsx(kw, city, df):
  log = print_log_header(CREATE_XLSX_MSG)
  row_cnt = len(df) + 1
  # ----- open instance of xlsx-file -----
  filename = f'{OUT_FILE_PRE}KW{kw}_{city}-{START_DT}.xlsx'
  writer = ExcelWriter(join(OUTPUT_DIR, filename), engine='xlsxwriter')
  df.to_excel(writer, 'Sheet1', index=False)
  # ----- format xlsx file -----
  workbook = writer.book
  fmt_dict = {f_key: workbook.add_format(f) for f_key, f in FMT_DICT.items()}
  worksheet = writer.sheets['Sheet1']
  worksheet.set_zoom(85)
  worksheet.autofilter('A1:U1')
  worksheet.freeze_panes(1, 2)
  worksheet.set_row(row_cnt, None, fmt_dict['border'])
  for column, width, fmt in XLS_REPORT_COL_FMT:
    worksheet.set_column(column, width, fmt_dict[fmt])
  # ----- add conditional formats -----
  for cols, cond, fmt in XLS_REPORT_COND_FMT:
    worksheet.conditional_format(
      f'{cols}{row_cnt}'
      , {**COND_FMT[cond], 'format': fmt_dict[fmt]} if fmt else COND_FMT[cond]
    )
  # ----------
  writer.save()
  return log + print_log(f'+++++ saved {filename}{BR}')
# -------------------------------------

# -------------------------------------
def shiftplan_check(city, kw, dirs, get_avails, merge_pngs, unzip_only):
  start = perf_counter()
  log = print_log_header(CITY_LOG_PRE + city, pre='=')
  log += zip_extract_screenshots(city, dirs, merge_pngs)
  if unzip_only:
    return log
  df, name_lists, dates, log = process_raw_xlsx_data_store_in_df(
    dirs, city, kw, log
  )
  if df is None:
    return log
  if get_avails and TESSERACT_AVAILABLE:
    df, log = process_screenshots_store_avails(
      df, name_lists, dates, dirs, log, city, kw
    )
  log += save_df_in_formated_xlsx(kw, city, df)
  return log + print_log_header(
    f'runtime {city = }: {perf_counter() - start:.2f} s', suf='='
  )
# -------------------------------------

# -------------------------------------
def zip_extract_screenshots(city, dirs, merge_pngs=False):
  log = print_log_header(UNZIP_MSG)
  log += print_log(ZIP_PNG_NAME_CHECK_MSG, '[X|O] ')
  idx_dict = defaultdict(int)
  for zip_file in zip_iter_city_files(city, dirs[0]):
    log += print_log(zip_file, TAB)
    with ZipFile(join(dirs[0], zip_file)) as zfile:
      for member in sorted(zfile.namelist()):
        f_name = basename(member)
        if f_name:
          f_name, idx_dict, log = zip_parse_png_filename(f_name, idx_dict, log)
          with open(join(dirs[3], f_name), "wb") as target:
            copyfileobj(zfile.open(member), target)
    log += print_log('-----')
  log += print_log(f'+++++ saved PNGs in: {dirs[3]}{BR}')
  if merge_pngs:
    log += zip_merge_png_files_per_day(city, dirs)
  return log
# -------------------------------------

# -------------------------------------
def zip_iter_city_files(city, kw_dir):
  for filename in listdir(kw_dir):
    if filename.endswith('.zip'):
      fn_cf = filename.casefold()
      if any(fuzz.partial_ratio(alias, fn_cf) > 86 for alias in ALIAS[city]):
        yield filename
# -------------------------------------

# -------------------------------------
def zip_merge_get_daily_files(day, png_dir):
  return [
      Image.open(join(png_dir, day_fn))
      for day_fn in sorted(fn for fn in listdir(png_dir) if day in fn)
    ]
# -------------------------------------

# -------------------------------------
def zip_merge_png_files_per_day(city, dirs):
  log = print_log_header(MERGE_FILES_MSG)
  for day in WEEKDAYS:
    images = zip_merge_get_daily_files(day, dirs[3])
    widths, heights = zip(*(img.size for img in images))
    new_image = Image.new('RGB', (max(widths), sum(heights)))
    y_offset = 0
    for img in images:
      new_image.paste(img, (0, y_offset))
      y_offset += img.size[1]
    daily_img_fn = f'{city}_{day}.png'
    new_image.save(join(dirs[2], daily_img_fn))
    log += print_log(f'+++++ saved {daily_img_fn}')
  return log + print_log('-----')
# -------------------------------------

# -------------------------------------
def zip_parse_png_filename(original, idx_dict, log):
  similarity = 0
  current_day = ''
  for weekday in WEEKDAYS:
    weekday_similarity = fuzz.partial_ratio(original, weekday)
    if weekday_similarity > similarity:
      similarity = weekday_similarity
      current_day = weekday
      if similarity > 90:
        break
  if similarity <= 90:
    for n, abrev in enumerate(WEEKDAY_ABREVATIONS):
      abrevation_similarity = fuzz.partial_ratio(original, abrev)
      if abrevation_similarity > similarity:
        similarity = abrevation_similarity
        current_day = WEEKDAYS[n]
        if similarity == 100:
          break
  idx_dict[current_day] += 1
  saved_as = f'{current_day}_{idx_dict[current_day]}.png'
  if similarity != 100:
    log += print_log(f'{TAB}- {original = }, {saved_as = }')
  return saved_as, idx_dict, log
# -------------------------------------
# =================================================================

# =================================================================
# ### MAIN FUNCTION ###
# =================================================================
# -------------------------------------
def main(start_kw, last_kw, cities, get_avails, merge_pngs, unzip_only):
  start = perf_counter()
  log = print_log_header(INITIAL_MSG, pre='=', suf='=')
  if last_kw < start_kw:
    last_kw = start_kw 
  for kw in range(start_kw, last_kw + 1):
    kw_dir = join(SPD_DIR, f'KW{kw}')
    if not exists(kw_dir):
      log += print_log(f'##### Couldn`t find "{kw_dir}"{BR}')
      continue
    log_dir = check_make_dir(kw_dir, 'logs')
    screen_dir = join(kw_dir, 'Screenshots')
    ree_dir = check_make_dir(FILE_DIR, EE)
    for city in cities:
      png_dir = check_make_dir(screen_dir, city)
      dirs = (kw_dir, log_dir, screen_dir, png_dir, ree_dir)
      log += shiftplan_check(
        city, kw, dirs, get_avails, merge_pngs, unzip_only)
    with open(join(log_dir, LOG_FN), 'w', encoding='utf-8') as logfile:
      logfile.write(log)
  log += print_log_header(
    f'TOTAL RUNTIME: {perf_counter() - start:.2f} s', pre='=', suf='=', brk=NL
  )
# -------------------------------------
# =================================================================

# =================================================================
# ### START SCRIPT ###
# =================================================================
# -------------------------------------
if __name__ == '__main__':
  from argparse import ArgumentParser
  parser = ArgumentParser()
  parser.add_argument('--kalenderwoche', '-kw', required=True, type=int)
  parser.add_argument('--last_kw', '-l', type=int, default=0)
  parser.add_argument('--cities', '-c', nargs='*', default=DEFAULT_CITIES)
  parser.add_argument('--getavails', '-a', action='store_true')
  parser.add_argument('--mergeperday', '-m', action='store_true')
  parser.add_argument('--unzip_only', '-z', action='store_true')
  main(*parser.parse_args().__dict__.values())
# -------------------------------------
# =================================================================
