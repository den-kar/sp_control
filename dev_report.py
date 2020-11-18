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
_factor = 20
SHARP_KERNEL = np.array(
  (
    [-_factor + 3, 0, -_factor + 1]
    , [-_factor + 2, 6 * _factor - 8, -_factor]
    , [-_factor + 3, 0, -_factor + 1]
  )
  , dtype='int'
)
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
CONVERT_COLS_AVAIL = (
  (ID, 'User ID')
  , (RID_NAM, USE_NAM)
  , (CON_TYP, USE_TYP)
  , (MAX, 'Total Availability')
  , (AVA, 'Hours Available')
)
CONVERT_COLS_MONTH = (
  (WOR, 'Worked hours')
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
# -------------------------------------

# -------------------------------------
# ### DICTS ###
# -------------------------------------
ALIAS = {
  'Frankfurt': {'frankfurt', 'ffm', 'frankfurt am main'}
  , 'Offenbach': {'offenbach', 'of', 'offenbach am main'}
  , AVA: ('Verfügbarkeit', 'Verfügbarkeiten')
  , MON: ('Monatsstunden', 'Stunden')
  , SHI: ('Schichtplan', 'Schichtplanung')
}
COLOR = {
  'black': 0
  , 'filled_av': 126
  , 'name_box': 64
  , 'NP': {222, 228, 238}
  , 'NP 8d': 222
  , 'NP 50d': 238
  , 'scroll bar': 241
  , 'thin line': 221
  , 'time_field_bg': 245
  , 'white': 255
}
CONTRACT_MIN_H = defaultdict(
  lambda: 'NO DATA'
  , {
    'Foodora_Minijob': 5
    , 'Foodora_Working Student': 12
    , 'Midijob': 12
    , 'Minijob': 5
    , 'Minijobber': 5
    , 'Mini-Jobber': 5
    , 'TE Midijob': 12
    , 'TE Minijob': 5
    , 'TE Teilzeit': 30
    , 'TE Werkstudent': 12
    , 'TE WS': 12
    , 'Teilzeit': 30
    , 'Vollzeit': 30
    , 'Werk Student': 12
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
def calc_given_hour_ratios(avail, given, max_h):
  return {
    GIV_MAX: round(given / max_h, 2)
    , GIV_AVA: round(given / avail, 2) if avail else 10
  }
# -------------------------------------

# -------------------------------------
def check_data_and_make_comment(rider_data):
  comment = []
  check = ''
  call = ''
  min_available = not isinstance(rider_data[MIN], str)
  if rider_data[GIV_MAX] > 1:
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
      call = False
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
def check_w_avails_wo_shift_and_vice_versa(df):
  log = ''
  for msg, op_1, op_2 in ((W_AV_WO_SHIFT, ne, eq), (W_SHIFT_WO_AV, eq, ne)):
    log += print_log(msg)
    for rider in df[op_1(df[AVA], 0) & op_2(df[GIV_SHI], '')][RID_NAM]:
      log += print_log(rider, '\t- ')
  return log + print_log('-----')
# -------------------------------------

# -------------------------------------
def get_availability_data(df_row):
  return {out_col: df_row[in_col] for out_col, in_col in CONVERT_COLS_AVAIL}
# -------------------------------------

# -------------------------------------
def get_rider_min_hours(df_ee, contract, rider_name):
  return {
    MIN: CONTRACT_MIN_H[contract]
    if 'TE' in contract or rider_name not in df_ee[RID_NAM].to_numpy()
    else df_ee[df_ee[RID_NAM] == rider_name][MIN].item()
  }
# -------------------------------------

# -------------------------------------
def get_rider_month_hours(df_month, rider_id):
  rider_data = {}
  if df_month is None:
    for output_col, _ in CONVERT_COLS_MONTH:
      rider_data[output_col] = 'N/A'
    rider_data[PAI_MAX] = 'N/A'
  else:
    df_h_rider = df_month[df_month[DR_ID] == rider_id]
    if df_h_rider.shape[0] == 0:
      print(df_h_rider.shape, rider_id, df_h_rider)
    for output_col, input_col in CONVERT_COLS_MONTH:
      try:
        rider_data[output_col] = df_h_rider[input_col].item()
      except ValueError:
        rider_data[output_col] = 0
      except KeyError as ex:
        print(ex)
    try:
      work_ratio = float(str(df_h_rider[WO_RA].item()).strip('%'))
      if work_ratio > 5: 
        work_ratio /= 100
    except ValueError:
      work_ratio = 0
    rider_data[PAI_MAX] = round(work_ratio, 2)
  return rider_data
# -------------------------------------

# -------------------------------------
def get_rider_shifts(df_shifts, kw_dates, rider_id):
  given = 0
  shifts = ''
  for _, d in df_shifts[df_shifts[DR_ID] == rider_id].iterrows():
    if len(kw_dates) < 7:
      kw_dates.add(d[SH_DA].isoformat())
    given += d[WO_HO]
    shifts += f'{d[SH_DA]} | {d[FR_HO]} - {d[TO_HO]} | {d[WO_HO]}h{NL}'
  return {GIV: given, GIV_SHI: shifts}, kw_dates
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
def load_xlsx_city_in_filename(city, filename):
  return (
    filename.endswith('xlsx')
    and filename[0].isalpha()
    and any(fuzz.partial_ratio(alias, filename) > 86 for alias in ALIAS[city])
  )
# -------------------------------------

# -------------------------------------
def load_xlsx_data_into_dfs(city, dirs, log):
  missing_files = [ALIAS[AVA][0], ALIAS[SHI][0]]
  dfs = {MON: None}
  for filename in listdir(dirs[0]):
    if not load_xlsx_city_in_filename(city, filename):
      continue
    if fuzz.WRatio(STD_REP, filename) > 86:
      log += print_log(f' {STD_REP} file available, {filename = }', '|O.O|')
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
def parse_progress_bar(bar_len, prog, pre, suf):
  done = int(bar_len * prog)
  return f'{pre} [{"#" * done + "-" * (bar_len - done)}] {prog:.2%} {suf}'
  # return f'{pre} [{"█" * done + "-" * (bar_len - done)}] {prog:.2%} {suf}'
# -------------------------------------

# -------------------------------------
def parse_break_line(fil='=', text=''):
  return text.center(80, fil)
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
  return any(
    img[y_test, x_test] == COLOR['filled_av'] for y_test in (top, bot)
  )
# -------------------------------------

# -------------------------------------
def png_capture_grid(image):
  columns = None
  rows = None
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
      if pixel_color == COLOR['thin line']:
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
      if len(rows) == 2 and y_value > rows[1] + 70:
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
def png_get_current_date(day_str, kw_dates:list):
  return date.fromisoformat(kw_dates[WEEKDAYS.index(day_str)]).strftime(MD)
# -------------------------------------

# -------------------------------------
def png_name_determination(rider_names, ocr_name, det_name):
  char_cnt = len(ocr_name)
  if char_cnt >= 5:
    for rider_name in rider_names:
      query = rider_name[:char_cnt]
      slice_similarity = fuzz.WRatio(ocr_name, query)
      if slice_similarity >= 87:
        det_name = rider_name
        break
      partial_similarity = fuzz.partial_ratio(ocr_name, query)
      if partial_similarity >= 88:
        det_name = rider_name
        break
      if slice_similarity >= 70:
        det_name.append((slice_similarity, rider_name, ocr_name, 'slice'))
      if partial_similarity >= 70:
        det_name.append((partial_similarity, rider_name, ocr_name, 'partial'))
  return det_name
# -------------------------------------

# -------------------------------------
def png_ocr_read_frame(top, bot, left, first_col, image):
  return pytesseract.image_to_string(
    image[top:bot, left:int(first_col * NAME_SHARE)]
  ).strip()
# -------------------------------------

# -------------------------------------
def png_one_row_determine_rider_name(
  png, img, alt_imgs, row_n, top, bot, left, first_col, rider_names, log
):
  det_name = []
  ocr_read = []
  for image in (img, *alt_imgs):
    ocr_name = png_ocr_read_frame(top, bot, left, first_col, image)
    if not ocr_name:
      continue
    ocr_read.append(ocr_name)
    det_name = png_name_determination(rider_names, ocr_name, det_name)
    if isinstance(det_name, str) or det_name and max(det_name)[0] >= 73:
      break
  if not det_name:
    log += print_no_name_determined(png, row_n, ocr_read)
  elif isinstance(det_name, list):
    if len(det_name) > len(alt_imgs) + 1:
      det_name, log = print_multi_match(png, row_n, det_name, ocr_read, log)
    else:
      det_name = max(det_name)[1]
  return det_name, ocr_read, log
# -------------------------------------

# -------------------------------------
def png_one_row_get_availabities(img, high, low, columns, col_cnt, date_str):
  daily_avail = ''
  daily_hours = 0
  extra_hours = 0
  hours_block = 0
  in_availablity_block = False
  for col_idx, column in enumerate(columns):
    if png_avail_cell_is_filled(img, high, low, column + 1):
      if col_idx == 0 and col_cnt == 22:
        extra_hours += .5
      hours_block += .5
      if not in_availablity_block:
        daily_avail += f'{date_str} | {TIMES[col_cnt][col_idx]}'
        in_availablity_block = True
      elif col_idx == col_cnt - 1:
        extra_hours += EXTRA_HOURS[col_cnt]
        daily_avail += f' - {TIMES[col_cnt][-1]} | {hours_block: .1f}h{NL}'
        daily_hours += hours_block
        break
    elif in_availablity_block:
      daily_avail += f' - {TIMES[col_cnt][col_idx]} | {hours_block: .1f}h{NL}'
      daily_hours += hours_block
      hours_block = 0
      in_availablity_block = False
  return daily_avail, daily_hours, extra_hours
# -------------------------------------

# -------------------------------------
def png_one_row_no_data(img, y_up, y_low, x_name, x_avail):
  return (
    all(
      img[y_val, x_name] == img[y_val, x_name + 4] == img[y_val, x_name + 8]
      for y_val in range(y_up, y_low)
    )
    or all(
      img[y_val, x_avail] == COLOR['white'] for y_val in range(y_up, y_low)
    )
  )
# -------------------------------------

# -------------------------------------
def png_read_out_one_row(
  png, img, alt_imgs, grid_values, row_n, row, next_row, data, log, args
):
  columns, col_cnt, first_x, first_col = grid_values
  if png_one_row_no_data(
    img
    , y_up=row + 1
    , y_low=next_row - 4
    , x_name=(39 * first_x + first_col) // 40
    , x_avail=first_col * 24 // 25
  ):
    data[CNT][NODA] += 1
  else:
    date_str, day, png_idx, city, kw, riders = args
    avail_str, daily_h, extra_h = png_one_row_get_availabities(
      img, row + 1, next_row - 4, columns, col_cnt, date_str
    )
    rider, ocr_read, log = png_one_row_determine_rider_name(
      png, img, alt_imgs, row_n, row, next_row, first_x, first_col, riders, log
    )
    data[DNA].append(
      (kw, city, day, png_idx, row_n, avail_str[:-1], rider, ocr_read)
    )
    if not avail_str:
      data[CNT][NOAV] += 1
    elif not rider:
      data[CNT][NOCR] += 1
    elif date_str in data[DON][rider]:
      data[CNT][DUPL] += 1
    else:
      data[CNT][LINK] += 1
      data[AVA][rider].append(avail_str)
      data[HRS][rider] += daily_h
      data[XTR][rider] += extra_h
      data[DON][rider].add(date_str)
    # data = data, avail_str, ocr_read
  return data, log
# -------------------------------------

# -------------------------------------
def png_read_out_screenshot(png_dir, png_cnt, png_n, png, data, log, *args):
  img = cv.imread(join(png_dir, png), cv.IMREAD_GRAYSCALE)
  alt_imgs = [
    cv.filter2D(img, -1, SHARP_KERNEL)
    , cv.threshold(img, 212, 255, cv.THRESH_BINARY)[1]
    , cv.threshold(img, 220, 255, cv.THRESH_BINARY)[1]
  ]
  rows, row_cnt, grid_values = png_capture_grid(img)
  data[CNT][SCAN] += row_cnt
  # print(f'{png = }, first x = {grid_values[2]}, first col = {grid_values[3]}, row count = {len(rows)}, {rows = }, columns = {grid_values[0]}')
  for row_n, row in enumerate(rows[:-1], 1):
    print_progress_bar(png_cnt, png_n, row_cnt, row_n, png)
    data, log = png_read_out_one_row(
      png, img, alt_imgs, grid_values, row_n, row, rows[row_n], data, log, args
    )
    # if isinstance(data, tuple):
      # data, avails, ocr_read = data
      # print(f'{row_n = }, {avails = }, {ocr_read = }')
  return data, log
# -------------------------------------

# -------------------------------------
def png_update_report_dataframe(df, data):
  for rider, daily_avail in data[AVA].items():
    rider_df_idx = df[df[RID_NAM] == rider].index[0]
    week_h = data[HRS][rider]
    max_h = week_h + data[XTR][rider]
    df.at[rider_df_idx, AVAILS] = (
      ''.join(sorted(daily_avail))
      + f'total: {week_h}h | avail <= {max_h}{NL}'
      + '' if week_h <= df.at[rider_df_idx, AVA] <= max_h else ' '
    )
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
def print_no_name_determined(png, row_n, ocr_read):
  print('\r', end='')
  return print_log(
    f'{NF}{png = }, {row_n = }{" " * 30}'
    + (f'{NL}|OCR| {ocr_read = }' if ocr_read else '')
    , '##### ', BR
  )
# -------------------------------------

# -------------------------------------
def print_multi_match(png, row_n, det_name, ocr_read, log):
  log += print_log(f'{MUL_MAT}{png = }, {row_n = }, {ocr_read = }', '[] [] ')
  for similarity, name, ocr, source in det_name:
    log += f'{TAB}{ocr = }, {name = }, {source = }, {similarity = }{NL}'
  det_name = max(det_name)[1]
  return det_name, log + print_log(f' stored at: {det_name}', TAB, BR)
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
def process_screenshots_and_store_avails_in_df(df, kw_dates, dirs, log, *args):
  log += print_log_header(PROCESS_PNG_MSG)
  data = deepcopy(PNG_PROCESSING_DICT)
  pngs = sorted(listdir(dirs[3]))
  png_cnt = len(pngs)
  for png_n, png in enumerate(pngs):
    day = png[:-5]
    date_s = png_get_current_date(day, kw_dates)
    data, log = png_read_out_screenshot(
      dirs[3], png_cnt, png_n, png, data, log, date_s, day, int(png[-5]), *args
    )
  log += print_log(parse_stats_msg(data[CNT]), '', BR)
  df_determined = DataFrame(data[DNA], columns=DF_DET_COLUMNS)
  df_determined.to_excel(
    join(dirs[1], f'det_names_{args[0]}_{START_DT}.xlsx')
    , columns=DF_DET_COLUMNS
    , index=False
  )
  return png_update_report_dataframe(df, data), log
# -------------------------------------

# -------------------------------------
def process_raw_xlsx_data_and_store_in_df(kw, city, dirs, log):
  log += print_log_header(PROCESS_XLSX_MSG)
  # ----- read weekly xlsx data, check availability of mendatory raw files ----
  dfs, log = load_xlsx_data_into_dfs(city, dirs, log)
  if dfs is None:
    return None, None, None, log
  dfs, log = rider_ersterfassung_update_names(kw, city, dirs[4], dfs, log)
  # ----- store data from mendatory xlsx files in dataframes -----
  kw_dates = set()
  data_list = []
  for _, df_avail_row in dfs[AVA].iterrows():
    d = {report_column: '' for report_column in REPORT_HEADER}
    d.update(get_availability_data(df_avail_row))
    rider_shifts, kw_dates = get_rider_shifts(dfs[SHI], kw_dates, d[ID])
    d.update(rider_shifts)
    d.update(calc_given_hour_ratios(d[AVA], d[GIV], d[MAX]))
    d.update(get_rider_month_hours(dfs[MON], d[ID]))
    d.update(get_rider_min_hours(dfs[EE], d[CON_TYP], d[RID_NAM]))
    d.update(check_data_and_make_comment(d))
    data_list.append(d)
  # ---- create dataframe, extract additional information, filter unneeded rows
  df = DataFrame(data_list)
  log += check_w_avails_wo_shift_and_vice_versa(df)
  # ----------
  return df, dfs[AVA][USE_NAM].to_numpy(), sorted(kw_dates), log
# -------------------------------------

# -------------------------------------
def rider_ee_new_entry(city, new_name, contract, kw_monday_date):
  return {
    RID_NAM: new_name
    , CON_TYP: contract
    , MIN: CONTRACT_MIN_H[contract]
    , CIT: city
    , FIR_ENT: kw_monday_date
    , LAS_ENT: kw_monday_date
  }
# -------------------------------------

# -------------------------------------
def rider_ersterfassung_update_names(kw, city, ree_dir, dfs, log):
  log += print_log_header(SYNCH_MIN_H_MSG)
  data_list = []
  kw_date = date.fromisocalendar(date.today().year, kw, 1)
  names_mon = set() if dfs[MON] is None else {*dfs[MON][DRI]}
  log += print_log(f'{MH_AVAIL_MSG}{bool(names_mon)}', 'CHECK', BR)
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
        log += print_log(f'{name = }, {contract = }', '\t- ')
      elif name in known:
        rider_ee_idx = dfs[EE][ree_names == name].index[0]
        if kw_date > dfs[EE].at[rider_ee_idx, LAS_ENT]:
          dfs[EE].at[rider_ee_idx, LAS_ENT] = kw_date
    if new_names:
      log += print_log('-----')
  dfs[EE] = dfs[EE].append(data_list, ignore_index=True)
  dfs[EE].sort_values([LAS_ENT, FIR_ENT, RID_NAM], inplace=True)
  rider_ersterfassung_style_and_save_xlsx(city, ree_dir, dfs[EE])
  return dfs, log
# -------------------------------------

# -------------------------------------
def rider_ersterfassung_style_and_save_xlsx(city, ree_dir, df_min):
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
def save_df_in_formated_xlsx(kw, city, df):
  log = print_log_header(CREATE_XLSX_MSG)
  row_cnt = len(df) + 1
  #  ----- open instance of xlsx-file -----
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
  # ----- save xlsx-file -----
  writer.save()
# ----------
  return log + print_log(f' saved {filename}', '+++++', BR)
# -------------------------------------

# -------------------------------------
def shiftplan_check(kw, city, get_avails, merge_pngs, unzip_only, dirs):
  start = perf_counter()
  log = print_log_header(CITY_LOG_PRE + city, pre='=')
  log += zip_extract_screenshots(city, dirs, merge_pngs)
  if unzip_only:
    return log
  df, rider_names, kw_dates, log = process_raw_xlsx_data_and_store_in_df(
    kw, city, dirs, log
  )
  if df is None:
    return log
  if get_avails and TESSERACT_AVAILABLE:
    df, log = process_screenshots_and_store_avails_in_df(
      df, kw_dates, dirs, log, city, kw, rider_names
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
  for zip_file in zip_get_city_files(city, dirs[0]):
    log += print_log(zip_file, TAB)
    with ZipFile(join(dirs[0], zip_file)) as zfile:
      day = WEEKDAYS[1]
      idx = 1
      for member in sorted(zfile.namelist()):
        f_name = basename(member)
        if not f_name:
          continue
        f_name, idx, day, log = zip_parse_png_filename(f_name, idx, day, log)
        source = zfile.open(member)
        with open(join(dirs[3], f_name), "wb") as target:
          copyfileobj(source, target)
    log += print_log('-----')
  log += print_log(f' saved PNGs in: {dirs[3]}', '+++++', BR)
  if merge_pngs:
    log += zip_merge_png_files_per_day(city, dirs)
  return log
# -------------------------------------

# -------------------------------------
def zip_get_city_files(city, kw_dir):
  return (
    fn for fn in listdir(kw_dir)
    if fn.endswith('.zip') and any(ac in fn.casefold() for ac in ALIAS[city])
  )
# -------------------------------------

# -------------------------------------
def zip_merge_get_daily_files(Image, png_dir, day):
  return [
      Image.open(join(png_dir, day_fn))
      for day_fn in sorted(fn for fn in listdir(png_dir) if day in fn)
    ]
# -------------------------------------

# -------------------------------------
def zip_merge_png_files_per_day(city, dirs):
  from PIL import Image
  log = print_log_header(MERGE_FILES_MSG)
  for day in WEEKDAYS:
    images = zip_merge_get_daily_files(Image, dirs[3], day)
    widths, heights = zip(*(img.size for img in images))
    new_image = Image.new('RGB', (max(widths), sum(heights)))
    y_offset = 0
    for img in images:
      new_image.paste(img, (0, y_offset))
      y_offset += img.size[1]
    daily_img_fn = f'{city}_{day}.png'
    new_image.save(join(dirs[2], daily_img_fn))
    log += print_log(f' saved {daily_img_fn}', '+++++')
  return log + print_log('-----')
# -------------------------------------

# -------------------------------------
def zip_parse_png_filename(original, file_idx, last_day, log):
  similarity = 0
  current_day = ''
  for weekday in WEEKDAYS:
    weekday_similarity = fuzz.partial_ratio(original[:-5], weekday)
    if weekday_similarity > similarity:
      similarity = weekday_similarity
      current_day = weekday
      if weekday_similarity == 100:
        break
  if current_day != last_day:
    file_idx = 1
  saved_as = f'{current_day}{file_idx}.png'
  if similarity != 100:
    log += print_log(f'{original = }, {saved_as = }', '\t- ')
  return saved_as, file_idx + 1, current_day, log
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
      print(f'##### Couldn`t find "{kw_dir}" ...')
      return None
    log_dir = check_make_dir(kw_dir, 'logs')
    screen_dir = join(kw_dir, 'Screenshots')
    ree_dir = check_make_dir(FILE_DIR, EE)
    for city in cities:
      png_dir = check_make_dir(screen_dir, city)
      dirs = (kw_dir, log_dir, screen_dir, png_dir, ree_dir)
      log += shiftplan_check(
        kw, city, get_avails, merge_pngs, unzip_only, dirs)
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
  args = parser.parse_args()
  main(*args.__dict__.values())
# -------------------------------------
# =================================================================
