# =================================================================
# ### IMPORTS ###
# =================================================================
# -------------------------------------
from collections import defaultdict
from copy import deepcopy
from datetime import date, datetime, timedelta
import getpass
import io
import json
from os import get_terminal_size, listdir, makedirs
from os.path import abspath, basename, dirname, exists, join
import shutil
import signal
import sys
from time import perf_counter
from zipfile import ZipFile
# -------------------------------------
import cv2 as cv
from fuzzywuzzy import fuzz
import msoffcrypto
import numpy as np
from pandas import DataFrame, ExcelWriter, read_excel, to_datetime
from PIL import Image
from pytesseract.pytesseract import (
  get_tesseract_version, image_to_string, TesseractNotFoundError
)
from xlrd.biffh import XLRDError
# -------------------------------------
# =================================================================


# =================================================================
# ### COMMENTS, CONSTANTS, FORMATS, GLOBAL PATHS, LOGS, PRINTS ###
# =================================================================
# -------------------------------------
# ### DATE AND TIME RELATED ###
# -------------------------------------
DMY = '%d.%m.%Y'
HM = '%H:%M'
HMS = HM + ':%S'
MD = '%m-%d'
YMD = '%Y-%m-%d'
Y_S = '%Y_%m_%d_%H_%M_%S'
_now = datetime.now()
START_DT = _now.strftime(Y_S)
TOO_OLD = timedelta(weeks=8)
YEAR =  _now.year
# -------------------------------------

# -------------------------------------
# ### OPENCV ###
# -------------------------------------
CLAHE_DEF = cv.createCLAHE()
CLAHE_L2_S6 = cv.createCLAHE(clipLimit=2.0, tileGridSize=(6, 6))
CLAHE_L3_S7 = cv.createCLAHE(clipLimit=3.0, tileGridSize=(7, 7))
CLAHE_L4_S7 = cv.createCLAHE(clipLimit=4.0, tileGridSize=(7, 7))
CLAHE_L4_S9 = cv.createCLAHE(clipLimit=4.0, tileGridSize=(9, 9))
KERNEL_SHARP = np.array(([-1, -1, -1], [-1, 9, -1], [-1, -1, -1]), dtype="int")
# -------------------------------------

# -------------------------------------
# ### STRINGS ###
# -------------------------------------
ALI = 'align'
AVA = 'avail'
AVAILS = 'availablities'
BAD_RESO = 'bad_resolution'
BG_COL = 'bg_color'
BOLD = 'bold'
BOR = 'border'
BR = '\n-----'
CAL = 'call'
CC_FE = 'current contract first entry'
CEN = 'center'
CHE = 'check'
CITY_LOG_PRE = ' LOG FOR CITY:'
CIT = 'city'
CMT = 'comment'
CNT = 'counter'
CONFIG_MISSING_MSG = '##### missing config file, using default values ...\n'
CON_H = 'Contracted hours'
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
FI_ENT = 'first entry'
FO_SI = 'font_size'
FR_HO = 'From Hour'
GIV = 'given'
GIV_AVA = 'given/avail'
GIV_MAX = 'given/max'
GIV_SHI = 'given shifts'
HRS = 'hours'
H_AV = 'Hours Available'
ID = 'ID'
IMG = 'img'
IMG_VARIATIONS = 'img_data'
INITIAL_MSG = f'SHIFTPLAN CHECK {START_DT}'
JPG_NAME_CHECK_MSG = 'Check available JPG file names'
PNG_NAME_CHECK_MSG = 'Check available PNG file names'
LA_ENT = 'last entry'
LEF = 'left'
LINK = 'linked'
MAX = 'max'
MAX_H = 'Total Availability'
MAX_HOURS_MSG = 'mehr als max. Std.'
MAX_T = 'max_type'
MAX_V = 'max_value'
MERGE_FILES_MSG = ' MERGING DAILY PNGS TO DAILY FILE '
MH_AVAIL_MSG = ' "Monatsstunden" file available: '
MID_T = 'mid_type'
MID_V = 'mid_value'
MINI_LIMIT_MSG = 'Minijobber Monatsmax Std. prüfen'
MIN = 'min'
MIN_HOURS_MSG = 'weniger als Min.Std.'
MIN_H_CHECK_MSG = 'Min.Std. prüfen'
MIN_T = 'min_type'
MIN_V = 'min_value'
MISSING_FILE_MSG = ' MISSING MEDATORY FILE '
MON = 'month'
MORE_HOURS_MSG = 'mehr Stunden'
MORE_THAN_AVAIL_MSG = 'mehr Std. als Verfügbarkeiten'
MUL_MAT = 'MULTIPLE MATCHES '
NF = 'NOT FOUND '
NL = '\n'
NOAV = 'no_avail'
NOCR = 'no_ocr'
NODA = 'no_data'
NOT_IN_MON = ' not in "Monatsstunden": '
NO_AVAILS_MSG = 'keine Verfügbarkeiten'
NO_DA = 'NO DATA'
NO_SCREENS_MSG = 'NO SCREENSHOTS AVAILABLE'
NP = 'not_planable'
NU = 'num'
NU_FO = 'num_format'
PAI = 'paid'
PAI_MAX = 'paid/max'
PRE_C = 'prev contracts'
PRE_F = 'prev first entries'
PRE_L = 'prev last entries'
PROCESS_PNG_MSG = ' SCAN AVAILABILITES FROM PNGS '
PROCESS_XLSX_MSG = ' PROCESS RAW XLSX DATA '
PW = None
REDUCE_HOURS = ' -> auf Min.Std. reduzieren'
REP = 'report'
RE_MA = 'resize_margin'
RE_ROW = 'resized rows'
RID_NAM = 'rider name'
RI = 'right'
ROI = 'region_of_interest'
SHI = 'shift'
SCAN = 'scanned'
SIC = 'sick'
SIM_NAM = 'similar names'
SH_DA = 'Shift Date'
STD_REP = 'stundenreports'
STRIP_CHARS = """ .,-_'`"()|"""
SYNCH_MIN_H_MSG = ' SYNCHRONIZE NAMES IN MINDESTSTUNDEN LIST '
TAB = '\t'
TIDY_JPG_MSG = ' TIDY JPGS IN WORKING DIRECTORY '
TIDY_PNG_MSG = ' TIDY PNGS IN WORKING DIRECTORY '
TOP = 'top'
TO_HO = 'To Hour'
TYP = 'type'
UNK = 'unknown'
UNP = 'unpaid'
UNZIP_MSG = ' UNZIP CITY PNG FILES '
USER_N = 'User Name'
USER_T = 'User Type'
U_ID = 'User ID'
VAC = 'vacation'
VALID = 'min_valid'
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
BASE_DIR = dirname(abspath(__file__))
CONFIG_FP = join(BASE_DIR, 'config_report.json')
EE_BACKUP = f'{EE}.xlsx'
LOG_FN = f'report_{START_DT}.log'
OUTPUT_DIR = join(BASE_DIR, 'Schichtplan_bearbeitet')
if not exists(OUTPUT_DIR):
  makedirs(OUTPUT_DIR)
OUT_FILE_PRE = ''
SPD_DIR = join(BASE_DIR, 'Schichtplan_Daten', str(YEAR))
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
  , PAI_MAX, WOR, VAC, SIC, PAI, UNP, CMT, CAL, 'cmt shift coordinator'
)
RIDER_MIN_HEADER = (
  RID_NAM, CON_TYP, MIN, CIT, FI_ENT, LA_ENT, CC_FE, PRE_C, SIM_NAM
)
TIMEBLOCK_STRINGS = (
  '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30'
  , '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30'
  , '19:00', '19:30', '20:00', '20:30', '21:00', '21:30', '22:00', '22:30'
  , '23:00', '23:30'
)
WEEKDAYS = [
  'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag'
  , 'Samstag', 'Sonntag'
]
WEEKDAY_ABREVATIONS = (
  ('mo', 'mon'), ('di', 'tue'), ('mi', 'wed'), ('do', 'thu'), ('fr', 'fri')
  , ('sa', 'sat'), ('so', 'sun')
)
# -------------------------------------

# -------------------------------------
# ### DICTS ###
# -------------------------------------
COLOR = {
  'empty_av_field': 225
  , 'filled': {*range(93, 111), 126}
  , 'name_box': 64
  , 'NP': {228, 238}
  , 'NP 8d': 222
  , 'NP 50d': 238
  , 'scroll bar': {*range(236, 246)}
  , 'shift': {*range(143, 151)}
  , 'line': {195, *range(208, 234)}
  , 'time field': {245}
  , 'white': {*range(246, 256)}
}
COLOR['rmv'] = COLOR['white'] | COLOR['NP']
CON_BY_H = defaultdict(
  lambda: NO_DA, {
    45: 'TE Minijob'
    , 47: 'TE Minijob'
    , 60: 'Foodora_Midijob'
    , 80: 'TE WS'
    , 87: 'TE WS'
    , 130: 'TE Teilzeit'
    , 136: 'TE Midijob'
    , 174: 'Vollzeit'  
  }
)
CON_BY_N = defaultdict(
  lambda: (NO_DA, NO_DA), {
    'Arbeitnehmerüberlassung': (0, 20)
    , 'Foodora_Midijob': (12, 40)
    , 'Foodora_Minijob': ('10h/Monat', 15)
    , 'Foodora_Working Student': (12, 20)
    , 'Midijob': (12, 40)
    , 'Minijob': ('10h/Monat', 15)
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
    , 'Working Student': (12, 20)
    , 'Working-Student': (12, 20)
  }
)
EXTRA_HOURS_BEGIN = defaultdict(int, {21: 1, 22: .5})
EXTRA_HOURS_END = {21: 1, 22: 1, 23: 1, 24: .5, 25: 0, 26: 0}
PNG_PROCESSING_DICT = {
  AVA: defaultdict(list)
  , DON: defaultdict(set)
  , HRS: defaultdict(int)
  , XTR: defaultdict(int)
  , CNT: defaultdict(int)
  , DNA: list()
}
TIMES = {
  21: TIMEBLOCK_STRINGS[2:-2]
  , 22: TIMEBLOCK_STRINGS[1:-2]
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
  , ('G:G', 25, 'text')
  , ('H:H', 35, 'text')
  , ('I:I', 15, 'text')
  # , ('J:J', 20, 'text')
)
RIDER_EE_CONDS = (('C2:C', MIN, 'red'), ('F2:F', 'date', 'old'))
RIDER_EE_FMTS = ('int', 'old', 'red', 'text')
XLS_REPORT_COL_FMT = (
  ('A:A', 5, 'int')
  , ('B:B', 24, 'text')
  , ('C:C', 11, 'text')
  , ('D:D', 4, 'int')
  , ('E:E', 8, 'int')
  , ('F:F', 4, 'int')
  , ('G:G', 5, 'int')
  , ('H:I', 9, 'ratio')
  , ('J:J', 31, 'text')
  , ('K:K', 27, 'text')
  , ('L:L', 8, 'ratio')
  , ('M:M', 6, 'int')
  , ('N:P', 7, 'int')
  , ('Q:Q', 6, 'int')
  , ('R:R', 27, 'comment')
  , ('S:S', 5, 'int')
  , ('T:T', 16, 'comment')
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
    from pytesseract import pytesseract
    pytesseract.tesseract_cmd = config['tesseract']['cmd_path']
  if 'password' in config:
    PW = config['password'] or None
try:
  get_tesseract_version()
  TESSERACT_AVAILABLE = True
except TesseractNotFoundError:
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
def check_make_dir(*args):
  dir_path = join(*args)
  if not exists(dir_path):
    makedirs(dir_path)
  return dir_path
# -------------------------------------

# -------------------------------------
def create_report_df(dfs, ref_data, log):
  data_list = []
  src = AVA if dfs[MON] is None else MON
  for name, df_row in dfs[src].iterrows():
    data = new_report_data_entry()
    data, contract, log = get_base_data(data, name, df_row, src, dfs[AVA], log)
    data.update(get_min_hours(dfs[EE], ref_data[3], contract, name))
    data.update(get_shifts_and_kw_dates(dfs[SHI], data[ID]))
    data.update(get_given_hour_ratios(data[AVA], data[GIV], data[MAX]))
    data.update(get_data_check_and_first_comment(data))
    data_list.append(data)
  return DataFrame(data_list).set_index(RID_NAM, drop=False)
# -------------------------------------

# -------------------------------------
def get_base_data(data, name, df_row, source, df_ava, log):
  if source == AVA:
    contract = df_row[USER_T]
    data.update(get_base_data_from_ava(contract, name, df_row))
    data.update(get_max_hours(contract, df_row))
  else:
    data, log = get_base_data_from_mon(data, name, df_row, log)
    data, contract = get_contract_and_av_h(data, name, df_row, df_ava)
    data.update(get_max_hours(contract, df_ava, name))
  return data, contract, log
# -------------------------------------

# -------------------------------------
def get_base_data_from_ava(contract, name, df_row):
  return {
    **{output_col: 'N/A' for output_col, _ in CONVERT_COLS_MONTH}
    , AVA: df_row[H_AV]
    , ID: df_row[U_ID]
    , RID_NAM: name
    , CON_TYP: contract
  }
# -------------------------------------

# -------------------------------------
def get_base_data_from_mon(data, name, df_row, log):
  if df_row is None:
    for output_col, _ in CONVERT_COLS_MONTH:
      data[output_col] = 'N/A'
    return data, log
  for output_col, input_col in CONVERT_COLS_MONTH[1:]:
    try:
      data[output_col] = df_row[input_col]
    except ValueError:
      data[output_col] = 0
  try:
    work_ratio = float(str(df_row[WO_RA]).strip('%'))
    if work_ratio > 5: 
      work_ratio /= 100
  except ValueError:
    work_ratio = 0
  finally:
    work_ratio = round(work_ratio, 2)
  data.update({PAI_MAX: work_ratio, ID: df_row[DR_ID], RID_NAM: name})
  return data, log
# -------------------------------------

# -------------------------------------
def get_contract_and_av_h(data, name, df_row, df_ava):
  try:
    ava_row = df_ava.loc[name]
  except KeyError:
    avail = ''
    contract = getattr(df_row, CO_TY, None)
    if contract is None:
      contract = CON_BY_H[df_row[CON_H]]
  else:
    avail = ava_row.squeeze()[H_AV]
    contract = ava_row[USER_T]
  data.update({CON_TYP: contract, AVA: avail})
  return data, contract
# -------------------------------------

# -------------------------------------
def get_data_check_and_first_comment(data):
  comment = []
  call = ''
  if isinstance(data[MIN], int):
    min_h = data[MIN] 
  elif 'h/' in data[MIN]:
    hours, period = data[MIN].split('h/')
    if period == 'Monat':
      min_h = max(3, (int(hours) * 6 // 13) / 2)
    else:
      min_h = int(hours)
  else:
    min_h = None
  if min_h is None:
    comment.append(MIN_H_CHECK_MSG)
  if data[GIV_MAX] > 1 and data[MAX] != 40:
    comment.append(MAX_HOURS_MSG)
    call = 'X'
  if data[GIV_AVA] == 10:
    comment.append(NO_AVAILS_MSG)
    if data[GIV] > min_h:
      comment[-1] += REDUCE_HOURS
      call = 'X'
  else:
    if data[GIV] < min_h:
      comment.append(MIN_HOURS_MSG)
      if data[GIV] != data[AVA]:
        call = 'X'
    if data[GIV_AVA] > 1:
      comment.append(MORE_THAN_AVAIL_MSG)
  if 'mini' not in data[CON_TYP].casefold():
    threshold = .75
  else:
    if not isinstance(data[PAI_MAX], str) and data[PAI_MAX] > .9:
      threshold = 0
      comment.append(MINI_LIMIT_MSG)
      call = ''
    else:
      threshold = .55
  if data[GIV_MAX] < threshold and data[GIV_AVA] < threshold:
    comment.append(MORE_HOURS_MSG)
  return {CMT: NL.join(comment), CAL: call, AVAILS: ''}
# -------------------------------------

# -------------------------------------
def get_given_hour_ratios(avail, given, max_h):
  if isinstance(avail, str):
    return {GIV_MAX: 0, GIV_AVA: 0}
  return {
    GIV_MAX: round(given / max_h, 2) if not isinstance(max_h, str) else 0
    , GIV_AVA: round(given / avail, 2) if avail else 10
  }
# -------------------------------------

# -------------------------------------
def get_max_hours(contract, df_row, name=None):
  try:
    if name is not None:
      df_row = df_row.loc[name]
    return {MAX: df_row[MAX_H]}
  except KeyError:
    return {MAX: CON_BY_N[contract][1]}
# -------------------------------------

# -------------------------------------
def get_min_hours(df_ee, ref_names_set, contract, name):
  return {
    MIN: (
      CON_BY_N[contract][0]
      if 'TE' in contract or name not in ref_names_set
      else df_ee.at[name, MIN]
    )
  }
# -------------------------------------

# -------------------------------------
def get_shifts_and_kw_dates(df_shifts, rider_id):
  given = 0
  shifts = ''
  for _, d in df_shifts[df_shifts[DR_ID] == rider_id].iterrows():
    given += d[WO_HO]
    shifts += f'{d[SH_DA]} | {d[FR_HO]} - {d[TO_HO]} | {d[WO_HO]}h{NL}'
  return {GIV: given, GIV_SHI: shifts}
# -------------------------------------

# -------------------------------------
def invalid_city_xlsx_filename(filename, city):
  return not (
    filename.endswith('xlsx')
    and filename[0].isalpha()
    and all(word not in filename for word in ('wochenstunden', 'gefahrene'))
    and any(fuzz.partial_ratio(alias, filename) > 86 for alias in ALIAS[city])
  )
# -------------------------------------

# -------------------------------------
def load_avail_xlsx_into_df(df):
  df[USER_T] = df[USER_T].apply(lambda x: x.strip().replace('Foodora_', ''))
  df.loc[df[USER_T] == 'TE Werkstudent', USER_T] = 'TE WS'
  return df.drop_duplicates().set_index(USER_N, drop=False).sort_index()
# -------------------------------------

# -------------------------------------
def load_decrpyted_xlsx(file_path):
  try:
    df = read_excel(file_path)
  except XLRDError:
    global PW
    print('file is encrypted... decrypting')
    file = msoffcrypto.OfficeFile(open(file_path, 'rb'))
    if PW is None:
      PW = getpass.getpass('enter password:')
    file.load_key(password=PW)
    decrypted = io.BytesIO()
    file.decrypt(decrypted)
    df = read_excel(decrypted)
  return df.rename(columns=lambda x: str(x).strip())
# -------------------------------------

# -------------------------------------
def load_ersterkennung_xlsx_into_df(city, ree_dir):
  try:
    df_re = read_excel(parse_city_ee_filepath(city, ree_dir))
  except FileNotFoundError:
    try:
      df_re = read_excel(EE_BACKUP, city)
      df_re[SIM_NAM] = ''
      df_re[SIM_NAM] = rider_ee_get_similar_names(df_re[RID_NAM], df_re)
    except (FileNotFoundError, XLRDError):
      df_re = DataFrame(columns=RIDER_MIN_HEADER)
  df_re[FI_ENT] = to_datetime(df_re[FI_ENT], format=YMD).dt.date
  df_re[LA_ENT] = to_datetime(df_re[LA_ENT], format=YMD).dt.date
  df_re[CC_FE] = to_datetime(df_re[CC_FE], format=YMD).dt.date
  return (
    df_re.fillna('').drop_duplicates([RID_NAM]).set_index(RID_NAM, drop=False)
  )
# -------------------------------------

# -------------------------------------
def load_month_xlsx_into_df(df):
  if CO_TY in df.columns:
    df[CO_TY] = df[CO_TY].apply(lambda x: x.strip().replace('Foodora_', ''))
    df.loc[df[CO_TY] == 'TE Werkstudent', CO_TY] = 'TE WS'
  return df.drop_duplicates(ignore_index=True).set_index(DRI, drop=False)
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
  return df.sort_values([DRI, SH_DA, FR_HO]).set_index(DR_ID, drop=False)
# -------------------------------------

# -------------------------------------
def load_xlsx_data_into_dfs(dirs, city, log):
  mendatory = [ALIAS[AVA][0], ALIAS[SHI][0]]
  dfs = {MON: None}
  for filename in listdir(dirs[0]):
    fn_cf = filename.casefold().replace('_', ' ').replace('-', ' ')
    if invalid_city_xlsx_filename(fn_cf, city):
      continue
    if fuzz.WRatio(STD_REP, fn_cf) > 86:
      log += print_log(f'|O.O| {STD_REP} file available, {filename = }')
      continue
    df = load_decrpyted_xlsx(join(dirs[0], filename))
    if any(fuzz.partial_ratio(alias, fn_cf) > 86 for alias in ALIAS[AVA]):
      dfs[AVA] = load_avail_xlsx_into_df(df)
      mendatory.remove(ALIAS[AVA][0])
    elif any(fuzz.partial_ratio(alias, fn_cf) > 86 for alias in ALIAS[SHI]):
      dfs[SHI] = load_shift_xlsx_into_df(df)
      mendatory.remove(ALIAS[SHI][0])
    elif any(fuzz.partial_ratio(alias, fn_cf) > 86 for alias in ALIAS[MON]):
      dfs[MON] = load_month_xlsx_into_df(df)
  if mendatory:
    dfs = None
    log += print_log_header(f'{MISSING_FILE_MSG}{mendatory}', '#', '', '')
  else:
    dfs[EE] = load_ersterkennung_xlsx_into_df(city, dirs[4])
  return dfs, log
# -------------------------------------

# -------------------------------------
def new_report_data_entry():
  return {report_column: '' for report_column in REPORT_HEADER}
# -------------------------------------

# -------------------------------------
def parse_availabilities_string(avails, week, extra, stored=0):
  max_h = week + extra
  if isinstance(stored, str):
    suf = 'NOT IN AVAILS'
  elif week <= stored <= max_h:
    suf = ''
  else:
    suf = ' '
  return f'{"".join(sorted(avails))}total: {week}h | avail <= {max_h}{NL}{suf}'
# -------------------------------------

# -------------------------------------
def parse_break_line(fil='=', text=''):
  return text.center(80, fil)
# -------------------------------------

# -------------------------------------
def parse_city_ee_filepath(city, ree_dir):
  return join(ree_dir, f'{EE}_{city}.xlsx')
# -------------------------------------

# -------------------------------------
def parse_city_runtime(city, start):
  return f'runtime {city = }: {perf_counter() - start:.2f} s'
# -------------------------------------

# -------------------------------------
def parse_date(day_str, kw_dates):
  return date.fromisoformat(kw_dates[WEEKDAYS.index(day_str)]).strftime(YMD)
# -------------------------------------

# -------------------------------------
def parse_progress_bar(bar_len, prog, pre, suf):
  done = int(bar_len * prog)
  return f'{pre} [{"#" * done + "-" * (bar_len - done)}] {prog:.2%} {suf}'
# -------------------------------------

# -------------------------------------
def parse_run_end_msg(start):
  return f'TOTAL RUNTIME: {perf_counter() - start:.2f} s'
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
def png_cv_data(rows, row_cnt, left, first_col, img):
  n_top = row_cnt // 2
  row_height = rows[n_top + 1] - rows[n_top]
  resize_factor = 109 / row_height
  margin = 0 if rows[-1] - rows[-2] < row_height * 2 // 3 else 3
  right = int(.29 * first_col)
  res_width = int(resize_factor * right)
  res_height = int(resize_factor * img.shape[0])
  res_img = cv.resize(img[:, left:right], (res_width, res_height))
  return {
    BAD_RESO: row_height < 21
    , IMG: img
    , IMG_VARIATIONS: (
      cv.threshold(res_img, 220, 255, cv.THRESH_BINARY)[1]
      , CLAHE_DEF.apply(res_img)
      , cv.threshold(res_img, 212, 255, cv.THRESH_BINARY)[1]
      , cv.filter2D(res_img, -1, KERNEL_SHARP)
      , CLAHE_L3_S7.apply(res_img)
    )
    , RE_MA: int(resize_factor * margin)
    , RE_ROW: [int(resize_factor * row) for row in rows]
    , RI: right
    , ROI: [0, 0, 0, res_width]
    , VALID: (
      33 if row_height < 19 else 66 if row_height > 65 else row_height * 11 / 6
    )
  }
# -------------------------------------

# -------------------------------------
def png_cv_data_update(row_n, rows, cv_data):
  top = rows[row_n - 1] + 2
  bot = rows[row_n] - 2
  cv_data[ROI][0] = cv_data[RE_ROW][row_n - 1]
  cv_data[ROI][1] = cv_data[RE_ROW][row_n] - cv_data[RE_MA]
  cv_data[NP] = cv_data['img'][top + 2, cv_data[RI]] in COLOR['NP']
  return top, bot, cv_data
# -------------------------------------

# -------------------------------------
def png_get_grid_values(img):
  rows, left = png_grid_capture_rows(img)
  cols, x_valid = png_grid_capture_cols(rows, left, img)
  if x_valid is None:
    return None
  rows = png_grid_remove_invalid_rows(rows, x_valid, img)
  return rows, len(rows) - 1, left, cols[0], cols, len(cols)
# -------------------------------------

# -------------------------------------
def png_grid_capture_cols(rows, left, img):
  img_width = img.shape[1]
  cols = []
  row_cnt = len(rows)
  for row_n in range(row_cnt // 2, row_cnt):
    cols.clear()
    row_height = rows[row_n] - rows[row_n - 1]
    bot_margin = 4 if row_n == row_cnt -1 else 2
    for x in range(left + 8 * row_height, img_width - 4):
      pixel_color = img[rows[row_n] - bot_margin, x]
      if pixel_color in COLOR['time field'] | COLOR['shift']:
        break
      elif pixel_color in COLOR['line']:
        if cols and x <= cols[-1] + 5:
          del cols[-1]
        cols.append(x)
        if len(cols) == 26:
          break
      elif cols and pixel_color in COLOR['scroll bar']:
        if x <= cols[-1] + 2:
          del cols[-1]
        break
    if len(cols) >= 21:
      break
  x_row_validation = left + (cols[0] - left) * 19 // 20 if cols else None
  return cols, x_row_validation
# -------------------------------------

# -------------------------------------
def png_grid_capture_rows(img):
  left = 0
  rows = []
  height = img.shape[0]
  height_thresh = height // 2
  for x in range(0, 50):
    vert_line_cnt = 0
    rows = [height]
    y = 0
    for y in range(height - 1, -1, -1):
      pixel_color = img[y, x]
      if pixel_color not in COLOR['line']:
        vert_line_cnt = 0
      elif all(img[y, x + i] in COLOR['NP'] for i in range(3)):
        continue
      elif all(img[y, x + i] in COLOR['line'] for i in range(50, 201, 50)):
        if y > height - 4:
          del rows[0]
        vert_line_cnt = 0
        rows.append(y)
      else:
        vert_line_cnt += 1
        if vert_line_cnt == 5:
          break
      if y <= rows[-1] - (40 if y < height_thresh else 110):
        if len(rows) > 1:
          rows.append(2 * rows[-1] - rows[-2] + 4)
        else:
          y = height
        break
    if y < height_thresh:
      left = x
      break
  if 15 < rows[-1] <= 35:
    rows.append(0)
  n_bot = len(rows) // 2
  row_height = rows[n_bot] - rows[n_bot + 1]
  if rows[0] >= rows[1] + 1.3 * row_height:
    del rows[0]
    rows.insert(0, rows[0] + row_height)
  return rows[::-1], left
# -------------------------------------

# -------------------------------------
def png_grid_remove_invalid_rows(rows, x, img):
  thresh = len(rows) // 2
  x_2 = x * 7 // 8
  del_rows = []
  for row_n in range(len(rows) - 2, -1, -1):
    row_height = rows[row_n + 1] - rows[row_n]
    if row_height < 15:
      del_rows.append((row_n + 1) if row_n > thresh else row_n)
      continue
    if row_n < thresh:
      del_row = row_n
      top = rows[row_n] + 2
      bot = rows[row_n] + row_height * 3 // 5
    else:
      del_row = row_n + 1
      top = rows[row_n] + row_height * 2 // 5
      bot = rows[row_n + 1] - 2
    if all(img[y, x] in COLOR['rmv'] | {img[y, x_2]} for y in range(top, bot)):
      del_rows.append(del_row)
  for del_row in del_rows:
    del rows[del_row]
  return rows
# -------------------------------------

# -------------------------------------
def png_name_determination(ref_data, cv_data):
  hit = ''
  img_n = 0
  ocr_read = []
  scores = defaultdict(int)
  score_list = []
  for _, ocr_name in png_ocr_yield_name_from_image_variations(cv_data):
    ocr_read.append(ocr_name)
    if len(ocr_name) < 6:
      if ocr_name:
        img_n += 1
      score_list.append({})
      continue
    img_n += 1
    try:
      img_score = score_list[ocr_read[:-1].index(ocr_name)]
      for name, score in img_score.items():
        scores[name] += score
    except ValueError:
      hit, scores, img_score = png_name_variations(ocr_name, scores, ref_data)
      if hit:
        return hit, ocr_read
    score_list.append(img_score)
    hit = png_name_score_check(scores, img_n, cv_data[VALID])
    if hit:
      return hit, ocr_read
  if not (hit or cv_data[BAD_RESO]):
    hit = png_name_score_check(scores, img_n, cv_data[VALID], 1.01)
    if not hit:
      hit = png_name_fallback_algo(ocr_read)
  return hit, ocr_read
# -------------------------------------

# -------------------------------------
def png_name_fallback_algo(ocr_read):
  for name in ocr_read:
    if len(name) < 6 or ocr_read.count(name) < 2 or ' ' not in name:
      continue
    if all(char.isalpha() for char in name):
      return name
  return ''
# -------------------------------------

# -------------------------------------
def png_name_score_check(scores, img_n, min_total_perc, leading_factor=1.35):
  if scores and img_n:
    msn, *rest = png_name_sorted_scores(scores)
    valid_score = img_n * 2 * min_total_perc
    if len(scores) > 1:
      valid_score = max(leading_factor * rest[0][1], valid_score)
    if msn[1] >= valid_score:
      return msn[0]
  return ''
# -------------------------------------

# -------------------------------------
def png_name_sorted_scores(scores, result_cnt=5):
  return sorted(scores.items(), key=lambda x: x[1], reverse=True)[:result_cnt]
# -------------------------------------

# -------------------------------------
def png_name_variations(ocr_name, scores, ref_data):
  char_cnt = len(ocr_name)
  hit = ''
  idx = 0
  img_score = defaultdict(int)
  query = ''
  for idx, name in enumerate(ref_data[0]):
    query = name[:char_cnt]
    slice_similarity = fuzz.WRatio(ocr_name, query)
    if slice_similarity >= 89:
      hit = name
      break
    partial_similarity = fuzz.partial_ratio(ocr_name, query)
    if partial_similarity >= 93:
      hit = name
      break
    if fuzz.partial_ratio(ocr_name, name) >= 95:
      hit = name
      break
    if slice_similarity > 44:
      scores[name] += slice_similarity
      img_score[name] += slice_similarity
    if partial_similarity > 44:
      scores[name] += partial_similarity
      img_score[name] += partial_similarity
  if hit and ref_data[2][idx]:
    highest = 0
    simil_names = ref_data[2][idx]
    if any(fuzz.WRatio(hit, name) == 100 for name in simil_names):
      method = 'ratio'
    else:
      method = 'WRatio'
    for name in (hit, *simil_names):
      similarity = getattr(fuzz, method)(ocr_name, name)
      if similarity > 95:
        hit = name
        break
      elif similarity > highest:
        hit = name
        highest = similarity
  return hit, scores, img_score
# -------------------------------------

# -------------------------------------
def png_ocr_yield_name_from_image_variations(cv_data):
  for idx, image in enumerate(cv_data[IMG_VARIATIONS], 1):
    top, bot, left, right = cv_data[ROI]
    ocr = image_to_string(image[top:bot,left:right], config='--psm 7').strip()
    if cv_data[NP]:
      for split_chars in ('NP', 'Np', 'DA', 'wp', 'pA', ' pa'):
        ocr = ocr.rsplit(split_chars, maxsplit=1)[0]
    yield idx, ocr.strip(STRIP_CHARS)
# -------------------------------------

# -------------------------------------
def png_read_row(av_args, cv_data, img, row_n, png, data, args):
  log = ''
  ref_data, date_str, kw, city, day, idx = args
  avails, hours, extra = png_row_availabities(*av_args, date_str, img)
  if not avails:
    data[CNT][NOAV] += 1
    return data, log
  name, ocr_read = png_name_determination(ref_data, cv_data)
  data[DNA].append((kw, city, day, idx, row_n, avails[:-1], name, ocr_read))
  if not name:
    data[CNT][NOCR] += 1
    log += print_no_name_determined(avails, name, ocr_read, row_n, png)
  elif date_str in data[DON][name]:
    data[CNT][DUPL] += 1
  else:
    data[CNT][LINK] += 1
    data[AVA][name].append(avails)
    data[HRS][name] += hours
    data[XTR][name] += extra
    data[DON][name].add(date_str)
  return data, log
# -------------------------------------

# -------------------------------------
def png_read_screenshot(png_cnt, png_n, png_dir, png, data, args):
  log = ''
  img = cv.imread(join(png_dir, png), cv.IMREAD_GRAYSCALE)
  grid_values = png_get_grid_values(img)
  if grid_values is None or grid_values[1] == 0:
    return data, log
  rows, row_cnt, left, x_time, cols, col_cnt = grid_values
  data[CNT][SCAN] += row_cnt
  cv_data = png_cv_data(rows, row_cnt, left, x_time, img)
  for row_n in range(1, row_cnt + 1):
    print_progress_bar(png_cnt, png_n, row_cnt, row_n, png)
    top, bot, cv_data = png_cv_data_update(row_n, rows, cv_data)
    av_args = (top, bot, cols, col_cnt)
    data, row_log = png_read_row(av_args, cv_data, img, row_n, png, data, args)
    log += row_log
  return data, log
# -------------------------------------

# -------------------------------------
def png_row_availabities(top, bot, cols, col_cnt, date_str, img):
  daily_avail = ''
  daily_hours = 0
  extra_hours = 0
  hours_block = 0
  in_availablity_block = False
  for col_idx, col in enumerate(cols):
    filled = png_row_check_cell_filling(col + 1, top, bot, img)
    if filled is None:
      break
    elif filled:
      hours_block += .5
      if col_idx == 0:
        extra_hours += EXTRA_HOURS_BEGIN[col_cnt]
      if not in_availablity_block:
        daily_avail += f'{date_str} | {TIMES[col_cnt][col_idx]}'
        in_availablity_block = True
      elif col_idx == col_cnt - 1:
        extra_hours += EXTRA_HOURS_END[col_cnt]
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
def png_row_check_cell_filling(x_test, top, bot, img):
  for y_test in (top, bot):
    pixel_color = img[y_test, x_test]
    if pixel_color in COLOR['filled']:
      return True
    elif pixel_color in (COLOR['shift'] | COLOR['NP']):
      return None
  return False
# -------------------------------------

# -------------------------------------
def print_log(text='', pre='', end=''):
  print(pre + text + end)
  return pre + text + end + NL
# -------------------------------------

# -------------------------------------
def print_log_city_runtime(city, start, log):
  return log + print_log_header(parse_city_runtime(city, start), suf='=')
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
def print_no_name_determined(avail_str, name, ocr_read, row_n, png):
  print('\r', end='')
  return print_log(
    f'##### {NF}{png = }, {row_n = }{" " * 30}{NL}'
    + (f'|OCR| {ocr_read = }{NL}' if ocr_read else '')
    + f'|AVA| {"|".join(avail_str.split(NL)).strip("|")}'
    + (f', {name = }' if name else '')
    + BR
  )
# -------------------------------------

# -------------------------------------
def print_progress_bar(total, progress, sub_total, sub_progress, suf):
  prog = (progress + sub_progress / sub_total) / total
  if prog < 1:
    pre = ' ==> '
    print_end = '\r'
    suf += f' ({sub_progress}/{sub_total})'
  else:
    prog = 1
    pre = '|FIN|'
    print_end = '\r\n-----\n'
    suf = f'... DONE'
  bar_str = parse_progress_bar(30, prog, pre, suf)
  padding = get_terminal_size().columns - len(bar_str) - 13
  if padding < 0:
    bar_str = '\r' + parse_progress_bar(30 + padding, prog, pre, suf)
  print(bar_str, ' ' * 12, sep='', end=print_end, flush=True)
# -------------------------------------

# -------------------------------------
def process_screenshots(ref_data, kw_dates, dirs, kw, city):
  log = print_log_header(PROCESS_PNG_MSG)
  data = deepcopy(PNG_PROCESSING_DICT)
  pngs = sorted(listdir(dirs[3]))
  png_cnt = len(pngs)
  if png_cnt == 0:
    return data, log + print_log(NO_SCREENS_MSG, 'XXXXX ', BR)
  day = pngs[0].split('_')[0]
  for png_n, png in enumerate(pngs):
    png_split = png.split('_')
    if len(png_split) != 2:
      continue
    day, file_suf = png_split
    file_idx = file_suf.split('.')[0]
    date_str = parse_date(day, kw_dates)
    args = (ref_data, date_str, kw, city, day, file_idx)
    data, s_log = png_read_screenshot(png_cnt, png_n, dirs[3], png, data, args)
    log += s_log
  log += print_log(parse_stats_msg(data[CNT]), end=BR)
  DataFrame(data[DNA], columns=DF_DET_COLUMNS).to_excel(
    join(dirs[1], f'det_names_{city}_{START_DT}.xlsx')
    , sheet_name=city
    , columns=DF_DET_COLUMNS
    , index=False
    , freeze_panes=(1, 0)
  )
  return data, log
# -------------------------------------

# -------------------------------------
def process_xlsx_data(dirs, city, kw_date):
  log = print_log_header(PROCESS_XLSX_MSG)
  # ----- read weekly xlsx data, check availability of mendatory raw files ----
  dfs, log = load_xlsx_data_into_dfs(dirs, city, log)
  if dfs is None:
    return None, None, log
  dfs, ee_log = rider_ee_update_names(kw_date, city, dfs)
  log += ee_log
  ref_data = reference_names_and_contract_data(dfs[EE], kw_date)
  dfs[REP] = create_report_df(dfs, ref_data, log)
  # ----------
  return dfs, ref_data, log
# -------------------------------------

# -------------------------------------
def reference_contract_list(dt, df_kw):
  ref_contracts = []
  for _, d in df_kw.iterrows():
    if d[CC_FE] <= dt:
      ref_contracts.append(d[CON_TYP])
    else:
      prev_data = d[PRE_C]
      for contract_data in prev_data.split(NL):
        fe_le, contract = contract_data.split(' | ')
        first_entry, last_entry = map(date.fromisoformat, fe_le.split(' - '))
        if first_entry <= dt and last_entry >= dt:
          ref_contracts.append(contract)
          break
  return ref_contracts
# -------------------------------------

# -------------------------------------
def reference_names_and_contract_data(df, kw_date):
  df_rel = df[(df[LA_ENT] > (kw_date - TOO_OLD)) & (df[FI_ENT] <= kw_date)]
  ref_names = df_rel[RID_NAM].to_list()
  return (
    ref_names
    , reference_contract_list(kw_date, df_rel)
    , [*map(lambda x: x.split(';') if x else '', df_rel[SIM_NAM].to_list())]
    , {*ref_names}
  )
# -------------------------------------

# -------------------------------------
def rider_ee_get_similar_names(new_names, df_ee):
  all_names = df_ee[RID_NAM]
  simils = dict(zip(all_names, df_ee[SIM_NAM]))
  for new in new_names:
    for name in all_names:
      if fuzz.ratio(new, name) <= 78 or name == new or name in simils[new]:
        continue
      simils[new] += ('' if simils[new] == '' else ';') + name
      print(f'-----{TAB} found similar name | {new=}, {name}')
      if new not in simils[name]:
        simils[name] += ('' if simils[name] == '' else ';') + new
  return [*simils.values()]
# -------------------------------------

# -------------------------------------
def rider_ee_insert_new_names(new_data, new_names, df_ee):
  df_ee = df_ee.append(new_data).set_index(RID_NAM, drop=False).sort_index()
  df_ee[SIM_NAM] = rider_ee_get_similar_names(new_names, df_ee)
  return df_ee
# -------------------------------------

# -------------------------------------
def rider_ee_new_entry_df(name, avails, data, ref):
  eee = new_report_data_entry()
  contract = ref if isinstance(ref, str) else ref[1][ref[0].index(name)]
  eee.update({RID_NAM: name, AVA: 0, CON_TYP: contract})
  eee[MIN], eee[MAX] = CON_BY_N[contract]
  avail = parse_availabilities_string(avails, data[HRS][name], data[XTR][name])
  eee[AVAILS] = avail + 'NOT IN AVAILS OR MONTH'
  for key in (GIV, GIV_AVA, GIV_MAX, PAI_MAX, WOR, VAC, SIC, PAI, UNP):
    eee[key] = NO_DA
  return eee
# -------------------------------------

# -------------------------------------
def rider_ee_new_entry_xlsx(city, new_name, contract, kw_monday_date):
  return {
    RID_NAM: new_name
    , CON_TYP: contract
    , MIN: CON_BY_N[contract][0]
    , CIT: city
    , FI_ENT: kw_monday_date
    , LA_ENT: kw_monday_date
    , CC_FE: kw_monday_date
    , PRE_C: ''
    , SIM_NAM: ''
  }
# -------------------------------------

# -------------------------------------
def rider_ee_pre_c_update(df_row):
  prev_contracts = df_row[PRE_C]
  return (
    f'{(prev_contracts + NL) if prev_contracts else ""}'
    f'{df_row[CC_FE]} - {df_row[LA_ENT]} | {df_row[CON_TYP]}'
  )
# -------------------------------------

# -------------------------------------
def rider_ee_pre_c_refresh(df_row, contract, kw_date):
  prev_contracts_string = df_row[PRE_C]
  if prev_contracts_string == '':
    return f'{kw_date} - {kw_date} | {contract}'
  contract_in_prev_c = contract in prev_contracts_string
  updated_prev_c = ''
  for contract_line in prev_contracts_string.split(NL):
    fe_le, prev_contract = contract_line.split(' | ')
    first_entry, last_entry = map(date.fromisoformat, fe_le.split(' - '))
    pre = NL if updated_prev_c else ''
    if prev_contract != contract:
      if contract_in_prev_c:
        new_line = contract_line
      elif kw_date > last_entry:
        new_line = f'{contract_line}{NL}{kw_date} - {kw_date} | {contract}'
        contract_in_prev_c = True
      elif kw_date < first_entry:
        new_line = f'{kw_date} - {kw_date} | {contract}{NL}{contract_line}'
        contract_in_prev_c = True
      else:
        new_line = contract_line
    else:
      if kw_date < first_entry:
        first_entry = kw_date
      elif kw_date > last_entry:
        last_entry = kw_date
      new_line = f'{first_entry} - {last_entry} | {contract}'
    updated_prev_c += pre + new_line
  return updated_prev_c
# -------------------------------------

# -------------------------------------
def rider_ee_to_formated_xlsx(city, ree_dir, df_ee):
  row_cnt = df_ee.shape[0] + 1
  writer = ExcelWriter(parse_city_ee_filepath(city, ree_dir), 'xlsxwriter')
  df_ee.to_excel(writer, city, index=False)
  workbook = writer.book
  worksheet = writer.sheets[city]
  worksheet.autofilter('A1:I1')
  worksheet.freeze_panes(1, 0)
  fmt = {key: workbook.add_format(FMT_DICT[key]) for key in RIDER_EE_FMTS}
  for column, width, fmt_key in RIDER_EE_COL_FMT:
    worksheet.set_column(column, width, fmt[fmt_key])
  for columns, cond_key, fmt_key in RIDER_EE_CONDS:
    fmts = {**COND_FMT[cond_key], 'format': fmt[fmt_key]}
    worksheet.conditional_format(f'{columns}{row_cnt}', fmts)
  writer.save()
# -------------------------------------

# -------------------------------------
def rider_ee_update_known_names(df_ee, name, contract, kw_date):
  df_row = df_ee.loc[name]
  if kw_date > df_row[LA_ENT]:
    if contract != df_row[CON_TYP]:
      df_ee.at[name, PRE_C] = rider_ee_pre_c_update(df_row)
      df_ee.at[name, CON_TYP] = contract
      df_ee.at[name, MIN] = CON_BY_N[contract][0]
      df_ee.at[name, CC_FE] = kw_date
    df_ee.at[name, LA_ENT] = kw_date
  elif kw_date < df_row[LA_ENT]:
    if kw_date < df_row[FI_ENT]:
      df_ee.at[name, FI_ENT] = kw_date
    if contract != df_row[CON_TYP]:
      df_ee.at[name, PRE_C] = rider_ee_pre_c_refresh(df_row, contract, kw_date)
    elif kw_date < df_row[CC_FE]:
      df_ee.at[name, CC_FE] = kw_date
  return df_ee
# -------------------------------------

# -------------------------------------
def rider_ee_update_names(kw_date, city, dfs):
  log = print_log_header(SYNCH_MIN_H_MSG)
  log += print_log(f'CHECK {MH_AVAIL_MSG}{dfs[MON] is not None} {BR}')
  new_data = []
  new_names = set()
  processed_names = set()
  known = {*dfs[EE].index.values}
  names_mon = set() if dfs[MON] is None else {*dfs[MON][DRI]}
  new_in_mon = names_mon - known
  names_av = {*dfs[AVA][USER_N]}
  new_in_av = names_av - (known | new_in_mon)
  params = [(MON, CO_TY, new_in_mon)] if new_in_mon else []
  params.append((AVA, USER_T, new_in_av))
  for src_key, contract_key, new_names_in_df in params:
    for name, contract in dfs[src_key][contract_key].items():
      if name in processed_names:
        continue
      if name in new_names_in_df:
        new_data.append(rider_ee_new_entry_xlsx(city, name, contract, kw_date))
        new_names.add(name)
        log += print_log(f'{TAB}- {name = }, {contract = }')
      elif name in known:
        dfs[EE] = rider_ee_update_known_names(dfs[EE], name, contract, kw_date)
      processed_names.add(name)
    if new_names_in_df:
      log += print_log('-----')
  if new_data:
    dfs[EE] = rider_ee_insert_new_names(new_data, new_names, dfs[EE])
  return dfs, log
# -------------------------------------

# -------------------------------------
def screenshots_list_of_daily_files(day, png_dir):
  return [
    Image.open(join(png_dir, day_fn))
    for day_fn in sorted(fn for fn in listdir(png_dir) if day in fn)
  ]
# -------------------------------------

# -------------------------------------
def screenshots_merge_daily_files(city, dirs):
  log = print_log_header(MERGE_FILES_MSG)
  for day in WEEKDAYS:
    images = screenshots_list_of_daily_files(day, dirs[3])
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
def shiftplan_check(year, city, kw, dirs, get_ava, merge, tidy_only, ee_only):
  start = perf_counter()
  log = print_log_header(f'{CITY_LOG_PRE} {city} | KW {kw} / {year}', pre='=')
  tidy_log = tidy_screenshot_files(city, dirs, merge)
  log += tidy_log
  if tidy_only:
    return print_log_city_runtime(city, start, log)
  kw_date = date.fromisocalendar(year, kw, 1)
  dfs, ref_data, xlsx_log = process_xlsx_data(dirs, city, kw_date)
  log += xlsx_log
  if dfs is None:
    return print_log_city_runtime(city, start, log)
  if get_ava and TESSERACT_AVAILABLE:
    kw_dates = [str(date.fromisocalendar(year, kw, i)) for i in range(1, 8)]
    data, screen_log = process_screenshots(ref_data, kw_dates, dirs, kw, city)
    log += screen_log
    dfs, sp_log = sp_report_png_data_update(data, dfs, kw_date, ref_data, city)
    log += sp_log
  rider_ee_to_formated_xlsx(city, dirs[4], dfs[EE])
  if ee_only:
    return print_log_city_runtime(city, start, log)
  dfs[REP] = sp_report_remove_irrelevant(dfs[REP])
  log += sp_report_to_formated_xlsx(dfs[REP], city, kw)
  return print_log_city_runtime(city, start, log)
# -------------------------------------

# -------------------------------------
def sp_report_png_data_update(data, dfs, kw_date, ref_data, city):
  log = ''
  only_ee = []
  only_screen = []
  os_names = []
  for name, avails in data[AVA].items():
    try:
      dfs[REP].at[name, AVAILS] = parse_availabilities_string(
        avails, data[HRS][name], data[XTR][name], dfs[REP].at[name, AVA]
      )
      if dfs[REP].at[name, AVA] == '':
        dfs[REP].at[name, AVA] = data[HRS][name]
    except KeyError:
      if name in ref_data[3]:
        only_ee.append(rider_ee_new_entry_df(name, avails, data, ref_data))
      else:
        only_ee.append(rider_ee_new_entry_df(name, avails, data, NO_DA))
        only_screen.append(rider_ee_new_entry_xlsx(city, name, NO_DA, kw_date))
        os_names.append(name)
  if only_ee:
    dfs[REP] = dfs[REP].append(only_ee).sort_values([RID_NAM])
  dfs[REP].loc[dfs[REP][AVA] == dfs[REP][AVAILS], AVA] = 0
  dfs[REP].loc[(dfs[REP][AVA] != 0) & (dfs[REP][AVAILS] == ''), AVAILS] = ' '
  if only_screen:
    dfs[EE] = rider_ee_insert_new_names(only_screen, os_names, dfs[EE])
  return dfs, log + (BR if log else '')
# -------------------------------------

# -------------------------------------
def sp_report_remove_irrelevant(df):
  return df[(df[AVA] != 0) | (df[AVAILS] != '') | (df[GIV_SHI] != '')]
# -------------------------------------

# -------------------------------------
def sp_report_to_formated_xlsx(df, city, kw):
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
  worksheet.autofilter('A1:T1')
  worksheet.freeze_panes(1, 2)
  worksheet.set_row(row_cnt, None, fmt_dict['border'])
  for column, width, fmt in XLS_REPORT_COL_FMT:
    worksheet.set_column(column, width, fmt_dict[fmt])
  # ----- add conditional formats -----
  for cols, con, fmt in XLS_REPORT_COND_FMT:
    fmts = {**COND_FMT[con], 'format': fmt_dict[fmt]} if fmt else COND_FMT[con]
    worksheet.conditional_format(f'{cols}{row_cnt}', fmts)
  # ----------
  writer.save()
  return log + print_log(f'+++++ saved {filename}{BR}')
# -------------------------------------

# -------------------------------------
def tidy_jpg_files(city, dirs):
  jpg_files = [fn for fn in listdir(dirs[0]) if fn.endswith('.jpg')]
  if not jpg_files:
    return ''
  log = print_log_header(TIDY_JPG_MSG)
  log += print_log(JPG_NAME_CHECK_MSG, '|JPG| ')
  idx_dict = defaultdict(int)
  raw_dir = check_make_dir(dirs[2], 'raw')
  vac_dir = check_make_dir(raw_dir, 'Urlaub')
  for fn in sorted(jpg_files):
    fn_cf = fn.casefold().replace('_', ' ')
    if not any(fuzz.WRatio(alias, fn_cf) > 86 for alias in ALIAS[city]):
      continue
    source = join(dirs[0], fn)
    if 'urlaub' in fn.casefold():
      target = join(vac_dir, fn)
    else:
      target = join(raw_dir, fn)
      proc_fn, idx_dict, log = tidy_screenshot_fn(fn, fn_cf, idx_dict, log)
      Image.open(source).save(join(dirs[3], proc_fn))
    shutil.move(source, target)
  return log
# -------------------------------------

# -------------------------------------
def tidy_png_files(city, dirs):
  png_files = [fn for fn in listdir(dirs[0]) if fn.endswith('.png')]
  if not png_files:
    return ''
  log = print_log_header(TIDY_PNG_MSG)
  log += print_log(PNG_NAME_CHECK_MSG, '|PNG| ')
  idx_dict = defaultdict(int)
  raw_dir = check_make_dir(dirs[2], 'raw')
  for fn in sorted(png_files):
    source = join(dirs[0], fn)
    if city == 'Frankfurt':
      fn = fn.replace('FF_', 'FFM_').replace('FF ', 'FFM ')
    fn_cf = fn.casefold().replace('_', ' ')
    if all(fuzz.partial_ratio(alias, fn_cf) < 87 for alias in ALIAS[city]):
      continue
    proc_fn, idx_dict, log = tidy_screenshot_fn(fn, fn_cf, idx_dict, log)
    shutil.copy(source, join(raw_dir, fn))
    shutil.move(source, join(dirs[3], proc_fn))
  return log
# -------------------------------------

# -------------------------------------
def tidy_screenshot_files(city, dirs, merge):
  log = tidy_jpg_files(city, dirs)
  if not log:
    log += tidy_png_files(city, dirs)
  if not log:
    log += tidy_zip_files(city, dirs)
  if log:
    log += print_log(f'+++++ saved PNGs in: {dirs[3]}{BR}')
    if merge:
      log += screenshots_merge_daily_files(city, dirs)
  return log
# -------------------------------------

# -------------------------------------
def tidy_screenshot_fn(original, fn_cf, idx_dict, log):
  similarity = 0
  current_day = ''
  for weekday in WEEKDAYS:
    weekday_similarity = fuzz.partial_ratio(weekday, original)
    if weekday_similarity > similarity:
      similarity = weekday_similarity
      current_day = weekday
      if similarity > 90:
        break
  if similarity <= 90:
    for weekday_n, abrevations in enumerate(WEEKDAY_ABREVATIONS):
      if any(fuzz.token_set_ratio(abre, fn_cf) == 100 for abre in abrevations):
        current_day = WEEKDAYS[weekday_n]
        similarity = 100
        break
  idx_dict[current_day] += 1
  screenshot_filename = f'{current_day}_{idx_dict[current_day]}.png'
  if similarity != 100:
    log += print_log(f'{TAB}- {original = }, saved as = {screenshot_filename}')
  return screenshot_filename, idx_dict, log
# -------------------------------------

# -------------------------------------
def tidy_zip_files(city, dirs):
  zip_files = [fn for fn in listdir(dirs[0]) if fn.endswith('.zip')]
  if not zip_files:
    return ''
  log = print_log_header(UNZIP_MSG)
  log += print_log(ZIP_PNG_NAME_CHECK_MSG, '[X|O] ')
  idx_dict = defaultdict(int)
  for zip_file in zip_files:
    zip_fn_cf = zip_file.casefold()
    if all(fuzz.partial_ratio(alias, zip_fn_cf) < 87 for alias in ALIAS[city]):
      continue
    log += print_log(zip_file, TAB)
    with ZipFile(join(dirs[0], zip_file)) as zfile:
      for member in sorted(zfile.namelist()):
        fn = basename(member)
        if fn:
          fn_cf = fn.casefold().replace('_', ' ')
          fn, idx_dict, log = tidy_screenshot_fn(fn, fn_cf, idx_dict, log)
          with open(join(dirs[3], fn), "wb") as target:
            shutil.copyfileobj(zfile.open(member), target)
    log += print_log('-----')
  return log
# -------------------------------------
# =================================================================


# =================================================================
# ### MAIN FUNCTION ###
# =================================================================
# -------------------------------------
def main(year, start_kw, last_kw, cities, *args):
  start = perf_counter()
  log = print_log_header(INITIAL_MSG, pre='=', suf='=')
  if last_kw < start_kw:
    last_kw = start_kw
  for kw in range(start_kw, last_kw + 1):
    kw_dir = join(BASE_DIR, 'Schichtplan_Daten', str(year), f'KW{kw}')
    if not exists(kw_dir):
      print_log(f'##### Couldn`t find "{kw_dir}"{BR}')
      log = print_log()
      continue
    log_dir = check_make_dir(kw_dir, 'logs')
    screen_dir = join(kw_dir, 'Screenshots')
    ree_dir = check_make_dir(BASE_DIR, EE)
    for city in cities:
      try:
        png_dir = check_make_dir(screen_dir, city)
        dirs = (kw_dir, log_dir, screen_dir, png_dir, ree_dir)
        log += shiftplan_check(year, city, kw, dirs, *args)
      except Exception as ex:
        log += f'{type(ex)=} | {repr(ex)=}{NL}{parse_break_line("#")}'
        raise ex
      finally:
        with open(join(log_dir, LOG_FN), 'w', encoding='utf-8') as logfile:
          logfile.write(log)
        log = print_log()
  print_log_header(parse_run_end_msg(start), pre='=', suf='=', brk=NL)
# -------------------------------------
# =================================================================

PARSER_KW = 'Kalenderwoche der zu prüfenden Daten, default: 1'
PARSER_Y = 'Jahr der zu prüfenden Daten, default: heutiges Jahr'
PARSER_LKW = 'Letzte zu bearbeitende Kalenderwoche als Zahl'
PARSER_C = (
  'Zu prüfende Stadt oder Städte, Stadtnamen trennen mit einem '
  'Leerzeichen, default: [Frankfurt Offenbach]'
)
PARSER_A = 'Aktiviert das Auslesen mitgeschickter Screenshots'
PARSER_TO = 'Räumt alle Verfügbarkeiten Screenshot Dateien auf'
PARSER_M = (
  'Erstellt je Stadt und Tag eine zusammegesetzte '
  'Verfügbarkeiten-Screenshot-Datei'
)
PARSER_EEO = 'Erstellt nur die Rider_Ersterkennung Datei, ohne SP-Report'
# =================================================================
# ### START SCRIPT ###
# =================================================================
# -------------------------------------
if __name__ == '__main__':
  from argparse import ArgumentParser
  parser = ArgumentParser()
  parser.add_argument(
    '-y', '--year', type=int, default=YEAR, help=PARSER_Y
  )
  parser.add_argument(
    '-kw', '--kalenderwoche', type=int, default=1, help=PARSER_KW
  )
  parser.add_argument(
    '-lkw','--last_kw',  type=int, default=0, help=PARSER_LKW
  )
  parser.add_argument(
    '-c', '--cities', nargs='*', default=DEFAULT_CITIES, help=PARSER_C
  )
  parser.add_argument(
    '-a', '--get_avail', action='store_true', help=PARSER_A
  )
  parser.add_argument(
    '-to', '--tidy_only', action='store_true', help=PARSER_TO
  )
  parser.add_argument(
    '-m', '--mergeperday', action='store_true', help=PARSER_M
  )
  parser.add_argument(
    '-eeo','--ersterkennung_only',  action='store_true', help=PARSER_EEO
  )
  main(*parser.parse_args().__dict__.values())
# -------------------------------------
# =================================================================
