# =================================================================
# ### IMPORTS ###
# =================================================================
# -------------------------------------
from collections import defaultdict
import copy
from datetime import date, datetime, timedelta
import getpass
import io
import os
from os.path import abspath, basename, dirname, exists, join
import shutil
import signal
import sys
import time
# -------------------------------------
import cv2 as cv
from fuzzywuzzy import fuzz
import matplotlib.pyplot as plt
import msoffcrypto
import numpy as np
import pandas as pd
from pytesseract import pytesseract
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
KW = _now.isocalendar()[1] + 1
START_DT = _now.strftime(Y_S)
TOO_OLD = timedelta(weeks=8)
YEAR =  _now.year
# -------------------------------------

# -------------------------------------
# ### OPENCV ###
# -------------------------------------
CLAHE_DEF = cv.createCLAHE()
CLAHE_L3_S7 = cv.createCLAHE(clipLimit=3.0, tileGridSize=(7, 7))
KERNEL = np.array(([-1, -1, -1], [-1, 9, -1], [-1, -1, -1]), dtype="int")
# -------------------------------------

# -------------------------------------
# ### STRINGS ###
# -------------------------------------
ALI = 'align'
AVA = 'avail'
AVAILS = 'availablities'
AV_ARGS = 'row_availability_args'
BAD_RESO = 'bad_resolution'
BAR = 'progress_bar_data'
BG_COL = 'bg_color'
BOLD = 'bold'
BOR = 'border'
BORDER = 'vertical_boundries'
BR = '\n-----'
CAL = 'call'
CC_FE = 'current contract first entry'
CEN = 'center'
CHE = 'check'
CITY_LOG_PRE = ' LOG FOR CITY:'
CIT = 'city'
CMT = 'comment'
CONFIG_MISSING_MSG = '##### missing config file, using default values ...\n'
CON_H = 'Contracted hours'
CON_TYP = 'contract type'
COUNTER = 'counter'
CO_SC = '3_color_scale'
CO_TY = 'Contract Type'
CREATE_XLSX_MSG = ' CREATE WEEKLY XLSX REPORT '
CRI = 'criteria'
DATE = 'date'
DEL = 'delete_color'
DONE = 'processed_day_and_rider'
DRI = 'Driver'
DR_ID = 'Driver ID'
DUPL = 'duplicates'
EE = 'Rider_Ersterfassung'
EMPTY = 'empty_availability_cell'
FILLED = 'filled_availability_cell'
FI_ENT = 'first entry'
FMT = 'format'
FO_SI = 'font_size'
FRAME = 'histgram_test_frame'
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
LA_ENT = 'last entry'
LEF = 'left'
LINE = 'thin_line'
LINK = 'linked'
LOG = 'log'
LOG_DATA = 'determined_names_log_data'
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
NA = 'N/A'
NAME_ARGS = 'cv_data_args'
NAME_BOX = 'name_box_in_av_cells'
NF = 'NOT FOUND '
NL = '\n'
NOAV = 'no_avail'
NOOCR = 'no_ocr'
NODA = 'no_data'
NOT_AV_MON = 'NOT IN AVAILS OR MONTH'
NOT_IN_AV = 'NOT IN AVAILS'
NOT_IN_MON = ' not in "Monatsstunden": '
NO_AVAILS_MSG = 'keine Verfügbarkeiten'
NO_DA = 'NO DATA'
NO_SCREENS_MSG = 'NO SCREENSHOTS AVAILABLE'
NP = 'not_planable'
NU = 'num'
NU_FO = 'num_format'
PAI = 'paid'
PAI_MAX = 'paid/max'
PLOT_SHIFTS_MSG = 'PLOT DISTRIBUTED SHIFTS'
PNG = 'png'
PNGS = 'available_png_list'
PNG_CNT = 'png_count'
PNG_N = 'png_number'
PNG_NAME_CHECK_MSG = 'Check available PNG file names'
PRE_C = 'prev contracts'
PRE_F = 'prev first entries'
PRE_L = 'prev last entries'
PROCESS_PNG_MSG = ' SCAN AVAILABILITES FROM PNGS '
PROCESS_XLSX_MSG = ' PROCESS RAW XLSX DATA '
PW = None
P_A = 'Aktiviert das Auslesen mitgeschickter Screenshots'
P_C = (
  'Zu prüfende Stadt oder Städte, Stadtnamen trennen mit einem '
  'Leerzeichen, default: [Frankfurt Offenbach]'
)
P_EEO = 'Erstellt nur die Rider_Ersterfassung Datei, ohne SP-Report'
P_KW = 'Kalenderwoche der zu prüfenden Daten, default: 1'
P_LKW = 'Letzte zu bearbeitende Kalenderwoche als Zahl'
P_M = (
  'Erstellt je Stadt und Tag eine zusammegesetzte '
  'Verfügbarkeiten-Screenshot-Datei'
)
P_TO = 'Räumt alle Verfügbarkeiten Screenshot Dateien auf'
P_V = 'Daten Visualisierung, erstellt Plots der vergebenen Schichten'
P_X = 'Letztes Jahr der zu prüfenden Daten, default: heutiges Jahr'
P_Y = 'Jahr der zu prüfenden Daten, default: heutiges Jahr'
REDUCE_HOURS = ' -> auf Min.Std. reduzieren'
REF = 'reference_data'
REP = 'report'
RE_MA = 'resize_margin'
RE_ROW = 'resized rows'
RID_NAM = 'rider name'
RI = 'right'
ROI = 'region_of_interest'
ROWS = 'row_y_values'
ROW_CNT = 'row_count'
ROW_N = 'row_number'
SCAN = 'scanned'
SCROLL_BAR = 'scroll_barf'
SHI = 'shift'
SIC = 'sick'
SIM_NAM = 'similar names'
SH_DA = 'Shift Date'
SH_DAY = 'Shift Day'
STD_REP = 'stundenreports'
STORE_TRUE = 'store_true'
STRIP_CHARS = """ .,-_'`"()|"""
SYNCH_MIN_H_MSG = ' SYNCHRONIZE NAMES IN MINDESTSTUNDEN LIST '
TAB = '\t'
TIDY_JPG_MSG = ' TIDY JPGS IN WORKING DIRECTORY '
TIDY_PNG_MSG = ' TIDY PNGS IN WORKING DIRECTORY '
TIME_CELL = 'header_time_cell'
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
WHITE = 'white_color_range'
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
  os.makedirs(OUTPUT_DIR)
OUT_FILE_PRE = ''
REE_DIR = join(BASE_DIR, EE)
if not exists(REE_DIR):
  os.makedirs(REE_DIR)
# -------------------------------------

# -------------------------------------
# ### SETS AND TUPLES ###
# -------------------------------------
CONVERT_COLS_MONTH = (
  (PAI_MAX, None)
  , (WOR, 'Worked hours')
  , (VAC, 'Paid leaves (hours)')
  , (SIC, 'Sick leaves (hours)')
  , (PAI, 'Total paid hours')
  , (UNP, 'Unpaid leaves (hours)')
)
DF_DET_COLS = ('kw', CIT, 'day', 'index', 'row', 'avail', 'name', 'ocr')
DIGITS = {*map(str, range(10))}
INVALID_WORDS = {'wochenstunden', 'gefahrene'}
REPORT_HEADER = (
  ID, RID_NAM, CON_TYP, MAX, MIN, AVA, GIV, GIV_AVA, GIV_MAX, GIV_SHI, AVAILS
  , PAI_MAX, WOR, VAC, SIC, PAI, UNP, CMT, CAL, 'cmt shift coordinator'
)
RIDER_MIN_HEADER = (
  RID_NAM, CON_TYP, MIN, CIT, FI_ENT, LA_ENT, CC_FE, PRE_C, SIM_NAM
)
SPLIT_CHARS = ('NP', 'Np', 'DA', 'wp', 'pA', ' pa')
TIMEBLOCK_STRINGS = (
  '11:00', '11:30', '12:00', '12:30', '13:00', '13:30', '14:00', '14:30'
  , '15:00', '15:30', '16:00', '16:30', '17:00', '17:30', '18:00', '18:30'
  , '19:00', '19:30', '20:00', '20:30', '21:00', '21:30', '22:00', '22:30'
  , '23:00', '23:30'
)
WEEKDAYS = (
  'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'
  , 'Sonntag'
)
WEEKDAYS_EN = (
  'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'
)
WEEKDAY_ABREVATIONS = (
  ('mo', 'mon'), ('di', 'tue'), ('mi', 'wed'), ('do', 'thu'), ('fr', 'fri')
  , ('sa', 'sat'), ('so', 'sun')
)
# -------------------------------------

# -------------------------------------
# ### DICTS ###
# -------------------------------------
COLOR = {
  EMPTY: 225
  , FILLED: {*range(93, 116), 126}
  , NAME_BOX: 64
  , NP: {228, 238}
  , SCROLL_BAR: {*range(236, 246)}
  , SHI: {*range(143, 151)}
  , LINE: {195, *range(208, 236)}
  , TIME_CELL: {245}
  , WHITE: {*range(246, 256)}
}
COLOR[DEL] = COLOR[WHITE] | COLOR[NP]
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
_mini_min_h = int(60 / 13) / 2
CON_BY_N = defaultdict(
  lambda: (NO_DA, NO_DA), {
    'Arbeitnehmerüberlassung': (0, 20)
    , 'Foodora_Minijob': (_mini_min_h, 15)
    , 'Minijob': (_mini_min_h, 15)
    , 'Minijobber': (5, 11)
    , 'Mini-Jobber': (5, 11)
    , 'TE Minijob': (5, 11)
    , 'Foodora_Working Student': (12, 20)
    , 'Werk Student': (12, 20)
    , 'Working Student': (12, 20)
    , 'Working-Student': (12, 20)
    , 'TE Werkstudent': (12, 20)
    , 'TE WS': (12, 20)
    , 'TE Midijob': (12, 28)
    , 'Foodora_Midijob': (12, 40)
    , 'Midijob': (12, 40)
    , 'TE Teilzeit': (30, 48)
    , 'Teilzeit': (30, 48)
    , 'Vollzeit': (30, 48)
  }
)
EXTRA_HOURS_BEGIN = defaultdict(int, {21: 1, 22: .5})
EXTRA_HOURS_END = {21: 1, 22: 1, 23: 1, 24: .5, 25: 0, 26: 0}
LGD_ARGS = {
  'bbox_to_anchor': (.5, 1.05)
  , 'loc': 'upper center'
  , 'fontsize': 'small'
  , 'facecolor': '#e5e5e5'
}
LGD_COLORS = dict(zip(
  CON_BY_N.keys()
  , plt.get_cmap('RdYlGn')(np.linspace(1, 0, len(CON_BY_N)))
))
PNG_PROCESSING_DICT = {
  AVA: defaultdict(list)
  , DONE: defaultdict(set)
  , HRS: defaultdict(int)
  , XTR: defaultdict(int)
  , COUNTER: defaultdict(int)
  , LOG_DATA: list()
}
TIMES = {
  21: TIMEBLOCK_STRINGS[2:-2]
  , 22: TIMEBLOCK_STRINGS[1:-2]
  , 23: TIMEBLOCK_STRINGS[:-2]
  , 24: TIMEBLOCK_STRINGS[:-1]
  , 25: TIMEBLOCK_STRINGS
  , 26: TIMEBLOCK_STRINGS
}
TITLE_FMT = {'fontsize': 20, 'fontweight': 'bold'}
TOTAL_FMT = {'fontsize': 'large', 'fontweight': 'bold', 'color': 'black'}
# -------------------------------------

# -------------------------------------
# ### XLSX FORMATS ###
# -------------------------------------
COND_FMT = {
  AVA: {TYP: 'text', CRI: 'ends with', VAL: ' '}
  , 'ee only': {TYP: 'text', CRI: 'ends with', VAL: NOT_AV_MON}
  , 'not avail': {TYP: 'text', CRI: 'ends with', VAL: NOT_IN_AV}
  , DATE: {TYP: 'formula', CRI: '$F2 + 60 < today()'}
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
)
RIDER_EE_CONDS = (('C2:C', MIN, 'red'), ('F2:F', DATE, 'old'))
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
if exists(CONFIG_FP):
  import json
  with open(CONFIG_FP, encoding='utf-8') as file_path:
    config = json.load(file_path)
  ALIAS = {k: tuple(x for x in alis) for k, alis in config['aliases'].items()}
  DEF_CITY = config['cities']
  if 'win' in sys.platform:
    pytesseract.tesseract_cmd = config['cmd_path']
  if 'password' in config:
    PW = config['password'] or None
else:
  ALIAS = {
    'Frankfurt': ('frankfurt', 'ffm', 'frankfurt am main')
    , 'Fürth': ['fürth', 'fuerth']
    , 'Nürnberg': ['nuernberg', 'nuremberg', 'nue']
    , 'Offenbach': ('offenbach', 'of', 'offenbach am main')
    , AVA: ('Verfügbarkeit', 'Verfügbarkeiten', 'Availabilities')
    , MON: ('Monatsstunden', 'Stunden', 'Working Hours')
    , SHI: ('Schichtplan', 'Schichtplanung', 'Working Shifts')
  }
  DEF_CITY = ('Frankfurt', 'Offenbach')
# -------------------------------------

# -------------------------------------
# ### CHECK TESSERACT AVAILABILITY ###
# -------------------------------------
try:
  pytesseract.get_tesseract_version()
except pytesseract.TesseractNotFoundError:
  TESSERACT_AVAILABLE = False
else:
  TESSERACT_AVAILABLE = True
# -------------------------------------

# -------------------------------------
# ### HANDLER FOR KEYBOARD INTERRUPT ###
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
def check_make_dir(*path_list):
  dir_path = join(*path_list)
  if not exists(dir_path):
    os.makedirs(dir_path)
  return dir_path
# -------------------------------------

# -------------------------------------
def get_base_data(name, df_row, source, df_ava):
  data = {}
  if source == AVA:
    data.update(get_base_data_from_ava(df_row[USER_T], name, df_row))
    data.update(get_max_hours(data[CON_TYP], df_row))
  else:
    data.update(get_base_data_from_mon(name, df_row))
    data.update(get_contract_and_avail_h(name, df_row, df_ava))
    data.update(get_max_hours(data[CON_TYP], df_ava, name))
  return data
# -------------------------------------

# -------------------------------------
def get_base_data_from_ava(contract, name, df_row):
  return {
    **{output_col: NA for output_col, _ in CONVERT_COLS_MONTH}
    , AVA: df_row[H_AV]
    , ID: df_row[U_ID]
    , RID_NAM: name
    , CON_TYP: contract
  }
# -------------------------------------

# -------------------------------------
def get_base_data_from_mon(name, df_row):
  base_data = {ID: df_row[DR_ID], RID_NAM: name}
  if df_row is None:
    for output_col, _ in CONVERT_COLS_MONTH:
      base_data[output_col] = NA
    return {**base_data, PAI_MAX: 0}
  for output_col, input_col in CONVERT_COLS_MONTH[1:]:
    try:
      base_data[output_col] = df_row[input_col]
    except ValueError:
      base_data[output_col] = 0
  try:
    work_ratio = float(str(df_row[WO_RA]).strip('%'))
    if work_ratio > 5: 
      work_ratio /= 100
  except ValueError:
    work_ratio = 0
  finally:
    work_ratio = round(work_ratio, 2)
  return {**base_data, PAI_MAX: work_ratio}
# -------------------------------------

# -------------------------------------
def get_contract_and_avail_h(name, df_row, df_ava):
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
  return {CON_TYP: contract, AVA: avail}
# -------------------------------------

# -------------------------------------
def get_data_check_and_first_comment(data):
  call = ''
  comment = []
  min_h = get_data_check_min_h(data)
  if data[GIV_MAX] > 1 and data[MAX] != 40:
    call = 'X'
    comment.append(MAX_HOURS_MSG)
  if data[GIV_AVA] == 10:
    comment.append(NO_AVAILS_MSG)
    if data[GIV] > min_h:
      call = 'X'
      comment[-1] += REDUCE_HOURS
  else:
    if data[GIV] < min_h:
      if data[GIV] != data[AVA]:
        call = 'X'
      comment.append(MIN_HOURS_MSG)
    if data[GIV_AVA] > 1:
      comment.append(MORE_THAN_AVAIL_MSG)
  ratio_threshold = get_data_check_ratio_threshold(data)
  if ratio_threshold == 0:
    call = ''
    comment.append(MINI_LIMIT_MSG)
  elif data[GIV_MAX] < ratio_threshold and data[GIV_AVA] < ratio_threshold:
    comment.append(MORE_HOURS_MSG)
  return {AVAILS: '', CAL: call, CMT: NL.join(comment)}
# -------------------------------------

# -------------------------------------
def get_data_check_min_h(data):
  if not isinstance(data[MIN], str):
    return data[MIN]
  if 'h/' not in data[MIN]:
    return data[MAX]
  hours, period = data[MIN].split('h/')
  if period == 'Monat':
    return max(3, (int(hours) * 6 // 13) / 2)
  return int(hours)
# -------------------------------------

# -------------------------------------
def get_data_check_ratio_threshold(data):
  if 'mini' not in data[CON_TYP].casefold():
    return .75
  elif isinstance(data[PAI_MAX], str) or data[PAI_MAX] <= .9:
    return .55
  return 0
# -------------------------------------

# -------------------------------------
def get_given_hour_ratios(avail, given, max_h):
  if isinstance(avail, str):
    avail_ratio = max_ratio = 0
  else:
    avail_ratio = round(given / avail, 2) if avail else 10
    max_ratio = round(given / max_h, 2) if not isinstance(max_h, str) else 0
  return {GIV_AVA: avail_ratio, GIV_MAX: max_ratio}
# -------------------------------------

# -------------------------------------
def get_max_hours(contract, df_row, name=None):
  try:
    if name is not None:
      df_row = df_row.loc[name]
    max_h = df_row[MAX_H]
  except KeyError:
    max_h = CON_BY_N[contract][1]
  return {MAX: max_h}
# -------------------------------------

# -------------------------------------
def get_min_hours(df_ee, ref_names_set, contract, name):
  if 'TE' in contract or name not in ref_names_set:
    min_h = CON_BY_N[contract][0]
  else:
    min_h = df_ee.at[name, MIN]
  return {MIN: min_h}
# -------------------------------------

# -------------------------------------
def get_new_df_entry():
  return {report_column: '' for report_column in REPORT_HEADER}
# -------------------------------------

# -------------------------------------
def get_shifts(df_shifts, rider_id):
  given = 0
  shifts = ''
  for _, d in df_shifts[df_shifts[DR_ID] == rider_id].iterrows():
    given += d[WO_HO]
    shifts += f'{d[SH_DA]} | {d[FR_HO]} - {d[TO_HO]} | {d[WO_HO]}h{NL}'
  return {GIV: given, GIV_SHI: shifts}
# -------------------------------------

# -------------------------------------
def invalid_city_xlsx_filename(filename, city):
  return (
    not filename.endswith('xlsx')
    or not filename[0].isalpha()
    or any(invalid_word in filename for invalid_word in INVALID_WORDS)
    or all(fuzz.partial_ratio(alias, filename) <= 86 for alias in ALIAS[city])
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
  global PW
  df = None
  print('file is encrypted... decrypting')
  tries = 5
  while tries != 0:
    try:
      if PW is None:
        PW = getpass.getpass('enter password:')
      decrypted = io.BytesIO()
      with open(file_path, 'rb') as f:
        file = msoffcrypto.OfficeFile(f)
        file.load_key(password=PW)
        file.decrypt(decrypted)
      df = pd.read_excel(decrypted)
      break
    except Exception:
      if tries:
        tries -= 1
        print(f'wrong password, {tries} tries left ...')
        PW = None
      else:
        print(f'could not open encrpyted xlsx file {file_path=}')
  return df
# -------------------------------------

# -------------------------------------
def load_ersterfassung_xlsx_into_df(city):
  try:
    df = pd.read_excel(parse_city_ee_filepath(city))
  except FileNotFoundError:
    try:
      df = pd.read_excel(EE_BACKUP, city)
      df[SIM_NAM] = ''
      df[SIM_NAM] = rider_ee_get_similar_names(df[RID_NAM], df)
    except (FileNotFoundError, XLRDError):
      df = pd.DataFrame(columns=RIDER_MIN_HEADER)
  df[FI_ENT] = pd.to_datetime(df[FI_ENT], format=YMD).dt.date
  df[LA_ENT] = pd.to_datetime(df[LA_ENT], format=YMD).dt.date
  df[CC_FE] = pd.to_datetime(df[CC_FE], format=YMD).dt.date
  return df.fillna('').drop_duplicates(RID_NAM).set_index(RID_NAM, drop=False)
# -------------------------------------

# -------------------------------------
def load_month_xlsx_into_df(df):
  if CO_TY in df.columns:
    df[CO_TY] = df[CO_TY].apply(lambda x: x.strip().replace('Foodora_', ''))
    df.loc[df[CO_TY] == 'TE Werkstudent', CO_TY] = 'TE WS'
  return df.drop_duplicates().set_index(DRI, drop=False)
# -------------------------------------

# -------------------------------------
def load_shift_xlsx_into_df(df):
  df[SH_DA] = pd.to_datetime(df[SH_DA], format=DMY).dt.date
  try:
    df[FR_HO] = pd.to_datetime(df[FR_HO], format=HM).dt.time
    df[TO_HO] = pd.to_datetime(df[TO_HO], format=HM).dt.time
  except (TypeError, ValueError):
    df[FR_HO] = pd.to_datetime(df[FR_HO], format=HMS).dt.time
    df[TO_HO] = pd.to_datetime(df[TO_HO], format=HMS).dt.time
  return df.sort_values([DRI, SH_DA, FR_HO]).set_index(SH_DAY, drop=False)
# -------------------------------------

# -------------------------------------
def load_xlsx_data_into_dfs(city, dirs, dfs):
  dfs[MON] = None
  mendatory = [ALIAS[AVA][0], ALIAS[SHI][0]]
  for filename in os.listdir(dirs[0]):
    fn_cf = filename.casefold().replace('_', ' ').replace('-', ' ')
    if invalid_city_xlsx_filename(fn_cf, city):
      continue
    if fuzz.WRatio(STD_REP, fn_cf) > 86:
      dfs[LOG] += print_log(f'|O.O| {STD_REP} file available, {filename = }')
      continue
    full_path = join(dirs[0], filename)
    try:
      df = pd.read_excel(full_path)
    except XLRDError:
      df = load_decrpyted_xlsx(full_path)
      if df is None:
        break
    df.rename(columns=lambda x: str(x).strip(), inplace=True)
    if any(fuzz.partial_ratio(alias, fn_cf) > 86 for alias in ALIAS[AVA]):
      dfs[AVA] = load_avail_xlsx_into_df(df)
      mendatory.remove(ALIAS[AVA][0])
    elif any(fuzz.partial_ratio(alias, fn_cf) > 86 for alias in ALIAS[SHI]):
      dfs[SHI] = load_shift_xlsx_into_df(df)
      mendatory.remove(ALIAS[SHI][0])
    elif any(fuzz.partial_ratio(alias, fn_cf) > 86 for alias in ALIAS[MON]):
      dfs[MON] = load_month_xlsx_into_df(df)
  if mendatory:
    dfs[AVA] = None
    dfs[LOG] += print_log_header(f'{MISSING_FILE_MSG}{mendatory}', '#', '', '')
  else:
    dfs[EE] = load_ersterfassung_xlsx_into_df(city)
  return dfs
# -------------------------------------

# -------------------------------------
def parse_availability_string(avails, week, extra, h_by_shifts=0):
  max_h = week + extra
  if isinstance(h_by_shifts, str):
    suf = NOT_IN_AV
  elif week <= h_by_shifts <= max_h:
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
def parse_city_ee_filepath(city):
  return join(REE_DIR, f'{EE}_{city}.xlsx')
# -------------------------------------

# -------------------------------------
def parse_city_runtime(city, start):
  return f'runtime {city = }: {time.perf_counter() - start:.2f} s'
# -------------------------------------

# -------------------------------------
def parse_date(day_str, kw_dates):
  return kw_dates[WEEKDAYS.index(day_str)]
# -------------------------------------

# -------------------------------------
def parse_plot_title(city, year, kw, date_str, day):
  return f'Verteilte Schichten {city} KW {kw}/{year}, {date_str} {day}'
# -------------------------------------

# -------------------------------------
def parse_progress_bar(bar_len, prog, pre, suf):
  done = int(bar_len * prog)
  return f'{pre} [{"#" * done + "-" * (bar_len - done)}] {prog:.2%} {suf}'
# -------------------------------------

# -------------------------------------
def parse_run_end_msg(start):
  return f'TOTAL RUNTIME: {time.perf_counter() - start:.2f} s'
# -------------------------------------

# -------------------------------------
def parse_sp_check_msg(city, kw, year):
  return f'{CITY_LOG_PRE} {city} | KW {kw} / {year}'
# -------------------------------------

# -------------------------------------
def parse_stats_msg(counter):
  return (
    f'| S | total rows count     : {counter[SCAN]:4d}{NL}'
    f'| T | duplicate rows       : {counter[DUPL]:4d}{NL}'
    f'| A | linked rows          : {counter[LINK]:4d}{NL}'
    f'| T | no availabilites     : {counter[NOAV]:4d}{NL}'
    f'| S | not readable         : {counter[NOOCR]:4d}'
  )
# -------------------------------------

# -------------------------------------
def plot_data_day_update(data, labels):
  day_labels = copy.deepcopy(labels)
  data = np.array([list(time_data.values()) for time_data in data.values()])
  for cntr, cnt in zip(range(len(labels) - 1, -1, -1), np.sum(data, 0)[::-1]):
    if int(cnt) == 0:
      data = np.delete(data, cntr, 1)
      del day_labels[cntr]
  return data, np.cumsum(data, axis=1), day_labels
# -------------------------------------

# -------------------------------------
def plot_initial_data_dict(dfs, labels, times):
  data_dict = {}
  times_set = {*times}
  for day in WEEKDAYS_EN:
    data_dict[day] = {}
    for time in times:
      data_dict[day][time] = {}
      for contract in labels:
        data_dict[day][time][contract] = 0
  for day, df_row in dfs[SHI].iterrows():
    try:
      contract = dfs[REP].at[df_row[DRI], CON_TYP]
    except KeyError:
      continue
    start_time = datetime.combine(_now, df_row[FR_HO])
    end_time = datetime.combine(_now, df_row[TO_HO])
    for dt_obj in pd.date_range(start_time, end_time, freq='30min'):
      time = dt_obj.strftime('%H:%M')
      if time in times_set:
        data_dict[day][time][contract] += 1
  return data_dict
# -------------------------------------

# -------------------------------------
def plot_init_subplot(city, year, kw, date_str, day):
  _, ax = plt.subplots(figsize=(18, 9))
  ax.invert_yaxis()
  ax.get_xaxis().set_visible(False)
  ax.set_facecolor('#e5e5e5')
  plot_title = parse_plot_title(city, year, kw, date_str, day)
  ax.set_title(plot_title, fontdict=TITLE_FMT, pad=35)
  return ax
# -------------------------------------

# -------------------------------------
def plot_partial_bar_count(ax, color, widths, xcenters):
  text_color = 'white' if np.average(color[:3]) < .5 else 'black'
  for y, (xt, ct) in enumerate(zip(xcenters, widths)):
    if ct == 0:
      continue
    ax.text(xt, y, str(int(ct)), ha=CEN, va=CEN, color=text_color)
  return ax
# -------------------------------------

# -------------------------------------
def plot_shift_distribution(dfs, city, kw, year, kw_dir):
  log = print_log_header(PLOT_SHIFTS_MSG)
  ana_dir = check_make_dir(kw_dir, 'Analyse')
  contracts = dfs[REP][CON_TYP].unique()
  dates_obj_list = [date.fromisocalendar(year, kw, i) for i in range(1, 8)]
  dates_DE = [date_obj.strftime(DMY) for date_obj in dates_obj_list]
  labels = [cont for ref_c in CON_BY_N for cont in contracts if cont == ref_c]
  times = TIMES[22][:-1]
  data_dict = plot_initial_data_dict(dfs, labels, times)
  for idx, day_dict in enumerate(data_dict.values()):
    data, data_cumsum, day_labels = plot_data_day_update(day_dict, labels)
    if data_cumsum.shape[1] == 0:
      continue
    target_fn = f'{city}_{dates_obj_list[idx]}.png'
    totals = data_cumsum[:, -1]
    peak_cnt = totals.max()
    ax = plot_init_subplot(city, year, kw, dates_DE[idx], WEEKDAYS[idx])
    ax = plot_stacked_bars(ax, data, data_cumsum, day_labels, times)
    ax = plot_total_bar_counts(ax, totals)
    ax.legend(day_labels, ncol=len(day_labels), **LGD_ARGS)
    plt.xlim(0, max(peak_cnt + 1, min(peak_cnt * 1.08, peak_cnt + 5)))
    plt.tight_layout()
    plt.savefig(join(ana_dir, target_fn))
    plt.close()
    log += print_log(f'{TAB} - saved plot = "{target_fn}"')
  log += print_log('-----')
  return log
# -------------------------------------

# -------------------------------------
def plot_stacked_bars(ax, data, data_cumsum, labels, times):
  for i, label in enumerate(labels):
    color = LGD_COLORS[label]
    widths = data[:, i]
    starts = data_cumsum[:, i] - widths
    xcenters = starts + widths / 2
    ax.barh(times, widths, left=starts, height=.7, color=color, label=label)
    ax = plot_partial_bar_count(ax, color, widths, xcenters)
  return ax
# -------------------------------------

# -------------------------------------
def plot_total_bar_counts(ax, totals):
  for i in range(len(totals)):
    tot = int(totals[i])
    if tot == 0:
      continue
    ax.text(tot + .5, i, tot, ha=LEF, va=CEN, fontdict=TOTAL_FMT)
  return ax
# -------------------------------------

# -------------------------------------
def png_grid_capture_cols(rows, left, img):
  row_cnt = len(rows)
  if row_cnt <= 1:
    return None, None, None
  cols = []
  img_width = img.shape[1]
  # ----- use found rows with offset as y value, start at the middle row -----
  for row_n in range(row_cnt // 2, row_cnt):
    cols.clear()
    row_height = rows[row_n] - rows[row_n - 1]
    bot_margin = 4 if row_n == (row_cnt - 1) else 2
  # ----- horizontal iteration, skip unnecessary parts, get x values of cols --
    for x in range(left + 8 * row_height, img_width - 4):
      pixel_color = img[rows[row_n] - bot_margin, x]
      if pixel_color in COLOR[TIME_CELL] | COLOR[SHI]:
        break
      if pixel_color in COLOR[LINE]:
        if cols and x <= cols[-1] + 5:
          del cols[-1]
        cols.append(x)
        if len(cols) == 26:
          break
  # ----- check break condition, continue if column count < 21 -----
      elif cols and pixel_color in COLOR[SCROLL_BAR]:
        if x <= cols[-1] + 2:
          del cols[-1]
        break
    if len(cols) >= 21:
      break
  col_cnt = len(cols)
  x_validation = left + (cols[0] - left) * 19 // 20 if col_cnt >= 21 else None
  return cols, col_cnt, x_validation
# -------------------------------------

# -------------------------------------
def png_grid_capture_rows(img):
  height = img.shape[0]
  height_thresh = height // 2
  left = 0
  rows = []
  skip = False
  # ----- find x value at which rows are separated by a thin line -----
  for x in range(0, 50):
    if skip is True:
      skip = False
      continue
    rows = [height]
    vert_line_cnt = y = 0
  # ----- vertical iteration, beginn at bottom, find y values of rows -----
    for y in range(height - 1, -1, -1):
      pixel_color = img[y, x]
      if pixel_color not in COLOR[LINE]:
        vert_line_cnt = 0
      elif all(img[y, x + i] in COLOR[NP] for i in range(3)):
        continue
      elif all(img[y, x + i] in COLOR[LINE] for i in range(50, 201, 50)):
        if y <= rows[-1] - 5:
          rows.append(y)
        elif rows[-1] == height:
          rows[0] = y
      else:
        vert_line_cnt += 1
        if vert_line_cnt == 5:
          skip = True
          break
  # ----- check break condition, continue if y > height threshold -----
      if y <= rows[-1] - (40 if y < height_thresh else 115):
        if len(rows) > 1:
          rows.append(2 * rows[-1] - rows[-2] + 4)
        else:
          y = height
        break
    if y < height_thresh:
      left = x
      break
  # ----- add 0 to rows if last found row value allows exactly one more row ---
  if 15 < rows[-1] <= 35:
    rows.append(0)
  return png_grid_check_bot_row_and_reverse(rows), left
# -------------------------------------

# -------------------------------------
def png_grid_check_bot_row_and_reverse(rows):
  raw_row_cnt = len(rows)
  if raw_row_cnt >= 2:
    mid_row_n = raw_row_cnt // 2
    row_height = rows[mid_row_n - 1] - rows[mid_row_n]
    if rows[0] >= rows[1] + 1.3 * row_height:
      rows[0] = rows[1] + row_height
  return rows[::-1]
# -------------------------------------

# -------------------------------------
def png_grid_remove_invalid_rows(rows, x, img):
  thresh = len(rows) // 2
  x_t = x * 7 // 8
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
    if all(img[y, x] in COLOR[DEL] | {img[y, x_t]} for y in range(top, bot)):
      del_rows.append(del_row)
  for del_row in del_rows:
    del rows[del_row]
  return rows, len(rows) - 1
# -------------------------------------

# -------------------------------------
def png_image_prep_ocr_string(frame, not_planable):
  ocr = pytesseract.image_to_string(frame, config='--psm 7').strip()
  if not_planable:
    for split_chars in SPLIT_CHARS:
      ocr = ocr.rsplit(split_chars, maxsplit=1)[0]
  return ocr.strip(STRIP_CHARS)
# -------------------------------------

# -------------------------------------
def png_image_variations_yield_ocr_name(cv_data):
  frame = cv_data[FRAME]
  if np.mean(frame) <= 245:
    for i in range(193, 208):
      frame[frame == i] = 255
    yield png_image_prep_ocr_string(frame, cv_data[NP])
    start_image_n = 2
  else:
    start_image_n = 0
  top, bot, left, right = cv_data[ROI]
  for image in cv_data[IMG_VARIATIONS][start_image_n:]:
    yield png_image_prep_ocr_string(image[top:bot,left:right], cv_data[NP])
# -------------------------------------

# -------------------------------------
def png_name_determination(ref_data, cv_data):
  hit = ''
  min_valid_perc = cv_data[VALID]
  ocr_read = []
  readable_img_cnt = 0
  scores = defaultdict(int)
  score_by_ocr = {}
  for ocr_name in png_image_variations_yield_ocr_name(cv_data):
    ocr_read.append(ocr_name)
    if ocr_name:
      readable_img_cnt += 1
    if len(ocr_name) < 6:
      continue
    try:
      for name, score in score_by_ocr[ocr_name].items():
        scores[name] += score
    except KeyError:
      hit, scores, img_score = png_name_main_algo(ocr_name, scores, ref_data)
      if hit:
        if hit in ref_data[2]:
          hit = png_name_similarity_check(hit, ocr_name, ref_data[2][hit])
        return hit, ocr_read
      score_by_ocr[ocr_name] = img_score
    hit = png_name_score_check(scores, readable_img_cnt, min_valid_perc)
    if hit:
      return hit, ocr_read
  if not (hit or cv_data[BAD_RESO]):
    hit = png_name_score_check(scores, readable_img_cnt, min_valid_perc, 1.01)
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
def png_name_main_algo(ocr_name, scores, ref_data):
  char_cnt = len(ocr_name)
  img_score = defaultdict(int)
  for name in ref_data[0]:
    query = name[:char_cnt]
    slice_similarity = fuzz.WRatio(ocr_name, query)
    if slice_similarity >= 89:
      return name, scores, img_score
    partial_similarity = fuzz.partial_ratio(ocr_name, query)
    if partial_similarity >= 93:
      return name, scores, img_score
    if fuzz.partial_ratio(ocr_name, name) >= 95:
      return name, scores, img_score
    if slice_similarity > 44:
      scores[name] += slice_similarity
      img_score[name] += slice_similarity
    if partial_similarity > 44:
      scores[name] += partial_similarity
      img_score[name] += partial_similarity
  return '', scores, img_score
# -------------------------------------

# -------------------------------------
def png_name_score_check(scores, img_n, min_valid_perc, leading_factor=1.35):
  if scores and img_n:
    msn, *rest = png_name_sorted_scores(scores)
    valid_score = img_n * 2 * min_valid_perc
    if len(scores) > 1:
      valid_score = max(leading_factor * rest[0][1], valid_score)
    if msn[1] >= valid_score:
      return msn[0]
  return ''
# -------------------------------------

# -------------------------------------
def png_name_similarity_check(hit, ocr_name, simil_names):
  highest = 0
  if any(fuzz.WRatio(hit, name) == 100 for name in simil_names):
    method = 'ratio'
  else:
    method = 'WRatio'
  for name in (hit, *simil_names):
    similarity = getattr(fuzz, method)(ocr_name, name)
    if similarity > 95:
      return name
    if similarity > highest:
      highest = similarity
      hit = name
  return hit
# -------------------------------------

# -------------------------------------
def png_name_sorted_scores(scores, result_cnt=5):
  return sorted(scores.items(), key=lambda x: x[1], reverse=True)[:result_cnt]
# -------------------------------------

# -------------------------------------
def png_row_availabities(top, bot, cols, col_cnt, date_str, img):
  daily_avail = ''
  daily_hours = extra_hours = hours_block = 0
  in_availablity_block = False
  for col_idx, col in enumerate(cols):
    filled = png_row_check_cell_filling(col + 1, top, bot, img)
    if filled is None:
      break
    if filled:
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
    if pixel_color in COLOR[FILLED]:
      return True
    elif pixel_color in (COLOR[SHI] | COLOR[NP]):
      return None
  return False
# -------------------------------------

# -------------------------------------
def png_row_get_data(data, png_vals):
  avails, hours, extra = png_row_availabities(*png_vals[AV_ARGS])
  if not avails:
    data[COUNTER][NOAV] += 1
    return data
  name, ocr_read = png_name_determination(*png_vals[NAME_ARGS])
  data[LOG_DATA].append((*png_vals[LOG_DATA], avails[:-1], name, ocr_read))
  if not name:
    data[COUNTER][NOOCR] += 1
    data[LOG] += print_no_name_determined(avails, name, ocr_read, png_vals)
  elif png_vals[DATE] in data[DONE][name]:
    data[COUNTER][DUPL] += 1
  else:
    data[COUNTER][LINK] += 1
    data[AVA][name].append(avails)
    data[HRS][name] += hours
    data[XTR][name] += extra
    data[DONE][name].add(png_vals[DATE])
  return data
# -------------------------------------

# -------------------------------------
def png_values_cv_data(rows, row_cnt, left, first_col, img):
  mid_row_n = row_cnt // 2
  row_height = rows[mid_row_n + 1] - rows[mid_row_n]
  resize_factor = 109 / row_height
  right = int(.29 * first_col)
  res_width = int(resize_factor * right)
  res_height = int(resize_factor * img.shape[0])
  res_img = cv.resize(img[:, left:right], (res_width, res_height))
  return {
    BAD_RESO: row_height < 21
    , BORDER: (left, right)
    , IMG_VARIATIONS: png_values_get_image_variations(res_img)
    , RE_MA: int(resize_factor * png_values_get_margin(rows, row_height))
    , RE_ROW: [int(resize_factor * row) for row in rows]
    , ROI: [0, 0, 0, res_width]
    , VALID: png_values_get_min_valid_perc(row_height)
  }
# -------------------------------------

# -------------------------------------
def png_values_get_image_variations(image):
  return (
    cv.threshold(image, 220, 255, cv.THRESH_BINARY)[1]
    , CLAHE_DEF.apply(image)
    , cv.threshold(image, 212, 255, cv.THRESH_BINARY)[1]
    , cv.filter2D(image, -1, KERNEL)
    , CLAHE_L3_S7.apply(image)
  )
# -------------------------------------

# -------------------------------------
def png_values_get_margin(rows, row_height):
  return 0 if rows[-1] - rows[-2] < row_height * 2 // 3 else 3
# -------------------------------------

# -------------------------------------
def png_values_get_min_valid_perc(height):
  return 33 if height < 19 else 66 if height > 65 else height * 11 / 6
# -------------------------------------

# -------------------------------------
def png_values_image_values(date_str, img, ref_data):
  rows, left = png_grid_capture_rows(img)
  cols, col_cnt, x_valid = png_grid_capture_cols(rows, left, img)
  if x_valid is None:
    return {ROW_CNT: 0}
  rows, row_cnt = png_grid_remove_invalid_rows(rows, x_valid, img)
  if row_cnt == 0:
    return {ROW_CNT: 0}
  cv_data = png_values_cv_data(rows, row_cnt, left, cols[0], img)
  return {
    ROWS: rows
    , ROW_CNT: row_cnt
    , NAME_ARGS: [ref_data, cv_data]
    , AV_ARGS: [None, None, cols, col_cnt, date_str, img]
  }
# -------------------------------------

# -------------------------------------
def png_values_imread(png, png_dir):
  image = cv.imread(join(png_dir, png), cv.IMREAD_GRAYSCALE)
  if image is None:
    shutil.copy(join(png_dir, png), png)
    image = cv.imread(png, cv.IMREAD_GRAYSCALE)
    os.remove(png)
  return image
# -------------------------------------

# -------------------------------------
def png_values_sort_key(png):
  day, idx_str = png.split('.')[0].split('_')
  return WEEKDAYS.index(day), int(idx_str)
# -------------------------------------

# -------------------------------------
def png_values_yield_pngs(ref_data, city, kw, year, png_dir):
  image_vals = {}
  kw_dates = [str(date.fromisocalendar(year, kw, i)) for i in range(1, 8)]
  pngs = os.listdir(png_dir)
  for png_n, png in enumerate(sorted(pngs, key=png_values_sort_key)):
    day, file_suf = png.split('_')
    date_str = parse_date(day, kw_dates)
    img = png_values_imread(png, png_dir)
    image_vals = png_values_image_values(date_str, img, ref_data)
    if image_vals[ROW_CNT] == 0:
      continue
    yield {
      **image_vals
      , DATE: date_str
      , IMG: img
      , LOG_DATA: [kw, city, day, int(file_suf.split('.')[0]), None]
      , BAR: [png, len(pngs), png_n]
      , PNG: png
      , PNG_N: png_n
    }
  if image_vals.get(ROWS, None) is None:
    return image_vals
# -------------------------------------

# -------------------------------------
def png_values_yield_rows(png_vals):
  cv_data = png_vals[NAME_ARGS][1]
  for row_n in range(1, png_vals[ROW_CNT] + 1):
    print_progress_bar(png_vals[BAR], png_vals[ROW_CNT], row_n)
    png_vals[LOG_DATA][4] = png_vals[ROW_N] = row_n
    top = png_vals[AV_ARGS][0] = png_vals[ROWS][row_n - 1] + 2
    bot = png_vals[AV_ARGS][1] = png_vals[ROWS][row_n] - 2
    cv_data[FRAME] = png_vals[IMG][top:bot, slice(*cv_data[BORDER])]
    cv_data[ROI][0] = cv_data[RE_ROW][row_n - 1]
    cv_data[ROI][1] = cv_data[RE_ROW][row_n] - cv_data[RE_MA]
    cv_data[NP] = png_vals[IMG][top, cv_data[BORDER][1]] in COLOR[NP]
    png_vals[NAME_ARGS][1] = cv_data
    yield png_vals
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
def print_no_name_determined(avail_str, name, ocr_read, png_vals):
  print('\r', end='')
  return print_log(
    f'##### {NF}png={png_vals[PNG]}, row={png_vals[ROW_N]}{" " * 30}{NL}'
    + (f'|OCR| {ocr_read = }{NL}' if ocr_read else '')
    + f'|AVA| {"|".join(avail_str.split(NL)).strip("|")}'
    + (f', {name = }' if name else '')
    + BR
  )
# -------------------------------------

# -------------------------------------
def print_progress_bar(bar_data, row_cnt, row_n):
  png, png_cnt, png_n = bar_data
  progress = (png_n + row_n / row_cnt) / png_cnt
  if progress < 1:
    pre = ' ==> '
    print_end = '\r'
    suf = f'{png} ({row_n}/{row_cnt})'
  else:
    progress = 1
    pre = '|FIN|'
    print_end = '\r\n-----\n'
    suf = f'... DONE'
  bar_str = parse_progress_bar(30, progress, pre, suf)
  padding = os.get_terminal_size().columns - len(bar_str) - 13
  if padding < 0:
    bar_str = parse_progress_bar(30 + padding, progress, pre, suf)
  print(bar_str, ' ' * 12, sep='', end=print_end, flush=True)
# -------------------------------------

# -------------------------------------
def processed_ocr_data_to_logfile(data, city, kw_dir):
  writer = pd.ExcelWriter(join(kw_dir, f'det_names_{city}_{START_DT}.xlsx'))
  pd.DataFrame(data, columns=DF_DET_COLS).to_excel(writer, city, index=False)
  worksheet = writer.sheets[city]
  worksheet.autofilter('A1:H1')
  worksheet.freeze_panes(1, 0)
  worksheet.set_column('B:C', 12)
  worksheet.set_column('F:H', 35)
  writer.save()
# -------------------------------------

# -------------------------------------
def processed_xlsx_data_to_report_df(dfs):
  data_list = []
  src = AVA if dfs[MON] is None else MON
  for name, df_row in dfs[src].iterrows():
    data = get_new_df_entry()
    data.update(get_base_data(name, df_row, src, dfs[AVA]))
    data.update(get_min_hours(dfs[EE], dfs[REF][3], data[CON_TYP], name))
    data.update(get_shifts(dfs[SHI], data[ID]))
    data.update(get_given_hour_ratios(data[AVA], data[GIV], data[MAX]))
    data.update(get_data_check_and_first_comment(data))
    data_list.append(data)
  return pd.DataFrame(data_list).set_index(RID_NAM, drop=False)
# -------------------------------------

# -------------------------------------
def process_screenshots(ref_data, city, kw, year, dirs):
  data = copy.deepcopy(PNG_PROCESSING_DICT)
  data[LOG] = print_log_header(PROCESS_PNG_MSG)
  for png_vals in png_values_yield_pngs(ref_data, city, kw, year, dirs[3]):
    data[COUNTER][SCAN] += png_vals[ROW_CNT]
    for png_vals in png_values_yield_rows(png_vals):
      data = png_row_get_data(data, png_vals)
  if data[LOG_DATA]:
    data[LOG] += print_log(parse_stats_msg(data[COUNTER]), end=BR)
    processed_ocr_data_to_logfile(data[LOG_DATA], city, dirs[1])
  else:
    data[LOG] += print_log(NO_SCREENS_MSG, 'XXXXX ', BR)
  return data
# -------------------------------------

# -------------------------------------
def process_xlsx_data(dfs, kw_date, city, dirs):
  dfs[LOG] += print_log_header(PROCESS_XLSX_MSG)
  dfs = load_xlsx_data_into_dfs(city, dirs, dfs)
  if dfs[AVA] is None:
    return dfs
  dfs = rider_ee_update_names(kw_date, city, dfs)
  dfs[REF] = reference_names_and_contract_data(dfs[EE], kw_date)
  dfs[REP] = processed_xlsx_data_to_report_df(dfs)
  return dfs
# -------------------------------------

# -------------------------------------
def reference_contract_list(kw_date, df_kw):
  ref_contracts = []
  for _, d in df_kw.iterrows():
    if kw_date >= d[CC_FE]:
      ref_contracts.append(d[CON_TYP])
      continue
    for contract_data in d[PRE_C].split(NL):
      fe_le, contract = contract_data.split(' | ')
      first_entry, last_entry = map(date.fromisoformat, fe_le.split(' - '))
      if kw_date < first_entry or kw_date > last_entry:
        continue
      ref_contracts.append(contract)
      break
  return ref_contracts
# -------------------------------------

# -------------------------------------
def reference_names_and_contract_data(df, kw_date):
  df_ref = df[(df[LA_ENT] > (kw_date - TOO_OLD)) & (df[FI_ENT] <= kw_date)]
  ref_names = df_ref[RID_NAM].to_list()
  ref_contracts = reference_contract_list(kw_date, df_ref)
  simil_dict = {name: x.split(';') for name, x in df_ref[SIM_NAM].items() if x}
  names_set = {*ref_names}
  return ref_names, ref_contracts, simil_dict, names_set
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
def rider_ee_new_df_entry(name, avails, data, ref):
  eee = get_new_df_entry()
  contract = ref if isinstance(ref, str) else ref[1][ref[0].index(name)]
  eee.update({RID_NAM: name, AVA: 0, CON_TYP: contract})
  eee[MIN], eee[MAX] = CON_BY_N[contract]
  avail = parse_availability_string(avails, data[HRS][name], data[XTR][name])
  eee[AVAILS] = avail + NOT_AV_MON
  for key in (GIV, GIV_AVA, GIV_MAX, PAI_MAX, WOR, VAC, SIC, PAI, UNP):
    eee[key] = NO_DA
  return eee
# -------------------------------------

# -------------------------------------
def rider_ee_new_xlsx_entry(city, new_name, contract, kw_monday_date):
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
def rider_ee_parse_pre_c(df_row):
  pre = (df_row[PRE_C] + NL) if df_row[PRE_C] else ""
  return f'{pre}{df_row[CC_FE]} - {df_row[LA_ENT]} | {df_row[CON_TYP]}'
# -------------------------------------

# -------------------------------------
def rider_ee_pre_c_update(df_row, contract, kw_date):
  prev_contracts_string = df_row[PRE_C]
  if prev_contracts_string == '':
    return f'{kw_date} - {kw_date} | {contract}'
  contract_in_prev_c = contract in prev_contracts_string
  updated_prev_c = ''
  for contract_line in prev_contracts_string.split(NL):
    fe_le, prev_contract = contract_line.split(' | ')
    first_entry, last_entry = map(date.fromisoformat, fe_le.split(' - '))
    if contract == prev_contract:
      if kw_date < first_entry:
        first_entry = kw_date
      elif kw_date > last_entry:
        last_entry = kw_date
      new_line = f'{first_entry} - {last_entry} | {contract}'
    elif contract_in_prev_c or kw_date == last_entry:
      new_line = contract_line
    elif kw_date > last_entry:
      new_line = f'{contract_line}{NL}{kw_date} - {kw_date} | {contract}'
      contract_in_prev_c = True
    else:
      new_line = f'{kw_date} - {kw_date} | {contract}{NL}{contract_line}'
      contract_in_prev_c = True
    updated_prev_c += ('' if updated_prev_c == '' else NL) + new_line
  return updated_prev_c
# -------------------------------------

# -------------------------------------
def rider_ee_to_formated_xlsx(city, df_ee):
  row_cnt = df_ee.shape[0] + 1
  writer = pd.ExcelWriter(parse_city_ee_filepath(city))
  df_ee.to_excel(writer, city, index=False)
  workbook = writer.book
  worksheet = writer.sheets[city]
  worksheet.autofilter('A1:I1')
  worksheet.freeze_panes(1, 0)
  fmt = {key: workbook.add_format(FMT_DICT[key]) for key in RIDER_EE_FMTS}
  for column, width, fmt_key in RIDER_EE_COL_FMT:
    worksheet.set_column(column, width, fmt[fmt_key])
  for columns, cond_key, fmt_key in RIDER_EE_CONDS:
    fmts = {**COND_FMT[cond_key], FMT: fmt[fmt_key]}
    worksheet.conditional_format(f'{columns}{row_cnt}', fmts)
  writer.save()
# -------------------------------------

# -------------------------------------
def rider_ee_update_known_names(df_ee, name, contract, kw_date):
  df_row = df_ee.loc[name]
  if kw_date > df_row[LA_ENT]:
    if contract != df_row[CON_TYP]:
      df_ee.at[name, PRE_C] = rider_ee_parse_pre_c(df_row)
      df_ee.at[name, CON_TYP] = contract
      df_ee.at[name, MIN] = CON_BY_N[contract][0]
      df_ee.at[name, CC_FE] = kw_date
    df_ee.at[name, LA_ENT] = kw_date
  elif kw_date < df_row[LA_ENT]:
    if kw_date < df_row[FI_ENT]:
      df_ee.at[name, FI_ENT] = kw_date
    if contract != df_row[CON_TYP]:
      df_ee.at[name, PRE_C] = rider_ee_pre_c_update(df_row, contract, kw_date)
    elif kw_date < df_row[CC_FE]:
      df_ee.at[name, CC_FE] = kw_date
  return df_ee
# -------------------------------------

# -------------------------------------
def rider_ee_update_names(kw_date, city, dfs):
  dfs[LOG] += print_log_header(SYNCH_MIN_H_MSG)
  dfs[LOG] += print_log(f'CHECK {MH_AVAIL_MSG}{dfs[MON] is not None} {BR}')
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
        new_data.append(rider_ee_new_xlsx_entry(city, name, contract, kw_date))
        new_names.add(name)
        dfs[LOG] += print_log(f'{TAB}- {name = }, {contract = }')
      elif name in known:
        dfs[EE] = rider_ee_update_known_names(dfs[EE], name, contract, kw_date)
      processed_names.add(name)
    if new_names_in_df:
      dfs[LOG] += print_log('-----')
  if new_data:
    dfs[EE] = rider_ee_insert_new_names(new_data, new_names, dfs[EE])
  return dfs
# -------------------------------------

# -------------------------------------
def screenshots_list_of_daily_files(day, png_dir, Image):
  return [
    Image.open(join(png_dir, day_fn))
    for day_fn in sorted(fn for fn in os.listdir(png_dir) if day in fn)
  ]
# -------------------------------------

# -------------------------------------
def screenshots_merge_daily_files(city, dirs):
  from PIL import Image
  log = print_log_header(MERGE_FILES_MSG)
  for day in WEEKDAYS:
    images = screenshots_list_of_daily_files(day, dirs[3], Image)
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
def shiftplan_check(city, kw, year, dirs, run_args):
  get_ava, merge, tidy_only, ee_only, visualize = run_args
  start = time.perf_counter()
  pre = '=' * os.get_terminal_size().columns + NL
  dfs = {LOG: print_log_header(parse_sp_check_msg(city, kw, year), pre=pre)}
  dfs[LOG] += tidy_screenshot_files(city, dirs, merge)
  if tidy_only:
    return shiftplan_check_log_city_runtime(city, start, dfs[LOG])
  kw_date = date.fromisocalendar(year, kw, 1)
  dfs = process_xlsx_data(dfs, kw_date, city, dirs)
  if dfs.get(AVA, None) is None:
    return shiftplan_check_log_city_runtime(city, start, dfs[LOG])
  if visualize and dfs.get(REP, None) is not None:
    dfs[LOG] += plot_shift_distribution(dfs, city, kw, year, dirs[0])
  if get_ava and TESSERACT_AVAILABLE:
    data = process_screenshots(dfs[REF], city, kw, year, dirs)
    dfs = shiftplan_report_png_data_update(dfs, data, kw_date, city)
  rider_ee_to_formated_xlsx(city, dfs[EE])
  if ee_only:
    return shiftplan_check_log_city_runtime(city, start, dfs[LOG])
  dfs[REP] = shiftplan_report_remove_unnecessary(dfs[REP])
  dfs[LOG] += shiftplan_report_to_formated_xlsx(dfs[REP], city, kw)
  return shiftplan_check_log_city_runtime(city, start, dfs[LOG])
# -------------------------------------

# -------------------------------------
def shiftplan_check_log_city_runtime(city, start, log):
  return log + print_log_header(parse_city_runtime(city, start), suf='=')
# -------------------------------------

# -------------------------------------
def shiftplan_report_png_data_update(dfs, data, kw_date, city):
  dfs[LOG] += data.pop(LOG)
  only_ee = []
  only_screen = []
  os_names = []
  # ----- iterate through read data, update availabilities in report df -----
  for name, avails in data[AVA].items():
  # ----- case 1: name is already in report df -----
    try:
      dfs[REP].at[name, AVAILS] = parse_availability_string(
        avails, data[HRS][name], data[XTR][name], dfs[REP].at[name, AVA]
      )
      if dfs[REP].at[name, AVA] == '':
        dfs[REP].at[name, AVA] = data[HRS][name]
    except KeyError:
  # ----- case 2: name not in report df, but in Rider_Ersterfassung -----
      if name in dfs[REF][3]:
        only_ee.append(rider_ee_new_df_entry(name, avails, data, dfs[REF]))
  # ----- case 3: name neither in report df, nor in Rider_Ersterfassung -----
      else:
        only_ee.append(rider_ee_new_df_entry(name, avails, data, NO_DA))
        only_screen.append(rider_ee_new_xlsx_entry(city, name, NO_DA, kw_date))
        os_names.append(name)
  if only_ee:
    dfs[REP] = dfs[REP].append(only_ee).sort_values([RID_NAM])
  dfs[REP].loc[dfs[REP][AVA] == dfs[REP][AVAILS], AVA] = 0
  dfs[REP].loc[(dfs[REP][AVA] != 0) & (dfs[REP][AVAILS] == ''), AVAILS] = ' '
  if only_screen:
    dfs[EE] = rider_ee_insert_new_names(only_screen, os_names, dfs[EE])
  return dfs
# -------------------------------------

# -------------------------------------
def shiftplan_report_remove_unnecessary(df):
  return df[(df[AVA] != 0) | (df[AVAILS] != '') | (df[GIV_SHI] != '')]
# -------------------------------------

# -------------------------------------
def shiftplan_report_to_formated_xlsx(df_report, city, kw):
  log = print_log_header(CREATE_XLSX_MSG)
  row_cnt = len(df_report) + 1
  # ----- open instance of xlsx-file -----
  filename = f'{OUT_FILE_PRE}KW{kw}_{city}-{START_DT}.xlsx'
  writer = pd.ExcelWriter(join(OUTPUT_DIR, filename), engine='xlsxwriter')
  df_report.to_excel(writer, 'Sheet1', index=False)
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
    fmts = {**COND_FMT[con], FMT: fmt_dict[fmt]} if fmt else COND_FMT[con]
    worksheet.conditional_format(f'{cols}{row_cnt}', fmts)
  # ----------
  writer.save()
  return log + print_log(f'+++++ saved {filename}{BR}')
# -------------------------------------

# -------------------------------------
def tidy_filename_query(fn):
  return ''.join(c for c in fn.casefold().replace('_', ' ') if c not in DIGITS)
# -------------------------------------

# -------------------------------------
def tidy_jpg_files(city, dirs):
  jpg_files = [fn for fn in os.listdir(dirs[0]) if fn.endswith('.jpg')]
  if not jpg_files:
    return ''
  log = print_log_header(TIDY_JPG_MSG)
  log += print_log(JPG_NAME_CHECK_MSG, '|JPG| ')
  idx_dict = defaultdict(int)
  raw_dir = check_make_dir(dirs[2], 'raw')
  vac_dir = check_make_dir(raw_dir, 'Urlaub')
  for fn in sorted(jpg_files):
    fn_cf = tidy_filename_query(fn)
    if not any(fuzz.WRatio(alias, fn_cf) > 86 for alias in ALIAS[city]):
      continue
    source = join(dirs[0], fn)
    if 'urlaub' in fn.casefold():
      shutil.move(source, join(vac_dir, fn))
      continue
    new_fn, idx_dict, log = tidy_screenshot_fn(fn, fn_cf, idx_dict, log, 'jpg')
    shutil.copy(source, join(dirs[3], new_fn))
    shutil.move(source, join(raw_dir, fn))
  return log
# -------------------------------------

# -------------------------------------
def tidy_png_files(city, dirs):
  png_files = [fn for fn in os.listdir(dirs[0]) if fn.endswith('.png')]
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
    fn_cf = tidy_filename_query(fn)
    if all(fuzz.partial_ratio(alias, fn_cf) < 87 for alias in ALIAS[city]):
      continue
    new_fn, idx_dict, log = tidy_screenshot_fn(fn, fn_cf, idx_dict, log)
    shutil.copy(source, join(raw_dir, fn))
    shutil.move(source, join(dirs[3], new_fn))
  return log
# -------------------------------------

# -------------------------------------
def tidy_screenshot_files(city, dirs, merge):
  log = ''
  log += tidy_jpg_files(city, dirs)
  log += tidy_png_files(city, dirs)
  log += tidy_zip_files(city, dirs)
  if log:
    log += print_log(f'+++++ saved unified screenshots in: {dirs[3]}{BR}')
    if merge:
      log += screenshots_merge_daily_files(city, dirs)
  return log
# -------------------------------------

# -------------------------------------
def tidy_screenshot_fn(original, fn_cf, idx_dict, log, suf='png'):
  similarity = 0
  current_day = ''
  for weekday in WEEKDAYS:
    weekday_similarity = fuzz.partial_ratio(weekday, original)
    if weekday_similarity > similarity:
      current_day = weekday
      similarity = weekday_similarity
      if similarity > 90:
        break
  if similarity <= 90:
    for weekday_n, abrevations in enumerate(WEEKDAY_ABREVATIONS):
      if any(fuzz.token_set_ratio(abr, fn_cf) == 100 for abr in abrevations):
        current_day = WEEKDAYS[weekday_n]
        similarity = 100
        break
  idx_dict[current_day] += 1
  tidy_png_fn = f'{current_day}_{idx_dict[current_day]}.{suf}'
  if similarity != 100:
    log += print_log(f'{TAB}- {original = }, saved as = {tidy_png_fn}')
  return tidy_png_fn, idx_dict, log
# -------------------------------------

# -------------------------------------
def tidy_zip_files(city, dirs):
  zip_files = [fn for fn in os.listdir(dirs[0]) if fn.endswith('.zip')]
  if not zip_files:
    return ''
  from zipfile import ZipFile
  log = print_log_header(UNZIP_MSG)
  log += print_log(ZIP_PNG_NAME_CHECK_MSG, '[X|O] ')
  idx_dict = defaultdict(int)
  for zip_filename in zip_files:
    zip_cf = tidy_filename_query(zip_filename)
    if all(fuzz.partial_ratio(alias, zip_cf) < 87 for alias in ALIAS[city]):
      continue
    log += print_log(zip_filename, TAB)
    with ZipFile(join(dirs[0], zip_filename)) as zfile:
      for member in sorted(zfile.namelist()):
        fn = basename(member)
        if not fn:
          continue
        fn_cf = tidy_filename_query(fn)
        fn, idx_dict, log = tidy_screenshot_fn(fn, fn_cf, idx_dict, log)
        with open(join(dirs[3], fn), "wb") as target:
          shutil.copyfileobj(zfile.open(member), target)
    log += print_log('-----')
  return log
# -------------------------------------

# -------------------------------------
def yield_run_kws(start_year, last_year, start_kw, last_kw):
  if last_year <= start_year:
    last_year = start_year
    if last_kw < start_kw:
      last_kw = start_kw
  first_kw = date.fromisocalendar(start_year, start_kw, 1)
  end_kw = date.fromisocalendar(last_year, last_kw + 1, 1)
  for current_kw in pd.date_range(first_kw, end_kw, freq='W'):
    yield current_kw.isocalendar()[:2]
# -------------------------------------
# =================================================================


# =================================================================
# ### MAIN FUNCTION ###
# =================================================================
# -------------------------------------
def main(start_year, last_year, start_kw, last_kw, cities, *run_args):
  start = time.perf_counter()
  log = print_log_header(INITIAL_MSG, pre='=', suf='=')
  for year, kw in yield_run_kws(start_year, last_year, start_kw, last_kw):
    kw_dir = join(BASE_DIR, 'Schichtplan_Daten', str(year), f'KW{kw}')
    if not exists(kw_dir):
      print(f'##### Couldn`t find "{kw_dir}"{BR}')
      continue
    log_dir = check_make_dir(kw_dir, 'logs')
    screen_dir = join(kw_dir, 'Screenshots')
    dirs = [kw_dir, log_dir, screen_dir, None]
    try:
      for city in cities:
        dirs[3] = check_make_dir(screen_dir, city)
        log += shiftplan_check(city, kw, year, dirs, run_args)
    except Exception as ex:
      log += f'{type(ex)=} | {repr(ex)=}{NL}{parse_break_line("#")}'
      raise ex
    finally:
      with open(join(log_dir, LOG_FN), 'w', encoding='utf-8') as logfile:
        logfile.write(log)
  print_log_header(parse_run_end_msg(start), pre='=', suf='=', brk=NL)
# -------------------------------------
# =================================================================


# =================================================================
# ### START SCRIPT ###
# =================================================================
# -------------------------------------
if __name__ == '__main__':
  from argparse import ArgumentParser
  parser = ArgumentParser()
  parser.add_argument('-y', '--start_year', type=int, default=YEAR, help=P_Y)
  parser.add_argument('-z', '--last_year', type=int, default=YEAR, help=P_X)
  parser.add_argument('-k', '--start_kw', type=int, default=KW, help=P_KW)
  parser.add_argument('-l', '--last_kw',  type=int, default=0, help=P_LKW)
  parser.add_argument('-c', '--cities', nargs='*', default=DEF_CITY, help=P_C)
  parser.add_argument('-a', '--get_avail', action=STORE_TRUE, help=P_A)
  parser.add_argument('-t', '--tidy_only', action=STORE_TRUE, help=P_TO)
  parser.add_argument('-m', '--mergeperday', action=STORE_TRUE, help=P_M)
  parser.add_argument('-e', '--ersterfassung',  action=STORE_TRUE, help=P_EEO)
  parser.add_argument('-v', '--visualize_shifts', action=STORE_TRUE, help=P_V)
  main(*parser.parse_args().__dict__.values())
# -------------------------------------
# =================================================================
