
# Standalone NPS report with five tabs (Business Center, District, TAs, Dealer, DOM)
# saved as Excel. Unlike DOD8.py, this script does NOT load/overwrite a pre-formatted
# "Base File" template -- it builds the entire workbook (headers, data, color scales)
# from scratch every run, in a fixed column layout:
#   <tab identification columns -- codes AND names>, NPS, FFV %,
#   Advocates, Neutrals, Detractors, Total, FFV Passed,
#   Prior Month NPS, Prior Month FFV %
# Filtered by an explicit start/end date range on date_survey_received (inclusive)
# Co-authored with CoCo

import calendar
import snowflake.connector
import openpyxl
import os
import time
import win32com.client
import pywintypes
from datetime import date
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter

# Folder where the dated report snapshot gets saved. No longer read back for
# "last report" history -- Prior Month NPS/FFV% now come from re-running each
# tab's query over the prior month's own date window (see PREV_START_DATE /
# PREV_END_DATE below), not from a previously saved report file.
REPORT_FOLDER = r"C:\Users\Ta23458\OneDrive - Stellantis\BC Folders\2026\Staff Folders\Jesse G\NPS Reports"

# --- Step 1: Connect to Snowflake and run five NPS queries (current + prior month) ---
print("Connecting to Snowflake...")
conn = snowflake.connector.connect(
    account="nu14656-stellantis_aws_us_east_dataservices_prod",
    user="TA23458@STELLANTIS.COM",
    authenticator="externalbrowser",
    role="N_RL_F_AFS_PUB_USER",
    warehouse="N_WH_AFTERSALES_REG",
    database="DATALAKE_PROD",
    client_store_temporary_credential=True
)

cur = conn.cursor()

# Explicit reporting window, filtered on date_survey_received (inclusive on both
# ends) instead of the report_month column.
START_DATE = date(2026, 11, 3)
END_DATE = date(2026, 11, 30)
print(f"Report period: {START_DATE} to {END_DATE} (date_survey_received)")


def shift_back_one_month(d):
    """Return d shifted back exactly one calendar month, clamping the day to
    the prior month's last day if needed (e.g. Mar 31 -> Feb 28)."""
    if d.month == 1:
        new_year, new_month = d.year - 1, 12
    else:
        new_year, new_month = d.year, d.month - 1
    last_day = calendar.monthrange(new_year, new_month)[1]
    return d.replace(year=new_year, month=new_month, day=min(d.day, last_day))


# Prior-month equivalent of START_DATE/END_DATE, used to re-run every query a
# second time so each tab's "Prior Month NPS"/"Prior Month FFV %" columns come
# straight from Snowflake instead of from a previously saved report file.
PREV_START_DATE = shift_back_one_month(START_DATE)
PREV_END_DATE = shift_back_one_month(END_DATE)
print(f"Prior month period: {PREV_START_DATE} to {PREV_END_DATE} (date_survey_received)")

# All five queries run off MPR_CX_SERVICE_SURVEY_DETAILS_VW (one row per survey) and
# report every survey received in the window above. The brand_code in ('C','D','J','R')
# filter is replicated by keeping dealers whose dealer-level FRANCHISE string contains
# any of those letters (a dealer carrying Fiat/Alfa alongside CDJR franchises is still
# included), matching prior behavior.
#
# NOTE: each query uses a two-stage aggregation (base -> agg -> final SELECT). The
# inner "agg" CTE sums raw columns into *_sum names that are DIFFERENT from the
# underlying table's column names (e.g. ffv_passed_sum, not ffv_passed). Only the
# outer, non-aggregating SELECT renames them back to ffv_passed/advocates/etc. This
# avoids a Snowflake alias-resolution bug where "sum(ffv_passed) as ffv_passed"
# followed by another reference to "ffv_passed" later in the same SELECT list (e.g.
# "ffv_passed / nullif(ffv_total, 0)") gets resolved to the alias instead of the base
# column, raising:
#   "'BASE.FFV_PASSED' in select clause is neither an aggregate nor in the group by
#   clause." Do not rename the aggregate output back to a name matching a raw source
#   column within the same SELECT list that also references that raw column.
BRAND_FILTER = """
    (franchise LIKE '%C%' OR franchise LIKE '%D%' OR franchise LIKE '%J%' OR franchise LIKE '%R%')
"""

# Hardcoded TA (Technical Advisor) name for each TA district code, from
# "WBC TA 11122026.csv" (Employee Name column).
TA_DISTRICT_NAME_MAP = {
    "A": "ROBERT SUNDIN",
    "B": "WALLACE H SHIMABUKURO",
    "C": "JEFFREY SHMANSKE",
    "D": "YALE T RENDLEMAN",
    "E": "TATE BUNDERSON",
    "F": "BRETT FANNING",
    "G": "JEREMY M RHOADES",
    "H": "MICHAEL I GOLDWASSER",
    "J": "MICHAEL BERLINGEN",
    "K": "SAMUEL MACKEY",
    "L": "CELERINO VALENTON",
    "M": "RILEY CAMPBELL",
    "N": "RUSTIN MOON",
    "P": "ARTIN ARSATOORIAN",
    "Q": "MAXIM BURTON",
    "R": "SHAWN WHITE",
    "S": "JACOB CLOUGH",
    "9": "TBD",
}

# Hardcoded Service Manager (displayed as "Area Manager") for each sales district
# code, from "dealer alignment wbc.csv" (Service Manager column, keyed by that
# file's District column). Districts 2 and O are currently under temporary/open
# coverage (no permanent Area Manager assigned), so those two are intentionally
# just labeled "TEMP Coverage" rather than naming who's filling in. Also reused
# as the Dealer tab's "AM Name" column.
DISTRICT_SERVICE_MANAGER_MAP = {
    "1": "Morgan Engelstad",
    "2": "TEMP Coverage",
    "3": "Stephen Langsam",
    "A": "Ryan Pike",
    "B": "Ashley Randall",
    "C": "Gerald McKnght",
    "D": "Cody Potter",
    "E": "Josh Florea",
    "F": "Corbin Jones",
    "G": "Jesser Johnston",
    "H": "Justin Russell",
    "I": "Pacifique Nsanzamahoro",
    "J": "Claire Paparella",
    "K": "Steven Rich",
    "L": "Jason Kearns",
    "M": "Matt Greenland",
    "N": "Brian Stoskopf",
    "O": "TEMP Coverage",
    "P": "Adam Smiarowski",
    "Q": "Avo Khano",
    "R": "Sheri Whitesides",
    "S": "Emmanuella Kariotis",
    "T": "Shelly Shropshire",
    "U": "Annette Overstrom",
    "V": "Hope Meinert Shukla",
    "W": "Skyler Harrington",
    "X": "Stephanie Britt",
    "Y": "Clancy Pogue",
    "Z": "DJ James-Brannon",
}

# Hardcoded FOM (Field Operations Manager) for each sales district code, from
# "dealer alignment wbc.csv" (Service FOM column, keyed by that file's District
# column). Verified 1:1 against Snowflake (every district code maps to exactly
# one FOM). First name only, per request.
DISTRICT_FOM_MAP = {
    "1": "Bob",
    "2": "Ryan",
    "3": "Ryan",
    "A": "Bob",
    "B": "Bob",
    "C": "Bob",
    "D": "Jerrod",
    "E": "Jerrod",
    "F": "Jerrod",
    "G": "Ryan",
    "H": "Ryan",
    "I": "Jerrod",
    "J": "Jerrod",
    "K": "Jerrod",
    "L": "Jerrod",
    "M": "Jerrod",
    "N": "Ryan",
    "O": "Ryan",
    "P": "Bob",
    "Q": "Bob",
    "R": "Bob",
    "S": "Bob",
    "T": "Jerrod",
    "U": "Bob",
    "V": "Ryan",
    "W": "Ryan",
    "X": "Jerrod",
    "Y": "Ryan",
    "Z": "Ryan",
}

# FOM for each domestic_code (DOM tab grain), derived from DISTRICT_FOM_MAP: every
# district_code grouped under a given domestic_code (21/22/23) belongs to the same
# single FOM, verified against DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW.
# First name only, per request.
DOM_FOM_MAP = {
    "21": "Bob",
    "22": "Jerrod",
    "23": "Ryan",
}

# Fixed metric column order applied after each tab's identification columns:
# NPS, FFV %, then counts (Advocates/Neutrals/Detractors/Total), then the FFV
# Passed count. The Advocate/Neutral/Detractor SHARE PERCENT columns are
# deliberately not included here (removed per request) -- they're still
# computed internally (e.g. to derive NPS and the Total row), just not shown
# as their own output columns. Every name here must exist in the raw query's
# own column list (case-insensitive).
METRIC_COLUMNS = [
    "NPS", "FFV_PCT",
    "ADVOCATES", "NEUTRALS", "DETRACTORS", "TOTAL",
    "FFV_PASSED",
]

# Simplified, human-friendly header labels shown in the workbook, keyed by the
# INTERNAL column name used everywhere else in this script (color-scale lookups,
# prior-month matching, etc. all keep using the internal names -- only the
# header row text shown to readers is affected). Any column not listed here
# falls back to its internal name as-is.
HEADER_LABELS = {
    "BUSINESS_CENTER_DESC": "Business Center",
    "DISTRICT_CODE": "District",
    "SERVICE_MANAGER": "Area Manager",
    "AM_NAME": "Area Manager",
    "DOM": "DOM",
    "FOM_NAME": "FOM",
    "TA_DISTRICT_CODE": "TA District",
    "TA_NAME": "Technical Advisor",
    "DEALER_CODE": "Dealer Code",
    "DEALER_NAME": "Dealer Name",
    "NPS": "NPS",
    "FFV_PCT": "FFV %",
    "ADVOCATES": "Advocates",
    "NEUTRALS": "Neutrals",
    "DETRACTORS": "Detractors",
    "TOTAL": "Total",
    "FFV_PASSED": "FFV Passed",
    "PRIOR_MONTH_NPS": "Prior Month NPS",
    "PRIOR_MONTH_FFV_PCT": "Prior Month FFV %",
}


def build_tab(raw_columns, raw_rows, id_specs):
    """Build a tab's final (columns, rows) in the fixed layout: the tab's
    identification columns (id_specs), followed by METRIC_COLUMNS pulled
    straight from the raw query's own columns/rows.

    id_specs is a list of (column_name, value_fn) pairs, where value_fn(row, idx)
    computes that identification column's value for a raw row -- either read
    directly off the row (via idx, the {lower_column_name: position} map for
    raw_columns) or looked up from one of the hardcoded maps above."""
    idx = {col.lower(): i for i, col in enumerate(raw_columns)}
    new_columns = [name for name, _ in id_specs] + list(METRIC_COLUMNS)
    new_rows = []
    for row in raw_rows:
        id_values = [fn(row, idx) for _, fn in id_specs]
        metric_values = [row[idx[m.lower()]] for m in METRIC_COLUMNS]
        new_rows.append(id_values + metric_values)
    return new_columns, new_rows


def build_total_row(columns, rows, label="Total"):
    idx = {col.lower(): i for i, col in enumerate(columns)}
    total_row = [None] * len(columns)
    total_row[0] = label

    advocates = sum(row[idx["advocates"]] for row in rows) if "advocates" in idx else None
    neutrals = sum(row[idx["neutrals"]] for row in rows) if "neutrals" in idx else None
    detractors = sum(row[idx["detractors"]] for row in rows) if "detractors" in idx else None
    total = advocates + neutrals + detractors

    if "advocates" in idx:
        total_row[idx["advocates"]] = advocates
    if "neutrals" in idx:
        total_row[idx["neutrals"]] = neutrals
    if "detractors" in idx:
        total_row[idx["detractors"]] = detractors
    if "total" in idx:
        total_row[idx["total"]] = total

    advocateshare = advocates / total
    neutralshare = neutrals / total
    detractorshare = detractors / total

    if "advocateshare" in idx:
        total_row[idx["advocateshare"]] = advocateshare
    if "neutralshare" in idx:
        total_row[idx["neutralshare"]] = neutralshare
    if "detractorshare" in idx:
        total_row[idx["detractorshare"]] = detractorshare
    if "nps" in idx:
        total_row[idx["nps"]] = round((advocateshare - detractorshare) * 100, 1)

    if "ffv_passed" in idx and "ffv_total" in idx:
        ffv_passed = sum(row[idx["ffv_passed"]] or 0 for row in rows)
        ffv_total = sum(row[idx["ffv_total"]] or 0 for row in rows)
        total_row[idx["ffv_passed"]] = ffv_passed
        total_row[idx["ffv_total"]] = ffv_total
        if "ffv_pct" in idx:
            total_row[idx["ffv_pct"]] = round(ffv_passed / ffv_total * 100, 1) if ffv_total else None

    return total_row


def build_prior_month_map(columns, rows, key_cols):
    """Build a dict mapping each key_cols tuple -> (nps, ffv_pct) from a tab's
    already-built (post build_tab) PRIOR MONTH columns/rows, for use by
    append_prior_month_values. The prior month's NPS/FFV % come directly from
    re-running this tab's query over PREV_START_DATE/PREV_END_DATE, not from a
    previously saved report file."""
    idx = {col.lower(): i for i, col in enumerate(columns)}
    key_idx = [idx[k.lower()] for k in key_cols]
    nps_idx = idx["nps"]
    ffv_pct_idx = idx["ffv_pct"]
    result = {}
    for row in rows:
        key = tuple(row[i] for i in key_idx)
        result[key] = (row[nps_idx], row[ffv_pct_idx])
    return result


def append_prior_month_values(columns, rows, key_cols, prior_month_map):
    """Add PRIOR_MONTH_NPS and PRIOR_MONTH_FFV_PCT columns using prior_month_map
    (built by build_prior_month_map from this run's own prior-month query),
    matched by key_cols. Rows with no match in prior_month_map (e.g. a
    brand-new dealer) get None/blank for both."""
    idx = {col.lower(): i for i, col in enumerate(columns)}
    key_idx = [idx[k.lower()] for k in key_cols]
    matched = 0
    new_rows = []
    for row in rows:
        row = list(row)
        key = tuple(row[i] for i in key_idx)
        last_nps, last_ffv_pct = prior_month_map.get(key, (None, None))
        if key in prior_month_map:
            matched += 1

        last_nps = float(last_nps) if last_nps is not None else None
        last_ffv_pct = float(last_ffv_pct) if last_ffv_pct is not None else None

        row.append(last_nps)
        row.append(last_ffv_pct)
        new_rows.append(row)

    print(f"  Matched {matched}/{len(rows)} rows against prior month (key_cols={key_cols})")
    new_columns = list(columns) + ["PRIOR_MONTH_NPS", "PRIOR_MONTH_FFV_PCT"]
    return new_columns, new_rows


def query_business_center(start_date, end_date, label):
    """By Business Center (with weighted-average FFV%)."""
    cur.execute(f"""
        WITH base AS (
            SELECT *
            FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW
            WHERE market = 'USA'
            AND {BRAND_FILTER}
            AND date_survey_received BETWEEN '{start_date}' AND '{end_date}'
        ),
        agg AS (
            SELECT
                business_center_desc,
                sum(nps_advocate) as advocates_sum,
                sum(nps_neutral) as neutrals_sum,
                sum(nps_detractor) as detractors_sum,
                sum(ffv_passed) as ffv_passed_sum,
                sum(ffv_eligible_count) as ffv_total_sum
            FROM base
            GROUP BY ALL
        )
        SELECT
            business_center_desc,
            advocates_sum as advocates,
            neutrals_sum as neutrals,
            detractors_sum as detractors,
            advocates + neutrals + detractors as total,
            advocates / total as advocateshare,
            neutrals / total as neutralshare,
            detractors / total as detractorshare,
            round((advocateshare - detractorshare) * 100, 1) as NPS,
            ffv_passed_sum as ffv_passed,
            ffv_total_sum as ffv_total,
            round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct
        FROM agg
        ORDER BY NPS DESC
    """)
    rows = cur.fetchall()
    columns = [desc[0] for desc in cur.description]
    print(f"Business Center query ({label}): {len(rows)} rows")
    return columns, rows


def query_district(start_date, end_date, label):
    """By District Code (with weighted-average FFV%)."""
    cur.execute(f"""
        WITH base AS (
            SELECT *
            FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW
            WHERE market = 'USA'
            AND {BRAND_FILTER}
            AND business_center_desc = 'West'
            AND date_survey_received BETWEEN '{start_date}' AND '{end_date}'
        ),
        agg AS (
            SELECT
                district_code,
                domestic_code as "DOM",
                sum(nps_advocate) as advocates_sum,
                sum(nps_neutral) as neutrals_sum,
                sum(nps_detractor) as detractors_sum,
                sum(ffv_passed) as ffv_passed_sum,
                sum(ffv_eligible_count) as ffv_total_sum
            FROM base
            GROUP BY ALL
        )
        SELECT
            district_code,
            "DOM",
            advocates_sum as advocates,
            neutrals_sum as neutrals,
            detractors_sum as detractors,
            advocates + neutrals + detractors as total,
            advocates / total as advocateshare,
            neutrals / total as neutralshare,
            detractors / total as detractorshare,
            round((advocateshare - detractorshare) * 100, 1) as NPS,
            ffv_passed_sum as ffv_passed,
            ffv_total_sum as ffv_total,
            round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct
        FROM agg
        ORDER BY NPS DESC
    """)
    rows = cur.fetchall()
    columns = [desc[0] for desc in cur.description]
    print(f"District query ({label}): {len(rows)} rows")
    return columns, rows


def query_ta(start_date, end_date, label):
    """By TA District Code (same metrics as the District tab, but grouped by
    the Technical Advisor district instead of the sales district). ta_district_code
    comes from DLR_DOM.DLR_DIST_TECHADV -- the Technical Advisor district flag
    assigned to each dealer (same field used for the Dealer tab's ta_district_code
    column) -- joined in here so dealer-grain surveys can be aggregated by it.
    Sorted by FFV% (not NPS), since FFV is the metric highlighted on this tab."""
    cur.execute(f"""
        WITH base AS (
            SELECT s.*, d.dlr_dist_techadv AS ta_district_code
            FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW s
            LEFT JOIN DATALAKE_PROD.CUSTOMER_EXPERIENCE.DLR_DOM d ON s.dealer_code = d.c_dealer
            WHERE s.market = 'USA'
            AND {BRAND_FILTER}
            AND s.business_center_desc = 'West'
            AND s.date_survey_received BETWEEN '{start_date}' AND '{end_date}'
        ),
        agg AS (
            SELECT
                ta_district_code,
                sum(nps_advocate) as advocates_sum,
                sum(nps_neutral) as neutrals_sum,
                sum(nps_detractor) as detractors_sum,
                sum(ffv_passed) as ffv_passed_sum,
                sum(ffv_eligible_count) as ffv_total_sum
            FROM base
            GROUP BY ALL
        )
        SELECT
            ta_district_code,
            advocates_sum as advocates,
            neutrals_sum as neutrals,
            detractors_sum as detractors,
            advocates + neutrals + detractors as total,
            advocates / total as advocateshare,
            neutrals / total as neutralshare,
            detractors / total as detractorshare,
            round((advocateshare - detractorshare) * 100, 1) as NPS,
            ffv_passed_sum as ffv_passed,
            ffv_total_sum as ffv_total,
            round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct
        FROM agg
        ORDER BY ffv_pct DESC
    """)
    rows = cur.fetchall()
    columns = [desc[0] for desc in cur.description]
    print(f"TA District query ({label}): {len(rows)} rows")
    return columns, rows


def query_dealer(start_date, end_date, label):
    """By Dealer Code (dealer-grain). ta_district_code comes from
    DLR_DOM.DLR_DIST_TECHADV -- the Technical Advisor district flag assigned to
    each dealer (same field used to join to DLR_DOM_DIST dist_typ='D' in the TA
    report) -- joined in here purely for display on this tab."""
    cur.execute(f"""
        WITH base AS (
            SELECT s.*, d.dlr_dist_techadv AS ta_district_code
            FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW s
            LEFT JOIN DATALAKE_PROD.CUSTOMER_EXPERIENCE.DLR_DOM d ON s.dealer_code = d.c_dealer
            WHERE s.market = 'USA'
            AND {BRAND_FILTER}
            AND s.business_center_desc = 'West'
            AND s.date_survey_received BETWEEN '{start_date}' AND '{end_date}'
        ),
        agg AS (
            SELECT
                dealer_code,
                dealer_name,
                district_code,
                domestic_code as "DOM",
                ta_district_code,
                sum(nps_advocate) as advocates_sum,
                sum(nps_neutral) as neutrals_sum,
                sum(nps_detractor) as detractors_sum,
                sum(ffv_passed) as ffv_passed_sum,
                sum(ffv_eligible_count) as ffv_total_sum
            FROM base
            GROUP BY ALL
        )
        SELECT
            dealer_code,
            dealer_name,
            district_code,
            "DOM",
            ta_district_code,
            advocates_sum as advocates,
            neutrals_sum as neutrals,
            detractors_sum as detractors,
            advocates + neutrals + detractors as total,
            advocates / total as advocateshare,
            neutrals / total as neutralshare,
            detractors / total as detractorshare,
            round((advocateshare - detractorshare) * 100, 1) as NPS,
            ffv_passed_sum as ffv_passed,
            ffv_total_sum as ffv_total,
            round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct
        FROM agg
        ORDER BY NPS DESC
    """)
    rows = cur.fetchall()
    columns = [desc[0] for desc in cur.description]
    print(f"Dealer query ({label}): {len(rows)} rows")
    return columns, rows


def query_dom(start_date, end_date, label):
    """By Domestic Code (DOM) (with weighted-average FFV%)."""
    cur.execute(f"""
        WITH base AS (
            SELECT *
            FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW
            WHERE market = 'USA'
            AND {BRAND_FILTER}
            AND business_center_desc = 'West'
            AND domestic_code in (21,22,23)
            AND date_survey_received BETWEEN '{start_date}' AND '{end_date}'
        ),
        agg AS (
            SELECT
                domestic_code as "DOM",
                sum(nps_advocate) as advocates_sum,
                sum(nps_neutral) as neutrals_sum,
                sum(nps_detractor) as detractors_sum,
                sum(ffv_passed) as ffv_passed_sum,
                sum(ffv_eligible_count) as ffv_total_sum
            FROM base
            GROUP BY ALL
        )
        SELECT
            "DOM",
            advocates_sum as advocates,
            neutrals_sum as neutrals,
            detractors_sum as detractors,
            advocates + neutrals + detractors as total,
            advocates / total as advocateshare,
            neutrals / total as neutralshare,
            detractors / total as detractorshare,
            round((advocateshare - detractorshare) * 100, 1) as NPS,
            ffv_passed_sum as ffv_passed,
            ffv_total_sum as ffv_total,
            round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct
        FROM agg
        ORDER BY NPS DESC
    """)
    rows = cur.fetchall()
    columns = [desc[0] for desc in cur.description]
    print(f"DOM query ({label}): {len(rows)} rows")
    return columns, rows


# Run every query twice: once for the current reporting window, once for the
# prior-month window, so Prior Month NPS/FFV% come straight from Snowflake.
bc_columns, bc_rows = query_business_center(START_DATE, END_DATE, "current")
bc_prev_columns, bc_prev_rows = query_business_center(PREV_START_DATE, PREV_END_DATE, "prior month")

dist_columns, dist_rows = query_district(START_DATE, END_DATE, "current")
dist_prev_columns, dist_prev_rows = query_district(PREV_START_DATE, PREV_END_DATE, "prior month")

ta_columns, ta_rows = query_ta(START_DATE, END_DATE, "current")
ta_prev_columns, ta_prev_rows = query_ta(PREV_START_DATE, PREV_END_DATE, "prior month")

dealer_columns, dealer_rows = query_dealer(START_DATE, END_DATE, "current")
dealer_prev_columns, dealer_prev_rows = query_dealer(PREV_START_DATE, PREV_END_DATE, "prior month")

dom_columns, dom_rows = query_dom(START_DATE, END_DATE, "current")
dom_prev_columns, dom_prev_rows = query_dom(PREV_START_DATE, PREV_END_DATE, "prior month")

cur.close()
conn.close()

# --- Step 1b: Compute the Total row (Business Center only, current + prior
# month), then build each tab's final (columns, rows) in the fixed ID +
# NPS/FFV%/counts/FFV-count layout -- for both the current and prior-month
# datasets, using the same identification columns (id_specs) each time. ---
bc_total_row = build_total_row(bc_columns, bc_rows)
bc_rows_with_total = bc_rows + [bc_total_row]
bc_prev_total_row = build_total_row(bc_prev_columns, bc_prev_rows)
bc_prev_rows_with_total = bc_prev_rows + [bc_prev_total_row]

BC_ID_SPECS = [("BUSINESS_CENTER_DESC", lambda row, idx: row[idx["business_center_desc"]])]
bc_columns, bc_rows = build_tab(bc_columns, bc_rows_with_total, BC_ID_SPECS)
bc_prev_columns, bc_prev_rows = build_tab(bc_prev_columns, bc_prev_rows_with_total, BC_ID_SPECS)

DIST_ID_SPECS = [
    ("DISTRICT_CODE", lambda row, idx: row[idx["district_code"]]),
    ("SERVICE_MANAGER", lambda row, idx: DISTRICT_SERVICE_MANAGER_MAP.get(row[idx["district_code"]])),
    ("DOM", lambda row, idx: row[idx["dom"]]),
    ("FOM_NAME", lambda row, idx: DISTRICT_FOM_MAP.get(row[idx["district_code"]])),
]
dist_columns, dist_rows = build_tab(dist_columns, dist_rows, DIST_ID_SPECS)
dist_prev_columns, dist_prev_rows = build_tab(dist_prev_columns, dist_prev_rows, DIST_ID_SPECS)

TA_ID_SPECS = [
    ("TA_DISTRICT_CODE", lambda row, idx: row[idx["ta_district_code"]]),
    ("TA_NAME", lambda row, idx: TA_DISTRICT_NAME_MAP.get(row[idx["ta_district_code"]])),
]
ta_columns, ta_rows = build_tab(ta_columns, ta_rows, TA_ID_SPECS)
ta_prev_columns, ta_prev_rows = build_tab(ta_prev_columns, ta_prev_rows, TA_ID_SPECS)

DEALER_ID_SPECS = [
    ("DEALER_CODE", lambda row, idx: row[idx["dealer_code"]]),
    ("DEALER_NAME", lambda row, idx: row[idx["dealer_name"]]),
    ("DISTRICT_CODE", lambda row, idx: row[idx["district_code"]]),
    ("AM_NAME", lambda row, idx: DISTRICT_SERVICE_MANAGER_MAP.get(row[idx["district_code"]])),
    ("DOM", lambda row, idx: row[idx["dom"]]),
    ("FOM_NAME", lambda row, idx: DISTRICT_FOM_MAP.get(row[idx["district_code"]])),
    ("TA_DISTRICT_CODE", lambda row, idx: row[idx["ta_district_code"]]),
]
dealer_columns, dealer_rows = build_tab(dealer_columns, dealer_rows, DEALER_ID_SPECS)
dealer_prev_columns, dealer_prev_rows = build_tab(dealer_prev_columns, dealer_prev_rows, DEALER_ID_SPECS)

DOM_ID_SPECS = [
    ("DOM", lambda row, idx: row[idx["dom"]]),
    ("FOM_NAME", lambda row, idx: DOM_FOM_MAP.get(str(row[idx["dom"]]).strip())),
]
dom_columns, dom_rows = build_tab(dom_columns, dom_rows, DOM_ID_SPECS)
dom_prev_columns, dom_prev_rows = build_tab(dom_prev_columns, dom_prev_rows, DOM_ID_SPECS)

# bc_rows now ends with the Total row -- pull it back out so it can be appended
# after the Prior Month columns are computed (matching DOD8.py's behavior).
bc_total_row = bc_rows[-1]
bc_rows = bc_rows[:-1]

# --- Step 1c: Build each tab's Prior Month NPS / Prior Month FFV % lookup
# directly from this run's own prior-month query results (bc_prev_rows etc.,
# already shaped into the same column layout via build_tab above), then append
# those two columns onto the current-period rows. ---
bc_prior_month_map = build_prior_month_map(bc_prev_columns, bc_prev_rows, ["BUSINESS_CENTER_DESC"])
dist_prior_month_map = build_prior_month_map(dist_prev_columns, dist_prev_rows, ["DISTRICT_CODE", "DOM"])
ta_prior_month_map = build_prior_month_map(ta_prev_columns, ta_prev_rows, ["TA_DISTRICT_CODE"])
dealer_prior_month_map = build_prior_month_map(dealer_prev_columns, dealer_prev_rows, ["DEALER_CODE"])
dom_prior_month_map = build_prior_month_map(dom_prev_columns, dom_prev_rows, ["DOM"])
print(f"  Business Center: {len(bc_prior_month_map)} keys found in prior month")
print(f"  District: {len(dist_prior_month_map)} keys found in prior month")
print(f"  TAs: {len(ta_prior_month_map)} keys found in prior month")
print(f"  Dealer: {len(dealer_prior_month_map)} keys found in prior month")
print(f"  DOM: {len(dom_prior_month_map)} keys found in prior month")

bc_columns, bc_all_rows = append_prior_month_values(bc_columns, bc_rows + [bc_total_row], ["BUSINESS_CENTER_DESC"], bc_prior_month_map)
bc_rows, bc_total_row = bc_all_rows[:-1], bc_all_rows[-1]

dist_columns, dist_rows = append_prior_month_values(dist_columns, dist_rows, ["DISTRICT_CODE", "DOM"], dist_prior_month_map)
ta_columns, ta_rows = append_prior_month_values(ta_columns, ta_rows, ["TA_DISTRICT_CODE"], ta_prior_month_map)
dealer_columns, dealer_rows = append_prior_month_values(dealer_columns, dealer_rows, ["DEALER_CODE"], dealer_prior_month_map)
dom_columns, dom_rows = append_prior_month_values(dom_columns, dom_rows, ["DOM"], dom_prior_month_map)

# --- Step 2: Build a brand-new workbook from scratch (no template) and write
# each tab's header row (dark blue fill, white bold text, simplified labels) +
# data rows, with thin gridlines around every cell for a cleaner look. ---
HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WEST_HIGHLIGHT_FILL = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
THIN_SIDE = Side(style="thin", color="000000")
GRID_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)


def write_sheet(wb, title, columns, rows):
    ws = wb.create_sheet(title=title)
    ws.append([HEADER_LABELS.get(col, col) for col in columns])
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    for row in rows:
        ws.append(list(row))
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = GRID_BORDER
    return ws


def highlight_row_by_value(ws, column_index, match_value, fill):
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if row[column_index - 1].value == match_value:
            for cell in row:
                cell.fill = fill


def apply_three_color_bands(ws, columns, column_name, red_max, green_min):
    """Apply a discrete 3-color band (solid red / solid yellow / solid green --
    NOT a smooth gradient) to column_name's data rows (row 2 through the
    sheet's current last row), located by position in columns (this run's own
    column list for the tab):
      - value < red_max               -> solid red
      - red_max <= value <= green_min -> solid yellow
      - value > green_min             -> solid green
    Using three separate CellIsRule conditional formats (each with a flat
    PatternFill) rather than Excel's native color-scale gradient means green
    appears immediately and fully once a value crosses green_min, instead of
    fading in gradually toward some far-off maximum (which is what made values
    just above the threshold still look yellow under the old gradient)."""
    idx = {col.lower(): i for i, col in enumerate(columns)}
    col_letter = get_column_letter(idx[column_name.lower()] + 1)
    last_row = ws.max_row
    if last_row < 2:
        return
    rng = f"{col_letter}2:{col_letter}{last_row}"
    red_fill = PatternFill(start_color="F8696B", end_color="F8696B", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB84", end_color="FFEB84", fill_type="solid")
    green_fill = PatternFill(start_color="63BE7B", end_color="63BE7B", fill_type="solid")
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=[str(red_max)], fill=red_fill))
    ws.conditional_formatting.add(rng, CellIsRule(operator="between", formula=[str(red_max), str(green_min)], fill=yellow_fill))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=[str(green_min)], fill=green_fill))


def get_versioned_path(path):
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    version = 2
    while True:
        candidate = f"{root} v{version}{ext}"
        if not os.path.exists(candidate):
            return candidate
        version += 1


wb = openpyxl.Workbook()
wb.remove(wb.active)  # drop the default blank sheet; we create our own five below

ws_bc = write_sheet(wb, "Business Center", bc_columns, bc_rows + [bc_total_row])
ws_dist = write_sheet(wb, "District", dist_columns, dist_rows)
ws_ta = write_sheet(wb, "TAs", ta_columns, ta_rows)
ws_dealer = write_sheet(wb, "Dealer", dealer_columns, dealer_rows)
ws_dom = write_sheet(wb, "DOM", dom_columns, dom_rows)

# NPS color bands: solid red below 70 (69.99 and below), solid yellow 70-79.99,
# solid green at/above 80 -- applies to every tab's NPS column. FFV % gets its
# own bands (solid red below 80, solid yellow 80-92.49, solid green at/above
# 92.5) on the DOM, Dealer, District, and TAs tabs. Business Center keeps NPS
# only, as before. All tabs mirror the same thresholds/styling.
apply_three_color_bands(ws_bc, bc_columns, "nps", 70, 79.99)

apply_three_color_bands(ws_dist, dist_columns, "nps", 70, 79.99)
apply_three_color_bands(ws_dist, dist_columns, "ffv_pct", 80, 92.49)

apply_three_color_bands(ws_ta, ta_columns, "nps", 70, 79.99)
apply_three_color_bands(ws_ta, ta_columns, "ffv_pct", 80, 92.49)

apply_three_color_bands(ws_dealer, dealer_columns, "nps", 70, 79.99)
apply_three_color_bands(ws_dealer, dealer_columns, "ffv_pct", 80, 92.49)

apply_three_color_bands(ws_dom, dom_columns, "nps", 70, 79.99)
apply_three_color_bands(ws_dom, dom_columns, "ffv_pct", 80, 92.49)

# Highlight the West row on the Business Center tab
highlight_row_by_value(ws_bc, 1, "West", WEST_HIGHLIGHT_FILL)

# Save directly as the dated file (versioned if one already exists today) -- this
# IS the only saved file now; there's no separate Base File to also overwrite.
today_str = date.today().strftime("%Y%m%d")
DATED_COPY_PATH = get_versioned_path(
    os.path.join(REPORT_FOLDER, f"NPS report West {today_str}.xlsx")
)
wb.save(DATED_COPY_PATH)
print(f"Saved report to {DATED_COPY_PATH}")

# --- Step 3: Email the file via Outlook ---
RECIPIENT = "JesseTest"

# Tabs to paste into the email body, in order. Each is copied directly out of
# the saved workbook via Excel COM (not hand-built from the Python rows), so the
# email shows exactly what Excel is currently rendering: the NPS/FFV% color
# scales, the highlighted West row, number formats, etc. A hand-built HTML table
# can't reproduce Excel's computed conditional formatting, since that's a live
# rule, not a value - only Excel itself knows the resulting color.
EMAIL_SHEET_NAMES = ["Business Center", "District", "TAs", "DOM"]


def com_retry(fn, attempts=6, delay=1.0):
    """Retry a COM call a few times with a short pause between attempts.

    Excel/Outlook/OneDrive sometimes reject or block a call transiently:
    pywintypes.com_error (e.g. RPC_E_CALL_REJECTED / 'Call was rejected by
    callee' while the target application is still starting up or busy), or a
    sharing-violation com_error (e.g. "We can't open '...'. It's possible the
    file is already open...") while OneDrive is mid-sync on a just-saved file.
    A brief retry loop papers over both races instead of failing the whole
    run."""
    last_err = None
    for attempt in range(1, attempts + 1):
        try:
            return fn()
        except pywintypes.com_error as e:
            last_err = e
            print(f"  COM call rejected (attempt {attempt}/{attempts}): {e}")
            time.sleep(delay)
    raise last_err


# Using plain win32com.client.Dispatch (dynamic dispatch) rather than
# gencache.EnsureDispatch: EnsureDispatch calls GetTypeInfo() on the Excel COM
# object immediately to build its cached type-library wrapper, and on this
# machine that call itself gets rejected ("Call was rejected by callee") --
# typically because Excel's RPC server isn't ready to respond yet (e.g. right
# after launch) or is still busy from a prior run. Dynamic dispatch defers
# member resolution, avoiding that upfront call, and DispatchEx forces a brand
# new Excel process instead of silently attaching to a possibly-stuck existing
# one. The actual property-set calls are wrapped in com_retry() for the same
# busy-server race.
excel_app = win32com.client.DispatchEx("Excel.Application")
com_retry(lambda: setattr(excel_app, "Visible", False))
com_retry(lambda: setattr(excel_app, "DisplayAlerts", False))
excel_wb = com_retry(lambda: excel_app.Workbooks.Open(os.path.abspath(DATED_COPY_PATH), ReadOnly=True))

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = RECIPIENT
mail.Subject = f"NPS Report {date.today().strftime('%m/%d/%Y')}"
mail.BodyFormat = 2  # olFormatHTML
mail.Display()  # Display() is required before GetInspector.WordEditor is available

word_doc = mail.GetInspector.WordEditor
word_app = word_doc.Application
sel = word_app.Selection

# Clear the default signature/body Outlook inserts on Display()
sel.WholeStory()
sel.Delete()

sel.TypeText("Hi Team,")
sel.TypeParagraph()
sel.TypeText("Please see todays NPS Report below and attached.")
sel.TypeParagraph()
sel.TypeParagraph()

for sheet_name in EMAIL_SHEET_NAMES:
    ws_excel = excel_wb.Sheets(sheet_name)
    last_row = ws_excel.UsedRange.Row + ws_excel.UsedRange.Rows.Count - 1
    last_col = ws_excel.UsedRange.Column + ws_excel.UsedRange.Columns.Count - 1
    full_range = ws_excel.Range(ws_excel.Cells(1, 1), ws_excel.Cells(last_row, last_col))
    # xlCellTypeVisible = 12 - skips any hidden columns automatically.
    visible_range = full_range.SpecialCells(12)
    visible_range.Copy()

    sel.Font.Bold = True
    sel.TypeText(sheet_name)
    sel.Font.Bold = False
    sel.TypeParagraph()
    sel.Paste()
    sel.EndKey(6)  # wdStory - move past the pasted table to the end of the document
    sel.TypeParagraph()
    sel.TypeParagraph()

sel.TypeText("Thanks")

excel_wb.Close(False)
excel_app.Quit()

mail.Attachments.Add(os.path.abspath(DATED_COPY_PATH))
mail.Send()

print(f"Email sent to {RECIPIENT}")
print("Done.")
