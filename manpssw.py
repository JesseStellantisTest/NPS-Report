# NPS report with five tabs (Business Center, District, Dealer, DOM, TAs) saved as Excel
# Filtered by an explicit start/end date range on date_survey_received (inclusive)
# Co-authored with CoCo

import snowflake.connector
import openpyxl
import os
import win32com.client
from datetime import date
from openpyxl.styles import PatternFill

# --- Step 1: Connect to Snowflake and run five NPS queries ---
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

# All five queries now run off MPR_CX_SERVICE_SURVEY_DETAILS_VW (one row per survey),
# filtering on the real per-survey BRAND_CODE column (A=Alfa, C=Chrysler, D=Dodge,
# F=Fiat, J=Jeep, R=Ram) instead of the dealer-level FRANCHISE string. The old
# FRANCHISE LIKE '%C/D/J/R%' filter against MPR_CX_SERVICE_SURVEY_SUMMARY_VW could
# only include/exclude a dealer wholesale -- it could not strip a CDJRF dealer's
# Fiat surveys (or a CDJRFA dealer's Alfa surveys) out of that dealer's pre-summed
# NPS totals, since SUMMARY_VW has already combined every brand a dealer sells into
# one row per dealer per month before any filter runs. Filtering brand_code at the
# survey level, before our own aggregation below, excludes those surveys correctly.
BRAND_FILTER = "brand_code IN ('C','D','J','R')"

# ta_district_code comes from DLR_SVWGEOS -- the Technical Advisor district flag
# assigned to each dealer -- joined in here so dealer-grain surveys can be
# aggregated by it. A dealer can have more than one active geo row on rare
# occasions, so we dedupe to the most recent active record per dealer before
# joining.
TA_DISTRICT_JOIN = """
    LEFT JOIN (
        SELECT c_dlr, c_dist_tech_advsr AS ta_district_code
        FROM DATALAKE_PROD.SCM.DLR_SVWGEOS
        WHERE c_actv = '1'
        QUALIFY ROW_NUMBER() OVER (PARTITION BY c_dlr ORDER BY d_actv DESC) = 1
    ) g ON s.dealer_code = g.c_dlr
"""

# Each tab's query now also computes a Maintenance NPS alongside the overall NPS:
# a second aggregation (maint_agg), filtered to SERVICE_CATEGORY = 'Maintenance'
# only (excludes 'Repair' and the combined 'Maintenance & Repair' category),
# left-joined back onto the overall aggregation by that tab's grouping key(s)
# via USING (so a grouping key with zero Maintenance surveys still shows up,
# with maintenance_* columns coalesced to 0/NULL instead of dropping the row).
MAINTENANCE_FILTER = "service_category = 'Maintenance'"

# Query 1: By Business Center (with weighted-average FFV% and Maintenance NPS)
cur.execute(f"""
    WITH base AS (
        SELECT *
        FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW
        WHERE market = 'USA'
          AND {BRAND_FILTER}
          AND date_survey_received BETWEEN '{START_DATE}' AND '{END_DATE}'
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
    ),
    maint_agg AS (
        SELECT
            business_center_desc,
            sum(nps_advocate) as m_advocates_sum,
            sum(nps_neutral) as m_neutrals_sum,
            sum(nps_detractor) as m_detractors_sum
        FROM base
        WHERE {MAINTENANCE_FILTER}
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
        round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct,
        coalesce(m_advocates_sum, 0) as maintenance_advocates,
        coalesce(m_neutrals_sum, 0) as maintenance_neutrals,
        coalesce(m_detractors_sum, 0) as maintenance_detractors,
        maintenance_advocates + maintenance_neutrals + maintenance_detractors as maintenance_total,
        maintenance_advocates / nullif(maintenance_total, 0) as maintenance_advocateshare,
        maintenance_neutrals / nullif(maintenance_total, 0) as maintenance_neutralshare,
        maintenance_detractors / nullif(maintenance_total, 0) as maintenance_detractorshare,
        round((maintenance_advocateshare - maintenance_detractorshare) * 100, 1) as MAINTENANCE_NPS
    FROM agg
    LEFT JOIN maint_agg USING (business_center_desc)
    ORDER BY NPS DESC
""")
bc_rows = cur.fetchall()
bc_columns = [desc[0] for desc in cur.description]
print(f"Business Center query: {len(bc_rows)} rows")
print(f"  Header: {bc_columns}")
for row in bc_rows[:5]:
    print(f"  {row}")

# Query 2: By District Code (with weighted-average FFV% and Maintenance NPS)
cur.execute(f"""
    WITH base AS (
        SELECT *
        FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW
        WHERE market = 'USA'
          AND {BRAND_FILTER}
          AND business_center_desc = 'Southwest'
          AND date_survey_received BETWEEN '{START_DATE}' AND '{END_DATE}'
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
    ),
    maint_agg AS (
        SELECT
            district_code,
            domestic_code as "DOM",
            sum(nps_advocate) as m_advocates_sum,
            sum(nps_neutral) as m_neutrals_sum,
            sum(nps_detractor) as m_detractors_sum
        FROM base
        WHERE {MAINTENANCE_FILTER}
        GROUP BY ALL
    )
    SELECT
        district_code, "DOM",
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
        round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct,
        coalesce(m_advocates_sum, 0) as maintenance_advocates,
        coalesce(m_neutrals_sum, 0) as maintenance_neutrals,
        coalesce(m_detractors_sum, 0) as maintenance_detractors,
        maintenance_advocates + maintenance_neutrals + maintenance_detractors as maintenance_total,
        maintenance_advocates / nullif(maintenance_total, 0) as maintenance_advocateshare,
        maintenance_neutrals / nullif(maintenance_total, 0) as maintenance_neutralshare,
        maintenance_detractors / nullif(maintenance_total, 0) as maintenance_detractorshare,
        round((maintenance_advocateshare - maintenance_detractorshare) * 100, 1) as MAINTENANCE_NPS
    FROM agg
    LEFT JOIN maint_agg USING (district_code, "DOM")
    ORDER BY NPS DESC
""")
dist_rows = cur.fetchall()
dist_columns = [desc[0] for desc in cur.description]
print(f"District query: {len(dist_rows)} rows")
print(f"  Header: {dist_columns}")
for row in dist_rows[:5]:
    print(f"  {row}")

# Query 3: By Dealer Code (dealer-grain, with Maintenance NPS)
cur.execute(f"""
    WITH base AS (
        SELECT s.*, g.ta_district_code
        FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW s
        {TA_DISTRICT_JOIN}
        WHERE s.market = 'USA'
          AND {BRAND_FILTER}
          AND s.business_center_desc = 'Southwest'
          AND s.date_survey_received BETWEEN '{START_DATE}' AND '{END_DATE}'
    ),
    agg AS (
        SELECT
            dealer_code, dealer_name, district_code, domestic_code as "DOM", ta_district_code,
            sum(nps_advocate) as advocates_sum,
            sum(nps_neutral) as neutrals_sum,
            sum(nps_detractor) as detractors_sum,
            sum(ffv_passed) as ffv_passed_sum,
            sum(ffv_eligible_count) as ffv_total_sum
        FROM base
        GROUP BY ALL
    ),
    maint_agg AS (
        SELECT
            dealer_code,
            sum(nps_advocate) as m_advocates_sum,
            sum(nps_neutral) as m_neutrals_sum,
            sum(nps_detractor) as m_detractors_sum
        FROM base
        WHERE {MAINTENANCE_FILTER}
        GROUP BY ALL
    )
    SELECT
        dealer_code, dealer_name, district_code, "DOM", ta_district_code,
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
        round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct,
        coalesce(m_advocates_sum, 0) as maintenance_advocates,
        coalesce(m_neutrals_sum, 0) as maintenance_neutrals,
        coalesce(m_detractors_sum, 0) as maintenance_detractors,
        maintenance_advocates + maintenance_neutrals + maintenance_detractors as maintenance_total,
        maintenance_advocates / nullif(maintenance_total, 0) as maintenance_advocateshare,
        maintenance_neutrals / nullif(maintenance_total, 0) as maintenance_neutralshare,
        maintenance_detractors / nullif(maintenance_total, 0) as maintenance_detractorshare,
        round((maintenance_advocateshare - maintenance_detractorshare) * 100, 1) as MAINTENANCE_NPS
    FROM agg
    LEFT JOIN maint_agg USING (dealer_code)
    ORDER BY NPS DESC
""")
dealer_rows = cur.fetchall()
dealer_columns = [desc[0] for desc in cur.description]
print(f"Dealer query: {len(dealer_rows)} rows")
print(f"  Header: {dealer_columns}")
for row in dealer_rows[:5]:
    print(f"  {row}")

# Query 4: By Domestic Code (DOM) (with weighted-average FFV% and Maintenance NPS)
cur.execute(f"""
    WITH base AS (
        SELECT *
        FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW
        WHERE market = 'USA'
          AND {BRAND_FILTER}
          AND business_center_desc = 'Southwest'
          AND domestic_code in (21,22,23)
          AND date_survey_received BETWEEN '{START_DATE}' AND '{END_DATE}'
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
    ),
    maint_agg AS (
        SELECT
            domestic_code as "DOM",
            sum(nps_advocate) as m_advocates_sum,
            sum(nps_neutral) as m_neutrals_sum,
            sum(nps_detractor) as m_detractors_sum
        FROM base
        WHERE {MAINTENANCE_FILTER}
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
        round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct,
        coalesce(m_advocates_sum, 0) as maintenance_advocates,
        coalesce(m_neutrals_sum, 0) as maintenance_neutrals,
        coalesce(m_detractors_sum, 0) as maintenance_detractors,
        maintenance_advocates + maintenance_neutrals + maintenance_detractors as maintenance_total,
        maintenance_advocates / nullif(maintenance_total, 0) as maintenance_advocateshare,
        maintenance_neutrals / nullif(maintenance_total, 0) as maintenance_neutralshare,
        maintenance_detractors / nullif(maintenance_total, 0) as maintenance_detractorshare,
        round((maintenance_advocateshare - maintenance_detractorshare) * 100, 1) as MAINTENANCE_NPS
    FROM agg
    LEFT JOIN maint_agg USING ("DOM")
    ORDER BY NPS DESC
""")
dom_rows = cur.fetchall()
dom_columns = [desc[0] for desc in cur.description]
print(f"DOM query: {len(dom_rows)} rows")
print(f"  Header: {dom_columns}")
for row in dom_rows[:5]:
    print(f"  {row}")

# Query 5: By TA District Code (with weighted-average FFV% and Maintenance NPS)
cur.execute(f"""
    WITH base AS (
        SELECT s.*, g.ta_district_code
        FROM DATALAKE_PROD.AFTERSALES.MPR_CX_SERVICE_SURVEY_DETAILS_VW s
        {TA_DISTRICT_JOIN}
        WHERE s.market = 'USA'
          AND {BRAND_FILTER}
          AND s.business_center_desc = 'Southwest'
          AND s.date_survey_received BETWEEN '{START_DATE}' AND '{END_DATE}'
    ),
    agg AS (
        SELECT
            ta_district_code,
            domestic_code as "DOM",
            sum(nps_advocate) as advocates_sum,
            sum(nps_neutral) as neutrals_sum,
            sum(nps_detractor) as detractors_sum,
            sum(ffv_passed) as ffv_passed_sum,
            sum(ffv_eligible_count) as ffv_total_sum
        FROM base
        GROUP BY ALL
    ),
    maint_agg AS (
        SELECT
            ta_district_code,
            domestic_code as "DOM",
            sum(nps_advocate) as m_advocates_sum,
            sum(nps_neutral) as m_neutrals_sum,
            sum(nps_detractor) as m_detractors_sum
        FROM base
        WHERE {MAINTENANCE_FILTER}
        GROUP BY ALL
    )
    SELECT
        ta_district_code, "DOM",
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
        round(ffv_passed / nullif(ffv_total, 0) * 100, 1) as ffv_pct,
        coalesce(m_advocates_sum, 0) as maintenance_advocates,
        coalesce(m_neutrals_sum, 0) as maintenance_neutrals,
        coalesce(m_detractors_sum, 0) as maintenance_detractors,
        maintenance_advocates + maintenance_neutrals + maintenance_detractors as maintenance_total,
        maintenance_advocates / nullif(maintenance_total, 0) as maintenance_advocateshare,
        maintenance_neutrals / nullif(maintenance_total, 0) as maintenance_neutralshare,
        maintenance_detractors / nullif(maintenance_total, 0) as maintenance_detractorshare,
        round((maintenance_advocateshare - maintenance_detractorshare) * 100, 1) as MAINTENANCE_NPS
    FROM agg
    LEFT JOIN maint_agg USING (ta_district_code, "DOM")
    ORDER BY NPS DESC
""")
ta_dist_rows = cur.fetchall()
ta_dist_columns = [desc[0] for desc in cur.description]
print(f"TA District query: {len(ta_dist_rows)} rows")
print(f"  Header: {ta_dist_columns}")
for row in ta_dist_rows[:5]:
    print(f"  {row}")

cur.close()
conn.close()

# --- Step 2: Overwrite the Base File's tabs with fresh data ---
# The base file already has the header row formatted (including the grouped/hidden
# advocateshare, neutralshare, and detractorshare columns), so we only clear and
# rewrite the data rows beneath the header, leaving row 1 formatting untouched.
# NOTE: all five tabs now have eight extra columns appended after ffv_pct --
# maintenance_advocates, maintenance_neutrals, maintenance_detractors,
# maintenance_total, maintenance_advocateshare, maintenance_neutralshare,
# maintenance_detractorshare, and MAINTENANCE_NPS -- on top of the earlier
# three FFV columns (ffv_passed, ffv_total, ffv_pct). The base file's header
# row for every tab needs matching columns added before this script is run.
# The TAs tab is new: if the base file doesn't have it yet, ensure_sheet_like()
# creates it with the District tab's header row/formatting copied over.
BASE_FILE_PATH = r"C:\Users\Ta23458\Documents\Base Files\NPS report Southwest 20261109 Base File.xlsx"

wb = openpyxl.load_workbook(BASE_FILE_PATH)
SOUTHWEST_HIGHLIGHT_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")


def overwrite_sheet_data(ws, rows):
    max_row = ws.max_row
    if max_row > 1:
        ws.delete_rows(2, max_row - 1)
    for row in rows:
        ws.append(list(row))


def highlight_row_by_value(ws, column_index, match_value, fill):
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if row[column_index - 1].value == match_value:
            for cell in row:
                cell.fill = fill


def ensure_sheet_like(wb, sheet_name, template_sheet_name):
    """Create sheet_name (with the header row/formatting copied from
    template_sheet_name) if the base file doesn't have it yet."""
    if sheet_name in wb.sheetnames:
        return wb[sheet_name]
    template = wb[template_sheet_name]
    ws = wb.create_sheet(sheet_name)
    for col_idx, cell in enumerate(template[1], start=1):
        new_cell = ws.cell(row=1, column=col_idx, value=cell.value)
        if cell.has_style:
            new_cell.font = cell.font.copy()
            new_cell.fill = cell.fill.copy()
            new_cell.border = cell.border.copy()
            new_cell.alignment = cell.alignment.copy()
            new_cell.number_format = cell.number_format
    return ws


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

    if "maintenance_advocates" in idx and "maintenance_neutrals" in idx and "maintenance_detractors" in idx:
        m_advocates = sum(row[idx["maintenance_advocates"]] or 0 for row in rows)
        m_neutrals = sum(row[idx["maintenance_neutrals"]] or 0 for row in rows)
        m_detractors = sum(row[idx["maintenance_detractors"]] or 0 for row in rows)
        m_total = m_advocates + m_neutrals + m_detractors

        total_row[idx["maintenance_advocates"]] = m_advocates
        total_row[idx["maintenance_neutrals"]] = m_neutrals
        total_row[idx["maintenance_detractors"]] = m_detractors
        if "maintenance_total" in idx:
            total_row[idx["maintenance_total"]] = m_total

        if m_total:
            m_advocateshare = m_advocates / m_total
            m_neutralshare = m_neutrals / m_total
            m_detractorshare = m_detractors / m_total
            if "maintenance_advocateshare" in idx:
                total_row[idx["maintenance_advocateshare"]] = m_advocateshare
            if "maintenance_neutralshare" in idx:
                total_row[idx["maintenance_neutralshare"]] = m_neutralshare
            if "maintenance_detractorshare" in idx:
                total_row[idx["maintenance_detractorshare"]] = m_detractorshare
            if "maintenance_nps" in idx:
                total_row[idx["maintenance_nps"]] = round((m_advocateshare - m_detractorshare) * 100, 1)

    return total_row


bc_total_row = build_total_row(bc_columns, bc_rows)

overwrite_sheet_data(wb["Business Center"], bc_rows + [bc_total_row])
overwrite_sheet_data(wb["District"], dist_rows)
overwrite_sheet_data(wb["Dealer"], dealer_rows)
overwrite_sheet_data(wb["DOM"], dom_rows)
overwrite_sheet_data(ensure_sheet_like(wb, "TAs", "District"), ta_dist_rows)

# Highlight the Southwest row on the Business Center tab
highlight_row_by_value(wb["Business Center"], 1, "Southwest", SOUTHWEST_HIGHLIGHT_FILL)

wb.save(BASE_FILE_PATH)
print(f"Overwrote {BASE_FILE_PATH}")

# Also save a dated copy alongside the base file (without "Base File" in the name).
# If a file with that name already exists, append v2, v3, etc. instead of overwriting it.
today_str = date.today().strftime("%Y%m%d")
DATED_COPY_PATH = get_versioned_path(
    os.path.join(os.path.dirname(BASE_FILE_PATH), f"NPS report Southwest {today_str}.xlsx")
)
wb.save(DATED_COPY_PATH)
print(f"Saved dated copy to {DATED_COPY_PATH}")

# --- Step 3: Email the file via Outlook ---
RECIPIENT = "JesseTest"
EXCLUDED_EMAIL_COLUMNS = {
    "advocateshare", "neutralshare", "detractorshare",
    "maintenance_advocateshare", "maintenance_neutralshare", "maintenance_detractorshare",
}


def drop_share_columns(columns, rows):
    keep_idx = [i for i, col in enumerate(columns) if col.lower() not in EXCLUDED_EMAIL_COLUMNS]
    filtered_columns = [columns[i] for i in keep_idx]
    filtered_rows = [[row[i] for i in keep_idx] for row in rows]
    return filtered_columns, filtered_rows


def rows_to_html_table(title, columns, rows):
    columns, rows = drop_share_columns(columns, rows)
    header_html = "".join(
        f"<th style='border:1px solid #ccc;padding:4px 8px;background:#f2f2f2;'>{col}</th>" for col in columns
    )
    body_html = ""
    for row in rows:
        cells = "".join(f"<td style='border:1px solid #ccc;padding:4px 8px;'>{val}</td>" for val in row)
        body_html += f"<tr>{cells}</tr>"
    return (
        f"<h3>{title}</h3>"
        "<table style='border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;'>"
        f"<tr>{header_html}</tr>{body_html}</table><br>"
    )


html_body = (
    "<p>Hi Team,</p>"
    "<p>Please see todays NPS Report below and attached.</p>"
    + rows_to_html_table("Business Center", bc_columns, bc_rows + [bc_total_row])
    + rows_to_html_table("District", dist_columns, dist_rows)
    + rows_to_html_table("TAs", ta_dist_columns, ta_dist_rows)
    + rows_to_html_table("DOM", dom_columns, dom_rows)
    + "<p>Thanks</p>"
)

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
mail.To = RECIPIENT
mail.Subject = f"NPS Report {date.today().strftime('%m/%d/%Y')}"
mail.HTMLBody = html_body
mail.Attachments.Add(os.path.abspath(DATED_COPY_PATH))
mail.Send()

print(f"Email sent to {RECIPIENT}")
print("Done.")
