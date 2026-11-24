import datetime
import os
import re
import time
import xml.etree.ElementTree as ET

import pandas as pd
from openpyxl.styles import Font, PatternFill

CASE_STATES = {
    "IE": 38,
    "CC": 39,
    "Retrival": 40,
}

SELECTED_RAW_COLUMNS = [1, 2, 3, 4, 6, 11, 26]

OUTPUT_HEADERS = [
    "Patient Name",
    "BC",
    "Payor",
    "NAD",
    "Sub-Stat",
    "PN Number",
    "Status",
]

PAYOR_RAW_COLUMN = 3  # 1-based raw column number for Payor


def retry(func, retries=3, delay=1):
    last_error = None

    for attempt in range(1, retries + 1):
        try:
            return func()

        except Exception as exc:
            last_error = exc

            print(f"Attempt {attempt}/{retries} failed: {exc}")

            if attempt < retries:
                time.sleep(delay)

    raise last_error


def local_name(tag):
    """Return the local name of an XML tag."""
    if "}" in tag:
        return tag.rsplit("}", 1)[1]

    if ":" in tag:
        return tag.rsplit(":", 1)[1]

    return tag


def read_xml_file(xml_file):
    with open(xml_file, "rb") as file:
        content = file.read()

    try:
        return ET.fromstring(content)
    except ET.ParseError:
        pass

    text = content.decode("utf-8-sig", errors="replace")

    element_prefixes = set(
        re.findall(r"<\s*/?\s*([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*", text)
    )
    attribute_prefixes = set(
        re.findall(r"\s([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\s*=", text)
    )
    used_prefixes = element_prefixes | attribute_prefixes
    declared_prefixes = set(re.findall(r"\bxmlns:([A-Za-z_][\w.-]*)\s*=", text))
    missing_prefixes = used_prefixes - declared_prefixes
    missing_prefixes.discard("xml")

    if not missing_prefixes:
        raise ValueError(f"XML parsing failed in {xml_file}.")

    root_match = re.search(r"<([A-Za-z_][\w.:-]*)(?=[\s>/])", text)
    if not root_match:
        raise ValueError(f"Could not identify XML root element in {xml_file}.")

    declarations = " ".join(
        f'xmlns:{prefix}="urn:temporary-repair:{prefix}"'
        for prefix in sorted(missing_prefixes)
    )
    insert_at = root_match.end()
    repaired_text = text[:insert_at] + " " + declarations + text[insert_at:]

    return ET.fromstring(repaired_text)


def extract_rows(root):
    """Extract table rows from ordinary XML or Excel-compatible XML."""
    extracted_rows = []

    for row in root.iter():
        if local_name(row.tag).lower() != "row":
            continue

        values = []
        children = list(row)
        cells = [
            child
            for child in children
            if local_name(child.tag).lower() == "cell"
        ]

        if cells:
            current_column = 1

            for cell in cells:
                index_value = None

                for attribute, value in cell.attrib.items():
                    if local_name(attribute).lower() == "index":
                        index_value = value
                        break

                if index_value is not None:
                    try:
                        target_column = int(index_value)

                        while current_column < target_column:
                            values.append("")
                            current_column += 1

                    except ValueError:
                        pass

                cell_value = ""

                for element in cell.iter():
                    if local_name(element.tag).lower() == "data":
                        cell_value = "".join(element.itertext()).strip()
                        break

                if not cell_value:
                    cell_value = "".join(cell.itertext()).strip()

                values.append(cell_value)
                current_column += 1

        else:
            # Ordinary XML: Row > child elements
            for child in children:
                values.append("".join(child.itertext()).strip())

        if values:
            extracted_rows.append(values)

    return extracted_rows


def find_header_candidates(rows):
    """Return (row_index, row_values) for rows that look like headers."""
    candidates = []

    for idx, row in enumerate(rows):
        non_empty = [str(val).strip() for val in row if str(val).strip() != ""]
        if len(non_empty) < 2:
            continue

        # Treat a row as a header candidate if at least half of non-empty cells are non-numeric text
        non_numeric_count = 0
        for cell in non_empty:
            cleaned = (
                cell.replace(",", "")
                .replace(".", "")
                .replace("-", "")
                .replace("/", "")
                .strip()
            )
            if not cleaned.isdigit():
                non_numeric_count += 1

        if non_numeric_count / len(non_empty) >= 0.5:
            candidates.append((idx, [str(val).strip() for val in row]))

    return candidates


def _resolve_selected_indices(total_cols):
    """Convert 1-based column numbers in SELECTED_RAW_COLUMNS to 0-based indices."""
    indices = []
    for col_num in SELECTED_RAW_COLUMNS:
        idx = int(col_num) - 1
        if 0 <= idx < total_cols and idx not in indices:
            indices.append(idx)
    return indices


def xml_to_dataframe(xml_file, case_state):
    root = read_xml_file(xml_file)
    rows = extract_rows(root)

    if not rows:
        raise ValueError(f"No data rows found in XML file: {xml_file}")

    trimmed_rows = rows[2:]
    if not trimmed_rows:
        raise ValueError(
            f"No rows remaining after removing top 2 rows in: {xml_file}"
        )

    data_rows = trimmed_rows[1:]
    source_col_numbers = SELECTED_RAW_COLUMNS[:6]
    mapped_records = []

    for row in data_rows:
        record = {}
        for header_name, col_num in zip(OUTPUT_HEADERS[:6], source_col_numbers):
            col_idx = int(col_num) - 1
            record[header_name] = (
                str(row[col_idx]).strip() if col_idx < len(row) else ""
            )
        record["Status"] = ""
        mapped_records.append(record)

    df = pd.DataFrame(mapped_records, columns=OUTPUT_HEADERS)
    print(f"[{case_state}] {len(df)} rows extracted.")
    return df


def click_more_filters(page):
    retry(
        lambda: page.locator(
            '//div[contains(@class, "textnode") '
            'and contains(., "More Filters")]'
        ).click()
    )
    time.sleep(2)


def load_filters(page):
    retry(
        lambda: page.locator(
            "//*[@id='id.DASH!!7.btnLoadFilters']/div"
        ).click()
    )
    time.sleep(2)


def select_case_state(page, case_state):
    row_number = CASE_STATES[case_state]

    print(f"Selecting case state: {case_state} (row {row_number})")

    retry(
        lambda: page.locator(
            f"//*[@id='id._RLS!!8.filterList']"
            f"/div/table/tbody/tr[{row_number}]"
        ).click()
    )
    time.sleep(2)

    retry(
        lambda: page.locator(
            "//*[@id='id._RLS!!8.btnFilterLoad']"
        ).click()
    )
    time.sleep(2)


def reset_filters_for_next_case(page):
    print("Resetting interface for the next case...")

    retry(lambda: page.keyboard.press("Escape"))
    time.sleep(2)

    click_more_filters(page)
    time.sleep(2)

    load_filters(page)
    time.sleep(2)


def apply_date_filters(page, start_date=None, end_date=None):
    """Apply date filters only when a date is provided."""
    if start_date is not None:
        print(f"Applying start date: {start_date}")
        retry(
            lambda: page.locator(
                '//input[@id="id.DASH!!7.penAuthNadFrom"]'
            ).fill(start_date)
        )

    if end_date is not None:
        print(f"Applying end date: {end_date}")
        retry(
            lambda: page.locator(
                '//input[@id="id.DASH!!7.penAuthNadTo"]'
            ).fill(end_date)
        )

    retry(
        lambda: page.locator(
            "//*[@id='id.DASH!!7.btnApplyFilters']/div"
        ).click()
    )
    time.sleep(2)


def process_case(page, case_state, start_date, end_date, download_path):
    select_case_state(page, case_state)
    click_more_filters(page)
    apply_date_filters(page, start_date, end_date)

    retry(
        lambda: page.locator(
            "(//*[@id='id.DASH!!5.lstPendAuth']//table//tr)[1]"
        ).click()
    )
    time.sleep(2)

    retry(
        lambda: page.locator(
            "//*[@id='ToolbarButton.&Print']/div/span"
        ).click()
    )
    time.sleep(2)

    retry(lambda: page.keyboard.press("Control+p"))
    time.sleep(2)

    download_button = page.locator(
        "//*[@id='cb']/div/div[4]/div/div[2]/div[2]"
    )
    time.sleep(2)

    def download_xml():
        with page.expect_download(timeout=60000) as download_info:
            download_button.click()
        download = download_info.value
        xml_file = os.path.join(download_path, f"{case_state}.xml")
        download.save_as(xml_file)
        return xml_file

    xml_file = retry(download_xml, retries=3, delay=2)
    return xml_to_dataframe(xml_file, case_state)


def _style_header_row(worksheet, num_cols):
    dark_yellow_fill = PatternFill(
        start_color="FFFF00",
        end_color="FFFF00",
        fill_type="solid",
    )
    bold_black_font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="000000",
    )

    for col_idx in range(1, num_cols + 1):
        cell = worksheet.cell(row=1, column=col_idx)
        cell.fill = dark_yellow_fill
        cell.font = bold_black_font


def save_combined_excel(results, output_file, summary_file):
    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
    ) as writer:
        for case_state in CASE_STATES:
            df = results.get(case_state)
            if df is None or df.empty:
                df = pd.DataFrame(columns=OUTPUT_HEADERS)
            else:
                df = df[OUTPUT_HEADERS]

            # startrow=0 writes headers on Row 1 and data starting on Row 2
            df.to_excel(
                writer,
                sheet_name=case_state,
                index=False,
                startrow=0,
            )

            worksheet = writer.sheets[case_state]
            _style_header_row(worksheet, len(OUTPUT_HEADERS))

    print(f"Agent Excel created: {output_file}")

    # ==========================================
    # COMBINE ALL THREE AGENTS FOR SUMMARY
    # ==========================================

    summary_frames = []
    for case_state in CASE_STATES:
        df = results.get(case_state)
        if df is not None and not df.empty:
            temp = df.copy()
            # Use the sheet's case_state for the summary crosstab since Status is blank in output
            temp["Status"] = case_state
            summary_frames.append(temp)

    combined_df = (
        pd.concat(summary_frames, ignore_index=True)
        if summary_frames
        else pd.DataFrame(columns=OUTPUT_HEADERS)
    )

    if not combined_df.empty:
        combined_df.columns = [str(col).strip() for col in combined_df.columns]

        if "Payor" not in combined_df.columns:
            raise ValueError("Payor column not found in combined data.")

        summary_data = combined_df[
            combined_df["Payor"].notna()
            & (combined_df["Payor"].astype(str).str.strip() != "")
            & combined_df["Status"].notna()
            & (combined_df["Status"].astype(str).str.strip() != "")
        ].copy()

        summary_data["Payor"] = summary_data["Payor"].astype(str).str.strip()
        summary_data["Status"] = summary_data["Status"].astype(str).str.strip()

        summary_df = pd.crosstab(
            summary_data["Payor"],
            summary_data["Status"],
        )

        for case_state in CASE_STATES:
            if case_state not in summary_df.columns:
                summary_df[case_state] = 0

        summary_df = summary_df[list(CASE_STATES.keys())]
        summary_df["Grand Total"] = summary_df.sum(axis=1)
        summary_df.loc["Grand Total"] = summary_df.sum(axis=0)
    else:
        summary_df = pd.DataFrame(
            columns=list(CASE_STATES.keys()) + ["Grand Total"]
        )

    # ==========================================
    # FILE 2: SEPARATE SUMMARY EXCEL
    # ==========================================

    with pd.ExcelWriter(
        summary_file,
        engine="openpyxl",
    ) as writer:
        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            startrow=0,
        )
        summary_ws = writer.sheets["Summary"]
        # +1 accounts for the Payor index column written by pandas
        _style_header_row(summary_ws, len(summary_df.columns) + 1)

    print(f"Summary Excel created: {summary_file}")

    # ==========================================
    # VERIFY BOTH EXCEL WORKBOOKS
    # ==========================================

    agent_excel = pd.ExcelFile(output_file)
    summary_excel = pd.ExcelFile(summary_file)

    expected_agent_sheets = list(CASE_STATES.keys())
    if agent_excel.sheet_names != expected_agent_sheets:
        raise ValueError(
            f"Agent workbook sheets incorrect. "
            f"Expected: {expected_agent_sheets}, "
            f"Found: {agent_excel.sheet_names}"
        )

    if summary_excel.sheet_names != ["Summary"]:
        raise ValueError(
            "Summary workbook must contain only the Summary worksheet."
        )

    print("\nExcel verification successful.")
    print("Agent workbook:", agent_excel.sheet_names)
    print("Summary workbook:", summary_excel.sheet_names)


def cleanup_output_folder(download_path, final_files):
    final_files = {os.path.abspath(file) for file in final_files}
    for item in os.listdir(download_path):
        item_path = os.path.abspath(os.path.join(download_path, item))
        if item_path in final_files:
            continue
        if item.lower().endswith(".xml"):
            try:
                os.remove(item_path)
            except Exception:
                pass


def build_output_filenames(start_date, end_date):
    if start_date and end_date:
        start_day = datetime.datetime.strptime(start_date, "%m-%d-%y").strftime("%d")
        end_day = datetime.datetime.strptime(end_date, "%m-%d-%y").strftime("%d")
        filename = f"NAD {start_day} - {end_day}"
    elif start_date:
        start_day = datetime.datetime.strptime(start_date, "%m-%d-%y").strftime("%d")
        filename = f"NAD From_{start_day}"
    elif end_date:
        end_day = datetime.datetime.strptime(end_date, "%m-%d-%y").strftime("%d")
        filename = f"NAD Till_{end_day}"
    else:
        filename = "All_Records"

    return f"{filename}.xlsx", f"Summary_{filename}.xlsx"


def normalize_date_str(date_str):
    """Convert slash-separated US dates (e.g. 11/01/26) to MM-DD-YY (e.g. 11-01-26)."""
    if not date_str:
        return None

    cleaned = date_str.strip()
    for fmt in ("%m/%d/%y", "%m/%d/%Y", "%m-%d-%y", "%m-%d-%Y"):
        try:
            parsed = datetime.datetime.strptime(cleaned, fmt)
            return parsed.strftime("%m-%d-%y")
        except ValueError:
            continue

    raise ValueError(
        f"Unsupported date format: '{date_str}'. Expected MM/DD/YY or MM-DD-YY."
    )