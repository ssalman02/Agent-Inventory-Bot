import shutil
import os
import re
import time
import xml.etree.ElementTree as ET
import datetime
import pandas as pd

from playwright.sync_api import sync_playwright

CASE_STATES = {
    "IE": 38,
    "CC": 39,
    "Retrival": 40,
}

def retry(func, retries=3, delay=1):
    last_error = None

    for attempt in range(1, retries + 1):
        try:
            return func()

        except Exception as exc:
            last_error = exc

            print(
                f"Attempt {attempt}/{retries} failed: {exc}"
            )

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
        root = ET.fromstring(content)

        print("XML parsed successfully.")

        return root

    except ET.ParseError as exc:
        print(f"Initial XML parsing failed: {exc}")

    text = content.decode(
        "utf-8-sig",
        errors="replace",
    )

    print("\n========== XML DIAGNOSTIC ==========")

    for line_number, line in enumerate(
        text.splitlines()[:20],
        start=1,
    ):
        print(f"{line_number:03}: {line[:300]}")

    print("====================================\n")

    # Identify namespace prefixes used in element names.
    element_prefixes = set(
        re.findall(
            r"<\s*/?\s*([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*",
            text,
        )
    )

    # Identify namespace prefixes used in attribute names.
    attribute_prefixes = set(
        re.findall(
            r"\s([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*\s*=",
            text,
        )
    )

    used_prefixes = element_prefixes | attribute_prefixes

    declared_prefixes = set(
        re.findall(
            r"\bxmlns:([A-Za-z_][\w.-]*)\s*=",
            text,
        )
    )

    missing_prefixes = used_prefixes - declared_prefixes

    # The xml namespace is predefined.
    missing_prefixes.discard("xml")

    print("Prefixes used:", sorted(used_prefixes))
    print("Prefixes declared:", sorted(declared_prefixes))
    print("Missing prefixes:", sorted(missing_prefixes))

    if not missing_prefixes:
        raise ValueError(
            "XML parsing failed, but no missing namespace prefix "
            "was detected. Inspect the XML diagnostic output."
        )

    root_match = re.search(
        r"<([A-Za-z_][\w.:-]*)(?=[\s>/])",
        text,
    )

    if not root_match:
        raise ValueError(
            "Could not identify the XML root element."
        )

    declarations = " ".join(
        f'xmlns:{prefix}="urn:temporary-repair:{prefix}"'
        for prefix in sorted(missing_prefixes)
    )

    insert_at = root_match.end()

    repaired_text = (
        text[:insert_at]
        + " "
        + declarations
        + text[insert_at:]
    )

    try:
        root = ET.fromstring(repaired_text)

    except ET.ParseError as exc:
        raise ValueError(
            f"XML repair failed: {exc}. "
            "Inspect the downloaded XML file."
        ) from exc

    repaired_file = (
        os.path.splitext(xml_file)[0]
        + "_repaired.xml"
    )

    with open(
        repaired_file,
        "w",
        encoding="utf-8",
    ) as file:
        file.write(repaired_text)

    print("XML parsed after temporary namespace repair.")
    print(f"Repaired XML saved to: {repaired_file}")

    return root

def extract_rows(root):
    """
    Extract table rows from ordinary XML or Excel-compatible XML.
    """

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
                        cell_value = "".join(
                            element.itertext()
                        ).strip()

                        break

                if not cell_value:
                    cell_value = "".join(
                        cell.itertext()
                    ).strip()

                values.append(cell_value)

                current_column += 1

        else:

            # Ordinary XML: Row > child elements
            for child in children:

                values.append(
                    "".join(child.itertext()).strip()
                )

        if values:
            extracted_rows.append(values)

    return extracted_rows

def xml_to_dataframe(xml_file, case_state):

    root = read_xml_file(xml_file)

    rows = extract_rows(root)

    if not rows:
        raise ValueError(
            f"No data rows found in XML file: {xml_file}"
        )

    print(f"Extracted XML rows: {len(rows)}")

    header_index = None
    last_vis_index = None

    for row_index, row in enumerate(rows):

        normalized_row = [
            str(value).strip().lower()
            for value in row
        ]

        if (
            "patient name" in normalized_row
            and "last vis" in normalized_row
        ):

            header_index = row_index

            last_vis_index = normalized_row.index(
                "last vis"
            )

            break

    if header_index is None:

        raise ValueError(
            "Could not identify the XML header row."
        )

    print(f"Header row found at: {header_index}")

    print(
        f"Last Vis column index: {last_vis_index}"
    )

    headers = rows[header_index][
        :last_vis_index + 1
    ]

    data_rows = rows[header_index + 1:]

    normalized_rows = []

    for row in data_rows:

        row = list(row[:last_vis_index + 1])

        # Ensure every row has the same number of columns.
        if len(row) < len(headers):

            row.extend(
                [""] * (len(headers) - len(row))
            )

        normalized_rows.append(row)

    df = pd.DataFrame(
        normalized_rows,
        columns=headers,
    )

    # Assign the current case state so crosstab counts IE, CC, and Retrival properly.
    df["Status"] = case_state

    print(
        f"{case_state} DataFrame created successfully."
    )

    print(f"Total rows: {len(df)}")
    print(f"Total columns: {len(df.columns)}")

    return df

def click_more_filters(page):
    """Open the More Filters panel."""

    retry(
        lambda: page.locator(
            '//div[contains(@class, "textnode") '
            'and contains(., "More Filters")]'
        ).click()
    )
    time.sleep(2)


def load_filters(page):
    """Load the saved filter list."""

    retry(
        lambda: page.locator(
            "//*[@id='id.DASH!!7.btnLoadFilters']/div"
        ).click()
    )
    time.sleep(2)



def select_case_state(page, case_state):
    """Select IE, CC or Retrival from the filter list."""

    row_number = CASE_STATES[case_state]

    print(
        f"Selecting case state: {case_state} "
        f"(row {row_number})"
    )

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

    retry(
        lambda: page.keyboard.press("Escape")
    )

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


def process_case(
    page,
    case_state,
    start_date,
    end_date,
    download_path,
):

    print("\n" + "=" * 60)
    print(f"STARTING SCENARIO: {case_state}")
    print("=" * 60)

    # Select the current case.
    select_case_state(page, case_state)

    # Open filters and apply the same date range.
    click_more_filters(page)

    apply_date_filters(
        page,
        start_date,
        end_date,
    )

    # Select the pending authorization record.
    retry(
        lambda: page.locator(
            "(//*[@id='id.DASH!!5.lstPendAuth']//table//tr)[1]"
        ).click()
    )
    time.sleep(2)

    # Open the print/export workflow.
    retry(
        lambda: page.locator(
            "//*[@id='ToolbarButton.&Print']/div/span"
        ).click()
    )
    time.sleep(2)

    # Retain this only if required by your application.
    retry(
        lambda: page.keyboard.press("Control+p")
    )
    time.sleep(2)

    download_button = page.locator(
        "//*[@id='cb']/div/div[4]/div/div[2]/div[2]"
    )
    time.sleep(2)

    # Download the XML.
    def download_xml():

        with page.expect_download(
            timeout=60000
        ) as download_info:

            download_button.click()

        download = download_info.value

        xml_file = os.path.join(
            download_path,
            f"{case_state}.xml",
        )

        download.save_as(xml_file)

        return xml_file

    xml_file = retry(
        download_xml,
        retries=3,
        delay=2,
    )

    print(
        f"{case_state} XML downloaded: {xml_file}"
    )

    # Parse the XML and create a dataframe.
    df = xml_to_dataframe(
        xml_file,
        case_state,
    )

    print(
        f"Scenario {case_state} completed successfully. "
        f"Rows: {len(df)}"
    )

    return df


def save_combined_excel(results, output_file, summary_file):

    # ==========================================
    # FILE 1: ALL THREE AGENTS
    # ==========================================

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
    ) as writer:

        for case_state in CASE_STATES:

            df = results.get(case_state)

            if df is None:
                df = pd.DataFrame()

            df.to_excel(
                writer,
                sheet_name=case_state,
                index=False,
            )

    print(f"Agent Excel created: {output_file}")

    # ==========================================
    # COMBINE ALL THREE AGENTS FOR SUMMARY
    # ==========================================

    combined_df = pd.concat(
        [
            df
            for df in results.values()
            if df is not None and not df.empty
        ],
        ignore_index=True,
    )

    if not combined_df.empty:

        combined_df.columns = [
            str(col).strip()
            for col in combined_df.columns
        ]

        if "Payor" not in combined_df.columns:
            raise ValueError(
                "Payor column not found in combined data."
            )

        if "Status" not in combined_df.columns:
            raise ValueError(
                "Status column not found in combined data."
            )

        # Exclude blank Payor and Status records.
        summary_data = combined_df[
            combined_df["Payor"].notna()
            & (
                combined_df["Payor"]
                .astype(str)
                .str.strip() != ""
            )
            & combined_df["Status"].notna()
            & (
                combined_df["Status"]
                .astype(str)
                .str.strip() != ""
            )
        ].copy()

        summary_data["Payor"] = summary_data["Payor"].astype(str).str.strip()
        summary_data["Status"] = summary_data["Status"].astype(str).str.strip()

        # Create Payor-wise summary.
        summary_df = pd.crosstab(
            summary_data["Payor"],
            summary_data["Status"],
        )

        # Ensure all three agent columns exist.
        for case_state in CASE_STATES:

            if case_state not in summary_df.columns:
                summary_df[case_state] = 0

        # Maintain agent column order.
        summary_df = summary_df[
            list(CASE_STATES.keys())
        ]

        # Add Grand Total column.
        summary_df["Grand Total"] = (
            summary_df.sum(axis=1)
        )

        # Add Grand Total row.
        grand_total = summary_df.sum(axis=0)

        summary_df.loc["Grand Total"] = grand_total

    else:

        summary_df = pd.DataFrame(
            columns=list(CASE_STATES.keys())
            + ["Grand Total"]
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
        )

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
            "Summary workbook must contain only "
            "the Summary worksheet."
        )

    print("\nExcel verification successful.")

    print("Agent workbook:", agent_excel.sheet_names)
    print("Summary workbook:", summary_excel.sheet_names)


def cleanup_output_folder(download_path, final_files):

    print("\nCleaning intermediate files...")

    final_files = {
        os.path.abspath(file)
        for file in final_files
    }

    for item in os.listdir(download_path):

        item_path = os.path.abspath(
            os.path.join(download_path, item)
        )

        if item_path in final_files:
            continue

        # Remove XML files generated by this automation.
        if item.lower().endswith(".xml"):

            try:
                os.remove(item_path)
                print(f"Deleted XML: {item}")

            except Exception as exc:
                print(f"Could not delete {item}: {exc}")

    print("\nCleanup completed.")
    print("Both final Excel workbooks are preserved.")

def run(
    playwright,
    start_date,
    end_date,
    download_path,
):
    try:    
        user_data_dir = os.path.abspath(os.path.join(os.getcwd(), "chrome-session"))
        os.makedirs(
                download_path,
                exist_ok=True,
        )
        context = playwright.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                headless=False,
                channel="chrome",
                chromium_sandbox=True,
                args=["--start-maximized"],
                no_viewport=True,
                accept_downloads=True,
            )

        print("Chromium launched successfully.")

        pages = context.pages

        if pages:
            page = pages[0]
        else:
            page = context.new_page()

        print("New browser page created.")

        retry(lambda: page.goto(
                "https://goldenbear.raintreeinc.com/webdocumentation",
                wait_until="domcontentloaded",
            ))
        print("Logging into Golden Bear...")
    
        retry(lambda:page.locator('//input[@name="user"]').fill("TAGO"))
        retry(lambda:page.locator('//input[@name="pass"]').fill("Therapy#167"))
        retry(lambda: page.locator("//input[@name='login']").click())
        time.sleep(3)
        retry(
            lambda: page.locator(
                '//div[@id="id.MMG!!1.0!sDashboard"]'
            ).click()
        )

        time.sleep(4)

        retry(
            lambda: page.locator(
                '//span[contains(., "Auth")]'
            ).nth(1).click()
        )

        time.sleep(3)

        print("Authorization page opened.")

        results = {}

        for index, case_state in enumerate(CASE_STATES):
            try:
                if index == 0:
                    click_more_filters(page)
                    load_filters(page)
                else:
                    reset_filters_for_next_case(page)

                df = process_case(
                    page=page,
                    case_state=case_state,
                    start_date=start_date,
                    end_date=end_date,
                    download_path=download_path,
                )

                results[case_state] = df

            except Exception as exc:

                print(
                    f"ERROR processing {case_state}: {exc}"
                )

                raise

            if index < len(CASE_STATES) - 1:

                reset_filters_for_next_case(page)

        if start_date and end_date:
            startdateP = datetime.datetime.strptime(
                start_date,
                "%m-%d-%y"
            ).strftime("%d")

            enddateP = datetime.datetime.strptime(
                end_date,
                "%m-%d-%y"
            ).strftime("%d")

            filename = f"NAD {startdateP} - {enddateP}"
            
        elif start_date:
            startdateP = datetime.datetime.strptime(
                start_date,
                "%m-%d-%y"
            ).strftime("%d")

            filename = f"NAD From_{startdateP}"

        elif end_date:

            enddateP = datetime.datetime.strptime(
                end_date,
                "%m-%d-%y"
            ).strftime("%d")

            filename = f"NAD Till_{enddateP}"

        else:

            filename = "All_Records"
        
        # ==========================================
        # GENERATE TWO EXCEL FILE NAMES
        # ==========================================

        # File 1: All three agents
        output_file = os.path.join(
            download_path,
            f"{filename}.xlsx",
        )

        # File 2: Separate Summary
        summary_file = os.path.join(
            download_path,
            f"Summary_{filename}.xlsx",
        )

        # ==========================================
        # SAVE BOTH EXCEL WORKBOOKS
        # ==========================================

        save_combined_excel(
            results=results,
            output_file=output_file,
            summary_file=summary_file,
        )

        # ==========================================
        # CLEANUP INTERMEDIATE FILES
        # ==========================================

        cleanup_output_folder(
            download_path,
            [
                output_file,
                summary_file,
            ],
        )

        # ==========================================
        # FINAL CONFIRMATION
        # ==========================================

        print("\n" + "=" * 60)
        print("PROCESS COMPLETED SUCCESSFULLY")
        print("=" * 60)

        print(f"Agent Excel: {output_file}")
        print(f"Summary Excel: {summary_file}")

        print("\nOnly two Excel files remain in downloads.")

        return output_file, summary_file
    
    finally:
        print("Browser closed successfully.")


if __name__ == "__main__":
    with open("config.txt", "r") as file:
        lines = file.readlines()

    START_DATE = lines[0].split("=")[1].strip()
    END_DATE = lines[1].split("=")[1].strip()  
    
    DOWNLOAD_PATH = os.path.join(
        os.getcwd(),
        "downloads",
    )

    with sync_playwright() as playwright:

        output_file, summary_file = run(
            playwright=playwright,
            start_date=START_DATE,
            end_date=END_DATE,
            download_path=DOWNLOAD_PATH,
        )

        print("\nALL THREE SCENARIOS COMPLETED.")
        time.sleep(1)
        print(f"Agent Excel: {output_file}")
        time.sleep(1)
        print(f"Summary Excel: {summary_file}")