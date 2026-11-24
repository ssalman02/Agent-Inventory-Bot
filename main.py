import os
import time

from playwright.sync_api import sync_playwright

from helper import (
    CASE_STATES,
    build_output_filenames,
    cleanup_output_folder,
    click_more_filters,
    load_filters,
    normalize_date_str,
    process_case,
    reset_filters_for_next_case,
    retry,
    save_combined_excel,
)


def run(
    playwright,
    start_date,
    end_date,
    download_path,
):
    try:
        user_data_dir = os.path.abspath(
            os.path.join(os.getcwd(), "chrome-session")
        )
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

        retry(
            lambda: page.goto(
                "https://goldenbear.raintreeinc.com/webdocumentation",
                wait_until="domcontentloaded",
            )
        )
        print("Logging into Golden Bear...")

        retry(lambda: page.locator('//input[@name="user"]').fill("TAGO"))
        retry(lambda: page.locator('//input[@name="pass"]').fill("Therapy#167"))
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
            )
            .nth(1)
            .click()
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
                print(f"ERROR processing {case_state}: {exc}")
                raise

            if index < len(CASE_STATES) - 1:
                reset_filters_for_next_case(page)

        output_name, summary_name = build_output_filenames(start_date, end_date)

        # ==========================================
        # GENERATE TWO EXCEL FILE NAMES
        # ==========================================

        # File 1: All three agents
        output_file = os.path.join(
            download_path,
            output_name,
        )

        # File 2: Separate Summary
        summary_file = os.path.join(
            download_path,
            summary_name,
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
    config_values = {}
    with open("config.txt", "r", encoding="utf-8") as file:
        for line in file:
            line = line.strip()
            if not line or "=" not in line:
                continue
            key, val = line.split("=", 1)
            cleaned_val = val.strip()
            config_values[key.strip().upper()] = cleaned_val if cleaned_val else None

    START_DATE = normalize_date_str(config_values.get("START_DATE"))
    END_DATE = normalize_date_str(config_values.get("END_DATE"))

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