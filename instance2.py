from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError
import pandas as pd
import ctypes
from datetime import datetime
import traceback
 
 
# ============================================================
# CONFIGURATION
# ============================================================
 
LOADER_CSS = "div.loader"
 
CHROME_DEBUG_URL = "http://127.0.0.1:9432"
 
EXCEL_PATH = "input2.xlsx"
 
# Change these only if your Excel column names are different
EXCEL_HOSTNAME_COLUMN = "hostname"
EXCEL_INSTANCE_TYPE_COLUMN = "instance"
 
# Main table XPath provided by you
MAIN_TABLE_XPATH = (
    "/html/body/div/div/div/div/div[2]/div/div[1]/div/"
    "div[4]/div[2]/div/div[2]/div[2]/div[1]/div/div[2]/table"
)
 
# Table header names
VM_OS_HEADER = "VM / OS"
INSTANCE_TYPE_HEADER = "INSTANCE TYPE"
 
# Pagination button
NEXT_BUTTON_SELECTOR = 'button:has-text("Next")'
 
# Save button shown in the screenshot
SAVE_ALL_BUTTON_SELECTOR = 'button:has-text("Save All Selection")'
 
# Set this to False temporarily if you do not want the final save click
SAVE_ALL_SELECTION_AT_END = True
 
 
# ============================================================
# EXECUTION COUNTERS
# ============================================================
 
automation_start_time = datetime.now()
 
total_excel_rows = 0
total_pages_processed = 0
total_website_rows_processed = 0
total_instance_types_updated = 0
total_already_matching = 0
total_not_found_in_excel = 0
total_skipped = 0
total_errors = 0
 
error_details = []
 
current_page_number = 1
 
 
# ============================================================
# LOGGING
# ============================================================
 
def log(message):
 
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
 
    print(
        f"[{current_time}] {message}",
        flush=True
    )
 
 
def record_error(
    vm_hostname,
    excel_instance_type,
    page_number,
    error_message
):
 
    global total_errors
 
    total_errors += 1
 
    error_details.append(
        {
            "vm_hostname": vm_hostname,
            "excel_instance_type": excel_instance_type,
            "page_number": page_number,
            "error": str(error_message)
        }
    )
 
    log(
        f"ERROR | Page: {page_number} | "
        f"VM: {vm_hostname} | "
        f"Excel Instance: {excel_instance_type} | "
        f"Error: {error_message}"
    )
 
 
# ============================================================
# WINDOWS SLEEP PREVENTION
# ============================================================
 
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002
 
 
def prevent_system_sleep():
 
    try:
 
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS
            | ES_SYSTEM_REQUIRED
            | ES_DISPLAY_REQUIRED
        )
 
        log(
            "Windows system sleep and display sleep "
            "have been disabled during automation."
        )
 
    except Exception as sleep_error:
 
        record_error(
            vm_hostname="N/A",
            excel_instance_type="N/A",
            page_number="N/A",
            error_message=(
                "Unable to prevent Windows from sleeping: "
                f"{sleep_error}"
            )
        )
 
 
def allow_system_sleep():
 
    try:
 
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS
        )
 
        log(
            "Normal Windows sleep behavior has been restored."
        )
 
    except Exception as sleep_error:
 
        log(
            "Unable to restore Windows sleep behavior: "
            f"{sleep_error}"
        )
 
 
# ============================================================
# NORMALIZATION FUNCTIONS
# ============================================================
 
def normalize_hostname(hostname):
 
    if hostname is None:
        return ""
 
    return (
        str(hostname)
        .strip()
        .lower()
    )
 
 
def normalize_instance_type(instance_type):
 
    if instance_type is None:
        return ""
 
    return (
        str(instance_type)
        .strip()
        .lower()
    )
 
 
# ============================================================
# READ EXCEL
# ============================================================
 
log(f"Reading Excel file: {EXCEL_PATH}")
 
excel_df = pd.read_excel(
    EXCEL_PATH,
    engine="openpyxl"
)
 
excel_df.columns = (
    excel_df.columns
    .astype(str)
    .str.strip()
)
 
 
if EXCEL_HOSTNAME_COLUMN not in excel_df.columns:
 
    raise ValueError(
        f"Excel column '{EXCEL_HOSTNAME_COLUMN}' was not found. "
        f"Available columns: {excel_df.columns.tolist()}"
    )
 
 
if EXCEL_INSTANCE_TYPE_COLUMN not in excel_df.columns:
 
    raise ValueError(
        f"Excel column '{EXCEL_INSTANCE_TYPE_COLUMN}' was not found. "
        f"Available columns: {excel_df.columns.tolist()}"
    )
 
 
hostname_instance_mapping = {}
 
 
for excel_index, excel_row in excel_df.iterrows():
 
    excel_hostname = excel_row[EXCEL_HOSTNAME_COLUMN]
 
    excel_instance_type = excel_row[
        EXCEL_INSTANCE_TYPE_COLUMN
    ]
 
 
    if pd.isna(excel_hostname):
 
        log(
            f"Skipping Excel row {excel_index + 2}: "
            f"hostname is empty."
        )
 
        continue
 
 
    if pd.isna(excel_instance_type):
 
        log(
            f"Skipping Excel row {excel_index + 2}: "
            f"instance type is empty."
        )
 
        continue
 
 
    excel_hostname = str(excel_hostname).strip()
 
    excel_instance_type = str(
        excel_instance_type
    ).strip()
 
 
    if not excel_hostname or not excel_instance_type:
 
        log(
            f"Skipping Excel row {excel_index + 2}: "
            f"hostname or instance type is blank."
        )
 
        continue
 
 
    normalized_hostname = normalize_hostname(
        excel_hostname
    )
 
 
    if normalized_hostname in hostname_instance_mapping:
 
        log(
            f"Duplicate Excel hostname found: "
            f"{excel_hostname}. "
            f"The last instance type will be used."
        )
 
 
    hostname_instance_mapping[
        normalized_hostname
    ] = excel_instance_type
 
 
total_excel_rows = len(hostname_instance_mapping)
 
 
log(
    f"Valid hostname and instance-type mappings "
    f"loaded from Excel: {total_excel_rows}"
)
 
 
# ============================================================
# LOADER HANDLING
# ============================================================
 
def wait_for_loader():
 
    loader = page.locator(LOADER_CSS)
 
    try:
 
        if loader.is_visible(timeout=3000):
 
            log("Loader is visible. Waiting for it to disappear.")
 
            loader.wait_for(
                state="hidden",
                timeout=120000
            )
 
            log("Loader disappeared.")
 
    except TimeoutError:
 
        # A loader may complete before Playwright detects it.
        # Check whether a visible loader still exists.
        try:
 
            if loader.is_visible():
 
                loader.wait_for(
                    state="hidden",
                    timeout=120000
                )
 
                log("Loader disappeared.")
 
        except TimeoutError:
 
            raise Exception(
                "Loader did not disappear within 120 seconds."
            )

# ============================================================
# WAIT FOR RECALCULATION AFTER INSTANCE CHANGE
# ============================================================

def wait_for_row_recalculation():
    pass
  
# ============================================================
# EXTRACT HOSTNAME FROM VM / OS CELL
# ============================================================
 
def get_vm_hostname(vm_os_cell):
 
    vm_os_text = vm_os_cell.inner_text().strip()
 
    text_lines = [
        line.strip()
        for line in vm_os_text.splitlines()
        if line.strip()
    ]
 
 
    if not text_lines:
 
        return ""
 
 
    # Example:
    #
    # VM 834
    # linux
    #
    # The first non-empty line is treated as the hostname.
    return text_lines[0]
 
 
# ============================================================
# FIND REQUIRED TABLE COLUMNS
# ============================================================
 
def get_required_column_indexes(table):
 
    headers = table.locator("thead th")
 
 
    if headers.count() == 0:
 
        headers = table.locator("th")
 
 
    vm_os_column_index = None
    instance_type_column_index = None
 
 
    for header_index in range(headers.count()):
 
        header_text = (
            headers
            .nth(header_index)
            .inner_text()
            .strip()
        )
 
 
        log(
            f"Header index {header_index}: "
            f"'{header_text}'"
        )
 
 
        if header_text.upper() == VM_OS_HEADER.upper():
 
            vm_os_column_index = header_index
 
 
        if (
            header_text.upper()
            == INSTANCE_TYPE_HEADER.upper()
        ):
 
            instance_type_column_index = header_index
 
 
    if vm_os_column_index is None:
 
        raise Exception(
            f"Unable to find the '{VM_OS_HEADER}' "
            f"column in the table."
        )
 
 
    if instance_type_column_index is None:
 
        raise Exception(
            f"Unable to find the "
            f"'{INSTANCE_TYPE_HEADER}' "
            f"column in the table."
        )
 
 
    log(
        f"VM / OS column index: "
        f"{vm_os_column_index}"
    )
 
 
    log(
        f"INSTANCE TYPE column index: "
        f"{instance_type_column_index}"
    )
 
 
    return (
        vm_os_column_index,
        instance_type_column_index
    )
 
 
# ============================================================
# LOCATE "ALL INSTANCE TYPES" DROPDOWN
# ============================================================

def get_all_instance_types_dropdown(
    instance_type_cell
):

    buttons = instance_type_cell.locator(
        "button.tss-trigger"
    )

    if buttons.count() < 2:
        raise Exception(
            "Expected 2 dropdown buttons "
            "but found fewer."
        )

    return buttons.nth(1)
# ============================================================
# GET CURRENT INSTANCE TYPE
# ============================================================

def get_current_instance_type(
    instance_dropdown
):

    try:

        selected_text = (
            instance_dropdown
            .inner_text()
            .strip()
        )

        if (
            selected_text
            and selected_text.lower()
            != "all instance types"
        ):
            return selected_text

    except Exception:
        pass

    return ""
# ============================================================
# SELECT INSTANCE TYPE
# ============================================================

def select_instance_type(
    instance_dropdown,
    required_instance_type
):

    required_instance_type = (
        str(required_instance_type)
        .strip()
    )

    log(
        f"Selecting instance type "
        f"'{required_instance_type}'"
    )

    instance_dropdown.scroll_into_view_if_needed()
    page.wait_for_timeout(100)

    instance_dropdown.click(
        force=True
    )

    search_box = page.locator(
        'input[placeholder="Search..."]'
    )

    try:

        search_box.wait_for(
            state="visible",
            timeout=2000
        )

    except TimeoutError:

        # Dropdown probably didn't open after scroll.
        instance_dropdown.click(
            force=True
        )

        search_box.wait_for(
            state="visible",
            timeout=10000
        )


    search_box.fill(required_instance_type)

    page.wait_for_timeout(500)

    try:
        option = page.get_by_role(
            "listitem"
        ).filter(
            has_text=required_instance_type
        ).first

        option.wait_for(
            state="visible",
            timeout=5000
        )

        option.scroll_into_view_if_needed()

        page.wait_for_timeout(300)

        option.click(force=True)

    except Exception:

        # Fallback for last row / viewport issues
        page.keyboard.press("ArrowDown")
        page.wait_for_timeout(300)
        page.keyboard.press("Enter")
# ============================================================
# PROCESS CURRENT TABLE PAGE
# ============================================================
 
def process_current_page(page_number):
 
    global total_pages_processed
    global total_website_rows_processed
    global total_instance_types_updated
    global total_already_matching
    global total_not_found_in_excel
    global total_skipped
 
 
    log(
        f"Starting processing for website page "
        f"{page_number}."
    )
 
 
    table = page.locator(
        f"xpath={MAIN_TABLE_XPATH}"
    )
 
 
    table.wait_for(
        state="visible",
        timeout=120000
    )
 
 
    (
        vm_os_column_index,
        instance_type_column_index
    ) = get_required_column_indexes(table)
 
 
    table_rows = table.locator("tbody tr:visible")
 
 
    if table_rows.count() == 0:
 
        all_rows = table.locator("tr:visible")
 
        row_start_index = 1
 
    else:
 
        all_rows = table_rows
 
        row_start_index = 0
 
 
    row_count = all_rows.count()
 
 
    log(
        f"Data rows found on page "
        f"{page_number}: {row_count}"
    )
 
 
    for row_index in range(
        row_start_index,
        row_count
    ):

        table_rows = table.locator("tbody tr:visible")

        log(
            f"STARTING ROW "
            f"{row_index + 1}"
        )

        table_rows = table.locator("tbody tr:visible")

        current_row = table_rows.nth(row_index)
        current_row.evaluate("""
                el => el.scrollIntoView({
                    behavior: 'instant',
                    block: 'center'
                })
                """)

        page.wait_for_timeout(100)
 
        vm_hostname = "Unable to read VM hostname"
 
        excel_instance_type = "N/A"
 
 
        try:
 
            row_cells = current_row.locator("td")
 
 
            if row_cells.count() == 0:
 
                log(
                    f"Page {page_number}, row "
                    f"{row_index + 1}: no TD cells found. "
                    f"Skipping row."
                )
 
                continue
 
 
            vm_os_cell = row_cells.nth(
                vm_os_column_index
            )
 
 
            vm_hostname = get_vm_hostname(
                vm_os_cell
            )
 
 
            if not vm_hostname:
 
                raise Exception(
                    "VM hostname is empty in the VM / OS cell."
                )
 
 
            total_website_rows_processed += 1
 
 
            log(
                f"Processing page {page_number}, "
                f"row {row_index + 1}, "
                f"VM hostname: '{vm_hostname}'."
            )
 
 
            excel_instance_type = (
                hostname_instance_mapping.get(
                    normalize_hostname(vm_hostname)
                )
            )
 
 
            if excel_instance_type is None:
 
                total_not_found_in_excel += 1
                total_skipped += 1
 
                log(
                    f"SKIPPED | VM hostname "
                    f"'{vm_hostname}' was not found "
                    f"in Excel."
                )
 
                continue
 
 
            log(
                f"Excel instance type for "
                f"'{vm_hostname}': "
                f"'{excel_instance_type}'."
            )
 
 
            instance_type_cell = row_cells.nth(
                instance_type_column_index
            )
 
            log(
                f"Obtaining dropdown for "
                f"'{vm_hostname}'"
            )
            instance_dropdown = (
                get_all_instance_types_dropdown(
                    instance_type_cell
                )
            )
            log(
                f"Dropdown obtained for "
                f"'{vm_hostname}'"
            )
 
            current_instance_type = (
                get_current_instance_type(
                    instance_dropdown
                )
            )
 
 
            log(
                f"Current website instance type for "
                f"'{vm_hostname}': "
                f"'{current_instance_type}'."
            )
 
 
            if (
                normalize_instance_type(
                    current_instance_type
                )
                == normalize_instance_type(
                    excel_instance_type
                )
            ):
 
                total_already_matching += 1
                total_skipped += 1
 
                log(
                    f"SKIPPED | VM '{vm_hostname}' "
                    f"already has matching instance type "
                    f"'{excel_instance_type}'."
                )
 
                continue
 
 
            log(
                f"UPDATING | VM: '{vm_hostname}' | "
                f"Current: '{current_instance_type}' | "
                f"New: '{excel_instance_type}'."
            )
 
            current_row.evaluate("""
                el => el.scrollIntoView({
                    behavior: 'instant',
                    block: 'center'
                })
                """)
            page.wait_for_timeout(200)

            table_rows = table.locator("tbody tr:visible")
            current_row = table_rows.nth(row_index)

            row_cells = current_row.locator("td")

            instance_type_cell = row_cells.nth(
                instance_type_column_index
            )

            instance_dropdown = (
                get_all_instance_types_dropdown(
                    instance_type_cell
                )
            )

            select_instance_type(
                instance_dropdown,
                excel_instance_type
            )
            
 
            total_instance_types_updated += 1
 
 
            log(
                f"UPDATED | VM '{vm_hostname}' "
                f"was updated to instance type "
                f"'{excel_instance_type}'."
            )

            log(
                f"FINISHED ROW "
                f"{row_index + 1}"
            )
 
 
        except Exception as row_error:
 
            record_error(
                vm_hostname=vm_hostname,
                excel_instance_type=(
                    excel_instance_type
                    if excel_instance_type
                    else "N/A"
                ),
                page_number=page_number,
                error_message=row_error
            )
 
            log(traceback.format_exc())
 
            continue
 
 
    total_pages_processed += 1
 
 
    log(
        f"Completed processing website page "
        f"{page_number}."
    )
 
 
# ============================================================
# FIND NEXT BUTTON
# ============================================================
 
def get_next_button():
 
    # Scope pagination lookup near the main table first.
    table = page.locator(
        f"xpath={MAIN_TABLE_XPATH}"
    )
 
 
    table_container = table.locator(
        "xpath=ancestor::div[contains("
        "@class,'border')][1]"
    )
 
 
    scoped_next_button = table_container.locator(
        NEXT_BUTTON_SELECTOR
    )
 
 
    if scoped_next_button.count() > 0:
 
        return scoped_next_button.last
 
 
    # Fallback if the pagination button is outside the
    # immediate table container.
    next_buttons = page.locator(
        NEXT_BUTTON_SELECTOR
    )
 
 
    if next_buttons.count() == 0:
 
        raise Exception(
            "Unable to find the Next pagination button."
        )
 
 
    return next_buttons.first
 
 
# ============================================================
# SAVE ALL SELECTION
# ============================================================
 
def save_all_selection():
 
    if not SAVE_ALL_SELECTION_AT_END:
 
        log(
            "SAVE_ALL_SELECTION_AT_END is False. "
            "Final save click was skipped."
        )
 
        return
 
 
    if total_instance_types_updated == 0:
 
        log(
            "No instance types were updated. "
            "Save All Selection will not be clicked."
        )
 
        return
 
 
    save_buttons = page.locator(
        SAVE_ALL_BUTTON_SELECTOR
    )
 
 
    if save_buttons.count() == 0:
 
        raise Exception(
            "The 'Save All Selection' button "
            "was not found."
        )
 
 
    save_button = save_buttons.first
 
 
    save_button.scroll_into_view_if_needed()
 
 
    if not save_button.is_enabled():
 
        raise Exception(
            "The 'Save All Selection' button "
            "is disabled after instance type updates."
        )
 
 
    log("Clicking Save All Selection.")
 
 
    save_button.click()
 
 
    wait_for_loader()
 
 
    log(
        "Save All Selection completed successfully."
    )
 
 
# ============================================================
# START AUTOMATION
# ============================================================
 
p = None
browser = None
page = None
 
 
try:
 
    prevent_system_sleep()
 
 
    log(
        f"Connecting to Chrome debugging session: "
        f"{CHROME_DEBUG_URL}"
    )
 
 
    p = sync_playwright().start()
 
 
    browser = p.chromium.connect_over_cdp(
        CHROME_DEBUG_URL
    )
 
 
    if len(browser.contexts) == 0:
 
        raise Exception(
            "No browser context was found in the "
            "Chrome debugging session."
        )
 
 
    contxt = browser.contexts[0]
 
 
    if len(contxt.pages) == 0:
 
        raise Exception(
            "No open page was found in the "
            "Chrome debugging browser context."
        )
 
 
    page = contxt.pages[0]
 
 
    log(
        f"Connected to browser page: {page.url}"
    )
 
 
    current_page_number = 1
 
 
    while True:
 
        process_current_page(
            current_page_number
        )
 
        log(
            f"Last row processed on page "
            f"{current_page_number}"
        )
        next_button = get_next_button()
 
 
        next_disabled_attribute = (
            next_button.get_attribute("disabled")
        )
 
 
        next_is_disabled = (
            next_disabled_attribute is not None
            or not next_button.is_enabled()
        )
 
 
        if next_is_disabled:
 
            log(
                f"Next button is disabled on page "
                f"{current_page_number}. "
                f"The last page has been reached."
            )
 
            break
 
 
        log(
            f"Clicking Next after completing page "
            f"{current_page_number}."
        )
 
 
        next_button.click()

        log(
            "Next button clicked. "
            "Waiting for page refresh."
        )

        wait_for_loader()

        page.locator(
            f"xpath={MAIN_TABLE_XPATH}"
        ).wait_for(
            state="visible",
            timeout=300000
        )

        page.wait_for_timeout(1000)

        current_page_number += 1

        log(
            f"Moved to website page "
            f"{current_page_number}."
        )
        
 
    save_all_selection()
 
 
except Exception as automation_error:
 
    record_error(
        vm_hostname="N/A",
        excel_instance_type="N/A",
        page_number=current_page_number,
        error_message=(
            f"Automation-level error: "
            f"{automation_error}"
        )
    )
 
    log(traceback.format_exc())
 
 
finally:
 
    allow_system_sleep()
 
 
    automation_end_time = datetime.now()
 
    automation_duration = (
        automation_end_time
        - automation_start_time
    )
 
 
    print("")
    print("=" * 76)
    print("INSTANCE TYPE AUTOMATION EXECUTION SUMMARY")
    print("=" * 76)
 
    print(
        f"Start time                       : "
        f"{automation_start_time.strftime('%Y-%m-%d %H:%M:%S')}"
    )
 
    print(
        f"End time                         : "
        f"{automation_end_time.strftime('%Y-%m-%d %H:%M:%S')}"
    )
 
    print(
        f"Execution duration               : "
        f"{str(automation_duration).split('.')[0]}"
    )
 
    print(
        f"Valid Excel hostname rows        : "
        f"{total_excel_rows}"
    )
 
    print(
        f"Website pages processed          : "
        f"{total_pages_processed}"
    )
 
    print(
        f"Website VM rows processed        : "
        f"{total_website_rows_processed}"
    )
 
    print(
        f"Instance types updated           : "
        f"{total_instance_types_updated}"
    )
 
    print(
        f"Already matching                 : "
        f"{total_already_matching}"
    )
 
    print(
        f"VMs not found in Excel           : "
        f"{total_not_found_in_excel}"
    )
 
    print(
        f"Total skipped                    : "
        f"{total_skipped}"
    )
 
    print(
        f"Total errors                     : "
        f"{total_errors}"
    )
 
    print("-" * 76)
 
 
    if error_details:
 
        print("ERROR DETAILS")
        print("-" * 76)
 
 
        for error_number, error_item in enumerate(
            error_details,
            start=1
        ):
 
            print(
                f"{error_number}. Page          : "
                f"{error_item['page_number']}"
            )
 
            print(
                f"   VM Hostname   : "
                f"{error_item['vm_hostname']}"
            )
 
            print(
                f"   Excel Instance: "
                f"{error_item['excel_instance_type']}"
            )
 
            print(
                f"   Error         : "
                f"{error_item['error']}"
            )
 
            print("-" * 76)
 
 
    else:
 
        print(
            "Error details                    : "
            "No errors"
        )
 
 
    if total_errors == 0:
 
        print(
            "Final status                     : "
            "SUCCESS"
        )
 
    elif total_instance_types_updated > 0:
 
        print(
            "Final status                     : "
            "COMPLETED WITH ERRORS"
        )
 
    else:
 
        print(
            "Final status                     : "
            "FAILED"
        )
 
 
    print("=" * 76)