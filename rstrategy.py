from click import option
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError
import pandas as pd
import ctypes
import builtins
from datetime import datetime


LOADER_CSS = "div.loader"


# ============================================================
# EXCEL CONFIGURATION
# ============================================================

EXCEL_PATH ="input1.xlsx"

# Change these only if the Excel column names are different
EXCEL_HOSTNAME_COLUMN = "hostname"
EXCEL_STRATEGY_COLUMN = "rstrategy"


# ============================================================
# EXECUTION SUMMARY VARIABLES
# ============================================================

automation_start_time = datetime.now()

total_applications_processed = 0
total_strategy_updates = 0
total_skipped = 0
total_already_matching = 0
total_not_found_in_excel = 0
total_errors = 0

error_details = []

current_application_name = "APP_16_1"


# ============================================================
# DISABLE EXISTING LOG OUTPUT
# ============================================================

# Save the original print function. It will be used for the
# final summary.
original_print = builtins.print


# All existing print statements remain in the code, but they
# will not display output during automation.
def print(*args, **kwargs):
    pass


# ============================================================
# PREVENT WINDOWS SYSTEM AND DISPLAY FROM SLEEPING
# ============================================================

ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002


def prevent_system_sleep():

    global total_errors

    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS
            | ES_SYSTEM_REQUIRED
            | ES_DISPLAY_REQUIRED
        )

    except Exception as sleep_error:

        total_errors += 1

        error_details.append(
            {
                "application": "System",
                "vm_hostname": "N/A",
                "error": (
                    "Unable to prevent the system from sleeping: "
                    f"{sleep_error}"
                )
            }
        )


def allow_system_sleep():

    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS
        )

    except Exception:
        pass


# ============================================================
# READ EXCEL
# ============================================================

excel_df = pd.read_excel(
    EXCEL_PATH,
    engine="openpyxl"
)

# Remove accidental spaces from Excel column headings
excel_df.columns = (
    excel_df.columns
    .astype(str)
    .str.strip()
)


# Validate hostname column
if EXCEL_HOSTNAME_COLUMN not in excel_df.columns:

    raise ValueError(
        f"Column '{EXCEL_HOSTNAME_COLUMN}' was not found in Excel. "
        f"Available columns: {excel_df.columns.tolist()}"
    )


# Validate strategy column
if EXCEL_STRATEGY_COLUMN not in excel_df.columns:

    raise ValueError(
        f"Column '{EXCEL_STRATEGY_COLUMN}' was not found in Excel. "
        f"Available columns: {excel_df.columns.tolist()}"
    )


# Create hostname and strategy lookup dictionary
hostname_strategy = {}

for _, excel_row in excel_df.iterrows():

    excel_hostname = excel_row[EXCEL_HOSTNAME_COLUMN]
    excel_strategy = excel_row[EXCEL_STRATEGY_COLUMN]

    if pd.isna(excel_hostname) or pd.isna(excel_strategy):
        continue

    excel_hostname = str(excel_hostname).strip()
    excel_strategy = str(excel_strategy).strip()

    if excel_hostname:
        hostname_strategy[excel_hostname.lower()] = excel_strategy


print(
    "Excel hostname strategy count:",
    len(hostname_strategy)
)


# ============================================================
# STRATEGY MAPPING
# ============================================================

def get_dropdown_strategy(excel_strategy):

    strategy_mapping = {
        "replatform": "Re-Platform"
    }

    cleaned_strategy = str(excel_strategy).strip()

    return strategy_mapping.get(
        cleaned_strategy.lower(),
        cleaned_strategy
    )


# This makes Excel "Replatform" equal to UI "Re-Platform"
def normalize_strategy(strategy):

    if strategy is None:
        return ""

    return (
        str(strategy)
        .strip()
        .lower()
        .replace("-", "")
        .replace(" ", "")
    )


# ============================================================
# FIND WEBSITE TABLE COLUMNS
# ============================================================

def get_required_column_indexes(table):

    headers = table.locator("th")

    vm_hostname_column = None
    platform_strategy_column = None

    for header_index in range(headers.count()):

        header_text = (
            headers
            .nth(header_index)
            .inner_text()
            .strip()
        )

        if header_text == "VM Hostname":
            vm_hostname_column = header_index

        if header_text == "Platform Migration Strategy":
            platform_strategy_column = header_index

    if vm_hostname_column is None:

        raise Exception(
            "Unable to find the 'VM Hostname' column "
            "in the website table."
        )

    if platform_strategy_column is None:

        raise Exception(
            "Unable to find the "
            "'Platform Migration Strategy' column "
            "in the website table."
        )

    print(
        "VM Hostname column:",
        vm_hostname_column
    )

    print(
        "Platform Migration Strategy column:",
        platform_strategy_column
    )

    return (
        vm_hostname_column,
        platform_strategy_column
    )


# ============================================================
# START PLAYWRIGHT
# ============================================================

p=sync_playwright().start()
browser=p.chromium.connect_over_cdp("http://127.0.0.1:9432")
contxt=browser.contexts[0]
page=contxt.pages[0]


select_app=page.wait_for_selector(
    '//span[text()="Select Application"]'
)

select_app.click()


app_list=page.locator(
    "ul.select-dropdown-list>li.select-dropdown-option"
)


lis=[]


print(app_list.count())


for i in range(app_list.count()):

    app_name=app_list.nth(i).get_attribute("title")
    lis.append(app_name)


print(lis)


page.locator(
    '//span[text()="Select Application"]'
).click()


# ============================================================
# PROCESS FIRST APPLICATION
# ============================================================

def process_first_app(app):

    global total_strategy_updates
    global total_skipped
    global total_already_matching
    global total_not_found_in_excel
    global total_errors

    next_btn=page.locator(
        "button:has(path[d='M6 12L10 8L6 4'])"
    )

    while True:

        table=page.locator(
            '//table[@class="min-w-full"]'
        )

        row=table.locator('tr')

        length=row.count()

        print(f"{length}")


        vm_hostname_column, platform_strategy_column = (
            get_required_column_indexes(table)
        )


        for i in range(1, length):

            data = row.nth(i)

            vm_hostname = "Unable to read VM hostname"

            try:

                print(f"Selecting row {i}")


                cells = data.locator("td")


                vm_hostname = (
                    cells
                    .nth(vm_hostname_column)
                    .inner_text()
                    .strip()
                )


                current_strategy = (
                    cells
                    .nth(platform_strategy_column)
                    .inner_text()
                    .strip()
                )


                excel_strategy = hostname_strategy.get(
                    vm_hostname.lower()
                )


                print("VM Hostname:", vm_hostname)

                print(
                    "Current UI Strategy:",
                    current_strategy
                )

                print(
                    "Excel Strategy:",
                    excel_strategy
                )


                # VM hostname not available in Excel
                if excel_strategy is None:

                    total_skipped += 1
                    total_not_found_in_excel += 1

                    print(
                        f"{vm_hostname} not found in Excel. "
                        f"Skipping."
                    )

                    continue


                # UI and Excel strategies are already matching
                if (
                    normalize_strategy(current_strategy)
                    == normalize_strategy(excel_strategy)
                ):

                    total_skipped += 1
                    total_already_matching += 1

                    print(
                        f"{vm_hostname} strategy is already "
                        f"matching. Skipping."
                    )

                    continue


                dropdown_strategy = get_dropdown_strategy(
                    excel_strategy
                )


                print(
                    f"Updating {vm_hostname}: "
                    f"{current_strategy} -> "
                    f"{dropdown_strategy}"
                )


                #page.wait_for_timeout(2000)

                print("Clicking Edit")

                data.locator(
                    'button[title="Edit"]'
                ).click()


                #page.wait_for_timeout(3000)

                drop=data.locator('select')

                drop.click()

                drop.select_option(
                    value=dropdown_strategy
                )


                comment = page.locator(
                    'textarea[placeholder="Please state reason."]'
                )


                print(
                    "Textarea count:",
                    comment.count()
                )

                print(
                    "Textarea visible:",
                    comment.is_visible()
                )


                #page.wait_for_timeout(2000)

                print("Typing comment")

                comment.fill(".")


                #page.wait_for_timeout(3000)

                print(
                    "Entered value:",
                    comment.input_value()
                )


                print("Clicking Save")


                page.locator(
                    'button[title="Save"]'
                ).click()


                page.locator(LOADER_CSS).wait_for(
                    state="visible",
                    timeout=3000
                ) if page.locator(LOADER_CSS).is_visible(
                    timeout=3000
                ) else None; page.locator(LOADER_CSS).wait_for(
                    state="hidden",
                    timeout=120000
                )


                total_strategy_updates += 1


                print(f"Completed row {i}")


            except Exception as row_error:

                total_errors += 1

                error_details.append(
                    {
                        "application": app,
                        "vm_hostname": vm_hostname,
                        "error": str(row_error)
                    }
                )

                continue


        if next_btn.is_enabled():

            next_btn.click()

        else:

            print(
                "Last Page Reached..Selecting "
                "Next Application"
            )

            break


# ============================================================
# RUN AUTOMATION
# ============================================================

prevent_system_sleep()


try:

    for i,app in enumerate(lis):

        current_application_name = app

        print(f"Selecting application: {app}")


        if i==0:

            process_first_app(app)

            total_applications_processed += 1


        else:

            page.locator(
                '//span[text()="Select Application"]'
            ).click()


            page.wait_for_selector(
                "ul.select-dropdown-list>li.select-dropdown-option"
            )


            page.locator(
                "ul.select-dropdown-list>li.select-dropdown-option"
            ).filter(
                has_text=app
            ).click()


            page.locator(LOADER_CSS).wait_for(
                state="visible",
                timeout=3000
            ) if page.locator(LOADER_CSS).is_visible(
                timeout=3000
            ) else None; page.locator(LOADER_CSS).wait_for(
                state="hidden",
                timeout=120000
            )


            next_btn=page.locator(
                "button:has(path[d='M6 12L10 8L6 4'])"
            )


            while True:

                table=page.locator(
                    '//table[@class="min-w-full"]'
                )


                row=table.locator('tr')


                length=row.count()


                print(f"{length}")


                vm_hostname_column, platform_strategy_column = (
                    get_required_column_indexes(table)
                )


                for i in range(1, length):

                    data = row.nth(i)

                    vm_hostname = "Unable to read VM hostname"

                    try:

                        print(f"Selecting row {i}")


                        cells = data.locator("td")


                        vm_hostname = (
                            cells
                            .nth(vm_hostname_column)
                            .inner_text()
                            .strip()
                        )


                        current_strategy = (
                            cells
                            .nth(platform_strategy_column)
                            .inner_text()
                            .strip()
                        )


                        excel_strategy = (
                            hostname_strategy.get(
                                vm_hostname.lower()
                            )
                        )


                        print(
                            "VM Hostname:",
                            vm_hostname
                        )


                        print(
                            "Current UI Strategy:",
                            current_strategy
                        )


                        print(
                            "Excel Strategy:",
                            excel_strategy
                        )


                        # VM hostname not available in Excel
                        if excel_strategy is None:

                            total_skipped += 1

                            total_not_found_in_excel += 1

                            print(
                                f"{vm_hostname} not found in "
                                f"Excel. Skipping."
                            )

                            continue


                        # UI and Excel strategies are matching
                        if (
                            normalize_strategy(
                                current_strategy
                            )
                            == normalize_strategy(
                                excel_strategy
                            )
                        ):

                            total_skipped += 1

                            total_already_matching += 1

                            print(
                                f"{vm_hostname} strategy is "
                                f"already matching. Skipping."
                            )

                            continue


                        dropdown_strategy = (
                            get_dropdown_strategy(
                                excel_strategy
                            )
                        )


                        print(
                            f"Updating {vm_hostname}: "
                            f"{current_strategy} -> "
                            f"{dropdown_strategy}"
                        )


                        #page.wait_for_timeout(2000)

                        print("Clicking Edit")


                        data.locator(
                            'button[title="Edit"]'
                        ).click()


                        #page.wait_for_timeout(3000)

                        drop=data.locator('select')


                        drop.click()


                        drop.select_option(
                            value=dropdown_strategy
                        )


                        comment = page.locator(
                            'textarea[placeholder="Please state reason."]'
                        )


                        print(
                            "Textarea count:",
                            comment.count()
                        )


                        print(
                            "Textarea visible:",
                            comment.is_visible()
                        )


                        #page.wait_for_timeout(2000)

                        print("Typing comment")


                        comment.fill(".")


                        #page.wait_for_timeout(3000)

                        print(
                            "Entered value:",
                            comment.input_value()
                        )


                        print("Clicking Save")


                        page.locator(
                            'button[title="Save"]'
                        ).click()


                        page.locator(LOADER_CSS).wait_for(
                            state="visible",
                            timeout=3000
                        ) if page.locator(
                            LOADER_CSS
                        ).is_visible(
                            timeout=3000
                        ) else None; page.locator(
                            LOADER_CSS
                        ).wait_for(
                            state="hidden",
                            timeout=120000
                        )


                        total_strategy_updates += 1


                        print(f"Completed row {i}")


                    except Exception as row_error:

                        total_errors += 1

                        error_details.append(
                            {
                                "application": app,
                                "vm_hostname": vm_hostname,
                                "error": str(row_error)
                            }
                        )

                        continue


                if next_btn.is_enabled():

                    next_btn.click()

                else:

                    print(
                        "Last Page Reached..Selecting "
                        "Next Application"
                    )

                    break


            total_applications_processed += 1


except Exception as automation_error:

    total_errors += 1

    error_details.append(
        {
            "application": current_application_name,
            "vm_hostname": "N/A",
            "error": str(automation_error)
        }
    )


finally:

    # Restore normal Windows sleep behavior
    allow_system_sleep()


    automation_end_time = datetime.now()

    automation_duration = (
        automation_end_time
        - automation_start_time
    )


    original_print("")

    original_print("=" * 72)

    original_print(
        "AUTOMATION EXECUTION SUMMARY"
    )

    original_print("=" * 72)


    original_print(
        f"Start time                    : "
        f"{automation_start_time.strftime('%Y-%m-%d %H:%M:%S')}"
    )


    original_print(
        f"End time                      : "
        f"{automation_end_time.strftime('%Y-%m-%d %H:%M:%S')}"
    )


    original_print(
        f"Total execution duration      : "
        f"{str(automation_duration).split('.')[0]}"
    )


    original_print(
        f"Applications available        : "
        f"{len(lis)}"
    )


    original_print(
        f"Applications processed        : "
        f"{total_applications_processed}"
    )


    original_print(
        f"Total strategy updates        : "
        f"{total_strategy_updates}"
    )


    original_print(
        f"Total VMs skipped             : "
        f"{total_skipped}"
    )


    original_print(
        f"Already matching strategies   : "
        f"{total_already_matching}"
    )


    original_print(
        f"VMs not found in Excel        : "
        f"{total_not_found_in_excel}"
    )


    original_print(
        f"Total errors                  : "
        f"{total_errors}"
    )


    original_print("-" * 72)


    if error_details:

        original_print("ERROR DETAILS")

        original_print("-" * 72)


        for error_number, error_item in enumerate(
            error_details,
            start=1
        ):

            original_print(
                f"{error_number}. Application : "
                f"{error_item['application']}"
            )


            original_print(
                f"   VM Hostname : "
                f"{error_item['vm_hostname']}"
            )


            original_print(
                f"   Error       : "
                f"{error_item['error']}"
            )


            original_print("-" * 72)


    else:

        original_print(
            "Error details                 : "
            "No errors"
        )


    if (
        total_applications_processed == len(lis)
        and total_errors == 0
    ):

        original_print(
            "Final status                  : SUCCESS"
        )


    elif total_applications_processed > 0:

        original_print(
            "Final status                  : "
            "COMPLETED WITH ERRORS"
        )


    else:

        original_print(
            "Final status                  : FAILED"
        )


    original_print("=" * 72)